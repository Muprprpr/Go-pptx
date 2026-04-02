// Package pptx 提供 PPTX 文件的高级操作接口
// 作为人类开发者和 AI 调用的绝对入口
package pptx

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"io/fs"
	"os"
	"path/filepath"
	"sync"

	"github.com/Muprprpr/Go-pptx/opc"
)

// ============================================================================
// 模板懒加载系统
// ============================================================================
//
// 设计原则：
// 1. 延迟加载 - 模板在首次使用时才加载
// 2. 零拷贝优化 - 不可变资源（图片、母版、布局）使用 zero-copy
// 3. 线程安全 - 使用 sync.Map 确保并发安全
// 4. 内存高效 - 模板只加载一次，后续通过 Clone() 复用
// ============================================================================

// TemplateType 模板类型
type TemplateType string

const (
	// TemplateBlank 空白模板
	TemplateBlank TemplateType = "blank.pptx"
	// TemplateDefault 默认模板（16:9 宽屏）
	TemplateDefault TemplateType = "default.pptx"
	// TemplateWide 宽屏模板
	TemplateWide TemplateType = "wide.pptx"
	// TemplateStandard 标准模板（4:3）
	TemplateStandard TemplateType = "standard.pptx"
)

// TemplateManager 模板管理器
// 负责模板的懒加载、缓存和克隆
type TemplateManager struct {
	// 模板缓存：模板名称 -> *opc.Package
	templates sync.Map

	// 默认模板
	defaultTemplate TemplateType

	// 母版管理器缓存
	masterCache *MasterCache

	// 模板目录路径（可选）
	templateDir string
}

// 全局模板管理器
var globalTemplateManager = NewTemplateManager()

// NewTemplateManager 创建新的模板管理器
func NewTemplateManager() *TemplateManager {
	return &TemplateManager{
		defaultTemplate: TemplateDefault,
	}
}

// NewTemplateManagerWithDir 创建带模板目录的模板管理器
func NewTemplateManagerWithDir(dir string) *TemplateManager {
	return &TemplateManager{
		defaultTemplate: TemplateDefault,
		templateDir:     dir,
	}
}

// ============================================================================
// 模板加载方法
// ============================================================================

// LoadTemplate 加载指定模板
// 如果模板已缓存，直接返回克隆副本；否则尝试从文件系统加载
func (tm *TemplateManager) LoadTemplate(name TemplateType) (*opc.Package, error) {
	// 检查缓存
	if cached, ok := tm.templates.Load(name); ok {
		pkg := cached.(*opc.Package)
		// 返回克隆副本，确保线程安全
		return pkg.Clone(), nil
	}

	// 尝试从模板目录加载
	data, err := tm.readTemplateFile(string(name))
	if err != nil {
		return nil, fmt.Errorf("加载模板 %s 失败: %w", name, err)
	}

	// 解析为 OPC 包
	pkg, err := tm.parseTemplate(data)
	if err != nil {
		return nil, fmt.Errorf("解析模板 %s 失败: %w", name, err)
	}

	// 缓存模板（存储原始模板，后续通过 Clone 复用）
	tm.templates.Store(name, pkg)

	// 返回克隆副本
	return pkg.Clone(), nil
}

// LoadDefault 加载默认模板
func (tm *TemplateManager) LoadDefault() (*opc.Package, error) {
	return tm.LoadTemplate(tm.defaultTemplate)
}

// readTemplateFile 从文件系统读取模板文件
func (tm *TemplateManager) readTemplateFile(name string) ([]byte, error) {
	// 尝试路径顺序：
	// 1. 模板目录（如果设置）
	// 2. 当前目录下的 templates/ 目录
	// 3. 可执行文件目录下的 templates/ 目录

	searchPaths := []string{}

	if tm.templateDir != "" {
		searchPaths = append(searchPaths, filepath.Join(tm.templateDir, name))
	}

	// 当前目录
	searchPaths = append(searchPaths, filepath.Join("templates", name))

	// 可执行文件目录
	if exePath, err := os.Executable(); err == nil {
		searchPaths = append(searchPaths, filepath.Join(filepath.Dir(exePath), "templates", name))
	}

	// 依次尝试
	for _, path := range searchPaths {
		if data, err := os.ReadFile(path); err == nil {
			return data, nil
		}
	}

	return nil, fmt.Errorf("模板文件 %s 不存在于任何搜索路径", name)
}

// parseTemplate 将模板数据解析为 OPC 包
func (tm *TemplateManager) parseTemplate(data []byte) (*opc.Package, error) {
	// 创建 ZIP reader
	reader := bytes.NewReader(data)
	zipReader, err := zip.NewReader(reader, int64(len(data)))
	if err != nil {
		return nil, fmt.Errorf("解析 ZIP 失败: %w", err)
	}

	// 使用 opc.Open 解析
	pkg, err := opc.Open(reader, int64(len(data)))
	if err != nil {
		return nil, fmt.Errorf("解析 OPC 包失败: %w", err)
	}

	// 初始化母版管理器（懒加载母版和布局信息）
	if tm.masterCache == nil {
		masterMgr := NewMasterManager()
		if err := masterMgr.LoadFromZip(zipReader); err != nil {
			// 母版加载失败不阻止模板使用，只是记录警告
		}
		tm.masterCache = masterMgr.Cache()
	}

	return pkg, nil
}

// ============================================================================
// 模板注册方法
// ============================================================================

// RegisterTemplate 从文件路径注册模板
func (tm *TemplateManager) RegisterTemplate(name TemplateType, path string) error {
	pkg, err := opc.OpenFile(path)
	if err != nil {
		return fmt.Errorf("打开模板文件失败: %w", err)
	}

	tm.templates.Store(name, pkg)
	return nil
}

// RegisterTemplateFromBytes 从字节数据注册模板
func (tm *TemplateManager) RegisterTemplateFromBytes(name TemplateType, data []byte) error {
	reader := bytes.NewReader(data)
	pkg, err := opc.Open(reader, int64(len(data)))
	if err != nil {
		return fmt.Errorf("解析模板数据失败: %w", err)
	}

	tm.templates.Store(name, pkg)
	return nil
}

// RegisterTemplateFromFS 从文件系统注册模板
func (tm *TemplateManager) RegisterTemplateFromFS(fsys fs.FS, name TemplateType, path string) error {
	file, err := fsys.Open(path)
	if err != nil {
		return fmt.Errorf("打开模板文件失败: %w", err)
	}
	defer file.Close()

	data, err := io.ReadAll(file)
	if err != nil {
		return fmt.Errorf("读取模板文件失败: %w", err)
	}

	return tm.RegisterTemplateFromBytes(name, data)
}

// ============================================================================
// 辅助方法
// ============================================================================

// SetDefaultTemplate 设置默认模板
func (tm *TemplateManager) SetDefaultTemplate(name TemplateType) {
	tm.defaultTemplate = name
}

// SetTemplateDir 设置模板目录
func (tm *TemplateManager) SetTemplateDir(dir string) {
	tm.templateDir = dir
}

// GetMasterCache 获取母版缓存
func (tm *TemplateManager) GetMasterCache() *MasterCache {
	return tm.masterCache
}

// HasTemplate 检查模板是否已加载
func (tm *TemplateManager) HasTemplate(name TemplateType) bool {
	_, ok := tm.templates.Load(name)
	return ok
}

// ClearCache 清空模板缓存
func (tm *TemplateManager) ClearCache() {
	tm.templates = sync.Map{}
	tm.masterCache = nil
}

// ============================================================================
// 全局便捷函数
// ============================================================================

// LoadTemplate 加载指定模板（使用全局管理器）
func LoadTemplate(name TemplateType) (*opc.Package, error) {
	return globalTemplateManager.LoadTemplate(name)
}

// LoadDefaultTemplate 加载默认模板（使用全局管理器）
func LoadDefaultTemplate() (*opc.Package, error) {
	return globalTemplateManager.LoadDefault()
}

// RegisterTemplate 注册模板（使用全局管理器）
func RegisterTemplate(name TemplateType, path string) error {
	return globalTemplateManager.RegisterTemplate(name, path)
}

// RegisterTemplateFromBytes 从字节数据注册模板（使用全局管理器）
func RegisterTemplateFromBytes(name TemplateType, data []byte) error {
	return globalTemplateManager.RegisterTemplateFromBytes(name, data)
}

// ============================================================================
// 模板构建器 - 用于程序化创建模板
// ============================================================================

// TemplateBuilder 模板构建器
// 用于从零开始创建 PPTX 模板
type TemplateBuilder struct {
	pkg *opc.Package
}

// NewTemplateBuilder 创建新的模板构建器
func NewTemplateBuilder() *TemplateBuilder {
	return &TemplateBuilder{
		pkg: opc.NewPackage(),
	}
}

// Package 返回底层 OPC 包
func (tb *TemplateBuilder) Package() *opc.Package {
	return tb.pkg
}

// Build 构建模板并返回 OPC 包
func (tb *TemplateBuilder) Build() *opc.Package {
	return tb.pkg
}

// BuildAndRegister 构建模板并注册到全局管理器
func (tb *TemplateBuilder) BuildAndRegister(name TemplateType) error {
	globalTemplateManager.templates.Store(name, tb.pkg)
	return nil
}
