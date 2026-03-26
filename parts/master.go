package parts

import (
	"archive/zip"
	"fmt"
	"io"
	"path"
	"sort"
	"strings"
)

// ============================================================================
// MasterManager - 母版/版式管理器（门面模式）
// ============================================================================
//
// 作为外部 API 调用的入口，负责：
// 1. 从 ZIP 文件加载母版和版式
// 2. 解析 XML 并转换为只读数据结构
// 3. 填充到 MasterCache 供高并发读取
// ============================================================================

// MasterManager 母版/版式管理器
type MasterManager struct {
	cache *MasterCache
}

// NewMasterManager 创建新的母版管理器
func NewMasterManager() *MasterManager {
	return &MasterManager{
		cache: NewMasterCache(),
	}
}

// NewMasterManagerWithCache 使用指定缓存创建母版管理器
func NewMasterManagerWithCache(cache *MasterCache) *MasterManager {
	return &MasterManager{
		cache: cache,
	}
}

// Cache 返回内部缓存（只读）
func (m *MasterManager) Cache() *MasterCache {
	return m.cache
}

// ============================================================================
// 从 ZIP 加载
// ============================================================================

// LoadFromZip 从 ZIP Reader 加载母版和版式
// 遍历 ZIP 内的 /ppt/slideMasters/ 和 /ppt/slideLayouts/ 目录
func (m *MasterManager) LoadFromZip(zipReader *zip.Reader) error {
	var masters []*SlideMasterData
	var layouts []*SlideLayoutData

	// 收集母版文件
	masterFiles := m.collectFiles(zipReader, "ppt/slideMasters/", "slideMaster")
	layoutFiles := m.collectFiles(zipReader, "ppt/slideLayouts/", "slideLayout")

	// 按文件名排序，确保顺序一致
	sort.Slice(masterFiles, func(i, j int) bool {
		return masterFiles[i].name < masterFiles[j].name
	})
	sort.Slice(layoutFiles, func(i, j int) bool {
		return layoutFiles[i].name < layoutFiles[j].name
	})

	// 解析母版
	for _, f := range masterFiles {
		data, err := m.readFile(f.file)
		if err != nil {
			return fmt.Errorf("读取母版文件 %s 失败: %w", f.name, err)
		}

		master, err := ParseMaster(data)
		if err != nil {
			return fmt.Errorf("解析母版 %s 失败: %w", f.name, err)
		}

		masters = append(masters, master)
	}

	// 解析版式
	for _, f := range layoutFiles {
		data, err := m.readFile(f.file)
		if err != nil {
			return fmt.Errorf("读取版式文件 %s 失败: %w", f.name, err)
		}

		layout, err := ParseLayout(data)
		if err != nil {
			return fmt.Errorf("解析版式 %s 失败: %w", f.name, err)
		}

		layouts = append(layouts, layout)
	}

	// 初始化缓存（仅执行一次）
	m.cache.Init(masters, layouts)

	return nil
}

// LoadFromZipFile 从 ZIP 文件路径加载
func (m *MasterManager) LoadFromZipFile(filePath string) error {
	reader, err := zip.OpenReader(filePath)
	if err != nil {
		return fmt.Errorf("打开 ZIP 文件失败: %w", err)
	}
	defer reader.Close()

	return m.LoadFromZip(&reader.Reader)
}

// ============================================================================
// 文件收集辅助
// ============================================================================

type zipFileEntry struct {
	name string
	file *zip.File
}

// collectFiles 收集指定目录下匹配前缀的 XML 文件
func (m *MasterManager) collectFiles(zipReader *zip.Reader, dir, prefix string) []zipFileEntry {
	var files []zipFileEntry

	for _, f := range zipReader.File {
		// 检查是否在目标目录
		fileDir := path.Dir(f.Name)
		if fileDir != dir && !strings.HasPrefix(fileDir, dir) {
			continue
		}

		// 检查文件名前缀和扩展名
		fileName := path.Base(f.Name)
		if !strings.HasPrefix(fileName, prefix) {
			continue
		}
		if path.Ext(fileName) != ".xml" {
			continue
		}

		files = append(files, zipFileEntry{
			name: f.Name,
			file: f,
		})
	}

	return files
}

// readFile 读取 ZIP 文件内容
func (m *MasterManager) readFile(f *zip.File) ([]byte, error) {
	rc, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	return io.ReadAll(rc)
}

// ============================================================================
// 便捷访问方法（委托给 Cache）
// ============================================================================

// GetLayout 获取版式
func (m *MasterManager) GetLayout(layoutID string) (*SlideLayoutData, bool) {
	return m.cache.GetLayout(layoutID)
}

// GetLayoutByName 根据名称获取版式
func (m *MasterManager) GetLayoutByName(name string) (*SlideLayoutData, bool) {
	return m.cache.GetLayoutByName(name)
}

// GetMaster 获取母版
func (m *MasterManager) GetMaster(masterID string) (*SlideMasterData, bool) {
	return m.cache.GetMaster(masterID)
}

// GetMasterByName 根据名称获取母版
func (m *MasterManager) GetMasterByName(name string) (*SlideMasterData, bool) {
	return m.cache.GetMasterByName(name)
}

// GetPlaceholder 获取占位符
func (m *MasterManager) GetPlaceholder(layoutID, phType string) (*Placeholder, bool) {
	return m.cache.GetPlaceholder(layoutID, phType)
}

// AllLayouts 返回所有版式
func (m *MasterManager) AllLayouts() map[string]*SlideLayoutData {
	return m.cache.AllLayouts()
}

// AllMasters 返回所有母版
func (m *MasterManager) AllMasters() map[string]*SlideMasterData {
	return m.cache.AllMasters()
}

// LayoutCount 返回版式数量
func (m *MasterManager) LayoutCount() int {
	return m.cache.LayoutCount()
}

// MasterCount 返回母版数量
func (m *MasterManager) MasterCount() int {
	return m.cache.MasterCount()
}

// ListLayoutIDs 列出所有版式 ID
func (m *MasterManager) ListLayoutIDs() []string {
	return m.cache.ListLayoutIDs()
}

// ListLayoutNames 列出所有版式名称
func (m *MasterManager) ListLayoutNames() []string {
	return m.cache.ListLayoutNames()
}

// ============================================================================
// 全局默认管理器
// ============================================================================

var defaultManager *MasterManager

// DefaultManager 返回全局默认管理器
func DefaultManager() *MasterManager {
	if defaultManager == nil {
		defaultManager = NewMasterManager()
	}
	return defaultManager
}

// InitDefaultManager 初始化全局默认管理器
func InitDefaultManager(zipReader *zip.Reader) error {
	mgr := NewMasterManager()
	if err := mgr.LoadFromZip(zipReader); err != nil {
		return err
	}
	defaultManager = mgr
	return nil
}

// InitDefaultManagerFromFile 从文件初始化全局默认管理器
func InitDefaultManagerFromFile(filePath string) error {
	mgr := NewMasterManager()
	if err := mgr.LoadFromZipFile(filePath); err != nil {
		return err
	}
	defaultManager = mgr
	return nil
}
