// Package pptx 提供 PPTX 文件的高级操作接口
package pptx

import (
	"fmt"
	"sync"

	"github.com/Muprprpr/Go-pptx/opc"
	"github.com/Muprprpr/Go-pptx/parts"
)

// ============================================================================
// 嵌入式模板系统
// ============================================================================
//
// 设计原则：
// 1. 程序化生成 - 使用代码创建最小模板，无需外部文件
// 2. 懒加载 - 使用 sync.Once 确保只创建一次
// 3. 零依赖 - 运行时无需外部模板文件
// 4. 高性能 - 模板预创建，Clone 复用
//
// 如果需要嵌入实际的 .pptx 模板文件：
// 1. 在 pptx/templates/ 目录下放置 .pptx 文件
// 2. 创建 embed_fs.go 文件，添加 //go:embed 指令
// ============================================================================

// EmbeddedTemplateManager 嵌入式模板管理器
// 使用程序化方式创建模板
type EmbeddedTemplateManager struct {
	// 模板缓存：模板名称 -> *opc.Package
	templates sync.Map

	// 初始化标记
	once sync.Once

	// 初始化错误
	initErr error

	// 母版缓存
	masterCache *MasterCache
}

// 全局嵌入式模板管理器
var globalEmbeddedTemplates = &EmbeddedTemplateManager{}

// GetEmbeddedTemplateManager 获取全局嵌入式模板管理器
func GetEmbeddedTemplateManager() *EmbeddedTemplateManager {
	return globalEmbeddedTemplates
}

// ============================================================================
// 初始化方法
// ============================================================================

// Init 初始化嵌入式模板（仅执行一次）
func (etm *EmbeddedTemplateManager) Init() error {
	etm.once.Do(func() {
		etm.initErr = etm.createAllTemplates()
	})
	return etm.initErr
}

// createAllTemplates 创建所有模板
func (etm *EmbeddedTemplateManager) createAllTemplates() error {
	// 创建默认模板（16:9 宽屏）
	defaultTmpl := etm.createMinimalTemplate()
	etm.templates.Store(TemplateDefault, defaultTmpl)

	// 创建空白模板（使用默认模板）
	etm.templates.Store(TemplateBlank, defaultTmpl)

	// 创建宽屏模板（16:9）
	etm.templates.Store(TemplateWide, defaultTmpl)

	// 创建标准模板（4:3）
	standardTmpl := etm.createStandardTemplate()
	etm.templates.Store(TemplateStandard, standardTmpl)

	return nil
}

// createMinimalTemplate 创建最小化模板（16:9 宽屏）
func (etm *EmbeddedTemplateManager) createMinimalTemplate() *opc.Package {
	pkg := opc.NewPackage()

	// 创建 presentation.xml
	presPart := parts.NewPresentationPart()
	presPart.SetSlideSize(parts.SlideSize{Cx: 12192000, Cy: 6858000}) // 16:9 (13.333" x 7.5")

	presURI := opc.NewPackURI("/ppt/presentation.xml")
	presBlob, _ := presPart.ToXML()
	_ = pkg.AddPart(opc.NewPart(presURI, opc.ContentTypePresentation, presBlob))

	// 添加包级别关系
	_, _ = pkg.AddRelationship(opc.RelTypeOfficeDocument, "ppt/presentation.xml", false)

	return pkg
}

// createStandardTemplate 创建标准模板（4:3）
func (etm *EmbeddedTemplateManager) createStandardTemplate() *opc.Package {
	pkg := opc.NewPackage()

	// 创建 presentation.xml
	presPart := parts.NewPresentationPart()
	presPart.SetSlideSize(parts.SlideSize{Cx: 9144000, Cy: 6858000}) // 4:3 (10" x 7.5")

	presURI := opc.NewPackURI("/ppt/presentation.xml")
	presBlob, _ := presPart.ToXML()
	_ = pkg.AddPart(opc.NewPart(presURI, opc.ContentTypePresentation, presBlob))

	// 添加包级别关系
	_, _ = pkg.AddRelationship(opc.RelTypeOfficeDocument, "ppt/presentation.xml", false)

	return pkg
}

// ============================================================================
// 模板获取方法
// ============================================================================

// GetTemplate 获取模板（返回克隆副本）
func (etm *EmbeddedTemplateManager) GetTemplate(name TemplateType) (*opc.Package, error) {
	// 确保已初始化
	if err := etm.Init(); err != nil {
		return nil, err
	}

	val, ok := etm.templates.Load(name)
	if !ok {
		return nil, fmt.Errorf("模板 %s 不存在", name)
	}

	pkg := val.(*opc.Package)
	return pkg.Clone(), nil
}

// GetDefaultTemplate 获取默认模板
func (etm *EmbeddedTemplateManager) GetDefaultTemplate() (*opc.Package, error) {
	return etm.GetTemplate(TemplateDefault)
}

// HasTemplate 检查模板是否存在
func (etm *EmbeddedTemplateManager) HasTemplate(name TemplateType) bool {
	_, ok := etm.templates.Load(name)
	return ok
}

// ============================================================================
// 全局便捷函数
// ============================================================================

// GetEmbeddedTemplate 获取嵌入式模板（使用全局管理器）
func GetEmbeddedTemplate(name TemplateType) (*opc.Package, error) {
	return globalEmbeddedTemplates.GetTemplate(name)
}

// GetEmbeddedDefaultTemplate 获取嵌入式默认模板
func GetEmbeddedDefaultTemplate() (*opc.Package, error) {
	return globalEmbeddedTemplates.GetDefaultTemplate()
}

// InitEmbeddedTemplates 初始化嵌入式模板
func InitEmbeddedTemplates() error {
	return globalEmbeddedTemplates.Init()
}
