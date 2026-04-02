// Package pptx 提供 PPTX 文件的高级操作接口
package pptx

import (
	"fmt"
	"sync"
	"sync/atomic"

	"github.com/Muprprpr/Go-pptx/opc"
	"github.com/Muprprpr/Go-pptx/parts"
)

// ============================================================================
// SlideContext - 组件渲染上下文
// ============================================================================
//
// SlideContext 是派给组件的"特派员"，提供组件调用的"有限特权"后门。
//
// 核心职责：
// 1. 分配绝对不冲突的形状 ID
// 2. 管理媒体资源（图片、音频、视频）
// 3. 管理图表 XML 并返回关系 ID
// 4. 将组件生成的 XML 结构挂载到幻灯片
//
// 使用示例：
//
//	func (c *MyComponent) Render(ctx *SlideContext) error {
//		// 1. 分配形状 ID
//		id := ctx.NextShapeID()
//
//		// 2. 添加图片并获取 rId
//		rId, err := ctx.AddMedia(imageBytes, "image.png")
//		if err != nil {
//			return err
//		}
//
//		// 3. 构建形状 XML
//		sp := &parts.XSp{...}
//
//		// 4. 挂载到幻灯片
//		ctx.AppendShape(sp)
//		return nil
//	}
//
// ============================================================================

// SlideContext 幻灯片渲染上下文
// 提供组件渲染所需的资源和能力
type SlideContext struct {
	// 关联的幻灯片
	slide *Slide

	// 形状 ID 缓存（已分配的 ID）
	allocatedIDs map[uint32]bool

	// 关系 ID 缓存
	allocatedRIDs map[string]bool

	// 并发安全
	mu sync.RWMutex
}

// NewSlideContext 创建幻灯片上下文
func NewSlideContext(s *Slide) *SlideContext {
	return &SlideContext{
		slide:         s,
		allocatedIDs:  make(map[uint32]bool),
		allocatedRIDs: make(map[string]bool),
	}
}

// ============================================================================
// 形状 ID 管理
// ============================================================================

// NextShapeID 分配下一个形状 ID
// 返回绝对不冲突的形状 ID（线程安全，使用原子操作）
func (ctx *SlideContext) NextShapeID() uint32 {
	// 使用原子操作分配 ID，保证并发安全
	// Add(1) 会以原子操作的方式将值 +1，并返回相加后的新值
	return ctx.slide.shapeIDCounter.Add(1)
}

// CurrentShapeID 返回当前形状 ID（最后分配的）
func (ctx *SlideContext) CurrentShapeID() uint32 {
	return ctx.slide.shapeIDCounter.Load()
}

// AllocateShapeIDBatch 批量分配形状 ID
func (ctx *SlideContext) AllocateShapeIDBatch(count int) []uint32 {
	ctx.mu.Lock()
	defer ctx.mu.Unlock()

	ids := ctx.slide.part.AllocateShapeIDBatch(count)
	for _, id := range ids {
		ctx.allocatedIDs[id] = true
	}
	return ids
}

// IsShapeIDAllocated 检查形状 ID 是否已分配
func (ctx *SlideContext) IsShapeIDAllocated(id uint32) bool {
	ctx.mu.RLock()
	defer ctx.mu.RUnlock()
	return ctx.allocatedIDs[id]
}

// ============================================================================
// 媒体资源管理
// ============================================================================

// AddMedia 添加媒体资源（图片、音频、视频）
// data: 媒体数据
// fileName: 文件名（用于推断 MIME 类型）
// 返回: 关系 ID 和错误
func (ctx *SlideContext) AddMedia(data []byte, fileName string) (string, error) {
	ctx.mu.Lock()
	defer ctx.mu.Unlock()

	// 使用 MediaManager 添加媒体
	_, resource := ctx.slide.mediaManager.AddMediaAuto(fileName, data)
	if resource == nil {
		return "", fmt.Errorf("添加媒体资源失败")
	}

	// 获取目标 URI
	targetURI := resource.Target()

	// 添加关系到幻灯片
	slideRID := ctx.slide.part.AddImageRel(targetURI)
	ctx.allocatedRIDs[slideRID] = true

	return slideRID, nil
}

// AddMediaWithMIME 添加媒体资源（指定 MIME 类型）
func (ctx *SlideContext) AddMediaWithMIME(data []byte, fileName, mimeType string) (string, error) {
	// 暂时使用 AddMedia，后续可以扩展支持指定 MIME 类型
	return ctx.AddMedia(data, fileName)
}

// AddImage 添加图片资源（AddMedia 的别名，语义更清晰）
func (ctx *SlideContext) AddImage(data []byte, fileName string) (string, error) {
	return ctx.AddMedia(data, fileName)
}

// AddVideo 添加视频资源
func (ctx *SlideContext) AddVideo(data []byte, fileName string) (string, error) {
	return ctx.AddMedia(data, fileName)
}

// AddAudio 添加音频资源
func (ctx *SlideContext) AddAudio(data []byte, fileName string) (string, error) {
	return ctx.AddMedia(data, fileName)
}

// ============================================================================
// 图表管理
// ============================================================================

// AddChartXML 添加图表 XML
// chartXML: 图表 XML 数据
// 返回: 关系 ID 和错误
//
// 这是路线 C 的实现：组件塞入图表 XML，Context 负责写入底层的 ChartPart 并返回 rId
func (ctx *SlideContext) AddChartXML(chartXML []byte) (string, error) {
	ctx.mu.Lock()
	defer ctx.mu.Unlock()

	// 获取 Presentation 引用
	pres := ctx.slide.presentation

	// 分配图表编号
	chartNum := int(atomic.AddInt32(&pres.chartCounter, 1))

	// 创建图表部件
	chartPart := parts.NewChartPart(chartNum)
	chartPart.SetRawXML(chartXML)

	// 添加到 OPC 包
	chartURI := chartPart.PartURI()
	chartBlob := []byte(chartPart.Template()) // 获取 XML 内容
	part := opc.NewPart(chartURI, opc.ContentTypeChart, chartBlob)
	if err := pres.pkg.AddPart(part); err != nil {
		return "", fmt.Errorf("添加图表部件失败: %w", err)
	}

	// 添加关系到幻灯片
	chartRelID := ctx.slide.part.AddChartRel(chartURI.RelPathFrom(ctx.slide.part.PartURI()))
	ctx.allocatedRIDs[chartRelID] = true

	return chartRelID, nil
}

// AddChart 添加图表（使用模板）
// chartType: 图表类型
// data: 图表数据
// 返回: 关系 ID 和错误
func (ctx *SlideContext) AddChart(chartType parts.ChartType, data map[string]interface{}) (string, error) {
	// 创建图表部件
	ctx.mu.Lock()
	defer ctx.mu.Unlock()

	pres := ctx.slide.presentation
	chartNum := int(atomic.AddInt32(&pres.chartCounter, 1))

	chartPart := parts.NewChartPartWithType(chartNum, chartType)

	// 替换占位符
	for key, value := range data {
		chartPart.ReplacePlaceholder(key, fmt.Sprint(value))
	}

	// 添加到 OPC 包
	chartURI := chartPart.PartURI()
	chartBlob := []byte(chartPart.Template())
	part := opc.NewPart(chartURI, opc.ContentTypeChart, chartBlob)
	if err := pres.pkg.AddPart(part); err != nil {
		return "", fmt.Errorf("添加图表部件失败: %w", err)
	}

	// 添加关系
	chartRelID := ctx.slide.part.AddChartRel(chartURI.RelPathFrom(ctx.slide.part.PartURI()))
	ctx.allocatedRIDs[chartRelID] = true

	return chartRelID, nil
}

// ============================================================================
// 形状挂载
// ============================================================================

// AppendShape 将形状追加到幻灯片
// shape: 形状结构体（*parts.XSp, *parts.XPicture, *parts.XGraphicFrame 等）
func (ctx *SlideContext) AppendShape(shape interface{}) {
	ctx.mu.Lock()
	defer ctx.mu.Unlock()

	ctx.slide.part.AppendShapeChild(shape)
}

// AppendShapes 批量追加形状
func (ctx *SlideContext) AppendShapes(shapes ...interface{}) {
	ctx.mu.Lock()
	defer ctx.mu.Unlock()

	for _, shape := range shapes {
		ctx.slide.part.AppendShapeChild(shape)
	}
}

// ============================================================================
// 关系管理
// ============================================================================

// AddImageRel 添加图片关系
func (ctx *SlideContext) AddImageRel(targetURI string) string {
	ctx.mu.Lock()
	defer ctx.mu.Unlock()

	rID := ctx.slide.part.AddImageRel(targetURI)
	ctx.allocatedRIDs[rID] = true
	return rID
}

// AddMediaRel 添加媒体关系
func (ctx *SlideContext) AddMediaRel(targetURI string) string {
	ctx.mu.Lock()
	defer ctx.mu.Unlock()

	rID := ctx.slide.part.AddMediaRel(targetURI)
	ctx.allocatedRIDs[rID] = true
	return rID
}

// AddChartRel 添加图表关系
func (ctx *SlideContext) AddChartRel(targetURI string) string {
	ctx.mu.Lock()
	defer ctx.mu.Unlock()

	rID := ctx.slide.part.AddChartRel(targetURI)
	ctx.allocatedRIDs[rID] = true
	return rID
}

// HasRelationship 检查关系是否存在
func (ctx *SlideContext) HasRelationship(rID string) bool {
	ctx.mu.RLock()
	defer ctx.mu.RUnlock()
	return ctx.allocatedRIDs[rID]
}

// ============================================================================
// 幻灯片信息访问
// ============================================================================

// SlideIndex 返回幻灯片索引
func (ctx *SlideContext) SlideIndex() int {
	return ctx.slide.Index()
}

// SlideSize 返回幻灯片尺寸 (cx, cy in EMU)
func (ctx *SlideContext) SlideSize() (cx, cy int) {
	return ctx.slide.presentation.SlideSize()
}

// SlidePart 返回底层 SlidePart（高级用法）
func (ctx *SlideContext) SlidePart() *parts.SlidePart {
	return ctx.slide.part
}

// Presentation 返回所属演示文稿（高级用法）
func (ctx *SlideContext) Presentation() *Presentation {
	return ctx.slide.presentation
}

// ============================================================================
// 单位转换方法 - 默认单位: px
// ============================================================================

// PxToEMU 将像素转换为 EMU（基于 96 DPI）
func (ctx *SlideContext) PxToEMU(px int) int {
	return PxToEMU(px)
}

// EMUToPx 将 EMU 转换为像素（基于 96 DPI）
func (ctx *SlideContext) EMUToPx(emu int) int {
	return EMUToPx(emu)
}

// ============================================================================
// 批量操作
// ============================================================================

// RenderComponents 批量渲染组件
func (ctx *SlideContext) RenderComponents(components ...Component) error {
	for i, c := range components {
		if err := c.Render(ctx); err != nil {
			return &ComponentRenderError{
				Index:      i,
				Component:  c,
				Underlying: err,
			}
		}
	}
	return nil
}
