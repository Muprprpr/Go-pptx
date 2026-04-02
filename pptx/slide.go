// Package pptx 提供 PPTX 文件的高级操作接口
package pptx

import (
	"io"
	"sync/atomic"

	"github.com/Muprprpr/Go-pptx/opc"
	"github.com/Muprprpr/Go-pptx/parts"
	"github.com/Muprprpr/Go-pptx/utils"
)

// ============================================================================
// Slide - 高层幻灯片封装
// ============================================================================
//
// Slide 是对底层 parts.SlidePart 的高层封装，提供：
// 1. 便捷的内容添加方法（文本框、图片、表格等）
// 2. 与 Presentation 的双向关联
// 3. 媒体资源的自动管理
//
// 单位说明：
// - 所有位置和尺寸参数默认为 px（像素，基于 96 DPI）
// - 内部自动转换为 EMU（English Metric Units）
// - 1 px = 9525 EMU (914400 / 96)，此比例固定不变
// - 坐标原点 (0, 0) 位于幻灯片左上角
//
// 特殊说明：
// - px 坐标允许为负数，这在绘图中是合法的（元素可以放置在 canvas 边界外）
// - 幻灯片尺寸不影响坐标系统，用户需根据尺寸自行调整坐标值
//
// 使用示例：
//
//	s := pres.AddSlide()
//	s.AddTextBox(100, 100, 500, 50, "Hello World")  // 单位: px
//	s.AddPicture(100, 200, 300, 200, "image.png")    // 单位: px
//
// 标准幻灯片尺寸 (px):
// - 16:9  宽屏: 1280 x 720
// - 4:3   标准: 960 x 720
// - 16:10 超宽: 1280 x 800
//
// ============================================================================

// ============================================================================
// EMU 常量 - 单位换算比例固定不变
// ============================================================================
//
// 换算比例: 1 px = 9525 EMU (基于 96 DPI)
// 此比例是固定的，不随幻灯片尺寸改变
// 坐标原点 (0, 0) 位于幻灯片左上角

const (
	// EMUsPerPixel 每像素对应的 EMU 数量 (96 DPI)
	// 1 英寸 = 914400 EMU, // 1 英寸 = 96 像素 (96 DPI)
	// 因此 1 px = 914400 / 96 = 9525 EMU
	EMUsPerPixel = 914400 / 96 // = 9525
)

// ============================================================================
// 标准幻灯片尺寸 - 以 px 为单位 (基于 96 DPI)
// ============================================================================
//
// 这些尺寸用于帮助用户了解画布大小
// 用户需要根据尺寸自行调整坐标值
//
// 注意: px 坐标允许为负数，这在绘图中是合法的
// 元素可以放置在 canvas 边界之外

// SlideSize A幻灯片尺寸
type SlideSize struct {
	Width  int // 宽度 (px)
	Height int // 高度 (px)
}

// 标准幻灯片尺寸变量
var (
	// SlideSize16x9 宽屏幻灯片尺寸 (16:9)
	// 宽度: 1280 px (13.333 英寸)
	// 高度: 720 px (7.5 英寸)
	SlideSize16x9 = SlideSize{Width: 1280, Height: 720}

	// SlideSize4x3 标准幻灯片尺寸 (4:3)
	// 宽度: 960 px (10 英寸)
	// 高度: 720 px (7.5 英寸)
	SlideSize4x3 = SlideSize{Width: 960, Height: 720}

	// SlideSize16x10 超宽屏幻灯片尺寸 (16:10)
	// 宽度: 1280 px (13.333 英寸)
	// 高度: 800 px (8.333 英寸)
	SlideSize16x10 = SlideSize{Width: 1280, Height: 800}
)

// ============================================================================
// 边界标记系统 - 视口检查
// ============================================================================
//
// 边界标记用于标记元素相对于幻灯片视口的位置状态：
// - 允许元素越界（不会阻止操作）
// - 提供边界检查结果供后续处理参考
// - 支持批量元素的边界检测

// BoundaryStatus 边界状态
type BoundaryStatus int

const (
	// BoundaryStatusInside 完全在边界内
	BoundaryStatusInside BoundaryStatus = iota
	// BoundaryStatusPartial 部分越界
	BoundaryStatusPartial
	// BoundaryStatusOutside 完全越界
	BoundaryStatusOutside
	// BoundaryStatusOverflowRight 右侧越界
	BoundaryStatusOverflowRight
	// BoundaryStatusOverflowLeft 左侧越界
	BoundaryStatusOverflowLeft
	// BoundaryStatusOverflowTop 顶部越界
	BoundaryStatusOverflowTop
	// BoundaryStatusOverflowBottom 底部越界
	BoundaryStatusOverflowBottom
)

// String 返回边界状态的字符串表示
func (bs BoundaryStatus) String() string {
	switch bs {
	case BoundaryStatusInside:
		return "Inside"
	case BoundaryStatusPartial:
		return "Partial"
	case BoundaryStatusOutside:
		return "Outside"
	case BoundaryStatusOverflowRight:
		return "OverflowRight"
	case BoundaryStatusOverflowLeft:
		return "OverflowLeft"
	case BoundaryStatusOverflowTop:
		return "OverflowTop"
	case BoundaryStatusOverflowBottom:
		return "OverflowBottom"
	default:
		return "Unknown"
	}
}

// BoundaryCheckResult 边界检查结果
type BoundaryCheckResult struct {
	// Status 边界状态
	Status BoundaryStatus
	// ElementRect 元素矩形 (x, y, cx, cy in px)
	ElementRect Rect
	// ViewportRect 视口矩形 (0, 0, width, height in px)
	ViewportRect Rect
	// OverflowX X 方向越界量 (正数表示越出右边界，负数表示越出左边界)
	OverflowX int
	// OverflowY Y 方向越界量 (正数表示越出下边界，负数表示越出上边界)
	OverflowY int
	// IsVisible 是否有部分可见（至少有部分在视口内）
	IsVisible bool
}

// Rect 矩形区域
type Rect struct {
	X, Y    int // 左上角坐标 (px)
	Cx, Cy  int // 宽度和高度 (px)
}

// SlideViewport 幻灯片视口
type SlideViewport struct {
	// Width 视口宽度 (px)
	Width int
	// Height 视口高度 (px)
	Height int
	// Size 标准尺寸名称（可选）
	SizeName string
}

// NewSlideViewport 创建幻灯片视口
func NewSlideViewport(width, height int) *SlideViewport {
	return &SlideViewport{
		Width:  width,
		Height: height,
	}
}

// NewSlideViewportFromSize 从 SlideSize 创建视口
func NewSlideViewportFromSize(size SlideSize) *SlideViewport {
	return &SlideViewport{
		Width:  size.Width,
		Height: size.Height,
	}
}

// Rect 返回视口矩形
func (vp *SlideViewport) Rect() Rect {
	return Rect{X: 0, Y: 0, Cx: vp.Width, Cy: vp.Height}
}

// CheckBoundary 检查元素边界
// x, y: 元素左上角坐标 (px)
// cx, cy: 元素宽度和高度 (px)
func (vp *SlideViewport) CheckBoundary(x, y, cx, cy int) BoundaryCheckResult {
	elementRect := Rect{X: x, Y: y, Cx: cx, Cy: cy}
	viewportRect := vp.Rect()

	result := BoundaryCheckResult{
		ElementRect:  elementRect,
		ViewportRect: viewportRect,
		IsVisible:    true,
	}

	// 计算元素右下角
	elementRight := x + cx
	elementBottom := y + cy

	// 计算越界量
	// 右边界越界（正数）或左边界越界（负数）
	result.OverflowX = elementRight - vp.Width
	if x < 0 {
		result.OverflowX = x - 0 // 负数表示左越界
	}

	// 下边界越界（正数）或上边界越界（负数）
	result.OverflowY = elementBottom - vp.Height
	if y < 0 {
		result.OverflowY = y - 0 // 负数表示上越界
	}

	// 判断是否可见（至少有部分在视口内）
	result.IsVisible = !(elementRight <= 0 || x >= vp.Width ||
		elementBottom <= 0 || y >= vp.Height)

	// 判断边界状态
	if x >= 0 && y >= 0 && elementRight <= vp.Width && elementBottom <= vp.Height {
		// 完全在边界内
		result.Status = BoundaryStatusInside
	} else if elementRight <= 0 || x >= vp.Width || elementBottom <= 0 || y >= vp.Height {
		// 完全越界
		result.Status = BoundaryStatusOutside
		result.IsVisible = false
	} else {
		// 部分越界，判断具体方向
		result.Status = BoundaryStatusPartial
	}

	return result
}

// CheckRect 检查矩形边界
func (vp *SlideViewport) CheckRect(rect Rect) BoundaryCheckResult {
	return vp.CheckBoundary(rect.X, rect.Y, rect.Cx, rect.Cy)
}

// IsInside 检查元素是否完全在边界内
func (vp *SlideViewport) IsInside(x, y, cx, cy int) bool {
	return vp.CheckBoundary(x, y, cx, cy).Status == BoundaryStatusInside
}

// IsVisible 检查元素是否有部分可见
func (vp *SlideViewport) IsVisible(x, y, cx, cy int) bool {
	return vp.CheckBoundary(x, y, cx, cy).IsVisible
}

// Slide 高层幻灯片对象
type Slide struct {
	// 所属演示文稿
	presentation *Presentation

	// 底层幻灯片部件
	part *parts.SlidePart

	// 幻灯片构建器
	builder *SlideBuilder

	// 媒体管理器（引用自 Presentation）
	mediaManager *MediaManager

	// 幻灯片索引（从 0 开始）
	index int

	// 核心护城河：高并发无锁原子计数器
	// 用于分配唯一的形状 ID，保证并发安全
	shapeIDCounter atomic.Uint32
}

// ============================================================================
// 基本信息
// ============================================================================

// Index 返回幻灯片索引（从 0 开始）
func (s *Slide) Index() int {
	return s.index
}

// Part 返回底层 SlidePart
func (s *Slide) Part() *parts.SlidePart {
	return s.part
}

// Builder 返回幻灯片构建器
func (s *Slide) Builder() *SlideBuilder {
	return s.builder
}

// PartURI 返回部件 URI
func (s *Slide) PartURI() *opc.PackURI {
	return s.part.PartURI()
}

// ============================================================================
// 视口与边界检查
// ============================================================================

// Viewport 返回幻灯片视口
func (s *Slide) Viewport() *SlideViewport {
	cx, cy := s.SlideSize()
	return NewSlideViewport(cx, cy)
}

// CheckBoundary 检查元素边界
// x, y: 元素左上角坐标 (px)
// cx, cy: 元素宽度和高度 (px)
// 返回边界检查结果，包含越界信息和可见性状态
func (s *Slide) CheckBoundary(x, y, cx, cy int) BoundaryCheckResult {
	return s.Viewport().CheckBoundary(x, y, cx, cy)
}

// IsInsideBoundary 检查元素是否完全在边界内
func (s *Slide) IsInsideBoundary(x, y, cx, cy int) bool {
	return s.Viewport().IsInside(x, y, cx, cy)
}

// IsVisible 检查元素是否有部分可见
func (s *Slide) IsVisible(x, y, cx, cy int) bool {
	return s.Viewport().IsVisible(x, y, cx, cy)
}

// ============================================================================
// 组件系统
// ============================================================================

// AddComponent 添加组件到幻灯片
// 接收任何实现了 Component 接口的积木（文本、图片、图表）
// 内部生成一个 SlideContext，并调用组件的 c.Render(ctx) 方法
func (s *Slide) AddComponent(c Component) error {
	ctx := NewSlideContext(s)
	return c.Render(ctx)
}

// AddComponents 批量添加组件
func (s *Slide) AddComponents(components ...Component) error {
	ctx := NewSlideContext(s)
	return ctx.RenderComponents(components...)
}

// NewContext 创建幻灯片上下文（用于手动组件渲染）
func (s *Slide) NewContext() *SlideContext {
	return NewSlideContext(s)
}

// ============================================================================
// 文本添加方法 - 默认单位: px
// ============================================================================

// AddTextBox 添加文本框
// x, y: 位置（px 单位）
// cx, cy: 尺寸（px 单位）
// text: 文本内容
func (s *Slide) AddTextBox(x, y, cx, cy int, text string) *parts.XSp {
	return s.builder.AddTextBox(
		PxToEMU(x), PxToEMU(y),
		PxToEMU(cx), PxToEMU(cy),
		text,
	)
}

// ============================================================================
// 形状添加方法 - 默认单位: px
// ============================================================================

// AddAutoShape 添加自动形状
// x, y: 位置（px 单位）
// cx, cy: 尺寸（px 单位）
// presetID: 预设形状类型（如 "rectangle", "ellipse", "roundRect"）
func (s *Slide) AddAutoShape(x, y, cx, cy int, presetID string) *parts.XSp {
	return s.builder.AddAutoShape(
		PxToEMU(x), PxToEMU(y),
		PxToEMU(cx), PxToEMU(cy),
		presetID,
	)
}

// AddRectangle 添加矩形
func (s *Slide) AddRectangle(x, y, cx, cy int) *parts.XSp {
	return s.AddAutoShape(x, y, cx, cy, "rect")
}

// AddEllipse 添加椭圆
func (s *Slide) AddEllipse(x, y, cx, cy int) *parts.XSp {
	return s.AddAutoShape(x, y, cx, cy, "ellipse")
}

// AddRoundRect 添加圆角矩形
func (s *Slide) AddRoundRect(x, y, cx, cy int) *parts.XSp {
	return s.AddAutoShape(x, y, cx, cy, "roundRect")
}

// ============================================================================
// 图片添加方法 - 默认单位: px
// ============================================================================

// AddPicture 添加图片
// x, y: 位置（px 单位）
// cx, cy: 尺寸（px 单位）
// imageRId: 图片关系 ID
func (s *Slide) AddPicture(x, y, cx, cy int, imageRId string) *parts.XPicture {
	return s.builder.AddPicture(
		PxToEMU(x), PxToEMU(y),
		PxToEMU(cx), PxToEMU(cy),
		imageRId,
	)
}

// AddPictureFromBytes 从字节数据添加图片
// 自动处理媒体资源的添加和关系 ID 分配
func (s *Slide) AddPictureFromBytes(x, y, cx, cy int, fileName string, data []byte) (*parts.XPicture, error) {
	// 添加媒体资源
	_, resource := s.mediaManager.AddMediaAuto(fileName, data)
	if resource == nil {
		return nil, nil
	}

	// 获取目标 URI
	targetURI := resource.Target()

	// 添加关系到幻灯片
	slideRID := s.builder.AddImage(targetURI)

	// 添加图片形状
	return s.builder.AddPicture(
		PxToEMU(x), PxToEMU(y),
		PxToEMU(cx), PxToEMU(cy),
		slideRID,
	), nil
}

// AddPictureFromFile 从文件添加图片
func (s *Slide) AddPictureFromFile(x, y, cx, cy int, path string) (*parts.XPicture, error) {
	// 读取文件
	data, err := io.ReadAll(nil) // TODO: 实际读取文件
	if err != nil {
		return nil, err
	}

	return s.AddPictureFromBytes(x, y, cx, cy, path, data)
}

// ============================================================================
// 表格添加方法 - 默认单位: px
// ============================================================================

// AddTable 添加表格
// x, y: 位置（px 单位）
// cx, cy: 尺寸（px 单位）
// rows, cols: 行列数
func (s *Slide) AddTable(x, y, cx, cy, rows, cols int) *parts.XGraphicFrame {
	return s.builder.AddTable(
		PxToEMU(x), PxToEMU(y),
		PxToEMU(cx), PxToEMU(cy),
		rows, cols,
	)
}

// SetTableCellText 设置表格单元格文本
func (s *Slide) SetTableCellText(gf *parts.XGraphicFrame, row, col int, text string) {
	s.builder.SetTableCellText(gf, row, col, text)
}

// ============================================================================
// 关系管理方法
// ============================================================================

// AddImageRel 添加图片关系
func (s *Slide) AddImageRel(targetURI string) string {
	return s.builder.AddImage(targetURI)
}

// AddMediaRel 添加媒体关系
func (s *Slide) AddMediaRel(targetURI string) string {
	return s.builder.AddMedia(targetURI)
}

// AddChartRel 添加图表关系
func (s *Slide) AddChartRel(targetURI string) string {
	return s.builder.AddChart(targetURI)
}

// HasImage 判断是否已存在某图片关系
func (s *Slide) HasImage(targetURI string) bool {
	return s.builder.HasImage(targetURI)
}

// GetImageRId 获取图片 rId，不存在则添加
func (s *Slide) GetImageRId(targetURI string) string {
	return s.builder.GetImageRId(targetURI)
}

// ============================================================================
// 幻灯片尺寸 - 默认单位: px
// ============================================================================

// SlideSize 返回幻灯片尺寸（px 单位）
func (s *Slide) SlideSize() (cx, cy int) {
	emuCX, emuCY := s.presentation.SlideSize()
	return EMUToPx(emuCX), EMUToPx(emuCY)
}

// SlideSizeEMU 返回幻灯片尺寸（EMU 单位，高级用法）
func (s *Slide) SlideSizeEMU() (cx, cy int) {
	return s.presentation.SlideSize()
}

// ============================================================================
// 单位转换 - 委托给 utils 包
// ============================================================================

// PxToEMU 将像素转换为 EMU（基于 96 DPI）
func PxToEMU(px int) int {
	return int(utils.PixelsToEMU(float64(px)))
}

// EMUToPx 将 EMU 转换为像素（基于 96 DPI）
func EMUToPx(emu int) int {
	return int(utils.EMUToPixels(int64(emu)))
}

// ============================================================================
// 颜色工具方法
// ============================================================================

// ValidateColor 验证颜色
func (s *Slide) ValidateColor(color string) ColorValidationResult {
	return ValidateColor(color)
}

// ResolveColor 解析颜色（支持名称、十六进制、RGB、主题色）
func (s *Slide) ResolveColor(color string) Color {
	return DefaultColorMap().Resolve(color)
}

// ============================================================================
// 布局管理方法
// ============================================================================

// SetLayout 设置幻灯片布局
// layoutName: 布局名称（如 "blank", "title", "titleAndContent" 等）
// 返回是否设置成功
func (s *Slide) SetLayout(layoutName string) bool {
	if s.presentation.masterCache == nil {
		return false
	}

	layoutData, ok := s.presentation.masterCache.GetLayoutByName(layoutName)
	if !ok {
		return false
	}

	// 设置布局关系 ID
	s.part.SetLayoutRId(layoutData.ID())
	return true
}

// Layout 返回当前布局名称
func (s *Slide) Layout() string {
	layoutRId := s.part.LayoutRId()
	if layoutRId == "" {
		return ""
	}

	// 从 masterCache 中查找布局名称
	if s.presentation.masterCache != nil {
		if layout, ok := s.presentation.masterCache.GetLayout(layoutRId); ok {
			return layout.Name()
		}
	}

	return layoutRId
}
