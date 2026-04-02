package pptx_test

import (
	"testing"

	"github.com/Muprprpr/Go-pptx/parts"
	"github.com/Muprprpr/Go-pptx/pptx"
)

// ============================================================================
// 测试用组件定义
// ============================================================================

// TextComponent 简单文本组件
type TextComponent struct {
	x, y, cx, cy int
	text         string
}

func NewTextComponent(x, y, cx, cy int, text string) *TextComponent {
	return &TextComponent{
		x:    x,
		y:    y,
		cx:   cx,
		cy:   cy,
		text: text,
	}
}

func (t *TextComponent) Render(ctx *pptx.SlideContext) error {
	// 分配形状 ID
	id := ctx.NextShapeID()

	// 构建形状 XML
	sp := &parts.XSp{
		NonVisual: parts.XNonVisualDrawingShape{
			CNvPr: &parts.XNvCxnSpPr{
				ID:   int(id),
				Name: "TextBox",
			},
			CNvSpPr: &parts.XNvSpPr{},
		},
		ShapeProperties: &parts.XShapeProperties{
			Transform2D: &parts.XTransform2D{
				Offset: &parts.XOv2DrOffset{X: t.x, Y: t.y},
				Extent: &parts.XOv2DrExtent{Cx: t.cx, Cy: t.cy},
			},
		},
		TextBody: &parts.XTextBody{
			BodyPr:   &parts.XBodyPr{},
			LstStyle: &parts.XTextParagraphList{},
			Paragraphs: []parts.XTextParagraph{
				{
					TextRuns: []parts.XTextRun{
						{Text: t.text},
					},
				},
			},
		},
	}

	// 挂载到幻灯片
	ctx.AppendShape(sp)
	return nil
}

// RectangleComponent 矩形组件
type RectangleComponent struct {
	x, y, cx, cy int
}

func NewRectangleComponent(x, y, cx, cy int) *RectangleComponent {
	return &RectangleComponent{x: x, y: y, cx: cx, cy: cy}
}

func (r *RectangleComponent) Render(ctx *pptx.SlideContext) error {
	id := ctx.NextShapeID()

	sp := &parts.XSp{
		NonVisual: parts.XNonVisualDrawingShape{
			CNvPr: &parts.XNvCxnSpPr{
				ID:   int(id),
				Name: "Rectangle",
			},
			CNvSpPr: &parts.XNvSpPr{},
		},
		ShapeProperties: &parts.XShapeProperties{
			Transform2D: &parts.XTransform2D{
				Offset: &parts.XOv2DrOffset{X: r.x, Y: r.y},
				Extent: &parts.XOv2DrExtent{Cx: r.cx, Cy: r.cy},
			},
		},
		ShapePreset: "rect",
	}

	ctx.AppendShape(sp)
	return nil
}

// CompositeTestComponent 组合组件
type CompositeTestComponent struct {
	children []pptx.Component
}

func NewCompositeTestComponent(children ...pptx.Component) *CompositeTestComponent {
	return &CompositeTestComponent{children: children}
}

func (c *CompositeTestComponent) Render(ctx *pptx.SlideContext) error {
	for i, child := range c.children {
		if err := child.Render(ctx); err != nil {
			return err
		}
		_ = i // 使用 i 避免编译警告
	}
	return nil
}

// ============================================================================
// 测试用例
// ============================================================================

// TestAddComponent 测试添加组件
func TestAddComponent(t *testing.T) {
	pres := pptx.New()
	slide := pres.AddSlide()

	// 添加文本组件
	textComp := NewTextComponent(100, 100, 500, 50, "Hello Component")
	err := slide.AddComponent(textComp)
	if err != nil {
		t.Fatalf("AddComponent 失败: %v", err)
	}

	// 验证形状 ID 分配
	ctx := slide.NewContext()
	if ctx.CurrentShapeID() < 1 {
		t.Error("形状 ID 未正确分配")
	}
}

// TestAddComponents 批量添加组件
func TestAddComponents(t *testing.T) {
	pres := pptx.New()
	slide := pres.AddSlide()

	// 批量添加组件
	err := slide.AddComponents(
		NewTextComponent(100, 100, 500, 50, "Title"),
		NewRectangleComponent(100, 200, 300, 150),
		NewTextComponent(100, 400, 500, 30, "Footer"),
	)
	if err != nil {
		t.Fatalf("AddComponents 失败: %v", err)
	}
}

// TestCompositeComponent 测试组合组件
func TestCompositeComponent(t *testing.T) {
	pres := pptx.New()
	slide := pres.AddSlide()

	// 创建组合组件
	composite := NewCompositeTestComponent(
		NewTextComponent(100, 100, 500, 50, "Header"),
		NewRectangleComponent(100, 200, 300, 150),
	)

	err := slide.AddComponent(composite)
	if err != nil {
		t.Fatalf("AddComponent(组合组件) 失败: %v", err)
	}
}

// TestFuncComponent 测试函数式组件
func TestFuncComponent(t *testing.T) {
	pres := pptx.New()
	slide := pres.AddSlide()

	// 使用函数式组件
	renderCount := 0
	funcComp := pptx.FuncComponent(func(ctx *pptx.SlideContext) error {
		renderCount++
		id := ctx.NextShapeID()

		sp := &parts.XSp{
			NonVisual: parts.XNonVisualDrawingShape{
				CNvPr: &parts.XNvCxnSpPr{
					ID:   int(id),
					Name: "FuncComponent",
				},
				CNvSpPr: &parts.XNvSpPr{},
			},
			ShapeProperties: &parts.XShapeProperties{
				Transform2D: &parts.XTransform2D{
					Offset: &parts.XOv2DrOffset{X: 100, Y: 100},
					Extent: &parts.XOv2DrExtent{Cx: 200, Cy: 100},
				},
			},
		}

		ctx.AppendShape(sp)
		return nil
	})

	err := slide.AddComponent(funcComp)
	if err != nil {
		t.Fatalf("AddComponent(函数式组件) 失败: %v", err)
	}

	if renderCount != 1 {
		t.Errorf("期望 renderCount=1, 实际=%d", renderCount)
	}
}

// TestSlideContextShapeID 测试形状 ID 分配
func TestSlideContextShapeID(t *testing.T) {
	pres := pptx.New()
	slide := pres.AddSlide()
	ctx := slide.NewContext()

	// 分配多个形状 ID
	id1 := ctx.NextShapeID()
	id2 := ctx.NextShapeID()
	id3 := ctx.NextShapeID()

	if id1 == id2 || id2 == id3 || id1 == id3 {
		t.Error("形状 ID 重复分配")
	}

	// 验证 ID 递增
	if id2 <= id1 || id3 <= id2 {
		t.Error("形状 ID 未正确递增")
	}
}

// TestSlideContextSlideSize 测试幻灯片尺寸获取
func TestSlideContextSlideSize(t *testing.T) {
	pres := pptx.New()
	slide := pres.AddSlide()
	ctx := slide.NewContext()

	cx, cy := ctx.SlideSize()
	if cx == 0 || cy == 0 {
		t.Error("幻灯片尺寸无效")
	}

	// 验证默认尺寸（16:9 宽屏）
	if cx != 12192000 || cy != 6858000 {
		t.Logf("幻灯片尺寸: cx=%d, cy=%d (非默认 16:9)", cx, cy)
	}
}

// TestSlideContextUnitConversion 测试单位转换
func TestSlideContextUnitConversion(t *testing.T) {
	pres := pptx.New()
	slide := pres.AddSlide()
	ctx := slide.NewContext()

	// 测试 px -> EMU 转换 (96 DPI: 1 px = 9525 EMU)
	px := 96
	emu := ctx.PxToEMU(px)
	if emu != 914400 { // 96 * 9525 = 914400 (即 1 英寸)
		t.Errorf("PxToEMU(96) = %d, 期望 914400", emu)
	}

	// 测试 EMU -> px 转换
	backPx := ctx.EMUToPx(914400)
	if backPx != 96 {
		t.Errorf("EMUToPx(914400) = %d, 期望 96", backPx)
	}

	// 测试 1 px = 9525 EMU
	emu = ctx.PxToEMU(1)
	if emu != 9525 {
		t.Errorf("PxToEMU(1) = %d, 期望 9525", emu)
	}
}

// TestComponentWithSave 测试组件渲染后保存
func TestComponentWithSave(t *testing.T) {
	pres := pptx.New()
	slide := pres.AddSlide()

	// 添加多个组件
	slide.AddComponent(NewTextComponent(100, 100, 500, 50, "Title"))
	slide.AddComponent(NewRectangleComponent(100, 200, 300, 150))
	slide.AddComponent(NewTextComponent(100, 400, 500, 30, "Footer"))

	// 保存到字节
	data, err := pres.WriteToBytes()
	if err != nil {
		t.Fatalf("WriteToBytes 失败: %v", err)
	}

	if len(data) == 0 {
		t.Error("输出数据为空")
	}

	t.Logf("生成 PPTX 大小: %d 字节", len(data))
}

// TestBuiltInCompositeComponent 测试内置组合组件
func TestBuiltInCompositeComponent(t *testing.T) {
	pres := pptx.New()
	slide := pres.AddSlide()

	// 使用内置的 CompositeComponent
	composite := pptx.NewCompositeComponent("test-composite",
		NewTextComponent(100, 100, 500, 50, "Header"),
		NewRectangleComponent(100, 200, 300, 150),
	)

	err := slide.AddComponent(composite)
	if err != nil {
		t.Fatalf("AddComponent(CompositeComponent) 失败: %v", err)
	}

	if composite.Name() != "test-composite" {
		t.Errorf("期望名称 'test-composite', 实际 '%s'", composite.Name())
	}
}

// TestConditionalComponent 测试条件组件
func TestConditionalComponent(t *testing.T) {
	pres := pptx.New()
	slide := pres.AddSlide()

	condition := true

	// 创建条件组件
	condComp := pptx.NewConditionalComponent(
		func() bool { return condition },
		NewTextComponent(100, 100, 500, 50, "Condition is true"),
		NewTextComponent(100, 100, 500, 50, "Condition is false"),
	)

	err := slide.AddComponent(condComp)
	if err != nil {
		t.Fatalf("AddComponent(ConditionalComponent) 失败: %v", err)
	}
}

// TestRepeatedComponent 测试重复组件
func TestRepeatedComponent(t *testing.T) {
	pres := pptx.New()
	slide := pres.AddSlide()

	// 创建重复组件（生成 3 个文本框）
	repeatedComp := pptx.NewRepeatedComponent(3, func(index int) pptx.Component {
		return NewTextComponent(100, 100+index*60, 500, 50, "Item")
	})

	err := slide.AddComponent(repeatedComp)
	if err != nil {
		t.Fatalf("AddComponent(RepeatedComponent) 失败: %v", err)
	}
}

// TestShapeComponent 测试形状组件
func TestShapeComponent(t *testing.T) {
	pres := pptx.New()
	slide := pres.AddSlide()

	// 创建形状组件
	sp := &parts.XSp{
		NonVisual: parts.XNonVisualDrawingShape{
			CNvPr: &parts.XNvCxnSpPr{
				ID:   1,
				Name: "TestShape",
			},
		},
		ShapeProperties: &parts.XShapeProperties{
			Transform2D: &parts.XTransform2D{
				Offset: &parts.XOv2DrOffset{X: 100, Y: 100},
				Extent: &parts.XOv2DrExtent{Cx: 200, Cy: 100},
			},
		},
	}

	shapeComp := pptx.NewShapeComponent(sp, 100, 100)
	shapeComp.SetName("TestShapeComponent")

	err := slide.AddComponent(shapeComp)
	if err != nil {
		t.Fatalf("AddComponent(ShapeComponent) 失败: %v", err)
	}

	if shapeComp.Name() != "TestShapeComponent" {
		t.Errorf("期望名称 'TestShapeComponent', 实际 '%s'", shapeComp.Name())
	}
}

// TestComponentList 组件列表测试
func TestComponentList(t *testing.T) {
	pres := pptx.New()
	slide := pres.AddSlide()

	// 创建组件列表
	var list pptx.ComponentList
	list.Add(NewTextComponent(100, 100, 500, 50, "Item 1"))
	list.Add(NewTextComponent(100, 160, 500, 50, "Item 2"))
	list.Add(NewTextComponent(100, 220, 500, 50, "Item 3"))

	if list.Count() != 3 {
		t.Errorf("期望 3 个组件, 实际 %d", list.Count())
	}

	// 使用上下文渲染
	ctx := slide.NewContext()
	err := list.RenderAll(ctx)
	if err != nil {
		t.Fatalf("RenderAll 失败: %v", err)
	}
}
