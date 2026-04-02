// Package pptx 提供 PPTX 文件的高级操作接口
package pptx

import (
	"github.com/Muprprpr/Go-pptx/parts"
)

// ============================================================================
// Component 接口 - 积木抽象
// ============================================================================
//
// Component 是所有可渲染到幻灯片上的组件的统一接口。
// 任何实现了此接口的类型都可以作为"积木"添加到幻灯片。
//
// 设计原则：
// 1. 单一职责 - 组件只负责生成自己的 XML 结构
// 2. 依赖注入 - 通过 SlideContext 获取所需资源（ID、媒体等）
// 3. 可组合性 - 多个组件可以组合成更复杂的组件
//
// 使用示例：
//
//	type TitleComponent struct {
//		text string
//	}
//
//	func (t *TitleComponent) Render(ctx *SlideContext) error {
//		id := ctx.NextShapeID()
//		sp := &parts.XSp{...}
//		ctx.AppendShape(sp)
//		return nil
//	}
//
//	slide.AddComponent(&TitleComponent{text: "Hello"})
//
// ============================================================================

// Component 组件接口
// 所有可渲染到幻灯片的积木必须实现此接口
type Component interface {
	// Render 将组件渲染到幻灯片
	// ctx: 提供组件所需的上下文和资源访问能力
	// 返回 error 表示渲染失败
	Render(ctx *SlideContext) error
}

// ============================================================================
// 常用组件接口扩展
// ============================================================================

// ComponentWithSize 带尺寸信息的组件
type ComponentWithSize interface {
	Component
	// Bounds 返回组件的边界框 (x, y, cx, cy in EMU)
	Bounds() (x, y, cx, cy int)
}

// ComponentWithName 带名称的组件
type ComponentWithName interface {
	Component
	// Name 返回组件名称（用于调试和日志）
	Name() string
}

// ComponentWithPosition 可定位的组件
type ComponentWithPosition interface {
	Component
	// SetPosition 设置组件位置 (EMU 单位)
	SetPosition(x, y int)
	// Position 返回组件位置
	Position() (x, y int)
}

// ComponentWithSizeSetter 可调整尺寸的组件
type ComponentWithSizeSetter interface {
	Component
	// SetSize 设置组件尺寸 (EMU 单位)
	SetSize(cx, cy int)
	// Size 返回组件尺寸
	Size() (cx, cy int)
}

// ============================================================================
// 组件工具函数
// ============================================================================

// ComponentList 组件列表
// 用于批量管理组件
type ComponentList []Component

// Add 添加组件到列表
func (cl *ComponentList) Add(c Component) {
	*cl = append(*cl, c)
}

// RenderAll 渲染所有组件
func (cl ComponentList) RenderAll(ctx *SlideContext) error {
	for i, c := range cl {
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

// Count 返回组件数量
func (cl ComponentList) Count() int {
	return len(cl)
}

// ============================================================================
// 错误类型
// ============================================================================

// ComponentRenderError 组件渲染错误
type ComponentRenderError struct {
	Index      int
	Component  Component
	Underlying error
}

// Error 实现 error 接口
func (e *ComponentRenderError) Error() string {
	name := "<unknown>"
	if n, ok := e.Component.(ComponentWithName); ok {
		name = n.Name()
	}
	return "component render error at index " + string(rune(e.Index)) + " (" + name + "): " + e.Underlying.Error()
}

// Unwrap 返回底层错误
func (e *ComponentRenderError) Unwrap() error {
	return e.Underlying
}

// ============================================================================
// 基础组件实现
// ============================================================================

// FuncComponent 函数式组件
// 将普通函数包装为 Component 接口
type FuncComponent func(ctx *SlideContext) error

// Render 实现 Component 接口
func (fc FuncComponent) Render(ctx *SlideContext) error {
	return fc(ctx)
}

// ============================================================================
// 组合组件
// ============================================================================

// CompositeComponent 组合组件
// 将多个组件组合为一个
type CompositeComponent struct {
	components []Component
	name       string
}

// NewCompositeComponent 创建组合组件
func NewCompositeComponent(name string, components ...Component) *CompositeComponent {
	return &CompositeComponent{
		components: components,
		name:       name,
	}
}

// Add 添加子组件
func (cc *CompositeComponent) Add(c Component) {
	cc.components = append(cc.components, c)
}

// Render 实现 Component 接口
func (cc *CompositeComponent) Render(ctx *SlideContext) error {
	for i, c := range cc.components {
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

// Name 实现 ComponentWithName 接口
func (cc *CompositeComponent) Name() string {
	return cc.name
}

// Components 返回所有子组件
func (cc *CompositeComponent) Components() []Component {
	return cc.components
}

// ============================================================================
// 条件组件
// ============================================================================

// ConditionalComponent 条件组件
// 根据条件决定是否渲染
type ConditionalComponent struct {
	condition  func() bool
	component  Component
	elseComponent Component
}

// NewConditionalComponent 创建条件组件
func NewConditionalComponent(condition func() bool, ifComponent, elseComponent Component) *ConditionalComponent {
	return &ConditionalComponent{
		condition:     condition,
		component:     ifComponent,
		elseComponent: elseComponent,
	}
}

// Render 实现 Component 接口
func (cc *ConditionalComponent) Render(ctx *SlideContext) error {
	if cc.condition() {
		if cc.component != nil {
			return cc.component.Render(ctx)
		}
	} else {
		if cc.elseComponent != nil {
			return cc.elseComponent.Render(ctx)
		}
	}
	return nil
}

// ============================================================================
// 重复组件
// ============================================================================

// RepeatedComponent 重复组件
// 根据数据切片重复渲染组件
type RepeatedComponent struct {
	template func(index int) Component
	count    int
}

// NewRepeatedComponent 创建重复组件
func NewRepeatedComponent(count int, template func(index int) Component) *RepeatedComponent {
	return &RepeatedComponent{
		template: template,
		count:    count,
	}
}

// Render 实现 Component 接口
func (rc *RepeatedComponent) Render(ctx *SlideContext) error {
	for i := 0; i < rc.count; i++ {
		c := rc.template(i)
		if c == nil {
			continue
		}
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

// ============================================================================
// 形状组件辅助
// ============================================================================

// ShapeComponent 形状组件
// 最基础的组件类型，直接包装 XSp
type ShapeComponent struct {
	sp   *parts.XSp
	x, y int
	name string
}

// NewShapeComponent 创建形状组件
func NewShapeComponent(sp *parts.XSp, x, y int) *ShapeComponent {
	return &ShapeComponent{
		sp: sp,
		x:  x,
		y:  y,
	}
}

// Render 实现 Component 接口
func (sc *ShapeComponent) Render(ctx *SlideContext) error {
	if sc.sp == nil {
		return nil
	}
	ctx.AppendShape(sc.sp)
	return nil
}

// Name 实现 ComponentWithName 接口
func (sc *ShapeComponent) Name() string {
	return sc.name
}

// SetName 设置名称
func (sc *ShapeComponent) SetName(name string) {
	sc.name = name
}

// Bounds 实现 ComponentWithSize 接口
func (sc *ShapeComponent) Bounds() (x, y, cx, cy int) {
	if sc.sp != nil && sc.sp.ShapeProperties != nil && sc.sp.ShapeProperties.Transform2D != nil {
		x = sc.sp.ShapeProperties.Transform2D.Offset.X
		y = sc.sp.ShapeProperties.Transform2D.Offset.Y
		cx = sc.sp.ShapeProperties.Transform2D.Extent.Cx
		cy = sc.sp.ShapeProperties.Transform2D.Extent.Cy
	}
	return
}

// SetPosition 实现 ComponentWithPosition 接口
func (sc *ShapeComponent) SetPosition(x, y int) {
	sc.x = x
	sc.y = y
	if sc.sp != nil && sc.sp.ShapeProperties != nil && sc.sp.ShapeProperties.Transform2D != nil {
		sc.sp.ShapeProperties.Transform2D.Offset.X = x
		sc.sp.ShapeProperties.Transform2D.Offset.Y = y
	}
}

// Position 实现 ComponentWithPosition 接口
func (sc *ShapeComponent) Position() (x, y int) {
	return sc.x, sc.y
}
