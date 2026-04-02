# Component - 组件系统

组件系统提供可复用的渲染积木，支持自定义组件和组合模式。

## 核心接口

### Component

所有可渲染到幻灯片的积木必须实现此接口。

```go
type Component interface {
    // Render 将组件渲染到幻灯片
    // ctx: 提供组件所需的上下文和资源访问能力
    // 返回 error 表示渲染失败
    Render(ctx *SlideContext) error
}
```

## 内置组件

### ShapeComponent

形状组件，最基础的组件类型，直接包装 XSp。

```go
type ShapeComponent struct {
    // Has unexported fields.
}
```

#### 构造函数

```go
func NewShapeComponent(sp *parts.XSp, x, y int) *ShapeComponent
```

**参数:**
- `sp`: 形状对象
- `x, y`: 位置（EMU 单位）

#### 方法

```go
// Render 实现 Component 接口
func (sc *ShapeComponent) Render(ctx *SlideContext) error

// Name 实现 ComponentWithName 接口
func (sc *ShapeComponent) Name() string

// SetName 设置名称
func (sc *ShapeComponent) SetName(name string)

// Position 实现 ComponentWithPosition 接口
func (sc *ShapeComponent) Position() (x, y int)

// SetPosition 实现 ComponentWithPosition 接口
func (sc *ShapeComponent) SetPosition(x, y int)

// Bounds 实现 ComponentWithSize 接口
func (sc *ShapeComponent) Bounds() (x, y, cx, cy int)
```

### CompositeComponent

组合组件，将多个组件组合为一个。

```go
type CompositeComponent struct {
    // Has unexported fields.
}
```

#### 构造函数

```go
func NewCompositeComponent(name string, components ...Component) *CompositeComponent
```

**示例:**

```go
// 创建组合组件
header := NewCompositeComponent("header",
    &TitleComponent{Text: "标题"},
    &SubtitleComponent{Text: "副标题"},
)
```

#### 方法

```go
// Render 实现 Component 接口
func (cc *CompositeComponent) Render(ctx *SlideContext) error

// Add 添加子组件
func (cc *CompositeComponent) Add(c Component)

// Components 返回所有子组件
func (cc *CompositeComponent) Components() []Component

// Name 实现 ComponentWithName 接口
func (cc *CompositeComponent) Name() string
```

### ConditionalComponent

条件组件，根据条件决定是否渲染。

```go
type ConditionalComponent struct {
    // Has unexported fields.
}
```

#### 构造函数

```go
func NewConditionalComponent(condition func() bool, ifComponent, elseComponent Component) *ConditionalComponent
```

**参数:**
- `condition`: 条件函数
- `ifComponent`: 条件为真时渲染的组件
- `elseComponent`: 条件为假时渲染的组件（可为 nil）

**示例:**

```go
// 根据条件显示不同的组件
condComp := NewConditionalComponent(
    func() bool { return showAdvanced },
    &AdvancedComponent{},
    &SimpleComponent{},
)
```

#### 方法

```go
// Render 实现 Component 接口
func (cc *ConditionalComponent) Render(ctx *SlideContext) error
```

### RepeatedComponent

重复组件，根据数据切片重复渲染组件。

```go
type RepeatedComponent struct {
    // Has unexported fields.
}
```

#### 构造函数

```go
func NewRepeatedComponent(count int, template func(index int) Component) *RepeatedComponent
```

**参数:**
- `count`: 重复次数
- `template`: 组件模板函数，接收索引返回组件

**示例:**

```go
// 创建列表组件
list := NewRepeatedComponent(5, func(index int) Component {
    return &ListItemComponent{
        Text:  fmt.Sprintf("项目 %d", index+1),
        Index: index,
    }
})
```

#### 方法

```go
// Render 实现 Component 接口
func (rc *RepeatedComponent) Render(ctx *SlideContext) error
```

### FuncComponent

函数式组件，将普通函数包装为 Component 接口。

```go
type FuncComponent func(ctx *SlideContext) error
```

**示例:**

```go
// 创建函数式组件
myFunc := FuncComponent(func(ctx *SlideContext) error {
    // 直接使用上下文渲染
    ctx.AppendShape(&parts.XSp{
        // ... 形状属性
    })
    return nil
})

// 添加到幻灯片
slide.AddComponent(myFunc)
```

#### 方法

```go
// Render 实现 Component 接口
func (fc FuncComponent) Render(ctx *SlideContext) error
```

## 扩展接口

### ComponentWithName

带名称的组件。

```go
type ComponentWithName interface {
    Component
    // Name 返回组件名称（用于调试和日志）
    Name() string
}
```

### ComponentWithPosition

可定位的组件。

```go
type ComponentWithPosition interface {
    Component
    // SetPosition 设置组件位置 (EMU 单位)
    SetPosition(x, y int)
    // Position 返回组件位置
    Position() (x, y int)
}
```

### ComponentWithSize

带尺寸信息的组件。

```go
type ComponentWithSize interface {
    Component
    // Bounds 返回组件的边界框 (x, y, cx, cy in EMU)
    Bounds() (x, y, cx, cy int)
}
```

### ComponentWithSizeSetter

可调整尺寸的组件。

```go
type ComponentWithSizeSetter interface {
    Component
    // SetSize 设置组件尺寸 (EMU 单位)
    SetSize(cx, cy int)
    // Size 返回组件尺寸
    Size() (cx, cy int)
}
```

## 组件列表

### ComponentList

组件列表，用于批量管理组件。

```go
type ComponentList []Component
```

#### 方法

```go
// Add 添加组件到列表
func (cl *ComponentList) Add(c Component)

// Count 返回组件数量
func (cl ComponentList) Count() int

// RenderAll 渲染所有组件
func (cl ComponentList) RenderAll(ctx *SlideContext) error
```

**示例:**

```go
var list ComponentList
list.Add(&TitleComponent{Text: "标题"})
list.Add(&BodyComponent{Text: "正文"})

fmt.Printf("组件数量: %d\n", list.Count())

// 渲染所有组件
ctx := slide.NewContext()
list.RenderAll(ctx)
```

## 错误处理

### ComponentRenderError

组件渲染错误。

```go
type ComponentRenderError struct {
    Index      int       // 组件索引
    Component  Component // 组件对象
    Underlying error     // 底层错误
}
```

#### 方法

```go
// Error 实现 error 接口
func (e *ComponentRenderError) Error() string

// Unwrap 返回底层错误
func (e *ComponentRenderError) Unwrap() error
```

## 自定义组件示例

### 基础自定义组件

```go
// TitleComponent 标题组件
type TitleComponent struct {
    Text string
    X, Y int
    FontSize int
    Color    string
}

func (t *TitleComponent) Render(ctx *SlideContext) error {
    // 默认值
    if t.FontSize == 0 {
        t.FontSize = 44
    }
    if t.Color == "" {
        t.Color = "000000"
    }

    // 创建文本形状
    sp := &parts.XSp{
        NvSpPr: &parts.XNvSpPr{
            CNvPr: &parts.XCNvPr{
                ID:   ctx.NextShapeID(),
                Name: "Title",
            },
        },
        SpPr: &parts.XSpPr{
            Xfrm: &parts.XXfrm{
                Off: &parts.XPoint{
                    X: ctx.PxToEMU(t.X),
                    Y: ctx.PxToEMU(t.Y),
                },
                Ext: &parts.XSize{
                    Cx: ctx.PxToEMU(600),
                    Cy: ctx.PxToEMU(60),
                },
            },
        },
        TxBody: &parts.XTxBody{
            BodyPr: &parts.XBodyPr{
                Wrap: "square",
            },
            P: []*parts.XP{
                {
                    R: []*parts.XR{
                        {
                            T:  t.Text,
                            RPr: &parts.XRPr{
                                Sz:     t.FontSize * 100,
                                SolidFill: &parts.XSolidFill{
                                    SrgbClr: &parts.XSrgbClr{
                                        Val: t.Color,
                                    },
                                },
                            },
                        },
                    },
                },
            },
        },
    }

    ctx.AppendShape(sp)
    return nil
}
```

### 带图片的组件

```go
// ImageComponent 图片组件
type ImageComponent struct {
    X, Y     int
    Cx, Cy   int
    ImagePath string
}

func (img *ImageComponent) Render(ctx *SlideContext) error {
    // 读取图片文件
    data, err := os.ReadFile(img.ImagePath)
    if err != nil {
        return err
    }

    // 添加媒体资源
    rId, err := ctx.AddImage(data, filepath.Base(img.ImagePath))
    if err != nil {
        return err
    }

    // 创建图片形状
    pic := &parts.XPicture{
        NvPicPr: &parts.XNvPicPr{
            CNvPr: &parts.XCNvPr{
                ID:   ctx.NextShapeID(),
                Name: "Picture",
            },
        },
        BlipFill: &parts.XBlipFill{
            Blip: &parts.XBlip{
                Embed: rId,
            },
        },
        SpPr: &parts.XSpPr{
            Xfrm: &parts.XXfrm{
                Off: &parts.XPoint{
                    X: ctx.PxToEMU(img.X),
                    Y: ctx.PxToEMU(img.Y),
                },
                Ext: &parts.XSize{
                    Cx: ctx.PxToEMU(img.Cx),
                    Cy: ctx.PxToEMU(img.Cy),
                },
            },
        },
    }

    ctx.AppendShape(pic)
    return nil
}
```

### 组合组件示例

```go
// CardComponent 卡片组件
type CardComponent struct {
    X, Y       int
    Width      int
    Title      string
    Content    string
    Background string
}

func (c *CardComponent) Render(ctx *SlideContext) error {
    // 创建组合组件
    card := NewCompositeComponent("Card")

    // 添加背景矩形
    bgColor := c.Background
    if bgColor == "" {
        bgColor = "F5F5F5"
    }

    bgShape := &parts.XSp{
        // ... 设置背景矩形属性
    }
    card.Add(NewShapeComponent(bgShape, ctx.PxToEMU(c.X), ctx.PxToEMU(c.Y)))

    // 添加标题
    card.Add(&TitleComponent{
        Text: c.Title,
        X:    c.X + 20,
        Y:    c.Y + 20,
    })

    // 添加内容
    card.Add(&BodyComponent{
        Text: c.Content,
        X:    c.X + 20,
        Y:    c.Y + 80,
    })

    return card.Render(ctx)
}
```

### 条件和重复组件组合

```go
// ListComponent 列表组件
type ListComponent struct {
    X, Y      int
    ItemWidth int
    Items     []string
    ShowIndex bool
}

func (l *ListComponent) Render(ctx *SlideContext) error {
    // 创建重复组件
    list := NewRepeatedComponent(len(l.Items), func(index int) Component {
        // 创建列表项组合
        item := NewCompositeComponent(fmt.Sprintf("Item%d", index))

        // 可选：添加序号
        if l.ShowIndex {
            item.Add(NewConditionalComponent(
                func() bool { return l.ShowIndex },
                &TextComponent{
                    Text: fmt.Sprintf("%d.", index+1),
                    X:    l.X,
                    Y:    l.Y + index*40,
                },
                nil,
            ))
        }

        // 添加文本
        item.Add(&TextComponent{
            Text: l.Items[index],
            X:    l.X + 30,
            Y:    l.Y + index*40,
        })

        return item
    })

    return list.Render(ctx)
}
```

## 最佳实践

### 1. 组件职责单一

每个组件应该只负责一个功能：

```go
// 好的做法
type TitleComponent struct { /* ... */ }
type SubtitleComponent struct { /* ... */ }
type BodyComponent struct { /* ... */ }

// 不好的做法
type EverythingComponent struct {
    Title, Subtitle, Body, Image string
}
```

### 2. 使用组合而非继承

Go 没有继承，使用组合来构建复杂组件：

```go
// 使用组合构建复杂组件
func NewSlideLayout(title, content string) Component {
    return NewCompositeComponent("SlideLayout",
        &TitleComponent{Text: title},
        &BodyComponent{Text: content},
    )
}
```

### 3. 提供合理的默认值

```go
func (t *TextComponent) Render(ctx *SlideContext) error {
    // 设置默认值
    if t.FontSize == 0 {
        t.FontSize = 18
    }
    if t.Color == "" {
        t.Color = "000000"
    }
    // ...
}
```

### 4. 错误处理

```go
func (c *ImageComponent) Render(ctx *SlideContext) error {
    if c.ImagePath == "" {
        return fmt.Errorf("image path is required")
    }

    data, err := os.ReadFile(c.ImagePath)
    if err != nil {
        return fmt.Errorf("failed to read image: %w", err)
    }

    // ...
}
```

### 5. 使用 ComponentList 批量管理

```go
func CreateSlide(slide *pptx.Slide) error {
    var components ComponentList

    components.Add(&HeaderComponent{})
    components.Add(&ContentComponent{})
    components.Add(&FooterComponent{})

    ctx := slide.NewContext()
    return components.RenderAll(ctx)
}
```
