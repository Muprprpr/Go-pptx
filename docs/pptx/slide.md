# Slide - 幻灯片

`Slide` 是高层幻灯片对象，提供添加文本、图片、表格、形状等元素的方法。

## 类型定义

```go
type Slide struct {
    // Has unexported fields.
}
```

## 基础信息

### Index

返回幻灯片索引（从 0 开始）。

```go
func (s *Slide) Index() int
```

### Layout

返回当前布局名称。

```go
func (s *Slide) Layout() string
```

### SetLayout

设置幻灯片布局。

```go
func (s *Slide) SetLayout(layoutName string) bool
```

**参数:**
- `layoutName`: 布局名称（如 "blank", "title", "titleAndContent" 等）

**返回:**
- 是否设置成功

### SlideSize

返回幻灯片尺寸（px 单位）。

```go
func (s *Slide) SlideSize() (cx, cy int)
```

### SlideSizeEMU

返回幻灯片尺寸（EMU 单位，高级用法）。

```go
func (s *Slide) SlideSizeEMU() (cx, cy int)
```

### Part

返回底层 SlidePart。

```go
func (s *Slide) Part() *parts.SlidePart
```

### PartURI

返回部件 URI。

```go
func (s *Slide) PartURI() *opc.PackURI
```

## 文本操作

### AddTextBox

添加文本框。

```go
func (s *Slide) AddTextBox(x, y, cx, cy int, text string) *parts.XSp
```

**参数:**
- `x, y`: 位置（px 单位）
- `cx, cy`: 尺寸（px 单位）
- `text`: 文本内容

**返回:**
- 形状对象 `*parts.XSp`

**示例:**

```go
// 添加标题
title := slide.AddTextBox(100, 50, 600, 60, "演示文稿标题")

// 添加正文
body := slide.AddTextBox(100, 150, 600, 400, "这是正文内容...")
```

## 形状操作

### AddRectangle

添加矩形。

```go
func (s *Slide) AddRectangle(x, y, cx, cy int) *parts.XSp
```

**示例:**

```go
rect := slide.AddRectangle(100, 100, 200, 150)
```

### AddEllipse

添加椭圆。

```go
func (s *Slide) AddEllipse(x, y, cx, cy int) *parts.XSp
```

**示例:**

```go
ellipse := slide.AddEllipse(100, 100, 200, 150)
```

### AddRoundRect

添加圆角矩形。

```go
func (s *Slide) AddRoundRect(x, y, cx, cy int) *parts.XSp
```

**示例:**

```go
roundRect := slide.AddRoundRect(100, 100, 200, 150)
```

### AddAutoShape

添加自动形状。

```go
func (s *Slide) AddAutoShape(x, y, cx, cy int, presetID string) *parts.XSp
```

**参数:**
- `presetID`: 预设形状类型（如 "rectangle", "ellipse", "roundRect"）

**示例:**

```go
// 添加矩形
shape1 := slide.AddAutoShape(100, 100, 200, 150, "rectangle")

// 添加椭圆
shape2 := slide.AddAutoShape(100, 100, 200, 150, "ellipse")

// 添加圆角矩形
shape3 := slide.AddAutoShape(100, 100, 200, 150, "roundRect")
```

## 图片操作

### AddPicture

添加图片。

```go
func (s *Slide) AddPicture(x, y, cx, cy int, imageRId string) *parts.XPicture
```

**参数:**
- `x, y`: 位置（px 单位）
- `cx, cy`: 尺寸（px 单位）
- `imageRId`: 图片关系 ID

**示例:**

```go
// 需要先获取 rId
rId := slide.AddImageRel("media/image1.png")
pic := slide.AddPicture(100, 100, 400, 300, rId)
```

### AddPictureFromBytes

从字节数据添加图片，自动处理媒体资源的添加和关系 ID 分配。

```go
func (s *Slide) AddPictureFromBytes(x, y, cx, cy int, fileName string, data []byte) (*parts.XPicture, error)
```

**示例:**

```go
data, _ := os.ReadFile("logo.png")
pic, err := slide.AddPictureFromBytes(100, 100, 200, 150, "logo.png", data)
if err != nil {
    panic(err)
}
```

### AddPictureFromFile

从文件添加图片。

```go
func (s *Slide) AddPictureFromFile(x, y, cx, cy int, path string) (*parts.XPicture, error)
```

**示例:**

```go
pic, err := slide.AddPictureFromFile(100, 100, 400, 300, "photo.png")
if err != nil {
    panic(err)
}
```

## 表格操作

### AddTable

添加表格。

```go
func (s *Slide) AddTable(x, y, cx, cy, rows, cols int) *parts.XGraphicFrame
```

**参数:**
- `x, y`: 位置（px 单位）
- `cx, cy`: 尺寸（px 单位）
- `rows, cols`: 行列数

**返回:**
- 图形框架对象 `*parts.XGraphicFrame`

**示例:**

```go
// 添加 3 行 4 列的表格
table := slide.AddTable(100, 100, 600, 400, 3, 4)

// 设置单元格内容
slide.SetTableCellText(table, 0, 0, "姓名")
slide.SetTableCellText(table, 0, 1, "年龄")
slide.SetTableCellText(table, 1, 0, "张三")
slide.SetTableCellText(table, 1, 1, "25")
```

### SetTableCellText

设置表格单元格文本。

```go
func (s *Slide) SetTableCellText(gf *parts.XGraphicFrame, row, col int, text string)
```

**参数:**
- `gf`: 图形框架对象
- `row, col`: 行列索引（从 0 开始）
- `text`: 文本内容

## 组件操作

### AddComponent

添加组件到幻灯片。

```go
func (s *Slide) AddComponent(c Component) error
```

**参数:**
- `c`: 实现 `Component` 接口的组件

**示例:**

```go
// 添加自定义组件
err := slide.AddComponent(&MyCustomComponent{
    Text: "Hello",
    X:    100,
    Y:    100,
})
if err != nil {
    panic(err)
}
```

### AddComponents

批量添加组件。

```go
func (s *Slide) AddComponents(components ...Component) error
```

**示例:**

```go
err := slide.AddComponents(
    &TitleComponent{Text: "标题"},
    &BodyComponent{Text: "正文"},
)
```

### NewContext

创建幻灯片上下文（用于手动组件渲染）。

```go
func (s *Slide) NewContext() *SlideContext
```

### Builder

返回幻灯片构建器。

```go
func (s *Slide) Builder() *SlideBuilder
```

## 关系管理

### AddImageRel

添加图片关系。

```go
func (s *Slide) AddImageRel(targetURI string) string
```

**返回:**
- 关系 ID (rId)

### AddMediaRel

添加媒体关系。

```go
func (s *Slide) AddMediaRel(targetURI string) string
```

### AddChartRel

添加图表关系。

```go
func (s *Slide) AddChartRel(targetURI string) string
```

### GetImageRId

获取图片 rId，不存在则添加。

```go
func (s *Slide) GetImageRId(targetURI string) string
```

### HasImage

判断是否已存在某图片关系。

```go
func (s *Slide) HasImage(targetURI string) bool
```

## 边界检查

### CheckBoundary

检查元素边界。

```go
func (s *Slide) CheckBoundary(x, y, cx, cy int) BoundaryCheckResult
```

**参数:**
- `x, y`: 元素左上角坐标 (px)
- `cx, cy`: 元素宽度和高度 (px)

**返回:**
- 边界检查结果，包含越界信息和可见性状态

**示例:**

```go
result := slide.CheckBoundary(100, 100, 200, 150)
switch result.Status {
case pptx.BoundaryStatusInside:
    fmt.Println("完全在边界内")
case pptx.BoundaryStatusPartial:
    fmt.Println("部分越界")
case pptx.BoundaryStatusOutside:
    fmt.Println("完全越界")
}
```

### IsInsideBoundary

检查元素是否完全在边界内。

```go
func (s *Slide) IsInsideBoundary(x, y, cx, cy int) bool
```

### IsVisible

检查元素是否有部分可见。

```go
func (s *Slide) IsVisible(x, y, cx, cy int) bool
```

### Viewport

返回幻灯片视口。

```go
func (s *Slide) Viewport() *SlideViewport
```

## 颜色处理

### ResolveColor

解析颜色（支持名称、十六进制、RGB、主题色）。

```go
func (s *Slide) ResolveColor(color string) Color
```

**示例:**

```go
// 解析十六进制颜色
c1 := slide.ResolveColor("#FF0000")

// 解析主题色
c2 := slide.ResolveColor("accent1")

// 解析 RGB
c3 := slide.ResolveColor("rgb(255, 0, 0)")
```

### ValidateColor

验证颜色。

```go
func (s *Slide) ValidateColor(color string) ColorValidationResult
```

---

# SlideBuilder - 幻灯片构建器

`SlideBuilder` 提供幻灯片构建功能，主要用于底层操作。

## 类型定义

```go
type SlideBuilder struct {
    // Has unexported fields.
}
```

## 构造函数

### NewSlideBuilder

创建幻灯片构建器。

```go
func NewSlideBuilder(slide *parts.SlidePart) *SlideBuilder
```

## 形状操作

### AddAutoShape

添加自动形状到幻灯片。

```go
func (b *SlideBuilder) AddAutoShape(x, y, cx, cy int, presetID string) *parts.XSp
```

**注意:** 使用 EMU 单位

### AddTextBox

添加文本框到幻灯片。

```go
func (b *SlideBuilder) AddTextBox(x, y, cx, cy int, text string) *parts.XSp
```

**注意:** 使用 EMU 单位

### AddPicture

添加图片到幻灯片。

```go
func (b *SlideBuilder) AddPicture(x, y, cx, cy int, imageRId string) *parts.XPicture
```

**注意:** 使用 EMU 单位

### AddTable

添加表格到幻灯片。

```go
func (b *SlideBuilder) AddTable(x, y, cx, cy, rows, cols int) *parts.XGraphicFrame
```

**注意:** 使用 EMU 单位

## 关系管理

### AddImage

添加图片关系并返回 rId。

```go
func (b *SlideBuilder) AddImage(targetURI string) string
```

### AddMedia

添加媒体关系并返回 rId。

```go
func (b *SlideBuilder) AddMedia(targetURI string) string
```

### AddChart

添加图表关系并返回 rId。

```go
func (b *SlideBuilder) AddChart(targetURI string) string
```

### GetImageRId

获取图片 rId，不存在则添加。

```go
func (b *SlideBuilder) GetImageRId(targetURI string) string
```

### GetChartRId

获取图表 rId，不存在则添加。

```go
func (b *SlideBuilder) GetChartRId(targetURI string) string
```

### GetMediaRId

获取媒体 rId，不存在则添加。

```go
func (b *SlideBuilder) GetMediaRId(targetURI string) string
```

### HasImage

判断是否已存在某图片关系。

```go
func (b *SlideBuilder) HasImage(targetURI string) bool
```

### HasMedia

判断是否已存在某媒体关系。

```go
func (b *SlideBuilder) HasMedia(targetURI string) bool
```

### GetRelationshipURI

根据 rId 获取目标 URI。

```go
func (b *SlideBuilder) GetRelationshipURI(rId string) string
```

## 辅助方法

### GetOrAddPicture

添加图片到幻灯片并返回 XPicture，自动处理图片关系 ID。

```go
func (b *SlideBuilder) GetOrAddPicture(x, y, cx, cy int, imageURI string) *parts.XPicture
```

### SetTableCellText

设置表格单元格文本。

```go
func (b *SlideBuilder) SetTableCellText(gf *parts.XGraphicFrame, row, col int, text string)
```

### Slide

返回底层 SlidePart。

```go
func (b *SlideBuilder) Slide() *parts.SlidePart
```

---

# SlideContext - 幻灯片渲染上下文

`SlideContext` 提供组件渲染所需的资源和能力。

## 类型定义

```go
type SlideContext struct {
    // Has unexported fields.
}
```

## 构造函数

### NewSlideContext

创建幻灯片上下文。

```go
func NewSlideContext(s *Slide) *SlideContext
```

## 形状 ID 管理

### NextShapeID

分配下一个形状 ID，返回绝对不冲突的形状 ID（线程安全）。

```go
func (ctx *SlideContext) NextShapeID() uint32
```

### CurrentShapeID

返回当前形状 ID（最后分配的）。

```go
func (ctx *SlideContext) CurrentShapeID() uint32
```

### AllocateShapeIDBatch

批量分配形状 ID。

```go
func (ctx *SlideContext) AllocateShapeIDBatch(count int) []uint32
```

### IsShapeIDAllocated

检查形状 ID 是否已分配。

```go
func (ctx *SlideContext) IsShapeIDAllocated(id uint32) bool
```

## 形状追加

### AppendShape

将形状追加到幻灯片。

```go
func (ctx *SlideContext) AppendShape(shape interface{})
```

**参数:**
- `shape`: 形状结构体（`*parts.XSp`, `*parts.XPicture`, `*parts.XGraphicFrame` 等）

**示例:**

```go
sp := &parts.XSp{
    // ... 设置形状属性
}
ctx.AppendShape(sp)
```

### AppendShapes

批量追加形状。

```go
func (ctx *SlideContext) AppendShapes(shapes ...interface{})
```

## 媒体添加

### AddMedia

添加媒体资源（图片、音频、视频）。

```go
func (ctx *SlideContext) AddMedia(data []byte, fileName string) (string, error)
```

**返回:**
- 关系 ID 和错误

### AddImage

添加图片资源（AddMedia 的别名）。

```go
func (ctx *SlideContext) AddImage(data []byte, fileName string) (string, error)
```

### AddAudio

添加音频资源。

```go
func (ctx *SlideContext) AddAudio(data []byte, fileName string) (string, error)
```

### AddVideo

添加视频资源。

```go
func (ctx *SlideContext) AddVideo(data []byte, fileName string) (string, error)
```

### AddMediaWithMIME

添加媒体资源（指定 MIME 类型）。

```go
func (ctx *SlideContext) AddMediaWithMIME(data []byte, fileName, mimeType string) (string, error)
```

## 图表添加

### AddChart

添加图表（使用模板）。

```go
func (ctx *SlideContext) AddChart(chartType parts.ChartType, data map[string]interface{}) (string, error)
```

**参数:**
- `chartType`: 图表类型
- `data`: 图表数据

**返回:**
- 关系 ID 和错误

### AddChartXML

添加图表 XML。

```go
func (ctx *SlideContext) AddChartXML(chartXML []byte) (string, error)
```

## 关系管理

### AddImageRel

添加图片关系。

```go
func (ctx *SlideContext) AddImageRel(targetURI string) string
```

### AddMediaRel

添加媒体关系。

```go
func (ctx *SlideContext) AddMediaRel(targetURI string) string
```

### AddChartRel

添加图表关系。

```go
func (ctx *SlideContext) AddChartRel(targetURI string) string
```

### HasRelationship

检查关系是否存在。

```go
func (ctx *SlideContext) HasRelationship(rID string) bool
```

## 单位转换

### PxToEMU

将像素转换为 EMU（基于 96 DPI）。

```go
func (ctx *SlideContext) PxToEMU(px int) int
```

### EMUToPx

将 EMU 转换为像素（基于 96 DPI）。

```go
func (ctx *SlideContext) EMUToPx(emu int) int
```

## 其他方法

### SlideIndex

返回幻灯片索引。

```go
func (ctx *SlideContext) SlideIndex() int
```

### SlideSize

返回幻灯片尺寸 (cx, cy in EMU)。

```go
func (ctx *SlideContext) SlideSize() (cx, cy int)
```

### SlidePart

返回底层 SlidePart（高级用法）。

```go
func (ctx *SlideContext) SlidePart() *parts.SlidePart
```

### Presentation

返回所属演示文稿（高级用法）。

```go
func (ctx *SlideContext) Presentation() *Presentation
```

### RenderComponents

批量渲染组件。

```go
func (ctx *SlideContext) RenderComponents(components ...Component) error
```
