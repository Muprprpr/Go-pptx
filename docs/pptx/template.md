# Template - 模板系统

模板系统提供模板的加载、缓存和管理功能，支持从文件系统、嵌入式资源等多种来源加载模板。

## TemplateType

模板类型。

```go
type TemplateType string
```

### 预定义模板类型

```go
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
```

## TemplateManager

模板管理器，负责模板的懒加载、缓存和克隆。

```go
type TemplateManager struct {
    // Has unexported fields.
}
```

### 构造函数

#### NewTemplateManager

创建新的模板管理器。

```go
func NewTemplateManager() *TemplateManager
```

#### NewTemplateManagerWithDir

创建带模板目录的模板管理器。

```go
func NewTemplateManagerWithDir(dir string) *TemplateManager
```

**参数:**
- `dir`: 模板文件目录路径

### 模板加载

#### LoadDefault

加载默认模板。

```go
func (tm *TemplateManager) LoadDefault() (*opc.Package, error)
```

#### LoadTemplate

加载指定模板。

```go
func (tm *TemplateManager) LoadTemplate(name TemplateType) (*opc.Package, error)
```

**说明:** 如果模板已缓存，直接返回克隆副本；否则尝试从文件系统加载。

**示例:**

```go
tm := pptx.NewTemplateManager()
pkg, err := tm.LoadTemplate(pptx.TemplateDefault)
if err != nil {
    panic(err)
}
```

### 模板注册

#### RegisterTemplate

从文件路径注册模板。

```go
func (tm *TemplateManager) RegisterTemplate(name TemplateType, path string) error
```

**示例:**

```go
tm := pptx.NewTemplateManager()
err := tm.RegisterTemplate("custom", "/path/to/custom.pptx")
if err != nil {
    panic(err)
}
```

#### RegisterTemplateFromBytes

从字节数据注册模板。

```go
func (tm *TemplateManager) RegisterTemplateFromBytes(name TemplateType, data []byte) error
```

**示例:**

```go
data, _ := os.ReadFile("custom.pptx")
err := tm.RegisterTemplateFromBytes("custom", data)
```

#### RegisterTemplateFromFS

从文件系统注册模板。

```go
func (tm *TemplateManager) RegisterTemplateFromFS(fsys fs.FS, name TemplateType, path string) error
```

**示例:**

```go
// 从嵌入的文件系统注册
//go:embed templates/*.pptx
var templateFS embed.FS

err := tm.RegisterTemplateFromFS(templateFS, "custom", "templates/custom.pptx")
```

### 配置方法

#### SetDefaultTemplate

设置默认模板。

```go
func (tm *TemplateManager) SetDefaultTemplate(name TemplateType)
```

#### SetTemplateDir

设置模板目录。

```go
func (tm *TemplateManager) SetTemplateDir(dir string)
```

### 缓存管理

#### ClearCache

清空模板缓存。

```go
func (tm *TemplateManager) ClearCache()
```

#### HasTemplate

检查模板是否已加载。

```go
func (tm *TemplateManager) HasTemplate(name TemplateType) bool
```

#### GetMasterCache

获取母版缓存。

```go
func (tm *TemplateManager) GetMasterCache() *MasterCache
```

---

## EmbeddedTemplateManager

嵌入式模板管理器，使用程序化方式创建模板。

```go
type EmbeddedTemplateManager struct {
    // Has unexported fields.
}
```

### 获取全局管理器

```go
func GetEmbeddedTemplateManager() *EmbeddedTemplateManager
```

### 方法

#### Init

初始化嵌入式模板（仅执行一次）。

```go
func (etm *EmbeddedTemplateManager) Init() error
```

#### HasTemplate

检查模板是否存在。

```go
func (etm *EmbeddedTemplateManager) HasTemplate(name TemplateType) bool
```

#### GetTemplate

获取模板（返回克隆副本）。

```go
func (etm *EmbeddedTemplateManager) GetTemplate(name TemplateType) (*opc.Package, error)
```

#### GetDefaultTemplate

获取默认模板。

```go
func (etm *EmbeddedTemplateManager) GetDefaultTemplate() (*opc.Package, error)
```

---

## TemplateBuilder

模板构建器，用于从零开始创建 PPTX 模板。

```go
type TemplateBuilder struct {
    // Has unexported fields.
}
```

### 构造函数

```go
func NewTemplateBuilder() *TemplateBuilder
```

### 方法

#### Package

返回底层 OPC 包。

```go
func (tb *TemplateBuilder) Package() *opc.Package
```

#### Build

构建模板并返回 OPC 包。

```go
func (tb *TemplateBuilder) Build() *opc.Package
```

#### BuildAndRegister

构建模板并注册到全局管理器。

```go
func (tb *TemplateBuilder) BuildAndRegister(name TemplateType) error
```

---

## 全局函数

### LoadDefaultTemplate

加载默认模板（使用全局管理器）。

```go
func LoadDefaultTemplate() (*opc.Package, error)
```

### LoadTemplate

加载指定模板（使用全局管理器）。

```go
func LoadTemplate(name TemplateType) (*opc.Package, error)
```

### RegisterTemplate

注册模板（使用全局管理器）。

```go
func RegisterTemplate(name TemplateType, path string) error
```

### RegisterTemplateFromBytes

从字节数据注册模板（使用全局管理器）。

```go
func RegisterTemplateFromBytes(name TemplateType, data []byte) error
```

### GetEmbeddedDefaultTemplate

获取嵌入式默认模板。

```go
func GetEmbeddedDefaultTemplate() (*opc.Package, error)
```

### GetEmbeddedTemplate

获取嵌入式模板（使用全局管理器）。

```go
func GetEmbeddedTemplate(name TemplateType) (*opc.Package, error)
```

### InitEmbeddedTemplates

初始化嵌入式模板。

```go
func InitEmbeddedTemplates() error
```

---

## 视口相关

### SlideViewport

幻灯片视口。

```go
type SlideViewport struct {
    // Width 视口宽度 (px)
    Width int
    // Height 视口高度 (px)
    Height int
    // Size 标准尺寸名称（可选）
    SizeName string
}
```

### 构造函数

```go
func NewSlideViewport(width, height int) *SlideViewport

func NewSlideViewportFromSize(size SlideSize) *SlideViewport
```

### 方法

#### Rect

返回视口矩形。

```go
func (vp *SlideViewport) Rect() Rect
```

#### CheckBoundary

检查元素边界。

```go
func (vp *SlideViewport) CheckBoundary(x, y, cx, cy int) BoundaryCheckResult
```

#### CheckRect

检查矩形边界。

```go
func (vp *SlideViewport) CheckRect(rect Rect) BoundaryCheckResult
```

#### IsInside

检查元素是否完全在边界内。

```go
func (vp *SlideViewport) IsInside(x, y, cx, cy int) bool
```

#### IsVisible

检查元素是否有部分可见。

```go
func (vp *SlideViewport) IsVisible(x, y, cx, cy int) bool
```

---

## 边界检查

### BoundaryStatus

边界状态。

```go
type BoundaryStatus int
```

**常量:**

```go
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
```

#### String

返回边界状态的字符串表示。

```go
func (bs BoundaryStatus) String() string
```

### BoundaryCheckResult

边界检查结果。

```go
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
```

### Rect

矩形区域。

```go
type Rect struct {
    X, Y   int // 左上角坐标 (px)
    Cx, Cy int // 宽度和高度 (px)
}
```

---

## SlideSize

幻灯片尺寸。

```go
type SlideSize struct {
    Width  int // 宽度 (px)
    Height int // 高度 (px)
}
```

### 预设尺寸

```go
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
```

---

## 使用示例

### 使用预定义模板

```go
// 使用默认模板创建演示文稿
pres, err := pptx.NewWithTemplate(pptx.TemplateDefault)
if err != nil {
    panic(err)
}

// 使用空白模板
pres, err = pptx.NewWithTemplate(pptx.TemplateBlank)

// 使用宽屏模板
pres, err = pptx.NewWithTemplate(pptx.TemplateWide)

// 使用标准模板（4:3）
pres, err = pptx.NewWithTemplate(pptx.TemplateStandard)
```

### 注册自定义模板

```go
// 从文件注册
err := pptx.RegisterTemplate("custom", "/path/to/custom.pptx")
if err != nil {
    panic(err)
}

// 从字节数据注册
data, _ := os.ReadFile("custom.pptx")
err = pptx.RegisterTemplateFromBytes("custom", data)

// 使用自定义模板
pres, err := pptx.NewWithTemplate("custom")
```

### 使用模板管理器

```go
// 创建模板管理器
tm := pptx.NewTemplateManagerWithDir("/path/to/templates")

// 注册模板
tm.RegisterTemplate("report", "report.pptx")
tm.RegisterTemplate("proposal", "proposal.pptx")

// 加载模板
pkg, err := tm.LoadTemplate("report")
if err != nil {
    panic(err)
}

// 检查模板是否存在
if tm.HasTemplate("proposal") {
    fmt.Println("模板已加载")
}

// 清空缓存
tm.ClearCache()
```

### 边界检查

```go
slide := pres.AddSlide()

// 检查元素边界
result := slide.CheckBoundary(100, 100, 200, 150)

switch result.Status {
case pptx.BoundaryStatusInside:
    fmt.Println("元素完全在边界内")
case pptx.BoundaryStatusPartial:
    fmt.Printf("元素部分越界: X=%d, Y=%d\n", result.OverflowX, result.OverflowY)
case pptx.BoundaryStatusOutside:
    fmt.Println("元素完全越界")
}

// 快速检查
if slide.IsInsideBoundary(100, 100, 200, 150) {
    fmt.Println("元素在边界内")
}

if slide.IsVisible(100, 100, 200, 150) {
    fmt.Println("元素至少部分可见")
}
```

### 使用视口

```go
// 创建视口
viewport := pptx.NewSlideViewport(1280, 720)

// 检查边界
rect := pptx.Rect{X: 100, Y: 100, Cx: 200, Cy: 150}
result := viewport.CheckRect(rect)

fmt.Printf("边界状态: %s\n", result.Status)
fmt.Printf("是否可见: %v\n", result.IsVisible)
```

### 从嵌入资源加载

```go
//go:embed templates/*.pptx
var templateFS embed.FS

func main() {
    // 初始化嵌入式模板
    err := pptx.InitEmbeddedTemplates()
    if err != nil {
        panic(err)
    }

    // 使用嵌入式模板
    pres, err := pptx.NewWithTemplate(pptx.TemplateDefault)
}
```

### 模板克隆和修改

```go
// 加载模板
pres, _ := pptx.NewWithTemplate(pptx.TemplateDefault)

// 克隆演示文稿
presCopy, err := pres.Clone()
if err != nil {
    panic(err)
}

// 修改克隆的版本
presCopy.AddSlide()

// 原始版本不受影响
fmt.Printf("原始: %d 页, 克隆: %d 页\n", pres.SlideCount(), presCopy.SlideCount())
```
