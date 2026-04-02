# Color - 颜色系统

颜色系统提供完整的颜色处理方案，支持 RGB、主题色、透明度等。

## 类型定义

### Color

颜色结构体。

```go
type Color struct {
    // Type 颜色类型
    Type ColorType
    // RGB 十六进制值 (6位，如 "FF0000"，无 # 前缀)
    RGB string
    // Scheme 主题颜色名称 (如 "accent1")
    Scheme string
    // Alpha 透明度 (0-100000，OOXML 标准)
    // 100000 = 100% 不透明，0 = 完全透明
    Alpha int
    // IsValid 是否有效
    IsValid bool
}
```

### ColorType

颜色类型。

```go
type ColorType int

const (
    // ColorTypeRGB RGB 颜色
    ColorTypeRGB ColorType = iota
    // ColorTypeScheme 主题颜色
    ColorTypeScheme
    // ColorTypeInvalid 无效颜色
    ColorTypeInvalid
)
```

## 颜色常量

### Alpha 常量

Alpha 常量（OOXML 范围：0-100000）。

```go
const (
    // AlphaOpaque 完全不透明
    AlphaOpaque = 100000
    // AlphaTransparent 完全透明
    AlphaTransparent = 0
    // AlphaDefault 默认透明度（100%）
    AlphaDefault = 100000
)
```

### 主题颜色名称

```go
const (
    // 背景色
    SchemeBg1 = "bg1" // 背景色 1 (通常是白色或浅色)
    SchemeBg2 = "bg2" // 背景色 2
    SchemeFg1 = "fg1" // 前景色/文本色 1 (通常是黑色或深色)
    SchemeFg2 = "fg2" // 前景色/文本色 2

    // 强调色
    SchemeAccent1 = "accent1" // 强调色 1
    SchemeAccent2 = "accent2" // 强调色 2
    SchemeAccent3 = "accent3" // 强调色 3
    SchemeAccent4 = "accent4" // 强调色 4
    SchemeAccent5 = "accent5" // 强调色 5
    SchemeAccent6 = "accent6" // 强调色 6

    // 超链接色
    SchemeHlink    = "hlink"    // 超链接颜色
    SchemeFolHlink = "folHlink" // 已访问超链接颜色

    // 特殊色
    SchemePhClr = "phClr" // 幻灯片标题颜色
    SchemeTx1   = "tx1"   // 文本色 1
    SchemeTx2   = "tx2"   // 文本色 2
)
```

### 预设颜色常量

预设颜色常量（6位十六进制，无 # 前缀）。

```go
var (
    // 基础颜色
    ColorBlack   = RGBColor("000000")
    ColorWhite   = RGBColor("FFFFFF")
    ColorRed     = RGBColor("FF0000")
    ColorGreen   = RGBColor("00FF00")
    ColorBlue    = RGBColor("0000FF")
    ColorYellow  = RGBColor("FFFF00")
    ColorCyan    = RGBColor("00FFFF")
    ColorMagenta = RGBColor("FF00FF")

    // 透明色
    ColorTransparent = Color{Type: ColorTypeRGB, RGB: "000000", Alpha: AlphaTransparent, IsValid: true}

    // 常用 UI 颜色
    ColorGray      = RGBColor("808080")
    ColorLightGray = RGBColor("C0C0C0")
    ColorDarkGray  = RGBColor("404040")
    ColorOrange    = RGBColor("FFA500")
    ColorPurple    = RGBColor("800080")
    ColorPink      = RGBColor("FFC0CB")
    ColorBrown     = RGBColor("A52A2A")
    ColorNavy      = RGBColor("000080")
    ColorTeal      = RGBColor("008080")
    ColorOlive     = RGBColor("808000")
    ColorMaroon    = RGBColor("800000")
    ColorLime      = RGBColor("00FF00")
    ColorAqua      = RGBColor("00FFFF")
    ColorSilver    = RGBColor("C0C0C0")
    ColorGold      = RGBColor("FFD700")
)
```

## 构造函数

### RGBColor

创建 RGB 颜色（无 alpha）。

```go
func RGBColor(hex string) Color
```

**参数:**
- `hex`: 6位十六进制 RGB（无 # 前缀）

**示例:**

```go
red := pptx.RGBColor("FF0000")
blue := pptx.RGBColor("0000FF")
```

### RGBAColor

创建带 alpha 的 RGB 颜色。

```go
func RGBAColor(hex string, alpha int) Color
```

**参数:**
- `hex`: 6位十六进制 RGB
- `alpha`: 0-100000 (OOXML 标准)

**示例:**

```go
// 半透明红色
semiRed := pptx.RGBAColor("FF0000", 50000) // 50% 透明
```

### SchemeColor

创建主题颜色。

```go
func SchemeColor(name string) Color
```

**示例:**

```go
// 使用主题强调色
accent1 := pptx.SchemeColor(pptx.SchemeAccent1)
bgColor := pptx.SchemeColor(pptx.SchemeBg1)
```

### ColorFromRGB

从 RGB 值创建颜色。

```go
func ColorFromRGB(r, g, b int) Color
```

**参数:**
- `r, g, b`: 0-255

**示例:**

```go
red := pptx.ColorFromRGB(255, 0, 0)
```

### ColorFromRGBA

从 RGBA 值创建颜色。

```go
func ColorFromRGBA(r, g, b, alpha int) Color
```

**参数:**
- `r, g, b`: 0-255
- `alpha`: 0-100000 (OOXML 标准)

**示例:**

```go
semiRed := pptx.ColorFromRGBA(255, 0, 0, 50000) // 50% 透明
```

### ParseColor

解析颜色字符串，支持多种格式。

```go
func ParseColor(s string) Color
```

**支持的格式:**
- `#FF0000` 或 `FF0000` (6位十六进制)
- `#FF0000FF` (8位十六进制，最后两位为 alpha)
- `rgb(255, 0, 0)` (RGB)
- `rgba(255, 0, 0, 0.5)` (RGBA，alpha 0-1)
- `accent1`, `bg1` 等主题颜色名称

**示例:**

```go
// 十六进制
c1 := pptx.ParseColor("#FF0000")
c2 := pptx.ParseColor("FF0000")

// RGB
c3 := pptx.ParseColor("rgb(255, 0, 0)")

// RGBA
c4 := pptx.ParseColor("rgba(255, 0, 0, 0.5)")

// 主题色
c5 := pptx.ParseColor("accent1")
```

## Color 方法

### WithAlpha

设置透明度并返回新颜色。

```go
func (c Color) WithAlpha(alpha int) Color
```

**参数:**
- `alpha`: 0-100000 (OOXML 标准)

**示例:**

```go
red := pptx.RGBColor("FF0000")
semiRed := red.WithAlpha(50000) // 50% 透明
```

### WithAlphaPercent

设置透明度（百分比）并返回新颜色。

```go
func (c Color) WithAlphaPercent(percent float64) Color
```

**参数:**
- `percent`: 0-100

**示例:**

```go
red := pptx.RGBColor("FF0000")
semiRed := red.WithAlphaPercent(50) // 50% 透明
```

### AlphaPercent

返回百分比形式的 alpha (0-100)。

```go
func (c Color) AlphaPercent() float64
```

### AlphaValue

返回 OOXML 格式的 alpha 值 (0-100000)。

```go
func (c Color) AlphaValue() int
```

**用途:** 用于 `<a:alpha val="50000"/>`

### ToRGB

转换为 RGB 十六进制字符串（6位，无 # 前缀）。

```go
func (c Color) ToRGB() string
```

**用途:** 符合 OOXML 规范：`<a:srgbClr val="FF0000"/>`

### ToHex

转换为带 # 前缀的 6 位十六进制字符串（用于显示）。

```go
func (c Color) ToHex() string
```

**示例:**

```go
red := pptx.RGBColor("FF0000")
fmt.Println(red.ToHex()) // 输出: #FF0000
```

### ToHexA

转换为带 # 前缀的 8 位十六进制字符串（包含 alpha）。

```go
func (c Color) ToHexA() string
```

### ToScheme

转换为主题颜色名称。

```go
func (c Color) ToScheme() string
```

### RGBComponents

返回 RGB 分量 (r, g, b)。

```go
func (c Color) RGBComponents() (r, g, b int, ok bool)
```

**示例:**

```go
red := pptx.RGBColor("FF0000")
r, g, b, ok := red.RGBComponents()
// r=255, g=0, b=0, ok=true
```

### String

返回颜色的字符串表示。

```go
func (c Color) String() string
```

## 辅助函数

### IsSchemeColor

检查是否为有效的主题颜色名称。

```go
func IsSchemeColor(name string) bool
```

**示例:**

```go
pptx.IsSchemeColor("accent1") // true
pptx.IsSchemeColor("FF0000")  // false
```

### ValidateColor

验证颜色。

```go
func ValidateColor(s string) ColorValidationResult
```

**返回:**
- `ColorValidationResult` 验证结果

**示例:**

```go
result := pptx.ValidateColor("#FF0000")
if result.IsValid {
    fmt.Println("有效颜色:", result.Color.ToHex())
} else {
    fmt.Println("无效颜色:", result.Message)
}
```

## ColorValidationResult

颜色验证结果。

```go
type ColorValidationResult struct {
    // IsValid 是否有效
    IsValid bool
    // Color 解析后的颜色
    Color Color
    // Original 原始输入
    Original string
    // Message 验证消息
    Message string
}
```

## ColorMap - 颜色映射表

颜色映射表用于将颜色名称映射到实际颜色值。

```go
type ColorMap struct {
    // Has unexported fields.
}
```

### 构造函数

```go
// NewColorMap 创建颜色映射表
func NewColorMap() *ColorMap

// DefaultColorMap 默认颜色映射表
func DefaultColorMap() *ColorMap
```

### 方法

```go
// Set 设置颜色映射
func (cm *ColorMap) Set(name string, color Color)

// Get 获取颜色映射
func (cm *ColorMap) Get(name string) (Color, bool)

// All 返回所有颜色映射
func (cm *ColorMap) All() map[string]Color

// Resolve 解析颜色（支持名称、十六进制、RGB、主题色）
func (cm *ColorMap) Resolve(s string) Color
```

**示例:**

```go
cm := pptx.NewColorMap()

// 设置自定义颜色
cm.Set("primary", pptx.RGBColor("007AFF"))
cm.Set("secondary", pptx.RGBColor("5856D6"))

// 获取颜色
if primary, ok := cm.Get("primary"); ok {
    fmt.Println("主色:", primary.ToHex())
}

// 解析颜色（支持多种格式）
c1 := cm.Resolve("primary")      // 自定义名称
c2 := cm.Resolve("#FF0000")      // 十六进制
c3 := cm.Resolve("accent1")      // 主题色
```

## 使用示例

### 基础颜色使用

```go
// 创建演示文稿
pres := pptx.New()
slide := pres.AddSlide()

// 使用预设颜色
red := pptx.ColorRed
blue := pptx.ColorBlue
green := pptx.ColorGreen

// 使用自定义颜色
custom := pptx.RGBColor("007AFF")
```

### 透明度处理

```go
// 创建半透明颜色
semiRed := pptx.RGBColor("FF0000").WithAlpha(50000)     // 50% 透明
semiBlue := pptx.ColorBlue.WithAlphaPercent(30)         // 30% 透明

// 完全透明
transparent := pptx.ColorTransparent
```

### 主题颜色

```go
// 使用主题颜色
accent1 := pptx.SchemeColor(pptx.SchemeAccent1)
accent2 := pptx.SchemeColor(pptx.SchemeAccent2)
bgColor := pptx.SchemeColor(pptx.SchemeBg1)

// 主题颜色也支持透明度
semiAccent1 := accent1.WithAlphaPercent(50)
```

### 颜色解析

```go
// 从字符串解析颜色
colors := []string{
    "#FF0000",
    "rgb(0, 255, 0)",
    "rgba(0, 0, 255, 0.5)",
    "accent1",
    "bg1",
}

for _, s := range colors {
    c := pptx.ParseColor(s)
    fmt.Printf("%s -> %s (alpha: %.0f%%)\n",
        s, c.ToHex(), c.AlphaPercent())
}
```

### 颜色映射表使用

```go
// 创建品牌颜色映射
brandColors := pptx.NewColorMap()
brandColors.Set("primary", pptx.RGBColor("007AFF"))
brandColors.Set("secondary", pptx.RGBColor("5856D6"))
brandColors.Set("success", pptx.RGBColor("34C759"))
brandColors.Set("warning", pptx.RGBColor("FF9500"))
brandColors.Set("danger", pptx.RGBColor("FF3B30"))

// 使用品牌颜色
primary := brandColors.Resolve("primary")
```

### 在形状中使用颜色

```go
// 添加带颜色的形状
rect := slide.AddRectangle(100, 100, 200, 150)

// 设置填充颜色
rect.SpPr.SolidFill = &parts.XSolidFill{
    SrgbClr: &parts.XSrgbClr{
        Val: pptx.ColorRed.ToRGB(),
    },
}

// 设置边框颜色
rect.SpPr.Ln = &parts.XLn{
    SolidFill: &parts.XSolidFill{
        SrgbClr: &parts.XSrgbClr{
            Val: pptx.ColorBlack.ToRGB(),
        },
    },
}
```

### 在文本中使用颜色

```go
// 添加带颜色的文本
textBox := slide.AddTextBox(100, 100, 400, 50, "Hello World")

// 设置文本颜色
textBox.TxBody.P[0].R[0].RPr.SolidFill = &parts.XSolidFill{
    SrgbClr: &parts.XSrgbClr{
        Val: pptx.ColorBlue.ToRGB(),
    },
}
```

## 单位转换

### PxToEMU / EMUToPx

像素与 EMU 的转换。

```go
// PxToEMU 将像素转换为 EMU（基于 96 DPI）
func PxToEMU(px int) int

// EMUToPx 将 EMU 转换为像素（基于 96 DPI）
func EMUToPx(emu int) int
```

**示例:**

```go
emu := pptx.PxToEMU(100)  // 914400 / 96 * 100 = 952500
px := pptx.EMUToPx(952500) // 100
```

### EMUsPerPixel 常量

每像素对应的 EMU 数量 (96 DPI)。

```go
const (
    // EMUsPerPixel 每像素对应的 EMU 数量 (96 DPI)
    // 1 英寸 = 914400 EMU
    // 1 英寸 = 96 像素 (96 DPI)
    // 因此 1 px = 914400 / 96 = 9525 EMU
    EMUsPerPixel = 914400 / 96 // = 9525
)
```
