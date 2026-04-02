# Presentation - 演示文稿

`Presentation` 是 PPTX 演示文稿的总控门面，提供创建、保存、导入导出等核心功能。

## 类型定义

```go
type Presentation struct {
    // Has unexported fields.
}
```

## 构造函数

### New

创建空白演示文稿，使用默认模板（16:9 宽屏）。

```go
func New() *Presentation
```

**示例:**

```go
pres := pptx.New()
```

### NewFromFile

从文件创建演示文稿。

```go
func NewFromFile(path string) (*Presentation, error)
```

**示例:**

```go
pres, err := pptx.NewFromFile("template.pptx")
if err != nil {
    panic(err)
}
```

### NewFromBytes

从字节数据创建演示文稿。

```go
func NewFromBytes(data []byte) (*Presentation, error)
```

**示例:**

```go
data, _ := os.ReadFile("template.pptx")
pres, err := pptx.NewFromBytes(data)
if err != nil {
    panic(err)
}
```

### NewWithTemplate

从模板创建演示文稿。

```go
func NewWithTemplate(name TemplateType) (*Presentation, error)
```

**参数:**
- `name`: 模板名称（如 `TemplateDefault`, `TemplateBlank` 等）

**示例:**

```go
// 使用默认模板
pres, err := pptx.NewWithTemplate(pptx.TemplateDefault)

// 使用空白模板
pres, err := pptx.NewWithTemplate(pptx.TemplateBlank)

// 使用宽屏模板
pres, err := pptx.NewWithTemplate(pptx.TemplateWide)

// 使用标准模板（4:3）
pres, err := pptx.NewWithTemplate(pptx.TemplateStandard)
```

## 幻灯片管理

### AddSlide

添加新幻灯片。

```go
func (p *Presentation) AddSlide(layout ...string) *Slide
```

**参数:**
- `layout`: 可选的布局名称（如 "title", "blank", "titleAndContent" 等），不指定则使用空白布局

**返回:**
- 新创建的 `*Slide` 对象

**示例:**

```go
// 添加空白幻灯片
slide := pres.AddSlide()

// 添加标题幻灯片
slide := pres.AddSlide("title")

// 添加标题和内容幻灯片
slide := pres.AddSlide("titleAndContent")
```

### AddSlideAt

在指定位置插入幻灯片。

```go
func (p *Presentation) AddSlideAt(index int, layout ...string) (*Slide, error)
```

**参数:**
- `index`: 插入位置（0 为起始位置）
- `layout`: 可选的布局名称

**示例:**

```go
// 在第二页位置插入幻灯片
slide, err := pres.AddSlideAt(1)
if err != nil {
    panic(err)
}
```

### GetSlide

获取指定索引的幻灯片。

```go
func (p *Presentation) GetSlide(index int) (*Slide, error)
```

**示例:**

```go
slide, err := pres.GetSlide(0) // 获取第一张幻灯片
if err != nil {
    panic(err)
}
```

### RemoveSlide

移除指定索引的幻灯片。

```go
func (p *Presentation) RemoveSlide(index int) error
```

**示例:**

```go
err := pres.RemoveSlide(0) // 删除第一张幻灯片
if err != nil {
    panic(err)
}
```

### Slides

返回所有幻灯片。

```go
func (p *Presentation) Slides() []*Slide
```

### SlideCount

返回幻灯片数量。

```go
func (p *Presentation) SlideCount() int
```

**示例:**

```go
count := pres.SlideCount()
fmt.Printf("共有 %d 张幻灯片\n", count)
```

## 尺寸设置

### SlideSize

返回当前幻灯片尺寸。

```go
func (p *Presentation) SlideSize() (int, int)
```

**返回:**
- 宽度和高度（px 单位）

### SetSlideSize

设置幻灯片尺寸。

```go
func (p *Presentation) SetSlideSize(cx, cy int)
```

**参数:**
- `cx`: 宽度（px 单位）
- `cy`: 高度（px 单位）

**示例:**

```go
pres.SetSlideSize(1280, 720) // 设置为 16:9 宽屏
```

### SetSlideSizeStandard

设置标准幻灯片尺寸。

```go
func (p *Presentation) SetSlideSizeStandard(name string)
```

**示例:**

```go
pres.SetSlideSizeStandard("16:9")
pres.SetSlideSizeStandard("4:3")
```

## 输出方法

### Save

将演示文稿保存到文件。

```go
func (p *Presentation) Save(path string) error
```

**示例:**

```go
err := pres.Save("output.pptx")
if err != nil {
    panic(err)
}
```

### Write

将演示文稿写入 `io.Writer`。

```go
func (p *Presentation) Write(w io.Writer) error
```

**用途:** 适用于高并发流式输出（如 HTTP 响应）

**示例:**

```go
// HTTP 响应示例
http.HandleFunc("/download", func(w http.ResponseWriter, r *http.Request) {
    pres := pptx.New()
    // ... 构建演示文稿

    w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation")
    w.Header().Set("Content-Disposition", "attachment; filename=output.pptx")
    pres.Write(w)
})
```

### WriteToBytes

将演示文稿写入字节数组。

```go
func (p *Presentation) WriteToBytes() ([]byte, error)
```

**示例:**

```go
data, err := pres.WriteToBytes()
if err != nil {
    panic(err)
}
// 可以保存到数据库或发送到网络
```

## 其他方法

### Clone

克隆演示文稿，返回完全独立的副本。

```go
func (p *Presentation) Clone() (*Presentation, error)
```

**用途:** 创建演示文稿的副本进行修改，不影响原始对象

**示例:**

```go
presCopy, err := pres.Clone()
if err != nil {
    panic(err)
}
presCopy.AddSlide() // 修改副本，不影响原始 pres
```

### Close

关闭演示文稿，释放资源。

```go
func (p *Presentation) Close() error
```

### Package

返回底层 OPC 包（高级用法）。

```go
func (p *Presentation) Package() *opc.Package
```

### PresentationPart

返回演示文稿部件。

```go
func (p *Presentation) PresentationPart() *parts.PresentationPart
```

### MasterCache

返回母版缓存。

```go
func (p *Presentation) MasterCache() *MasterCache
```

### MediaManager

返回媒体管理器。

```go
func (p *Presentation) MediaManager() *MediaManager
```

## 完整示例

```go
package main

import (
    "github.com/Muprprpr/Go-pptx/pptx"
)

func main() {
    // 创建演示文稿
    pres := pptx.New()

    // 设置幻灯片尺寸
    pres.SetSlideSizeStandard("16:9")

    // 添加标题幻灯片
    titleSlide := pres.AddSlide("title")
    titleSlide.AddTextBox(100, 100, 400, 50, "演示文稿标题")

    // 添加内容幻灯片
    contentSlide := pres.AddSlide("titleAndContent")
    contentSlide.AddTextBox(100, 100, 400, 50, "第一章")
    contentSlide.AddTextBox(100, 200, 600, 300, "这里是正文内容...")

    // 添加带图片的幻灯片
    imageSlide := pres.AddSlide()
    imageSlide.AddPictureFromFile(100, 100, 400, 300, "photo.png")

    // 保存文件
    err := pres.Save("demo.pptx")
    if err != nil {
        panic(err)
    }
}
```
