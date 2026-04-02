# Package pptx - PPTX 高级操作接口

## 概述

`pptx` 包提供 PPTX 文件的高级操作接口，作为人类开发者和 AI 调用的绝对入口。该包封装了底层 OPC 包和 parts 层的复杂性，提供简洁易用的 API。

## 核心设计理念

- **高层抽象**: 提供面向业务的 API，隐藏 OOXML 规范细节
- **组件化**: 支持可复用的组件系统，便于组合和扩展
- **类型安全**: 强类型接口，减少运行时错误
- **并发安全**: 核心结构支持并发访问

## 主要模块

### 1. [Presentation](presentation.md) - 演示文稿
演示文稿的总控门面，提供创建、保存、导入导出等核心功能。

### 2. [Slide](slide.md) - 幻灯片
幻灯片操作接口，支持添加文本、图片、表格、形状等元素。

### 3. [Component](component.md) - 组件系统
可复用的渲染组件，支持自定义组件和组合模式。

### 4. [Color](color.md) - 颜色系统
完整的颜色处理方案，支持 RGB、主题色、透明度等。

### 5. [Media](media.md) - 媒体管理
图片、视频、音频等媒体资源的管理和去重。

### 6. [Template](template.md) - 模板系统
模板加载、缓存和管理，支持嵌入式模板。

## 快速开始

### 安装

```bash
go get github.com/Muprprpr/Go-pptx/pptx
```

### 创建演示文稿

```go
package main

import (
    "github.com/Muprprpr/Go-pptx/pptx"
)

func main() {
    // 创建空白演示文稿
    pres := pptx.New()

    // 添加幻灯片
    slide := pres.AddSlide()

    // 添加文本框
    slide.AddTextBox(100, 100, 400, 50, "Hello, World!")

    // 保存文件
    pres.Save("output.pptx")
}
```

### 从模板创建

```go
// 使用默认模板
pres, err := pptx.NewWithTemplate(pptx.TemplateDefault)
if err != nil {
    panic(err)
}

// 使用空白模板
pres, err := pptx.NewWithTemplate(pptx.TemplateBlank)
```

### 添加图片

```go
slide := pres.AddSlide()

// 从文件添加图片
pic, err := slide.AddPictureFromFile(100, 100, 400, 300, "image.png")
if err != nil {
    panic(err)
}

// 从字节数据添加图片
data, _ := os.ReadFile("logo.png")
pic, err = slide.AddPictureFromBytes(100, 100, 200, 150, "logo.png", data)
```

### 添加表格

```go
slide := pres.AddSlide()

// 添加 3x4 表格
table := slide.AddTable(100, 100, 600, 400, 3, 4)

// 设置单元格文本
slide.SetTableCellText(table, 0, 0, "Header 1")
slide.SetTableCellText(table, 0, 1, "Header 2")
```

### 使用组件系统

```go
// 创建自定义组件
type TitleComponent struct {
    Text string
    X, Y int
}

func (t *TitleComponent) Render(ctx *pptx.SlideContext) error {
    // 使用 SlideContext 的能力渲染组件
    return nil
}

// 添加组件到幻灯片
slide.AddComponent(&TitleComponent{
    Text: "My Title",
    X:    100,
    Y:    50,
})
```

## 单位说明

本包主要使用两种单位：

1. **像素 (px)**: 大部分高层 API 使用像素作为单位，基于 96 DPI
2. **EMU**: 底层 OOXML 使用的单位 (1 px = 9525 EMU)

转换函数：
- `PxToEMU(px int) int` - 像素转 EMU
- `EMUToPx(emu int) int` - EMU 转像素

## 标准幻灯片尺寸

```go
// 16:9 宽屏 (1280 x 720 px)
pptx.SlideSize16x9

// 4:3 标准 (960 x 720 px)
pptx.SlideSize4x3

// 16:10 超宽屏 (1280 x 800 px)
pptx.SlideSize16x10
```

## 架构层次

```
┌─────────────────────────────────────┐
│           pptx (高层 API)            │
│  Presentation, Slide, Component     │
├─────────────────────────────────────┤
│           slide (业务构建)           │
│        SlideBuilder, Helper         │
├─────────────────────────────────────┤
│           parts (XML 结构)           │
│    SlidePart, MasterPart, etc.      │
├─────────────────────────────────────┤
│            opc (底层包)              │
│      Package, Relationships         │
└─────────────────────────────────────┘
```

## 更多文档

- [Presentation 演示文稿](presentation.md)
- [Slide 幻灯片](slide.md)
- [Component 组件系统](component.md)
- [Color 颜色系统](color.md)
- [Media 媒体管理](media.md)
- [Template 模板系统](template.md)
