# Go-PPTX

[English](#english) | [中文](#中文)

---

<a name="english"></a>

## English

A Go library for creating, reading, and modifying PowerPoint (PPTX) files with streaming support for large files.

### Features

- **Full PPTX Support**: Create, read, and modify PPTX files
- **Streaming I/O**: Handle large files efficiently with lazy loading
- **OPC Implementation**: Complete Open Packaging Convention implementation
- **Thread Safe**: Safe for concurrent use
- **Zero Dependencies**: Only uses Go standard library

### Installation

```bash
go get github.com/Muprprpr/Go-pptx
```

### Quick Start

#### Traditional Usage (Small Files)

```go
package main

import (
    "github.com/Muprprpr/Go-pptx/opc"
)

func main() {
    // Open existing file
    pkg, err := opc.OpenFile("presentation.pptx")
    if err != nil {
        panic(err)
    }
    defer pkg.Close()

    // Access parts
    slides := pkg.GetPartsByType(opc.ContentTypeSlide)

    // Save changes
    pkg.SaveFile("output.pptx")
}
```

#### Streaming Usage (Large Files)

```go
package main

import (
    "github.com/Muprprpr/Go-pptx/opc"
)

func main() {
    // Open with lazy loading - only metadata is loaded
    pkg, err := opc.OpenStream("large.presentation.pptx")
    if err != nil {
        panic(err)
    }
    defer pkg.Close()

    // Get a part - content not loaded yet
    slide := pkg.GetPart(slideURI)

    // Load only when needed
    if needsModification {
        blob, _ := slide.Blob()  // Now loaded
        // ... modify blob
        slide.SetBlob(modifiedBlob)
    }

    // Stream save - no buffering of complete XML
    pkg.StreamSaveFile("output.pptx")
}
```

### When to Use Which Mode

| Scenario | Recommended Mode |
|----------|-----------------|
| File size < 10MB | Traditional |
| File size > 50MB | Streaming |
| Only reading metadata | Streaming |
| Modifying many parts | Traditional |
| Modifying few parts | Streaming |
| Random access to all content | Traditional |

### Documentation

- [Streaming Design](docs/streaming-design.md) - Detailed streaming architecture
- [OPC Package API](opc/package.go) - Traditional package API
- [Stream Package API](opc/streampkg.go) - Streaming package API

### Project Structure

```
go-pptx/
├── opc/                    # Open Packaging Convention implementation
│   ├── constants.go        # Content types and relationship types
│   ├── packuri.go          # Pack URI handling
│   ├── part.go             # Part and PartCollection
│   ├── package.go          # Traditional Package
│   ├── contenttypes.go     # [Content_Types].xml
│   ├── coreprops.go        # Core properties
│   ├── relation.go         # Relationships
│   ├── stream.go           # Streaming types
│   └── streampkg.go        # Streaming Package
├── test/
│   └── utils/              # Test utilities and examples
└── docs/
    └── streaming-design.md # Streaming design documentation
```

### License

MIT License

---

<a name="中文"></a>

## 中文

一个用于创建、读取和修改 PowerPoint (PPTX) 文件的 Go 库，支持大文件的流式处理。

### 特性

- **完整 PPTX 支持**：创建、读取和修改 PPTX 文件
- **流式 I/O**：通过懒加载高效处理大文件
- **OPC 实现**：完整的 Open Packaging Convention 实现
- **线程安全**：支持并发使用
- **零依赖**：只使用 Go 标准库

### 安装

```bash
go get github.com/Muprprpr/Go-pptx
```

### 快速开始

#### 传统用法（小文件）

```go
package main

import (
    "github.com/Muprprpr/Go-pptx/opc"
)

func main() {
    // 打开现有文件
    pkg, err := opc.OpenFile("presentation.pptx")
    if err != nil {
        panic(err)
    }
    defer pkg.Close()

    // 访问部件
    slides := pkg.GetPartsByType(opc.ContentTypeSlide)

    // 保存更改
    pkg.SaveFile("output.pptx")
}
```

#### 流式用法（大文件）

```go
package main

import (
    "github.com/Muprprpr/Go-pptx/opc"
)

func main() {
    // 懒加载打开 - 只加载元数据
    pkg, err := opc.OpenStream("large.presentation.pptx")
    if err != nil {
        panic(err)
    }
    defer pkg.Close()

    // 获取部件 - 内容尚未加载
    slide := pkg.GetPart(slideURI)

    // 只在需要时加载
    if needsModification {
        blob, _ := slide.Blob()  // 现在已加载
        // ... 修改 blob
        slide.SetBlob(modifiedBlob)
    }

    // 流式保存 - 不缓冲完整 XML
    pkg.StreamSaveFile("output.pptx")
}
```

### 何时使用哪种模式

| 场景 | 推荐模式 |
|------|---------|
| 文件大小 < 10MB | 传统 |
| 文件大小 > 50MB | 流式 |
| 只读取元数据 | 流式 |
| 修改大量部件 | 传统 |
| 修改少量部件 | 流式 |
| 随机访问所有内容 | 传统 |

### 文档

- [流式设计](docs/streaming-design.md) - 详细的流式架构说明
- [OPC 包 API](opc/package.go) - 传统包 API
- [流式包 API](opc/streampkg.go) - 流式包 API

### 项目结构

```
go-pptx/
├── opc/                    # Open Packaging Convention 实现
│   ├── constants.go        # 内容类型和关系类型
│   ├── packuri.go          # Pack URI 处理
│   ├── part.go             # Part 和 PartCollection
│   ├── package.go          # 传统 Package
│   ├── contenttypes.go     # [Content_Types].xml
│   ├── coreprops.go        # 核心属性
│   ├── relation.go         # 关系
│   ├── stream.go           # 流式类型
│   └── streampkg.go        # 流式 Package
├── test/
│   └── utils/              # 测试工具和示例
└── docs/
    └── streaming-design.md # 流式设计文档
```

### 许可证

MIT License
