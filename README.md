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
- **Parts Layer**: High-level PPTX content handling
  - Presentation, Slide, Master, Layout, Media parts
  - XML namespace handling for OOXML compatibility
  - Shape ID allocation and relationship management
- **Thread Safe**: Safe for concurrent use
  - `sync/atomic` for relationship ID allocation
  - `sync.RWMutex` for thread-safe operations
  - `sync.Map` for resource deduplication
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

### Thread-Safe Relationship ID Allocation

`Relationships` uses `sync/atomic.Int32` for thread-safe relationship ID allocation when multiple goroutines call `AddRelationship()` concurrently.

```go
// Automatic atomic ID allocation
rels := opc.NewRelationships(sourceURI)
rel1, _ := rels.AddNew(opc.RelTypeSlide, "/ppt/slides/slide1.xml", false)  // rId1
rel2, _ := rels.AddNew(opc.RelTypeSlide, "/ppt/slides/slide2.xml", false)  // rId2

// Preview next ID without consuming
nextID := rels.NextRID()  // "rId3"

// Thread-safe for concurrent calls
var wg sync.WaitGroup
for i := 0; i < 10; i++ {
    wg.Add(1)
    go func() {
        defer wg.Done()
        rels.AddNew(opc.RelTypeSlide, "/ppt/slides/slide.xml", false)
    }()
}
wg.Wait()
// All IDs are unique, no duplicates
```

**Key features:**
- `AddNew()` uses atomic operations, safe for concurrent calls
- Counter auto-initializes from existing relationships when loading from XML
- `NextRID()` previews the next ID without consuming

### Concurrent Streaming (Advanced)

For high-performance scenarios, the library provides concurrent streaming capabilities:

| Feature | Description |
|---------|-------------|
| `PartDataChannel` | Channel-based concurrent part writing |
| `ResourceDedupPool` | `sync.Map` based image/media deduplication |
| `ConcurrentZipCollector` | Goroutine-based ZIP writing |
| `ConcurrentStreamSave` | Worker-based concurrent save |

See [Streaming Design](docs/streaming-design.md) for detailed documentation.

### Documentation

- [Streaming Design](docs/streaming-design.md) - Detailed streaming architecture
- [OPC Package API](docs/opc/README.md) - OPC API reference (go doc)
- [OPC Relationship Resolution](docs/opc/relationship_resolution.md) - How relative paths are resolved
- [Presentation Part](docs/parts/presentation.md) - Presentation.xml handling
- [Slide Part](docs/parts/slide.md) - Slide content and shapes
- [Master Part](docs/parts/master.md) - Slide master and layouts
- [Media Manager](docs/parts/media.md) - Media handling and deduplication
- [XML Utilities](docs/parts/xmlutils.md) - XML namespace handling for OOXML

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
├── parts/                  # PPTX content parts
│   ├── presentation.go     # Presentation part (presentation.xml)
│   ├── slide.go            # Slide part (slideN.xml)
│   ├── slide_types.go      # Slide XML structures
│   ├── master.go           # Slide master
│   ├── master_types.go     # Master XML structures
│   ├── master_cache.go     # Thread-safe master cache
│   ├── media.go            # Media handling
│   ├── media_manager.go    # Media manager with deduplication
│   ├── coreprops.go        # Core properties part
│   ├── relationship.go     # XML relationship structures
│   └── xmlutils.go         # XML namespace utilities
├── test/
│   ├── parts/              # Parts tests
│   ├── utils/              # Test utilities and examples
│   └── pipeline_test.go    # Integration tests
└── docs/
    ├── streaming-design.md # Streaming design documentation
    ├── opc/                # OPC documentation
    │   ├── README.md       # OPC API reference
    │   └── relationship_resolution.md
    └── parts/              # Parts documentation
        ├── presentation.md
        ├── slide.md
        ├── master.md
        ├── media.md
        ├── relationship.md
        └── xmlutils.md
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
- **部件层**：高层 PPTX 内容处理
  - Presentation、Slide、Master、Layout、Media 部件
  - OOXML 兼容的 XML 命名空间处理
  - Shape ID 分配和关系管理
- **线程安全**：支持并发使用
  - `sync/atomic` 用于关系 ID 分配
  - `sync.RWMutex` 用于线程安全操作
  - `sync.Map` 用于资源去重
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

### 线程安全的关系 ID 分配

`Relationships` 使用 `sync/atomic.Int32` 实现线程安全的关系 ID 分配，确保多个 Goroutine 并发调用 `AddRelationship()` 时不会产生冲突。

```go
// 自动原子 ID 分配
rels := opc.NewRelationships(sourceURI)
rel1, _ := rels.AddNew(opc.RelTypeSlide, "/ppt/slides/slide1.xml", false)  // rId1
rel2, _ := rels.AddNew(opc.RelTypeSlide, "/ppt/slides/slide2.xml", false)  // rId2

// 预览下一个 ID（不消耗）
nextID := rels.NextRID()  // "rId3"

// 并发调用线程安全
var wg sync.WaitGroup
for i := 0; i < 10; i++ {
    wg.Add(1)
    go func() {
        defer wg.Done()
        rels.AddNew(opc.RelTypeSlide, "/ppt/slides/slide.xml", false)
    }()
}
wg.Wait()
// 所有 ID 都是唯一的，无重复
```

**核心特性：**
- `AddNew()` 使用原子操作，并发调用安全
- 从 XML 加载时自动初始化计数器为现有最大 ID
- `NextRID()` 预览下一个 ID 而不消耗

### 并发流式处理（高级）

对于高性能场景，库提供了并发流式处理能力：

| 功能 | 描述 |
|------|------|
| `PartDataChannel` | 基于 channel 的并发部件写入 |
| `ResourceDedupPool` | 基于 `sync.Map` 的图片/媒体去重 |
| `ConcurrentZipCollector` | 基于 Goroutine 的 ZIP 写入器 |
| `ConcurrentStreamSave` | 基于 Worker 的并发保存 |

详细文档请参阅[流式设计](docs/streaming-design.md)。

### 文档

- [流式设计](docs/streaming-design.md) - 详细的流式架构说明
- [OPC 包 API](docs/opc/README.md) - OPC API 参考 (go doc)
- [OPC 关系解析](docs/opc/relationship_resolution.md) - 相对路径如何解析
- [Presentation 部件](docs/parts/presentation.md) - presentation.xml 处理
- [Slide 部件](docs/parts/slide.md) - 幻灯片内容和形状
- [Master 部件](docs/parts/master.md) - 母版和版式
- [媒体管理器](docs/parts/media.md) - 媒体处理和去重
- [XML 工具](docs/parts/xmlutils.md) - OOXML 命名空间处理

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
├── parts/                  # PPTX 内容部件
│   ├── presentation.go     # 演示文稿部件 (presentation.xml)
│   ├── slide.go            # 幻灯片部件 (slideN.xml)
│   ├── slide_types.go      # 幻灯片 XML 结构
│   ├── master.go           # 母版
│   ├── master_types.go     # 母版 XML 结构
│   ├── master_cache.go     # 线程安全的母版缓存
│   ├── media.go            # 媒体处理
│   ├── media_manager.go    # 带去重功能的媒体管理器
│   ├── coreprops.go        # 核心属性部件
│   ├── relationship.go     # XML 关系结构
│   └── xmlutils.go         # XML 命名空间工具
├── test/
│   ├── parts/              # Parts 测试
│   ├── utils/              # 测试工具和示例
│   └── pipeline_test.go    # 集成测试
└── docs/
    ├── streaming-design.md # 流式设计文档
    ├── opc/                # OPC 文档
    │   ├── README.md       # OPC API 参考
    │   └── relationship_resolution.md
    └── parts/              # Parts 文档
        ├── presentation.md
        ├── slide.md
        ├── master.md
        ├── media.md
        ├── relationship.md
        └── xmlutils.md
```

### 许可证

MIT License
