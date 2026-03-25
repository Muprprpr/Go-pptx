# OPC Streaming Design / OPC 流式处理设计

[English](#english) | [中文](#中文)

---

<a name="english"></a>

## English

### Design Philosophy

#### Core Principles

1. **Lazy Loading**: Load data only when needed, not upfront
2. **Streaming I/O**: Process data in streams, not in memory buffers
3. **Zero-Copy When Possible**: Avoid unnecessary data copying
4. **Backward Compatibility**: Existing APIs continue to work

#### Memory Efficiency Goals

| Scenario | Traditional Approach | Streaming Approach |
|----------|---------------------|-------------------|
| Open 100MB PPTX | Load 100MB into memory | Load only metadata (~1MB) |
| Modify one slide | Keep all parts in memory | Load only modified parts |
| Save modified file | Build complete XML in memory | Stream XML directly to ZIP |

### Architecture Overview

```
┌────────────────────────────────────────────────────────────────┐
│                        StreamPackage                            │
│                                                                 │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────────────────┐ │
│  │ ContentTypes│  │  Rels       │  │      Parts              │ │
│  │ (loaded)    │  │ (loaded)    │  │  ┌─────────────────┐   │ │
│  │             │  │             │  │  │ StreamPart      │   │ │
│  │             │  │             │  │  │ ┌─────────────┐ │   │ │
│  │             │  │             │  │  │ │ PartSource  │ │   │ │
│  │             │  │             │  │  │ │ (not loaded)│ │   │ │
│  │             │  │             │  │  │ └─────────────┘ │   │ │
│  │             │  │             │  │  └─────────────────┘   │ │
│  └─────────────┘  └─────────────┘  └─────────────────────────┘ │
│                                                                 │
└────────────────────────────────────────────────────────────────┘
```

### Key Components

#### 1. PartSource Interface

`PartSource` is an abstraction for part data sources, enabling lazy loading from various sources.

```go
type PartSource interface {
    Open() (io.ReadCloser, error)  // Open stream to read data
    Size() int64                    // Return data size (or -1 if unknown)
}
```

**Implementations:**
- `ZipFileSource`: Data from ZIP file entry (lazy read)
- `BytesSource`: Data from memory ([]byte)
- `ReaderSource`: Data from io.Reader

#### 2. StreamPart

`StreamPart` is a part that supports lazy loading. It only loads content into memory when explicitly requested.

```go
type StreamPart struct {
    uri           *PackURI
    contentType   string
    source        PartSource      // Data source (lazy)
    relationships *Relationships
    dirty         bool
    loaded        bool            // Is content loaded?
    blob          []byte          // Cached content (if loaded)
}
```

**Key Methods:**
- `Open() (io.ReadCloser, error)`: Open stream without loading to memory
- `Load() error`: Load content into memory
- `Blob() ([]byte, error)`: Get content (loads if not already loaded)
- `IsLoaded() bool`: Check if content is in memory

**Lazy Loading Flow:**
```
NewStreamPart() ──▶ source set, loaded=false
       │
       ▼
  Open() ──▶ Returns stream from source (no memory load)
       │
       ▼
  Load() ──▶ Reads stream into blob, sets loaded=true
       │
       ▼
  Blob() ──▶ Returns blob (calls Load() if needed)
```

#### 3. StreamingZipWriter

`StreamingZipWriter` enables streaming writes to ZIP files without buffering entire entries in memory.

```go
type StreamingZipWriter struct {
    zipWriter *zip.Writer
}
```

**Key Methods:**
- `WriteFromReader(path, reader)`: Stream from io.Reader to ZIP entry
- `WriteFromStreamer(path, streamer)`: Stream from StreamWriter to ZIP entry
- `WriteStreamPart(part)`: Stream a StreamPart to ZIP
- `WriteXML(path, data)`: Write XML with automatic header

#### 4. StreamWriter Interface

`StreamWriter` is implemented by types that can stream their content directly to an io.Writer.

```go
type StreamWriter interface {
    StreamWriteTo(w io.Writer) error
}
```

**Implementations:**
- `RelationshipsStreamer`: Stream XML for relationships
- `ContentTypesStreamer`: Stream XML for [Content_Types].xml

### StreamPackage

`StreamPackage` is the main package type for streaming operations.

#### Opening a Package (Lazy Load)

```go
// Open with lazy loading - only metadata is loaded
pkg, err := OpenStream("presentation.pptx")

// Get a part - content not loaded yet
part := pkg.GetPart(slideURI)

// Check if loaded
fmt.Println(part.IsLoaded()) // false

// Access content triggers loading
blob, err := part.Blob()      // Now loaded
fmt.Println(part.IsLoaded()) // true
```

#### Saving a Package (Stream Write)

```go
// Create streaming writer
file, _ := os.Create("output.pptx")
defer file.Close()

// Stream save - no buffering of complete XML
err := pkg.StreamSave(file)
```

### Complete Streaming Flow

#### Reading Flow

```
┌─────────────────────────────────────────────────────────────────┐
│                     OpenStream(path)                             │
│                            │                                     │
│                            ▼                                     │
│         ┌──────────────────────────────────────┐                │
│         │  1. Open ZIP file (keep handle open) │                │
│         │  2. Parse [Content_Types].xml        │                │
│         │  3. Parse _rels/.rels                │                │
│         │  4. Scan part URIs (no content load) │                │
│         └──────────────────────────────────────┘                │
│                            │                                     │
│                            ▼                                     │
│         ┌──────────────────────────────────────┐                │
│         │  StreamPackage                        │                │
│         │  - parts: map[URI]*StreamPart         │                │
│         │  - Each StreamPart.loaded = false     │                │
│         │  - Each StreamPart.source = ZipFile   │                │
│         └──────────────────────────────────────┘                │
│                            │                                     │
│         ┌──────────────────┴──────────────────┐                 │
│         ▼                                     ▼                 │
│   part.Open()                          part.Blob()              │
│   (stream read,                        (load to memory,         │
│    no memory load)                      can modify)             │
│                                                                   │
└─────────────────────────────────────────────────────────────────┘
```

#### Writing Flow

```
┌─────────────────────────────────────────────────────────────────┐
│                     StreamSave(writer)                           │
│                            │                                     │
│                            ▼                                     │
│         ┌──────────────────────────────────────┐                │
│         │  StreamingZipWriter                   │                │
│         └──────────────────────────────────────┘                │
│                            │                                     │
│         ┌──────────────────┼──────────────────┐                 │
│         ▼                  ▼                  ▼                 │
│   [Content_Types]    _rels/.rels       Part 1, Part 2...        │
│         │                  │                  │                 │
│         ▼                  ▼                  ▼                 │
│   ContentTypes-      Relationships-     StreamPart              │
│   Streamer           Streamer           .Open()                 │
│         │                  │                  │                 │
│         └──────────────────┴──────────────────┘                 │
│                            │                                     │
│                            ▼                                     │
│         ┌──────────────────────────────────────┐                │
│         │  xml.Encoder writes directly to ZIP  │                │
│         │  No buffering of complete XML        │                │
│         └──────────────────────────────────────┘                │
│                                                                   │
└─────────────────────────────────────────────────────────────────┘
```

### Comparison: Traditional vs Streaming

#### Traditional Package

```go
// Opens entire file into memory
pkg, _ := OpenFile("large.pptx")

// All parts are in memory
for _, part := range pkg.AllParts() {
    blob := part.Blob()  // Already in memory
}
```

#### Streaming Package

```go
// Only opens metadata
pkg, _ := OpenStream("large.pptx")

// Parts are not loaded
for _, part := range pkg.AllParts() {
    if needContent {
        blob, _ := part.Blob()  // Load on demand
    }
}
```

### Best Practices

1. **Use StreamPackage for large files**
   - When file size > 50MB
   - When only reading metadata
   - When only modifying few parts

2. **Use traditional Package for small files**
   - When file size < 10MB
   - When modifying many parts
   - When you need random access to all content

3. **Keep file handle open for lazy loading**
   - StreamPackage keeps the ZIP file open
   - Call `Close()` when done to release resources

4. **Load only what you need**
   ```go
   // Only load specific part types
   slides := pkg.GetPartsByType(ContentTypeSlide)
   for _, slide := range slides {
       if needsModification(slide) {
           slide.Load()  // Only load slides that need changes
       }
   }
   ```

### Thread Safety

Both `StreamPart` and `StreamPackage` are thread-safe:
- Internal `sync.RWMutex` protects all operations
- Multiple goroutines can safely access different parts
- Loading is atomic and idempotent

---

<a name="中文"></a>

## 中文

### 设计哲学

#### 核心原则

1. **懒加载**：只在需要时加载数据，而不是预先加载
2. **流式 I/O**：以流的方式处理数据，而不是在内存缓冲区中处理
3. **尽可能零拷贝**：避免不必要的数据复制
4. **向后兼容**：现有 API 继续工作

#### 内存效率目标

| 场景 | 传统方式 | 流式方式 |
|------|---------|---------|
| 打开 100MB PPTX | 加载 100MB 到内存 | 只加载元数据（~1MB） |
| 修改一张幻灯片 | 保留所有部件在内存 | 只加载修改的部件 |
| 保存修改的文件 | 在内存中构建完整 XML | 直接流式写入 ZIP |

### 架构概览

```
┌────────────────────────────────────────────────────────────────┐
│                        StreamPackage                            │
│                                                                 │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────────────────┐ │
│  │ ContentTypes│  │  关系       │  │      部件               │ │
│  │ (已加载)    │  │ (已加载)    │  │  ┌─────────────────┐   │ │
│  │             │  │             │  │  │ StreamPart      │   │ │
│  │             │  │             │  │  │ ┌─────────────┐ │   │ │
│  │             │  │             │  │  │ │ PartSource  │ │   │ │
│  │             │  │             │  │  │ │ (未加载)    │ │   │ │
│  │             │  │             │  │  │ └─────────────┘ │   │ │
│  │             │  │             │  │  └─────────────────┘   │ │
│  └─────────────┘  └─────────────┘  └─────────────────────────┘ │
│                                                                 │
└────────────────────────────────────────────────────────────────┘
```

### 关键组件

#### 1. PartSource 接口

`PartSource` 是部件数据源的抽象，支持从各种来源进行懒加载。

```go
type PartSource interface {
    Open() (io.ReadCloser, error)  // 打开流读取数据
    Size() int64                    // 返回数据大小（未知返回 -1）
}
```

**实现类型：**
- `ZipFileSource`：来自 ZIP 文件条目的数据（懒读取）
- `BytesSource`：来自内存的数据（[]byte）
- `ReaderSource`：来自 io.Reader 的数据

#### 2. StreamPart

`StreamPart` 是支持懒加载的部件。它只在显式请求时才将内容加载到内存。

```go
type StreamPart struct {
    uri           *PackURI
    contentType   string
    source        PartSource      // 数据源（懒加载）
    relationships *Relationships
    dirty         bool
    loaded        bool            // 内容是否已加载？
    blob          []byte          // 缓存的内容（如果已加载）
}
```

**关键方法：**
- `Open() (io.ReadCloser, error)`：打开流但不加载到内存
- `Load() error`：将内容加载到内存
- `Blob() ([]byte, error)`：获取内容（如果未加载则先加载）
- `IsLoaded() bool`：检查内容是否在内存中

**懒加载流程：**
```
NewStreamPart() ──▶ 设置 source，loaded=false
       │
       ▼
  Open() ──▶ 返回来自 source 的流（不加载到内存）
       │
       ▼
  Load() ──▶ 将流读入 blob，设置 loaded=true
       │
       ▼
  Blob() ──▶ 返回 blob（如果需要则调用 Load()）
```

#### 3. StreamingZipWriter

`StreamingZipWriter` 支持流式写入 ZIP 文件，无需在内存中缓冲整个条目。

```go
type StreamingZipWriter struct {
    zipWriter *zip.Writer
}
```

**关键方法：**
- `WriteFromReader(path, reader)`：从 io.Reader 流式写入 ZIP 条目
- `WriteFromStreamer(path, streamer)`：从 StreamWriter 流式写入 ZIP 条目
- `WriteStreamPart(part)`：将 StreamPart 流式写入 ZIP
- `WriteXML(path, data)`：写入 XML 并自动添加头

#### 4. StreamWriter 接口

`StreamWriter` 由可以直接将其内容流式传输到 io.Writer 的类型实现。

```go
type StreamWriter interface {
    StreamWriteTo(w io.Writer) error
}
```

**实现类型：**
- `RelationshipsStreamer`：流式写入关系的 XML
- `ContentTypesStreamer`：流式写入 [Content_Types].xml

### StreamPackage

`StreamPackage` 是用于流式操作的主要包类型。

#### 打开包（懒加载）

```go
// 懒加载打开 - 只加载元数据
pkg, err := OpenStream("presentation.pptx")

// 获取部件 - 内容尚未加载
part := pkg.GetPart(slideURI)

// 检查是否已加载
fmt.Println(part.IsLoaded()) // false

// 访问内容触发加载
blob, err := part.Blob()      // 现在已加载
fmt.Println(part.IsLoaded()) // true
```

#### 保存包（流式写入）

```go
// 创建流式写入器
file, _ := os.Create("output.pptx")
defer file.Close()

// 流式保存 - 不缓冲完整 XML
err := pkg.StreamSave(file)
```

### 完整流式流程

#### 读取流程

```
┌─────────────────────────────────────────────────────────────────┐
│                     OpenStream(path)                             │
│                            │                                     │
│                            ▼                                     │
│         ┌──────────────────────────────────────┐                │
│         │  1. 打开 ZIP 文件（保持句柄打开）     │                │
│         │  2. 解析 [Content_Types].xml         │                │
│         │  3. 解析 _rels/.rels                 │                │
│         │  4. 扫描部件 URI（不加载内容）       │                │
│         └──────────────────────────────────────┘                │
│                            │                                     │
│                            ▼                                     │
│         ┌──────────────────────────────────────┐                │
│         │  StreamPackage                        │                │
│         │  - parts: map[URI]*StreamPart         │                │
│         │  - 每个 StreamPart.loaded = false     │                │
│         │  - 每个 StreamPart.source = ZipFile   │                │
│         └──────────────────────────────────────┘                │
│                            │                                     │
│         ┌──────────────────┴──────────────────┐                 │
│         ▼                                     ▼                 │
│   part.Open()                          part.Blob()              │
│   （流式读取，                          （加载到内存，           │
│    不加载到内存）                       可以修改）               │
│                                                                   │
└─────────────────────────────────────────────────────────────────┘
```

#### 写入流程

```
┌─────────────────────────────────────────────────────────────────┐
│                     StreamSave(writer)                           │
│                            │                                     │
│                            ▼                                     │
│         ┌──────────────────────────────────────┐                │
│         │  StreamingZipWriter                   │                │
│         └──────────────────────────────────────┘                │
│                            │                                     │
│         ┌──────────────────┼──────────────────┐                 │
│         ▼                  ▼                  ▼                 │
│   [Content_Types]    _rels/.rels       部件 1, 部件 2...        │
│         │                  │                  │                 │
│         ▼                  ▼                  ▼                 │
│   ContentTypes-      Relationships-     StreamPart              │
│   Streamer           Streamer           .Open()                 │
│         │                  │                  │                 │
│         └──────────────────┴──────────────────┘                 │
│                            │                                     │
│                            ▼                                     │
│         ┌──────────────────────────────────────┐                │
│         │  xml.Encoder 直接写入 ZIP            │                │
│         │  不缓冲完整 XML                      │                │
│         └──────────────────────────────────────┘                │
│                                                                   │
└─────────────────────────────────────────────────────────────────┘
```

### 对比：传统 vs 流式

#### 传统 Package

```go
// 将整个文件打开到内存
pkg, _ := OpenFile("large.pptx")

// 所有部件都在内存中
for _, part := range pkg.AllParts() {
    blob := part.Blob()  // 已经在内存中
}
```

#### 流式 Package

```go
// 只打开元数据
pkg, _ := OpenStream("large.pptx")

// 部件未加载
for _, part := range pkg.AllParts() {
    if needContent {
        blob, _ := part.Blob()  // 按需加载
    }
}
```

### 最佳实践

1. **大文件使用 StreamPackage**
   - 当文件大小 > 50MB
   - 当只读取元数据
   - 当只修改少量部件

2. **小文件使用传统 Package**
   - 当文件大小 < 10MB
   - 当修改大量部件
   - 当需要随机访问所有内容

3. **保持文件句柄打开以支持懒加载**
   - StreamPackage 保持 ZIP 文件打开
   - 完成后调用 `Close()` 释放资源

4. **只加载需要的内容**
   ```go
   // 只加载特定类型的部件
   slides := pkg.GetPartsByType(ContentTypeSlide)
   for _, slide := range slides {
       if needsModification(slide) {
           slide.Load()  // 只加载需要修改的幻灯片
       }
   }
   ```

### 线程安全

`StreamPart` 和 `StreamPackage` 都是线程安全的：
- 内部 `sync.RWMutex` 保护所有操作
- 多个 goroutine 可以安全地访问不同部件
- 加载是原子且幂等的

### API 速查

| 操作 | 传统 Package | 流式 Package |
|------|-------------|-------------|
| 打开 | `OpenFile(path)` | `OpenStream(path)` |
| 获取部件 | `pkg.GetPart(uri)` | `pkg.GetPart(uri)` |
| 读取内容 | `part.Blob()` | `part.Blob()` 或 `part.Open()` |
| 检查加载 | N/A | `part.IsLoaded()` |
| 保存 | `pkg.SaveFile(path)` | `pkg.StreamSaveFile(path)` |
| 关闭 | `pkg.Close()` | `pkg.Close()` |

### 性能建议

1. **使用迭代器处理大量部件**
   ```go
   iter := pkg.NewPartIterator().FilterByType(ContentTypeSlide)
   for iter.Next() {
       slide := iter.Part()
       // 处理幻灯片
   }
   ```

2. **使用流式读取处理大部件**
   ```go
   rc, _ := part.Open()
   defer rc.Close()

   decoder := xml.NewDecoder(rc)
   // 流式解析 XML，不需要加载完整内容
   ```

3. **避免不必要的加载**
   ```go
   // 检查大小而不加载
   size := part.Size()

   // 检查是否已加载
   if !part.IsLoaded() {
       // 决定是否需要加载
   }
   ```
