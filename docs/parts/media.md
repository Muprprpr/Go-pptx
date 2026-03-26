# Media 模块接口文档

> 统一处理 PPTX 中的图片、音频、视频等媒体文件

---

## 枚举类型

### MediaType

媒体类型枚举。

| 常量 | 值 | 说明 |
|------|-----|------|
| `MediaTypeUnknown` | `0` | 未知类型 |
| `MediaTypeImage` | `1` | 图片 |
| `MediaTypeAudio` | `2` | 音频 |
| `MediaTypeVideo` | `3` | 视频 |

#### 方法

```go
func (mt MediaType) String() string
```
返回媒体类型的字符串表示：`"image"` / `"audio"` / `"video"` / `"unknown"`

---

## MediaResource

媒体资源结构体（只读），统一处理图片、音频、视频。

### 字段

| 字段 | 类型 | 访问器 | 说明 |
|------|------|--------|------|
| `fileName` | `string` | `FileName()` | 文件名（如 `image1.png`） |
| `contentType` | `string` | `ContentType()` | MIME 类型（如 `image/png`） |
| `mediaType` | `MediaType` | `MediaType()` | 媒体类型枚举 |
| `target` | `string` | `Target()` | ZIP 中完整路径（如 `ppt/media/image1.png`） |
| `data` | `[]byte` | `Data()` | 小文件字节数据（可为 nil） |
| `dataSize` | `int64` | `DataSize()` | 数据大小（字节） |
| `reader` | `io.Reader` | `Reader()` | 大文件 Reader（可为 nil） |
| `rId` | `string` | `RID()` | 关系 ID |
| `extension` | `string` | `Extension()` | 文件扩展名（如 `.png`） |
| `hash` | `string` | `Hash()` | 内容 Hash（MD5） |

### 类型判断方法

```go
func (m *MediaResource) HasData() bool      // 是否有字节数据
func (m *MediaResource) HasReader() bool     // 是否有 Reader
func (m *MediaResource) IsImage() bool       // 是否为图片
func (m *MediaResource) IsAudio() bool       // 是否为音频
func (m *MediaResource) IsVideo() bool       // 是否为视频
```

### 设置方法

```go
func (m *MediaResource) SetRID(rId string)   // 设置关系 ID
func (m *MediaResource) SetHash(hash string)  // 设置内容 Hash
```

---

## 构造函数

### NewMediaResourceFromBytes

```go
func NewMediaResourceFromBytes(fileName, contentType, target string, data []byte) *MediaResource
```

从字节数据创建媒体资源，适用于小文件（如小图片）。

### NewMediaResourceFromReader

```go
func NewMediaResourceFromReader(fileName, contentType, target string, reader io.Reader, size int64) *MediaResource
```

从 Reader 创建媒体资源，适用于大文件（如视频、大图片）。

---

## MIME 类型辅助函数

### 支持的图片类型

```go
image/png, image/jpeg, image/gif, image/bmp, image/tiff,
image/svg+xml, image/webp, image/x-emf, image/x-wmf
```

### 支持的音频类型

```go
audio/mpeg, audio/wav, audio/ogg, audio/aac, audio/mp4
```

### 支持的视频类型

```go
video/mp4, video/webm, video/ogg, video/quicktime,
video/x-msvideo, video/x-ms-wmv
```

---

## MediaManager

媒体资源管理器（并发安全缓存）。

### 设计原则

1. 一次写入，到处读取 - 初始化后主要操作是读取
2. 读取优化 - 使用 `sync.RWMutex`，读操作无需阻塞
3. 双重索引 - 按 rID 和 fileName 都能快速查找

### 索引结构

| 索引 | key | value |
|------|-----|-------|
| `byRID` | rID | `*MediaResource` |
| `byName` | fileName | rID |
| `byTarget` | target | rID |
| `byHash` | contentHash | rID |

### 全局实例

```go
func DefaultMediaManager() *MediaManager
var defaultMediaManager *MediaManager
```

### 构造函数

```go
func NewMediaManager() *MediaManager
```

---

## MediaManager 写入方法

### AddMedia

```go
func (m *MediaManager) AddMedia(resource *MediaResource) string
```

添加媒体资源到缓存，返回资源的 rID，如果已存在则返回现有 rID。

### AddMediaWithBytes

```go
func (m *MediaManager) AddMediaWithBytes(rID, fileName, contentType, target string, data []byte) *MediaResource
```

从字节数据添加媒体资源。

### AddMediaWithReader

```go
func (m *MediaManager) AddMediaWithReader(rID, fileName, contentType, target string, reader io.Reader, size int64) *MediaResource
```

从 Reader 添加媒体资源。

### AddMediaAuto

```go
func (m *MediaManager) AddMediaAuto(fileName string, data []byte) (string, *MediaResource)
```

自动推断 MIME 类型并生成自增 rID。如果相同内容已存在（基于 Hash），则返回已有资源（去重）。

### RemoveMedia

```go
func (m *MediaManager) RemoveMedia(rID string) bool
```

移除媒体资源。

### Clear

```go
func (m *MediaManager) Clear()
```

清空所有媒体资源。

---

## MediaManager 读取方法

### GetMedia

```go
func (m *MediaManager) GetMedia(rID string) *MediaResource
```

根据 rID 获取媒体资源。

### GetMediaByFileName

```go
func (m *MediaManager) GetMediaByFileName(fileName string) *MediaResource
```

根据文件名获取媒体资源。

### GetMediaByTarget

```go
func (m *MediaManager) GetMediaByTarget(target string) *MediaResource
```

根据目标路径获取媒体资源。

### GetMediaByHash

```go
func (m *MediaManager) GetMediaByHash(hash string) *MediaResource
```

根据内容 Hash 获取媒体资源（用于去重）。

### HasMedia

```go
func (m *MediaManager) HasMedia(rID string) bool
```

检查媒体资源是否存在。

### HasMediaByFileName

```go
func (m *MediaManager) HasMediaByFileName(fileName string) bool
```

检查文件名是否存在。

---

## MediaManager 批量读取方法

### AllMedia

```go
func (m *MediaManager) AllMedia() []*MediaResource
```

返回所有媒体资源。

### AllMediaByType

```go
func (m *MediaManager) AllMediaByType(mediaType MediaType) []*MediaResource
```

返回指定类型的所有媒体资源。

### AllImages

```go
func (m *MediaManager) AllImages() []*MediaResource
```

返回所有图片资源。

### AllAudio

```go
func (m *MediaManager) AllAudio() []*MediaResource
```

返回所有音频资源。

### AllVideo

```go
func (m *MediaManager) AllVideo() []*MediaResource
```

返回所有视频资源。

---

## MediaManager 统计方法

### Count

```go
func (m *MediaManager) Count() int64
```

返回媒体资源总数。

### CountByType

```go
func (m *MediaManager) CountByType(mediaType MediaType) int64
```

返回指定类型的媒体资源数量。

### CountImages

```go
func (m *MediaManager) CountImages() int64
```

返回图片数量。

### CountAudio

```go
func (m *MediaManager) CountAudio() int64
```

返回音频数量。

### CountVideo

```go
func (m *MediaManager) CountVideo() int64
```

返回视频数量。

---

## MediaManager 列表方法

### ListRIDs

```go
func (m *MediaManager) ListRIDs() []string
```

返回所有 rID。

### ListFileNames

```go
func (m *MediaManager) ListFileNames() []string
```

返回所有文件名。

### ListTargets

```go
func (m *MediaManager) ListTargets() []string
```

返回所有目标路径。

---

## 全局便捷函数

```go
func AddMedia(resource *MediaResource) string
func GetMedia(rID string) *MediaResource
func GetMediaByFileName(fileName string) *MediaResource
func GetMediaByTarget(target string) *MediaResource
func ClearMedia()
```

操作全局默认管理器 `defaultMediaManager`。
