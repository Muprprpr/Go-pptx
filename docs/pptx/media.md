# Media - 媒体管理

媒体资源管理器维护 PPTX 中所有媒体资源的并发安全缓存，支持图片、音频、视频等资源的自动去重。

## MediaManager

媒体资源管理器。

```go
type MediaManager struct {
    // Has unexported fields.
}
```

### 构造函数

```go
func NewMediaManager() *MediaManager
```

### 基础操作

#### AddMedia

添加媒体资源到缓存，返回资源的 rID，如果已存在则返回现有 rID。

```go
func (m *MediaManager) AddMedia(resource *parts.MediaResource) string
```

#### AddMediaAuto

自动推断 MIME 类型并生成自增 rID，如果相同内容已存在（基于 Hash），则返回已有资源（去重）。

```go
func (m *MediaManager) AddMediaAuto(fileName string, data []byte) (string, *parts.MediaResource)
```

**参数:**
- `fileName`: 文件名（用于推断 MIME 类型）
- `data`: 媒体数据

**返回:**
- 生成的 rID 和创建的 MediaResource

**示例:**

```go
mm := pptx.NewMediaManager()
data, _ := os.ReadFile("logo.png")

rId, resource := mm.AddMediaAuto("logo.png", data)
fmt.Printf("添加媒体: rId=%s, 类型=%s\n", rId, resource.ContentType)
```

#### AddMediaForSlide

为指定幻灯片添加媒体（支持跨幻灯片去重）。

```go
func (m *MediaManager) AddMediaForSlide(slideIndex int, data []byte, fileName string) (string, *parts.MediaResource)
```

**参数:**
- `slideIndex`: 幻灯片索引
- `data`: 媒体数据
- `fileName`: 文件名

**返回:**
- 该幻灯片的本地 rId 和全局媒体资源

**示例:**

```go
// 第 1 页插入 Logo
rId1, _ := mm.AddMediaForSlide(0, logoData, "logo.png")
// 返回: rId1="rId1", 全局存储 image1.png

// 第 2 页插入同一个 Logo
rId2, _ := mm.AddMediaForSlide(1, logoData, "logo.png")
// 返回: rId2="rId1"（该幻灯片的本地 rId）, 复用 image1.png

// 最终 ZIP 包中只有一份 image1.png，但两张幻灯片都有各自的 rId 引用
```

#### AddMediaWithBytes

从字节数据添加媒体资源。

```go
func (m *MediaManager) AddMediaWithBytes(rID, fileName, contentType, target string, data []byte) *parts.MediaResource
```

**参数:**
- `rID`: 关系 ID
- `fileName`: 文件名
- `contentType`: MIME 类型
- `target`: 目标路径
- `data`: 媒体数据

#### AddMediaWithReader

从 Reader 添加媒体资源。

```go
func (m *MediaManager) AddMediaWithReader(rID, fileName, contentType, target string, reader io.Reader, size int64) *parts.MediaResource
```

### 查询操作

#### GetMedia

根据 rID 获取媒体资源。

```go
func (m *MediaManager) GetMedia(rID string) *parts.MediaResource
```

#### GetMediaByFileName

根据文件名获取媒体资源。

```go
func (m *MediaManager) GetMediaByFileName(fileName string) *parts.MediaResource
```

#### GetMediaByHash

根据内容 Hash 获取媒体资源（用于去重）。

```go
func (m *MediaManager) GetMediaByHash(hash string) *parts.MediaResource
```

#### GetMediaByTarget

根据目标路径获取媒体资源。

```go
func (m *MediaManager) GetMediaByTarget(target string) *parts.MediaResource
```

#### GetGlobalMediaByHash

根据 Hash 获取全局媒体资源。

```go
func (m *MediaManager) GetGlobalMediaByHash(hash string) *parts.MediaResource
```

#### GetSlideMediaIndex

获取幻灯片媒体索引。

```go
func (m *MediaManager) GetSlideMediaIndex(slideIndex int) *SlideMediaIndex
```

### 列表操作

#### AllMedia

返回所有媒体资源（返回新切片，线程安全）。

```go
func (m *MediaManager) AllMedia() []*parts.MediaResource
```

#### AllGlobalMedia

返回所有全局媒体资源（去重后的）。

```go
func (m *MediaManager) AllGlobalMedia() []*parts.MediaResource
```

#### AllImages

返回所有图片资源。

```go
func (m *MediaManager) AllImages() []*parts.MediaResource
```

#### AllVideo

返回所有视频资源。

```go
func (m *MediaManager) AllVideo() []*parts.MediaResource
```

#### AllAudio

返回所有音频资源。

```go
func (m *MediaManager) AllAudio() []*parts.MediaResource
```

#### AllMediaByType

返回指定类型的所有媒体资源。

```go
func (m *MediaManager) AllMediaByType(mediaType parts.MediaType) []*parts.MediaResource
```

#### ListRIDs

返回所有 rID。

```go
func (m *MediaManager) ListRIDs() []string
```

#### ListFileNames

返回所有文件名。

```go
func (m *MediaManager) ListFileNames() []string
```

#### ListTargets

返回所有目标路径。

```go
func (m *MediaManager) ListTargets() []string
```

### 统计操作

#### Count

返回媒体资源总数。

```go
func (m *MediaManager) Count() int64
```

#### GlobalMediaCount

返回全局媒体资源数量（去重后）。

```go
func (m *MediaManager) GlobalMediaCount() int64
```

#### CountImages

返回图片数量。

```go
func (m *MediaManager) CountImages() int64
```

#### CountVideo

返回视频数量。

```go
func (m *MediaManager) CountVideo() int64
```

#### CountAudio

返回音频数量。

```go
func (m *MediaManager) CountAudio() int64
```

#### CountByType

返回指定类型的媒体资源数量。

```go
func (m *MediaManager) CountByType(mediaType parts.MediaType) int64
```

#### SlideCount

返回引用媒体的幻灯片数量。

```go
func (m *MediaManager) SlideCount() int64
```

### 其他操作

#### HasMedia

检查媒体资源是否存在。

```go
func (m *MediaManager) HasMedia(rID string) bool
```

#### HasMediaByFileName

检查文件名是否存在。

```go
func (m *MediaManager) HasMediaByFileName(fileName string) bool
```

#### RemoveMedia

移除媒体资源。

```go
func (m *MediaManager) RemoveMedia(rID string) bool
```

#### Clear

清空所有媒体资源。

```go
func (m *MediaManager) Clear()
```

### 去重统计

#### GetDeduplicationStats

获取去重统计信息。

```go
func (m *MediaManager) GetDeduplicationStats() DeduplicationStats
```

## DeduplicationStats

去重统计信息。

```go
type DeduplicationStats struct {
    // 全局媒体数量（实际存储）
    GlobalMediaCount int64

    // 总引用次数（所有幻灯片的引用总和）
    TotalReferences int64

    // 幻灯片数量
    SlideCount int64

    // 节省的存储空间（字节）
    SavedBytes int64

    // 去重率（0.0 - 1.0）
    DeduplicationRate float64
}
```

**示例:**

```go
stats := mm.GetDeduplicationStats()
fmt.Printf("全局媒体数: %d\n", stats.GlobalMediaCount)
fmt.Printf("总引用次数: %d\n", stats.TotalReferences)
fmt.Printf("节省空间: %d 字节\n", stats.SavedBytes)
fmt.Printf("去重率: %.2f%%\n", stats.DeduplicationRate*100)
```

## SlideMediaIndex

幻灯片媒体索引，管理单个幻灯片的媒体引用。

```go
type SlideMediaIndex struct {
    // Has unexported fields.
}
```

### 构造函数

```go
func NewSlideMediaIndex(slideIndex int) *SlideMediaIndex
```

### 方法

#### GetLocalRIDByHash

根据 Hash 获取本地 rId。

```go
func (smi *SlideMediaIndex) GetLocalRIDByHash(hash string) string
```

#### GetHashByLocalRID

根据本地 rId 获取 Hash。

```go
func (smi *SlideMediaIndex) GetHashByLocalRID(localRID string) string
```

#### AllLocalRIDs

返回所有本地 rId。

```go
func (smi *SlideMediaIndex) AllLocalRIDs() []string
```

#### LocalRefCount

返回本地引用数量。

```go
func (smi *SlideMediaIndex) LocalRefCount() int64
```

---

# MasterManager - 母版/版式管理器

母版/版式管理器。

```go
type MasterManager struct {
    // Has unexported fields.
}
```

### 构造函数

```go
func NewMasterManager() *MasterManager

func NewMasterManagerWithCache(cache *MasterCache) *MasterManager
```

### 加载方法

#### LoadFromZipFile

从 ZIP 文件路径加载。

```go
func (m *MasterManager) LoadFromZipFile(filePath string) error
```

#### LoadFromZip

从 ZIP Reader 加载母版和版式。

```go
func (m *MasterManager) LoadFromZip(zipReader *zip.Reader) error
```

**说明:** 遍历 ZIP 内的 `/ppt/slideMasters/` 和 `/ppt/slideLayouts/` 目录

### 查询方法

#### GetMaster

获取母版。

```go
func (m *MasterManager) GetMaster(masterID string) (*parts.SlideMasterData, bool)
```

#### GetMasterByName

根据名称获取母版。

```go
func (m *MasterManager) GetMasterByName(name string) (*parts.SlideMasterData, bool)
```

#### GetLayout

获取版式。

```go
func (m *MasterManager) GetLayout(layoutID string) (*parts.SlideLayoutData, bool)
```

#### GetLayoutByName

根据名称获取版式。

```go
func (m *MasterManager) GetLayoutByName(name string) (*parts.SlideLayoutData, bool)
```

#### GetPlaceholder

获取占位符。

```go
func (m *MasterManager) GetPlaceholder(layoutID, phType string) (*parts.Placeholder, bool)
```

### 列表方法

#### AllMasters

返回所有母版。

```go
func (m *MasterManager) AllMasters() map[string]*parts.SlideMasterData
```

#### AllLayouts

返回所有版式。

```go
func (m *MasterManager) AllLayouts() map[string]*parts.SlideLayoutData
```

#### ListLayoutIDs

列出所有版式 ID。

```go
func (m *MasterManager) ListLayoutIDs() []string
```

#### ListLayoutNames

列出所有版式名称。

```go
func (m *MasterManager) ListLayoutNames() []string
```

### 统计方法

#### MasterCount

返回母版数量。

```go
func (m *MasterManager) MasterCount() int
```

#### LayoutCount

返回版式数量。

```go
func (m *MasterManager) LayoutCount() int
```

#### Cache

返回内部缓存（只读）。

```go
func (m *MasterManager) Cache() *MasterCache
```

---

# MasterCache - 母版缓存

母版/版式只读缓存，初始化后所有字段只读，支持无锁并发访问。

```go
type MasterCache struct {
    // Has unexported fields.
}
```

### 构造函数

```go
func NewMasterCache() *MasterCache
```

### 初始化方法

#### Init

使用提供的数据初始化缓存（仅执行一次），后续调用将被忽略。

```go
func (c *MasterCache) Init(masters []*parts.SlideMasterData, layouts []*parts.SlideLayoutData)
```

#### InitFunc

延迟初始化，接受初始化函数，函数仅在第一次访问时执行。

```go
func (c *MasterCache) InitFunc(initFn func() ([]*parts.SlideMasterData, []*parts.SlideLayoutData))
```

### 查询方法

#### GetMaster

根据 ID 获取母版。

```go
func (c *MasterCache) GetMaster(masterID string) (*parts.SlideMasterData, bool)
```

#### GetMasterByName

根据名称获取母版。

```go
func (c *MasterCache) GetMasterByName(name string) (*parts.SlideMasterData, bool)
```

#### GetLayout

根据 ID 获取版式。

```go
func (c *MasterCache) GetLayout(layoutID string) (*parts.SlideLayoutData, bool)
```

#### GetLayoutByName

根据名称获取版式。

```go
func (c *MasterCache) GetLayoutByName(name string) (*parts.SlideLayoutData, bool)
```

#### GetPlaceholder

根据版式 ID 和占位符类型获取占位符。

```go
func (c *MasterCache) GetPlaceholder(layoutID, phType string) (*parts.Placeholder, bool)
```

**参数:**
- `phType`: 可以是 `PlaceholderType.String()` 的值，如 "title", "body" 等

#### GetPlaceholderByID

根据版式 ID 和占位符 ID 获取占位符。

```go
func (c *MasterCache) GetPlaceholderByID(layoutID, placeholderID string) (*parts.Placeholder, bool)
```

#### GetMasterPlaceholder

根据母版 ID 和占位符类型获取占位符。

```go
func (c *MasterCache) GetMasterPlaceholder(masterID, phType string) (*parts.Placeholder, bool)
```

### 列表方法

#### AllMasters

返回所有母版（只读）。

```go
func (c *MasterCache) AllMasters() map[string]*parts.SlideMasterData
```

#### AllLayouts

返回所有版式（只读）。

```go
func (c *MasterCache) AllLayouts() map[string]*parts.SlideLayoutData
```

#### ListMasterIDs

列出所有母版 ID。

```go
func (c *MasterCache) ListMasterIDs() []string
```

#### ListLayoutIDs

列出所有版式 ID。

```go
func (c *MasterCache) ListLayoutIDs() []string
```

#### ListLayoutNames

列出所有版式名称。

```go
func (c *MasterCache) ListLayoutNames() []string
```

### 检查方法

#### MasterExists

检查母版是否存在。

```go
func (c *MasterCache) MasterExists(masterID string) bool
```

#### LayoutExists

检查版式是否存在。

```go
func (c *MasterCache) LayoutExists(layoutID string) bool
```

### 统计方法

#### MasterCount

返回母版数量。

```go
func (c *MasterCache) MasterCount() int
```

#### LayoutCount

返回版式数量。

```go
func (c *MasterCache) LayoutCount() int
```

## 使用示例

### 基础媒体管理

```go
// 获取媒体管理器
mm := pres.MediaManager()

// 添加图片
data, _ := os.ReadFile("logo.png")
rId, resource := mm.AddMediaAuto("logo.png", data)

// 在幻灯片上使用
slide := pres.AddSlide()
slide.AddPicture(100, 100, 200, 150, rId)
```

### 跨幻灯片去重

```go
// 同一张图片在多页使用
logoData, _ := os.ReadFile("logo.png")

// 在多个幻灯片上添加相同的 logo
for i := 0; i < 5; i++ {
    slide := pres.AddSlide()
    rId, _ := mm.AddMediaForSlide(i, logoData, "logo.png")
    slide.AddPicture(100, 100, 200, 150, rId)
}

// 最终只存储一份 logo 图片
stats := mm.GetDeduplicationStats()
fmt.Printf("节省空间: %d 字节\n", stats.SavedBytes)
```

### 查询媒体信息

```go
// 按类型获取媒体
images := mm.AllImages()
videos := mm.AllVideo()
audio := mm.AllAudio()

// 统计
fmt.Printf("图片: %d, 视频: %d, 音频: %d\n",
    mm.CountImages(), mm.CountVideo(), mm.CountAudio())
```

### 使用母版缓存

```go
// 获取母版缓存
cache := pres.MasterCache()

// 获取版式
layout, ok := cache.GetLayoutByName("title")
if ok {
    fmt.Println("找到标题版式:", layout.Name)
}

// 获取占位符
ph, ok := cache.GetPlaceholder(layout.ID, "title")
if ok {
    fmt.Printf("标题占位符位置: (%d, %d)\n", ph.X, ph.Y)
}

// 列出所有版式
for _, name := range cache.ListLayoutNames() {
    fmt.Println("版式:", name)
}
```
