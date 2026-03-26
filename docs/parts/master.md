# Master 模块接口文档

> 母版和版式的只读数据结构、解析器和缓存管理

## 设计原则

1. 所有结构体字段均为只读（小写字段通过构造函数初始化，大写字段为不可变值）
2. 针对高并发读取优化，无需加锁即可安全读取
3. 数据在解析时一次性构建，之后不再修改

---

## 枚举类型

### PlaceholderType

占位符类型枚举，对应 XML: `<p:ph type="...">`

| 常量 | 值 | XML 类型 |
|------|-----|----------|
| `PlaceholderTypeNone` | `0` | 未指定 |
| `PlaceholderTypeTitle` | `1` | `title` |
| `PlaceholderTypeBody` | `2` | `body` |
| `PlaceholderTypeCenterTitle` | `3` | `ctrTitle` |
| `PlaceholderTypeSubTitle` | `4` | `subTitle` |
| `PlaceholderTypeDateTime` | `5` | `dt` |
| `PlaceholderTypeSlideNumber` | `6` | `sldNum` |
| `PlaceholderTypeFooter` | `7` | `ftr` |
| `PlaceholderTypeHeader` | `8` | `hdr` |
| `PlaceholderTypeObject` | `9` | `obj` |
| `PlaceholderTypeChart` | `10` | `chart` |
| `PlaceholderTypeTable` | `11` | `tbl` |
| `PlaceholderTypeClipArt` | `12` | `clipArt` |
| `PlaceholderTypeOrgChart` | `13` | `dgm` |
| `PlaceholderTypeMedia` | `14` | `media` |
| `PlaceholderTypeSlideImage` | `15` | `sldImg` |
| `PlaceholderTypePicture` | `16` | `pic` |

#### 方法

```go
func (t PlaceholderType) String() string
```
返回占位符类型的字符串表示，如 `"title"`, `"body"` 等。

### BackgroundType

背景类型枚举，对应 XML: `<p:bg>` 下的不同子元素。

| 常量 | 值 | 说明 |
|------|-----|------|
| `BackgroundTypeNone` | `0` | 无背景 |
| `BackgroundTypeSolidColor` | `1` | 纯色背景 |
| `BackgroundTypeGradient` | `2` | 渐变背景 |
| `BackgroundTypePattern` | `3` | 图案填充 |
| `BackgroundTypePicture` | `4` | 图片背景 |
| `BackgroundTypeThemeColor` | `5` | 主题色背景 |

### SlideLayoutType

版式类型枚举。

| 常量 | 值 | 说明 |
|------|-----|------|
| `SlideLayoutBlank` | `0` | 空白布局 |
| `SlideLayoutTitle` | `1` | 标题布局 |
| `SlideLayoutTitleAndContent` | `2` | 标题和内容布局 |
| `SlideLayoutTwoContent` | `3` | 两栏内容布局 |
| `SlideLayoutComparison` | `4` | 比较布局 |
| `SlideLayoutTitleOnly` | `5` | 仅标题布局 |
| `SlideLayoutBlankVertical` | `6` | 空白垂直布局 |
| `SlideLayoutObject` | `7` | 对象布局 |
| `SlideLayoutPictureAndCaption` | `8` | 图片和标题布局 |

---

## 只读数据结构

### TextStyle

文本样式，用于定义占位符中文本的默认字体、大小、颜色等。

| 字段 | 类型 | 访问器 | 说明 |
|------|------|--------|------|
| `fontName` | `string` | `FontName()` | 字体名称 |
| `fontSize` | `int32` | `FontSize()` | 字体大小（百分之一磅，100 = 1pt） |
| `bold` | `bool` | `Bold()` | 是否粗体 |
| `italic` | `bool` | `Italic()` | 是否斜体 |
| `underline` | `bool` | `Underline()` | 是否下划线 |
| `colorRGB` | `string` | `ColorRGB()` | 文本颜色（RGB 十六进制，如 `"FF0000"`） |

### Placeholder

占位符，母版/版式中定义的可填充区域。对应 XML: `<p:sp>` with `<p:nvSpPr><p:nvPr><p:ph ...>`

| 字段 | 类型 | 访问器 | 说明 |
|------|------|--------|------|
| `id` | `string` | `ID()` | 占位符唯一标识符 |
| `placeholderType` | `PlaceholderType` | `Type()` | 占位符类型 |
| `x` | `int64` | `X()` | X 坐标（EMU 单位） |
| `y` | `int64` | `Y()` | Y 坐标（EMU 单位） |
| `cx` | `int64` | `Cx()` | 宽度（EMU 单位） |
| `cy` | `int64` | `Cy()` | 高度（EMU 单位） |
| `rotation` | `int32` | `Rotation()` | 旋转角度（1/60000 度） |
| `defaultStyle` | `*TextStyle` | `DefaultStyle()` | 默认文本样式（可为 nil） |

#### 方法

```go
func (p *Placeholder) Bounds() (x, y, cx, cy int64)
```
返回边界矩形 (x, y, cx, cy)。

### Background

背景定义，对应 XML: `<p:bg>` 或 `<p:cSld><p:bg>`

| 字段 | 类型 | 访问器 | 说明 |
|------|------|--------|------|
| `backgroundType` | `BackgroundType` | `Type()` | 背景类型 |
| `solidColorRGB` | `string` | `SolidColorRGB()` | RGB 十六进制颜色值（仅纯色背景有效） |
| `gradientAngle` | `int32` | `GradientAngle()` | 渐变角度（度，仅渐变背景有效） |
| `gradientColors` | `[]GradientStop` | `GradientColors()` | 渐变色标列表（仅渐变背景有效） |
| `pictureRId` | `string` | `PictureRId()` | 图片关系 ID（仅图片背景有效） |
| `pictureURI` | `string` | `PictureURI()` | 图片内部 URI 路径（仅图片背景有效） |
| `opacity` | `float32` | `Opacity()` | 不透明度 (0.0 - 1.0) |

### GradientStop

渐变色标，对应 XML: `<a:gs>`

| 字段 | 类型 | 访问器 | 说明 |
|------|------|--------|------|
| `position` | `float32` | `Position()` | 位置 (0.0 - 1.0) |
| `colorRGB` | `string` | `ColorRGB()` | RGB 十六进制颜色值 |

### SlideLayoutData

版式只读数据，对应 XML: `/ppt/slideLayouts/slideLayoutN.xml`

| 字段 | 类型 | 访问器 | 说明 |
|------|------|--------|------|
| `id` | `string` | `ID()` | 版式唯一标识符 |
| `name` | `string` | `Name()` | 版式名称 |
| `layoutType` | `SlideLayoutType` | `LayoutType()` | 版式类型 |
| `background` | `*Background` | `Background()` | 背景（可为 nil） |
| `masterId` | `string` | `MasterID()` | 所属母版的 ID |
| `placeholders` | `map[string]*Placeholder` | `Placeholders()` | 占位符集合 |

#### 方法

```go
func (l *SlideLayoutData) PlaceholderByID(id string) *Placeholder
func (l *SlideLayoutData) PlaceholderCount() int
func (l *SlideLayoutData) PlaceholderByType(phType PlaceholderType) *Placeholder
func (l *SlideLayoutData) TitlePlaceholder() *Placeholder
func (l *SlideLayoutData) BodyPlaceholder() *Placeholder
```

### SlideMasterData

母版只读数据，对应 XML: `/ppt/slideMasters/slideMasterN.xml`

| 字段 | 类型 | 访问器 | 说明 |
|------|------|--------|------|
| `id` | `string` | `ID()` | 母版唯一标识符 |
| `name` | `string` | `Name()` | 母版名称 |
| `background` | `*Background` | `Background()` | 背景（可为 nil） |
| `placeholders` | `map[string]*Placeholder` | `Placeholders()` | 母版级占位符 |
| `layouts` | `[]*SlideLayoutData` | `Layouts()` | 包含的版式列表 |

#### 方法

```go
func (m *SlideMasterData) PlaceholderByID(id string) *Placeholder
func (m *SlideMasterData) PlaceholderCount() int
func (m *SlideMasterData) LayoutCount() int
func (m *SlideMasterData) LayoutByID(id string) *SlideLayoutData
```

---

## MasterCache

母版/版式只读缓存。初始化后所有字段只读，支持无锁并发访问。

### 创建

```go
func NewMasterCache() *MasterCache
```

### 初始化

```go
func (c *MasterCache) Init(masters []*SlideMasterData, layouts []*SlideLayoutData)
func (c *MasterCache) InitFunc(initFn func() ([]*SlideMasterData, []*SlideLayoutData))
```

### 读取接口

```go
func (c *MasterCache) GetMaster(masterID string) (*SlideMasterData, bool)
func (c *MasterCache) GetMasterByName(name string) (*SlideMasterData, bool)
func (c *MasterCache) GetLayout(layoutID string) (*SlideLayoutData, bool)
func (c *MasterCache) GetLayoutByName(name string) (*SlideLayoutData, bool)
func (c *MasterCache) GetPlaceholder(layoutID, phType string) (*Placeholder, bool)
func (c *MasterCache) GetPlaceholderByID(layoutID, placeholderID string) (*Placeholder, bool)
func (c *MasterCache) GetMasterPlaceholder(masterID, phType string) (*Placeholder, bool)
```

### 批量读取

```go
func (c *MasterCache) AllMasters() map[string]*SlideMasterData
func (c *MasterCache) AllLayouts() map[string]*SlideLayoutData
func (c *MasterCache) MasterCount() int
func (c *MasterCache) LayoutCount() int
```

### 辅助方法

```go
func (c *MasterCache) LayoutExists(layoutID string) bool
func (c *MasterCache) MasterExists(masterID string) bool
func (c *MasterCache) ListLayoutIDs() []string
func (c *MasterCache) ListMasterIDs() []string
func (c *MasterCache) ListLayoutNames() []string
```

### 全局缓存

```go
func DefaultCache() *MasterCache
func InitDefaultCache(masters []*SlideMasterData, layouts []*SlideLayoutData)
func GetLayout(layoutID string) (*SlideLayoutData, bool)
func GetLayoutByName(name string) (*SlideLayoutData, bool)
func GetMaster(masterID string) (*SlideMasterData, bool)
func GetPlaceholder(layoutID, phType string) (*Placeholder, bool)
```

---

## MasterManager

母版/版式管理器（门面模式），负责从 ZIP 文件加载母版和版式。

### 创建

```go
func NewMasterManager() *MasterManager
func NewMasterManagerWithCache(cache *MasterCache) *MasterManager
```

### 加载

```go
func (m *MasterManager) LoadFromZip(zipReader *zip.Reader) error
func (m *MasterManager) LoadFromZipFile(filePath string) error
```

### 访问器

```go
func (m *MasterManager) Cache() *MasterCache
func (m *MasterManager) GetLayout(layoutID string) (*SlideLayoutData, bool)
func (m *MasterManager) GetLayoutByName(name string) (*SlideLayoutData, bool)
func (m *MasterManager) GetMaster(masterID string) (*SlideMasterData, bool)
func (m *MasterManager) GetMasterByName(name string) (*SlideMasterData, bool)
func (m *MasterManager) GetPlaceholder(layoutID, phType string) (*Placeholder, bool)
func (m *MasterManager) AllLayouts() map[string]*SlideLayoutData
func (m *MasterManager) AllMasters() map[string]*SlideMasterData
func (m *MasterManager) LayoutCount() int
func (m *MasterManager) MasterCount() int
func (m *MasterManager) ListLayoutIDs() []string
func (m *MasterManager) ListLayoutNames() []string
```

### 全局管理器

```go
func DefaultManager() *MasterManager
func InitDefaultManager(zipReader *zip.Reader) error
func InitDefaultManagerFromFile(filePath string) error
```

---

## XML 解析器

### ParseLayout

```go
func ParseLayout(xmlData []byte) (*SlideLayoutData, error)
```

解析幻灯片版式 XML，返回版式数据。

### ParseMaster

```go
func ParseMaster(xmlData []byte) (*SlideMasterData, error)
```

解析幻灯片母版 XML，返回母版数据。

---

## 单位转换

```go
var EMUToPixels      = utils.EMUToPixels      // EMU -> 像素 (96 DPI)
var EMUToPoints       = utils.EMUToPoints       // EMU -> 磅
var EMUToInches       = utils.EMUToInches       // EMU -> 英寸
var EMUToCentimeters  = utils.EMUToCentimeters  // EMU -> 厘米
```
