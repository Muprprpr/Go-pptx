# Presentation 模块接口文档

> 对应 `/ppt/presentation.xml`，是整个 PPTX 逻辑上的根节点

---

## 常量

```go
SlideIDStart = 256  // Slide ID 的起始值
```

---

## 数据结构

### SlideSize

幻灯片尺寸，单位 EMU (English Metric Units)。

| 字段 | 类型 | 说明 |
|------|------|------|
| `Cx` | `int` | 宽度 |
| `Cy` | `int` | 高度 |

### StandardSlideSizes

标准幻灯片尺寸。

| 字段 | 类型 | 尺寸 |
|------|------|------|
| `Wide16x9` | `SlideSize` | 12192000 x 6858000 EMU |
| `Standard4x3` | `SlideSize` | 9144000 x 6858000 EMU |

### PresentationPart

演示文稿部件，对应 `/ppt/presentation.xml`。

| 字段 | 类型 | 说明 |
|------|------|------|
| `uri` | `*opc.PackURI` | 部件 URI |
| `slideIDs` | `[]uint32` | 分配过的 slide ID 列表 |
| `slideIDNext` | `uint32` | 下一个可分配的 slide ID |
| `slideCount` | `int32` | 当前幻灯片数量 |
| `slideMasterIDs` | `[]string` | 母版 rId 列表 |
| `slideLayoutIDs` | `[]string` | 布局 rId 列表 |
| `slideSize` | `SlideSize` | 幻灯片尺寸 |
| `notesMasterID` | `string` | 备注母版 rId |
| `themeID` | `string` | 主题 rId |

---

## XML 结构类型

### XPresentation

对应 `presentation.xml` 的完整 XML 结构。

| 字段 | XML 路径 | 类型 | 说明 |
|------|----------|------|------|
| `XmlnsA` | `xmlns:a` | `string` | DrawingML 命名空间 |
| `XmlnsR` | `xmlns:r` | `string` | 关系命名空间 |
| `XmlnsP` | `xmlns:p` | `string` | PresentationML 命名空间 |
| `Compatibility` | `p:compatSpt` | `*XCompatibility` | 兼容设置 |
| `SldSz` | `p:sldSz` | `*XSldSz` | 幻灯片尺寸 |
| `NotesSz` | `p:notesSz` | `*XSldSz` | 备注尺寸 |
| `SldIdLst` | `p:sldIdLst` | `*XSldIdLst` | 幻灯片 ID 列表 |
| `SldMasterIdLst` | `p:sldMasterIdLst` | `*XSldMasterIdLst` | 母版 ID 列表 |
| `NotesMasterIdLst` | `p:notesMasterIdLst` | `*XSldMasterIdLst` | 备注母版 ID 列表 |
| `PrintSettings` | `p:printSettings` | `*XPrintSettings` | 打印设置 |

### XCompatibility

兼容设置，对应 `p:compatSpt`。

| 字段 | XML 属性 | 类型 | 说明 |
|------|----------|------|------|
| `CompatMode` | `compatMode` | `string` | 兼容模式 |

### XSldSz

幻灯片尺寸，对应 `p:sldSz`。

| 字段 | XML 属性 | 类型 | 说明 |
|------|----------|------|------|
| `Cx` | `cx` | `int` | 宽度 |
| `Cy` | `cy` | `int` | 高度 |

### XSldIdLst

幻灯片 ID 列表，对应 `p:sldIdLst`。

| 字段 | 类型 | 说明 |
|------|------|------|
| `SldIds` | `[]XSldId` | 幻灯片 ID 列表 |

### XSldId

单个幻灯片 ID，对应 `p:sldId`。

| 字段 | XML 属性 | 类型 | 说明 |
|------|----------|------|------|
| `Id` | `id` | `uint32` | 幻灯片 ID |
| `RId` | `rid` | `string` | 关系 ID |

> **注意**: XML 属性 `r:id` 在解析时会转换为 `rid`（Go xml 标签为 `xml:"rid,attr"`）。

### XSldMasterIdLst

母版 ID 列表，对应 `p:sldMasterIdLst`。

| 字段 | 类型 | 说明 |
|------|------|------|
| `SldMasterIds` | `[]XSldMasterId` | 母版 ID 列表 |

### XSldMasterId

单个母版 ID，对应 `p:sldMasterId`。

| 字段 | XML 属性 | 类型 | 说明 |
|------|----------|------|------|
| `Id` | `id` | `uint32` | 母版 ID |
| `RId` | `rid` | `string` | 关系 ID |

> **注意**: XML 属性 `r:id` 在解析时会转换为 `rid`。

### XPrintSettings

打印设置，对应 `p:printSettings`。

| 字段 | 类型 | 说明 |
|------|------|------|
| `OutputOptions` | `*XOutputOptions` | 输出选项 |

### XOutputOptions

输出选项，对应 `p:outputOptions`。

| 字段 | XML 属性 | 类型 | 说明 |
|------|----------|------|------|
| `UsePrintFml` | `usePrintFml` | `*bool` | 使用打印格式 |
| `CloneLinkedObjs` | `cloneLinkedObjs` | `*bool` | 克隆链接对象 |

---

## 构造函数

### NewPresentationPart

```go
func NewPresentationPart() *PresentationPart
```

创建演示文稿部件，使用默认 16:9 宽屏尺寸。

### NewPresentationPartWithSize

```go
func NewPresentationPartWithSize(size SlideSize) *PresentationPart
```

创建演示文稿并设置指定尺寸。

---

## PresentationPart 方法

### URI 方法

```go
func (p *PresentationPart) PartURI() *opc.PackURI
```

### 尺寸方法

```go
func (p *PresentationPart) SlideSize() SlideSize
func (p *PresentationPart) SetSlideSize(size SlideSize)
```

### 幻灯片管理

```go
func (p *PresentationPart) SlideCount() int32
func (p *PresentationPart) SlideIDAt(index int) (uint32, error)
func (p *PresentationPart) SlideIDs() []uint32
func (p *PresentationPart) AddSlide(layoutRId string, slidePart *SlidePart) error
func (p *PresentationPart) RemoveSlide(index int) error
```

### 母版管理

```go
func (p *PresentationPart) SlideMasterIDs() []string
func (p *PresentationPart) AddSlideMaster(rId string)
```

### XML 序列化/反序列化

```go
func (p *PresentationPart) ToXML() ([]byte, error)
func (p *PresentationPart) FromXML(data []byte) error
```

> **注意**: `FromXML` 内部会自动调用 `StripNamespacePrefixes` 处理命名空间前缀问题。详见 [xmlutils.md](xmlutils.md)。

---

## 辅助函数

### NewSlideSizeFromStandard

```go
func NewSlideSizeFromStandard(name string) SlideSize
```

根据标准尺寸名称创建 SlideSize。

| 参数 | 说明 |
|------|------|
| `"16:9"`, `"wide"`, `"widescreen"` | 返回 Wide16x9 |
| `"4:3"`, `"standard"` | 返回 Standard4x3 |
| 其他 | 默认返回 Wide16x9 |

---

## EMU 单位转换

### EMUFromPoints / PointsFromEMU

```go
func EMUFromPoints(points float64) int
func PointsFromEMU(emu int) float64
```

磅值与 EMU 互转（1 pt = 12700 EMU）。

### EMUFromInches / InchesFromEMU

```go
func EMUFromInches(inches float64) int
func InchesFromEMU(emu int) float64
```

英寸与 EMU 互转（1 inch = 914400 EMU）。

### EMUFromMM / MMFromEMU

```go
func EMUFromMM(mm float64) int
func MMFromEMU(emu int) float64
```

毫米与 EMU 互转（1 mm = 36000 EMU）。
