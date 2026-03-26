# Slide 模块接口文档

> 对应 `/ppt/slides/slideN.xml`，包含幻灯片、版式、形状、文本、图片、表格等 XML 结构

---

## 枚举类型

### SlideLayoutType

幻灯片布局类型，对应 `slideLayoutN.xml`。

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

## 关系类型常量

```go
const (
    RelationshipTypeImage       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    RelationshipTypeMedia      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"
    RelationshipTypeChart      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
    RelationshipTypeSlideLayout = "http://schemas.openxmlformats.org/presentationml/2006/relationships/slideLayout"
    RelationshipTypeSlideMaster = "http://schemas.openxmlformats.org/presentationml/2006/relationships/slideMaster"
    RelationshipTypeTable      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
)
```

---

## SlidePart

幻灯片部件，对应 `/ppt/slides/slideN.xml`。

### 创建

```go
func NewSlidePart(id int) *SlidePart
func NewSlidePartWithURI(uri *opc.PackURI) *SlidePart
```

### URI 方法

```go
func (s *SlidePart) PartURI() *opc.PackURI
func (s *SlidePart) SetURI(uri *opc.PackURI)
```

### 布局/母版关联

```go
func (s *SlidePart) LayoutRId() string
func (s *SlidePart) SetLayoutRId(rId string)
func (s *SlidePart) MasterRId() string
func (s *SlidePart) SetMasterRId(rId string)
```

### 关系管理

```go
func (s *SlidePart) Relationships() *SlideRelationships
func (s *SlidePart) AddImage(targetURI string) string
func (s *SlidePart) AddMedia(targetURI string) string
func (s *SlidePart) AddChart(targetURI string) string
func (s *SlidePart) GetRelationshipURI(rId string) string
func (s *SlidePart) HasImage(targetURI string) bool
func (s *SlidePart) HasMedia(targetURI string) bool
func (s *SlidePart) GetImageRId(targetURI string) string
func (s *SlidePart) GetMediaRId(targetURI string) string
func (s *SlidePart) GetChartRId(targetURI string) string
func (s *SlidePart) GetOrAddPicture(x, y, cx, cy int, imageURI string) *XPicture
```

### Shape ID 管理

```go
func (s *SlidePart) Allocator() *ShapeIDAllocator
func (s *SlidePart) NextShapeID() uint32
func (s *SlidePart) AllocateShapeID() uint32
func (s *SlidePart) AllocateShapeIDBatch(count int) []uint32
func (s *SlidePart) AllocateShapeIDWithOffset(offset uint32) uint32
func (s *SlidePart) PeekNextShapeID() uint32
func (s *SlidePart) CurrentShapeID() uint32
func (s *SlidePart) ResetShapeID()
func (s *SlidePart) SetShapeIDStart(startID uint32)
func (s *SlidePart) ShapeIDCount() uint32
```

### 添加形状

```go
func (s *SlidePart) AddShape(shape any)
func (s *SlidePart) AddTextBox(x, y, cx, cy int, text string) *XSp
func (s *SlidePart) AddAutoShape(x, y, cx, cy int, presetID string) *XSp
func (s *SlidePart) AddPicture(x, y, cx, cy int, imageRId string) *XPicture
func (s *SlidePart) AddTable(x, y, cx, cy, rows, cols int) *XGraphicFrame
func (s *SlidePart) SetTableCellText(gf *XGraphicFrame, row, col int, text string)
```

### XML 序列化

```go
func (s *SlidePart) ToXML() ([]byte, error)
func (s *SlidePart) FromXML(data []byte) error
```

> **注意**:
> - `ToXML` 使用 `ToXMLFast()` 进行高效序列化，输出带命名空间前缀的标准 OOXML 格式
> - `FromXML` 内部自动调用 `StripNamespacePrefixes` 处理命名空间问题。详见 [xmlutils.md](xmlutils.md)

---

## SlideLayoutPart

幻灯片版式部件，对应 `/ppt/slideLayouts/slideLayoutN.xml`。

### 创建

```go
func NewSlideLayoutPart(id int) *SlideLayoutPart
```

### 方法

```go
func (s *SlideLayoutPart) PartURI() *opc.PackURI
func (s *SlideLayoutPart) LayoutType() SlideLayoutType
func (s *SlideLayoutPart) SetLayoutType(t SlideLayoutType)
func (s *SlideLayoutPart) MasterRId() string
func (s *SlideLayoutPart) SetMasterRId(rId string)
```

---

## SlideRelationships

页面级 Relationship 管理，维护图片、图表、布局等 rId 映射。

### 创建

```go
func NewSlideRelationships() *SlideRelationships
```

### 添加关系

```go
func (sr *SlideRelationships) AddImageRel(targetURI string) string
func (sr *SlideRelationships) AddMediaRel(targetURI string) string
func (sr *SlideRelationships) AddChartRel(targetURI string) string
func (sr *SlideRelationships) AddTableRel(targetURI string) string
```

### 查询关系

```go
func (sr *SlideRelationships) ImageRels() map[string]string
func (sr *SlideRelationships) MediaRels() map[string]string
func (sr *SlideRelationships) ChartRels() map[string]string
func (sr *SlideRelationships) TableRels() map[string]string
func (sr *SlideRelationships) LayoutRId() string
func (sr *SlideRelationships) SetLayoutRId(rId string)
func (sr *SlideRelationships) MasterRId() string
func (sr *SlideRelationships) SetMasterRId(rId string)
func (sr *SlideRelationships) GetImageRelByURI(targetURI string) string
func (sr *SlideRelationships) GetMediaRelByURI(targetURI string) string
func (sr *SlideRelationships) RelationshipCount() int
```

### 序列化

```go
func (sr *SlideRelationships) ToRelationshipsXML() ([]byte, error)
```

---

## ShapeIDAllocator

形状 ID 分配器（单线程使用）。

### 创建

```go
func NewShapeIDAllocator(reservedID uint32) *ShapeIDAllocator
func NewShapeIDAllocatorWithMax(reservedID, maxID uint32) *ShapeIDAllocator
```

### 分配方法

```go
func (a *ShapeIDAllocator) Next() uint32                    // 分配下一个 ID
func (a *ShapeIDAllocator) NextBatch(count int) []uint32    // 批量分配
func (a *ShapeIDAllocator) Peek() uint32                    // 查看下一个 ID（不分配）
func (a *ShapeIDAllocator) Current() uint32                 // 返回当前 ID
func (a *ShapeIDAllocator) Reset()                         // 重置
func (a *ShapeIDAllocator) ResetFrom(startID uint32)       // 从指定 ID 重置
func (a *ShapeIDAllocator) SetReserved(reservedID uint32)   // 设置保留起始 ID
func (a *ShapeIDAllocator) Remaining() uint32               // 剩余可分配数量
func (a *ShapeIDAllocator) IsExhausted() bool               // 检查是否耗尽
func (a *ShapeIDAllocator) UsedCount() uint32              // 已使用数量
```

---

## ShapeIDAllocatorSync

线程安全的形状 ID 分配器。

### 创建

```go
func NewShapeIDAllocatorSync(reservedID uint32) *ShapeIDAllocatorSync
func NewShapeIDAllocatorSyncWithMax(reservedID, maxID uint32) *ShapeIDAllocatorSync
```

### 分配方法

```go
func (a *ShapeIDAllocatorSync) Next() uint32
func (a *ShapeIDAllocatorSync) NextBatch(count int) []uint32
func (a *ShapeIDAllocatorSync) TryNext() (uint32, bool)  // 尝试分配，失败返回 false
func (a *ShapeIDAllocatorSync) Peek() uint32
func (a *ShapeIDAllocatorSync) Reset()
func (a *ShapeIDAllocatorSync) ResetFrom(startID uint32)
```

---

## XML 结构类型

### XSpTree

形状树，对应 `<p:spTree>`。

```go
func NewXSpTree() *XSpTree
func (xst *XSpTree) WriteXML(xw *XMLWriter) error
func (xst *XSpTree) ToXMLFast() ([]byte, error)
```

### XSp

形状，对应 `<p:sp>`。

```go
func (xs *XSp) WriteXML(xw *XMLWriter) error
func (xs *XSp) ToXMLFast() ([]byte, error)
```

### XPicture

图片，对应 `<p:pic>`。

```go
func (xp *XPicture) WriteXML(xw *XMLWriter) error
func (xp *XPicture) ToXMLFast() ([]byte, error)
```

### XGraphicFrame

图形框架，对应 `<p:graphicFrame>`。

```go
func (xgf *XGraphicFrame) WriteXML(xw *XMLWriter) error
func (xgf *XGraphicFrame) ToXMLFast() ([]byte, error)
```

### XTextBody

文本主体，对应 `<p:txBody>`。

```go
func (xtb *XTextBody) WriteXML(xw *XMLWriter) error
func (xtb *XTextBody) ToXMLFast() ([]byte, error)
```

### XTextParagraph

文本段落，对应 `<a:p>`。

```go
func (xtp *XTextParagraph) WriteXML(xw *XMLWriter) error
func (xtp *XTextParagraph) ToXMLFast() ([]byte, error)
```

### XTextRun

文本片段，对应 `<a:r>`。

```go
func (xtr *XTextRun) WriteXML(xw *XMLWriter) error
func (xtr *XTextRun) ToXMLFast() ([]byte, error)
```

### XTable

表格，对应 `<a:tbl>`。

```go
func (xt *XTable) WriteXML(xw *XMLWriter) error
func (xt *XTable) ToXMLFast() ([]byte, error)
```

### XTableRow

表格行，对应 `<a:tr>`。

```go
func (xtr *XTableRow) WriteXML(xw *XMLWriter) error
func (xtr *XTableRow) ToXMLFast() ([]byte, error)
```

### XTableCell

表格单元格，对应 `<a:tc>`。

```go
func (xtc *XTableCell) WriteXML(xw *XMLWriter) error
func (xtc *XTableCell) ToXMLFast() ([]byte, error)
```

### XTransform2D

二维变换，对应 `<a:xfrm>`。

```go
func (xt *XTransform2D) WriteXML(xw *XMLWriter) error
func (xt *XTransform2D) ToXMLFast() ([]byte, error)
```

### XSlide

幻灯片 XML 结构，对应 `<p:sld>`。

```go
func (xs *XSlide) WriteXML(xw *XMLWriter) error
func (xs *XSlide) ToXMLFast() ([]byte, error)
```

### XCSld

公共幻灯片数据，对应 `<p:cSld>`。包含幻灯片的实际内容（形状树等）。

```go
type XCSld struct {
    SpTree *XSpTree `xml:"spTree"`  // 形状树
}
```

**XML 结构：**

```xml
<p:sld>
  <p:cSld>
    <p:spTree>
      <!-- 形状内容 -->
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <!-- 颜色映射覆盖 -->
  </p:clrMapOvr>
</p:sld>
```

**反序列化示例：**

```go
// 读取 slide XML
data := slidePart.Blob()

// 去除命名空间前缀（必需，Go xml.Unmarshal 不支持命名空间）
cleanData, err := parts.StripNamespacePrefixes(data)
if err != nil {
    return err
}

// 解析
var xs XSlide
if err := xml.Unmarshal(cleanData, &xs); err != nil {
    return err
}

// 访问形状树
if xs.CSld != nil && xs.CSld.SpTree != nil {
    for _, child := range xs.CSld.SpTree.Children {
        // 处理子元素
    }
}
```

### XBlipFillProperties

图片填充属性，对应 `<p:blipFill>`。

```go
func (xbfp *XBlipFillProperties) WriteXML(xw *XMLWriter) error
func (xbfp *XBlipFillProperties) ToXMLFast() ([]byte, error)
```

### XBlip

图片引用，对应 `<a:blip r:embed="..."/>`。

### XBodyPr

主体属性，对应 `<a:bodyPr>`。

| 字段 | XML 属性 | 类型 | 说明 |
|------|----------|------|------|
| `Wrap` | `wrap` | `string` | 自动换行 |
| `Rotation` | `rot` | `int` | 旋转角度 |
| `Vertical` | `vert` | `string` | 垂直方向 |
| `Anchor` | `anchor` | `string` | 锚点位置 |
| `AnchorCtr` | `anchorCtr` | `bool` | 居中锚点 |

### XClrMap

颜色映射，对应 `<p:clrMap>`。

| 字段 | XML 属性 | 类型 |
|------|----------|------|
| `BG1` | `bg1` | `string` |
| `T1` | `t1` | `string` |
| `BG2` | `bg2` | `string` |
| `T2` | `t2` | `string` |
| `Accent1-6` | `accent1-6` | `string` |
| `HLink` | `hlink` | `string` |
| `HLink1` | `hlink1` | `string` |
| `HLink2` | `hlink2` | `string` |
| `FollClr` | `follClr` | `string` |
| `LastClr` | `lastClr` | `string` |

### XSlideRelationships

幻灯片关系，对应 `_rels/slideN.xml.rels`。

```go
func (xsr *XSlideRelationships) WriteXML(xw *XMLWriter) error
func (xsr *XSlideRelationships) ToXMLFast() ([]byte, error)
```

---

## XMLWriter

流式 XML 写入辅助，提供高效的 XML 生成。

### 创建

```go
func NewXMLWriter(w io.Writer) *XMLWriter
func NewXMLWriterWithIndent(w io.Writer, indentStr string) *XMLWriter
func NewXMLWriterBuffered(cap int) *XMLWriter
```

### 配置

```go
func (xw *XMLWriter) SetAutoFlush(enable bool)
func (xw *XMLWriter) SetIndent(indentStr string)
func (xw *XMLWriter) SetUseIndent(use bool)
func (xw *XMLWriter) Reset(w io.Writer)
func (xw *XMLWriter) ResetBuffer()
```

### XML 写入

```go
func (xw *XMLWriter) Declaration() error
func (xw *XMLWriter) DeclarationWithEncoding(encoding string) error
func (xw *XMLWriter) StartElement(prefix, localName string) error
func (xw *XMLWriter) StartElementNS(prefix, localName, ns string) error
func (xw *XMLWriter) StartElementWithAttrs(prefix, localName string, attrs ...string) error
func (xw *XMLWriter) StartElementNSWithAttrs(prefix, localName, ns string, attrs ...string) error
func (xw *XMLWriter) StartElementRaw(prefix, localName string, attrs ...string) error
func (xw *XMLWriter) EndElement(prefix, localName string) error
func (xw *XMLWriter) EmptyElement(prefix, localName string) error
func (xw *XMLWriter) EmptyElementWithAttrs(prefix, localName string, attrs ...string) error
```

### 内容写入

```go
func (xw *XMLWriter) Text(content string) error
func (xw *XMLWriter) TextRaw(content string) error
func (xw *XMLWriter) CharData(data []byte) error
func (xw *XMLWriter) Comment(content string) error
func (xw *XMLWriter) CData(content string) error
func (xw *XMLWriter) ProcessingInstruction(target, data string) error
func (xw *XMLWriter) Newline() error
func (xw *XMLWriter) Raw(content string) error
```

### 缩进控制

```go
func (xw *XMLWriter) Indent()
func (xw *XMLWriter) Dedent()
func (xw *XMLWriter) WithIndent(fn func())
```

### 数值写入

```go
func (xw *XMLWriter) WriteInt(val int) error
func (xw *XMLWriter) WriteInt64(val int64) error
func (xw *XMLWriter) WriteUint64(val uint64) error
func (xw *XMLWriter) WriteFloat64(val float64, prec int) error
func (xw *XMLWriter) WriteBool(val bool) error
func (xw *XMLWriter) WriteBoolStr(val bool) error
```

### EMU 单位写入

```go
func (xw *XMLWriter) WriteEMUs(val int64) error
func (xw *XMLWriter) WriteEMUsWithUnit(val int64) error
func (xw *XMLWriter) WriteEMUsF(val float64) error
func (xw *XMLWriter) WriteInchesAsEMU(inches float64) error
func (xw *XMLWriter) WriteCentimetersAsEMU(cm float64) error
func (xw *XMLWriter) WriteMillimetersAsEMU(mm float64) error
func (xw *XMLWriter) WritePointsAsEMU(points float64) error
func (xw *XMLWriter) WritePixelsAsEMU(pixels float64) error
func (xw *XMLWriter) WritePercentage(val int) error
```

### 输出

```go
func (xw *XMLWriter) Flush() error
func (xw *XMLWriter) Bytes() []byte
func (xw *XMLWriter) String() string
func (xw *XMLWriter) Size() int
func (xw *XMLWriter) Capacity() int
```

---

## XMLWriterPool

XMLWriter 对象池，用于减少内存分配。

### 创建

```go
func NewXMLWriterPool() *XMLWriterPool
```

### 方法

```go
func (p *XMLWriterPool) Get() *XMLWriter
func (p *XMLWriterPool) Put(xw *XMLWriter)
func (p *XMLWriterPool) GetWithWriter(w io.Writer) *XMLWriter
func (p *XMLWriterPool) GetBuffered() *XMLWriter
```
