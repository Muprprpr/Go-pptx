# XML Master Models 接口文档

> 母版和版式解析相关的 XML 结构体定义

---

## 基础类型

### XMLOffset

偏移量结构体，对应 XML: `<a:off x="..." y="..."/>`

| 字段 | XML 属性 | 类型 | 说明 |
|------|----------|------|------|
| `X` | `x` | `int64` | X 坐标 |
| `Y` | `y` | `int64` | Y 坐标 |

#### 方法

```go
func (o *XMLOffset) IsValid() bool  // 检查是否有效（x 和 y 属性必须存在）
func (o *XMLOffset) IsZero() bool   // 检查是否为零值
```

---

### XMLExtents

尺寸结构体，对应 XML: `<a:ext cx="..." cy="..."/>`

| 字段 | XML 属性 | 类型 | 说明 |
|------|----------|------|------|
| `Cx` | `cx` | `int64` | 宽度（EMU 单位） |
| `Cy` | `cy` | `int64` | 高度（EMU 单位） |

#### 方法

```go
func (e *XMLExtents) IsValid() bool  // 检查是否有效（OpenXML 规范要求 cx 和 cy 必须为正数）
func (e *XMLExtents) IsZero() bool   // 检查是否为零值
```

---

### XMLTransform

二维变换结构体，对应 XML: `<a:xfrm>...</a:xfrm>`

| 字段 | 类型 | 说明 |
|------|------|------|
| `Off` | `*XMLOffset` | 位置偏移 |
| `Ext` | `*XMLExtents` | 尺寸扩展 |

---

### XMLPlaceholder

占位符结构体，对应 XML: `<p:ph type="..." idx="..."/>`

| 字段 | XML 属性 | 类型 | 说明 |
|------|----------|------|------|
| `Type` | `type` | `string` | 占位符类型 |
| `Idx` | `idx` | `string` | 占位符索引 |

---

## 非视觉属性

### XMLCNvPr

通用非视觉属性，对应 XML: `<p:cNvPr id="..." name="..."/>`

| 字段 | XML 属性 | 类型 |
|------|----------|------|
| `ID` | `id` | `int` |
| `Name` | `name` | `string` |

---

### XMLNvPr

非视觉属性，对应 XML: `<p:nvPr>...</p:nvPr>`

| 字段 | 类型 | 说明 |
|------|------|------|
| `Ph` | `*XMLPlaceholder` | 占位符定义（若存在） |

---

### XMLNvSpPr

非视觉形状属性，对应 XML: `<p:nvSpPr>...</p:nvSpPr>`

| 字段 | 类型 | 说明 |
|------|------|------|
| `CNvPr` | `*XMLCNvPr` | 通用非视觉属性 |
| `NvPr` | `*XMLNvPr` | 非视觉属性 |

---

### XMLSpPr

视觉形状属性，对应 XML: `<p:spPr>...</p:spPr>`

| 字段 | 类型 | 说明 |
|------|------|------|
| `Xfrm` | `*XMLTransform` | 变换信息 |

---

## 组形状

### XMLNvGrpSpPr

非视觉组属性，对应 XML: `<p:nvGrpSpPr>...</p:nvGrpSpPr>`

| 字段 | 类型 |
|------|------|
| `CNvPr` | `*XMLCNvPr` |
| `CNvGrpSpPr` | `*XMLCNvGrpSpPr` |

---

### XMLCNvGrpSpPr

组形状非视觉属性，对应 XML: `<p:cNvGrpSpPr>...</p:cNvGrpSpPr>`

---

### XMLGrpSpPr

组形状属性，对应 XML: `<p:grpSpPr>...</p:grpSpPr>`

| 字段 | 类型 | 说明 |
|------|------|------|
| `Xfrm` | `*XMLTransform` | 组变换 |

---

### XMLGroupShape

组形状，对应 XML: `<p:grpSp>...</p:grpSp>`

| 字段 | 类型 | 说明 |
|------|------|------|
| `NvGrpSpPr` | `*XMLNvGrpSpPr` | 非视觉组属性 |
| `GrpSpPr` | `*XMLGrpSpPr` | 组形状属性 |
| `Shapes` | `[]XMLShape` | 子形状列表 |

---

## 背景

### XMLBackground

背景结构体，对应 XML: `<p:bg>...</p:bg>`

| 字段 | 类型 | 说明 |
|------|------|------|
| `BgPr` | `*XMLBackgroundPr` | 背景属性 |
| `BgRef` | `*XMLBackgroundRef` | 背景引用 |

---

### XMLBackgroundRef

背景引用，对应 XML: `<p:bgRef idx="..."><a:schemeClr val="..."/>`

| 字段 | XML 属性 | 类型 | 说明 |
|------|----------|------|------|
| `Idx` | `idx` | `string` | 背景索引 |
| `Clr` | `schemeClr` | `*XMLSchemeColor` | 主题颜色 |

---

### XMLBackgroundPr

背景属性，对应 XML: `<p:bgPr>...</p:bgPr>`

| 字段 | 类型 | 说明 |
|------|------|------|
| `Fill` | `*XMLFillProperties` | 填充属性 |

---

## 填充

### XMLFillProperties

填充属性（联合类型），对应 XML: `<a:solidFill>` / `<a:gradFill>` / `<a:blipFill>` 等

| 字段 | XML 元素 | 类型 |
|------|----------|------|
| `SolidFill` | `a:solidFill` | `*XMLSolidFill` |
| `GradFill` | `a:gradFill` | `*XMLGradFill` |
| `BlipFill` | `a:blipFill` | `*XMLBlipFill` |
| `NoFill` | `a:noFill` | `*struct{}` |

---

### XMLSolidFill

纯色填充，对应 XML: `<a:solidFill>...</a:solidFill>`

| 字段 | 类型 | 说明 |
|------|------|------|
| `SrgbClr` | `*XMLSRgbColor` | RGB 颜色 |
| `SchemeClr` | `*XMLSchemeColor` | 主题颜色 |

---

### XMLSRgbColor

RGB 颜色，对应 XML: `<a:srgbClr val="..."/>`

| 字段 | XML 属性 | 类型 |
|------|----------|------|
| `Val` | `val` | `string` |

---

### XMLSchemeColor

主题颜色，对应 XML: `<a:schemeClr val="..."/>`

| 字段 | XML 属性 | 类型 |
|------|----------|------|
| `Val` | `val` | `string` |

---

### XMLGradFill

渐变填充，对应 XML: `<a:gradFill>...</a:gradFill>`

| 字段 | 类型 | 说明 |
|------|------|------|
| `GsLst` | `*XMLGradientStopList` | 色标列表 |
| `Lin` | `*XMLLinearGradient` | 线性渐变 |

---

### XMLGradientStopList

渐变色标列表，对应 XML: `<a:gsLst>...</a:gsLst>`

| 字段 | 类型 | 说明 |
|------|------|------|
| `Stops` | `[]XMLGradientStop` | 色标 |

---

### XMLGradientStop

渐变色标，对应 XML: `<a:gs pos="...">...</a:gs>`

| 字段 | XML 属性 | 类型 | 说明 |
|------|----------|------|------|
| `Pos` | `pos` | `int64` | 位置 |
| `SolidFill` | `a:solidFill` | `*XMLSolidFill` | 颜色 |

---

### XMLLinearGradient

线性渐变，对应 XML: `<a:lin ang="..." scaled="..."/>`

| 字段 | XML 属性 | 类型 |
|------|----------|------|
| `Ang` | `ang` | `int64` |
| `Scaled` | `scaled` | `bool` |

---

### XMLBlipFill

图片填充，对应 XML: `<a:blipFill>...</a:blipFill>`

| 字段 | 类型 | 说明 |
|------|------|------|
| `Blip` | `*XMLBlip` | 图片引用 |

---

### XMLBlip

图片引用，对应 XML: `<a:blip r:embed="..."/>`

| 字段 | XML 属性 | 类型 |
|------|----------|------|
| `Embed` | `r:embed` | `string` |

---

## 形状

### XMLShape

形状结构体，对应 XML: `<p:sp>...</p:sp>`

| 字段 | 类型 | 说明 |
|------|------|------|
| `NvSpPr` | `*XMLNvSpPr` | 非视觉形状属性 |
| `SpPr` | `*XMLSpPr` | 视觉形状属性 |

---

### XMLShapeTree

形状树结构体，对应 XML: `<p:spTree>...</p:spTree>`

| 字段 | 类型 | 说明 |
|------|------|------|
| `NvGrpSpPr` | `*XMLNvGrpSpPr` | 非视觉组属性 |
| `GrpSpPr` | `*XMLGrpSpPr` | 组形状属性 |
| `Shapes` | `[]XMLShape` | 形状列表 |
| `GroupShapes` | `[]XMLGroupShape` | 组形状列表 |

---

## 幻灯片结构

### XMLCommonSlideData

通用幻灯片数据，对应 XML: `<p:cSld>...</p:cSld>`

| 字段 | 类型 | 说明 |
|------|------|------|
| `Bg` | `*XMLBackground` | 背景 |
| `SpTree` | `*XMLShapeTree` | 形状树 |

---

### XMLSlideLayout

幻灯片版式，对应 XML: `<p:sldLayout>...</p:sldLayout>`

| 字段 | 类型 | 说明 |
|------|------|------|
| `XmlnsA` | `string` | DrawingML 命名空间 |
| `XmlnsR` | `string` | 关系命名空间 |
| `XmlnsP` | `string` | PresentationML 命名空间 |
| `CSld` | `*XMLCommonSlideData` | 通用幻灯片数据 |

---

### XMLSlideMaster

幻灯片母版，对应 XML: `<p:sldMaster>...</p:sldMaster>`

| 字段 | 类型 | 说明 |
|------|------|------|
| `XmlnsA` | `string` | DrawingML 命名空间 |
| `XmlnsR` | `string` | 关系命名空间 |
| `XmlnsP` | `string` | PresentationML 命名空间 |
| `CSld` | `*XMLCommonSlideData` | 通用幻灯片数据 |

---

## 线条

### XLineProperties

线条属性，对应 XML: `<a:ln w="...">`

| 字段 | XML 属性 | 类型 | 说明 |
|------|----------|------|------|
| `Width` | `w` | `int` | 线条宽度 |
| `SolidFill` | `a:solidFill` | `*XPresetFill` | 填充 |

---

### XPresetFill

预设填充，对应 XML: `<a:solidFill>` 或 `<a:schemeClr>`

| 字段 | 类型 | 说明 |
|------|------|------|
| `SrgbClr` | `*XSrgbClr` | RGB 颜色 |
| `SchemeClr` | `*XSchemeClr` | 主题颜色 |

---

## 文本属性

### XTextProperties

文本属性，对应 `<a:rPr>`

| 字段 | XML 属性 | 类型 | 说明 |
|------|----------|------|------|
| `FontSize` | `sz` | `int` | 字体大小（百分之一磅） |
| `Bold` | `b` | `bool` | 粗体 |
| `Italic` | `i` | `bool` | 斜体 |
| `Underline` | `u` | `string` | 下划线 |
| `FontFace` | `typeface` | `string` | 字体 |
| `Color` | `solidFill` | `string` | 颜色 |

---

## 表格

### XTableGrid

表格网格，对应 XML: `<a:tblGrid>`

| 字段 | 类型 | 说明 |
|------|------|------|
| `GridCols` | `[]XTableColumn` | 列定义 |

---

### XTableColumn

表格列，对应 XML: `<a:gridCol w="..."/>`

| 字段 | XML 属性 | 类型 | 说明 |
|------|----------|------|------|
| `W` | `w` | `int` | 列宽 |

---

## 其他

### XFillRectProperties

填充矩形属性，对应 XML: `<a:fillRect/>`

### XStretchProperties

拉伸填充属性，对应 XML: `<a:stretch><a:fillRect/></a:stretch>`

| 字段 | 类型 | 说明 |
|------|------|------|
| `FillRect` | `*XFillRectProperties` | 填充矩形 |

### XGraphic

图形，对应 XML: `<a:graphic>`

| 字段 | 类型 | 说明 |
|------|------|------|
| `Table` | `*XTable` | 表格 |

### XNonVisualGraphicFrame

图形框架非视觉属性，对应 `<p:nvGraphicFramePr>`

| 字段 | 类型 | 说明 |
|------|------|------|
| `CNvPr` | `*XNvCxnSpPr` | 通用非视觉属性 |
| `CNvGraphicFramePr` | `*XNvGraphicFramePr` | 图形框架非视觉属性 |

### XNvGraphicFramePr

图形框架非视觉属性，对应 `<p:cNvGraphicFramePr>`

| 字段 | 类型 | 说明 |
|------|------|------|
| `CNvPr` | `*XNvPr` | 非视觉属性 |

### XNvCxnSpPr

连接形状非视觉属性，对应 `<p:cNvCxnSpPr>`

| 字段 | XML 属性 | 类型 |
|------|----------|------|
| `ID` | `id` | `int` |
| `Name` | `name` | `string` |

### XNvPicPr

图片非视觉属性，对应 `<p:cNvPicPr>`

| 字段 | 类型 | 说明 |
|------|------|------|
| `CNvPr` | `*XNvPr` | 非视觉属性 |

### XTextParagraphList

文本段落列表，对应 `<a:lstStyle/>`

### XOutputOptions

输出选项，对应 `<p:outputOptions>`

| 字段 | XML 属性 | 类型 | 说明 |
|------|----------|------|------|
| `UsePrintFml` | `usePrintFml` | `*bool` | 使用打印格式 |
| `CloneLinkedObjs` | `cloneLinkedObjs` | `*bool` | 克隆链接对象 |
