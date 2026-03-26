package parts

import "encoding/xml"

// ============================================================================
// OpenXML 基础结构体 - 用于 encoding/xml 反序列化
// ============================================================================
//
// 本文件包含母版/版式解析所需的最底层 XML 结构体定义
// 遵循自底向上的原则，这些结构体可被上层结构体组合使用
// ============================================================================

// XMLOffset 偏移量结构体
// 对应 XML: <a:off x="..." y="..."/>
// 用于表示形状或对象的 X/Y 坐标位置（EMU 单位）
type XMLOffset struct {
	X int64 `xml:"x,attr"`
	Y int64 `xml:"y,attr"`
}

// IsZero 检查是否为零值（未设置或缺失属性）
func (o *XMLOffset) IsZero() bool {
	return o.X == 0 && o.Y == 0
}

// IsValid 检查是否有效（OpenXML 规范要求 x 和 y 属性必须存在）
// 注意：零值坐标 (0, 0) 在技术上是有效的位置
func (o *XMLOffset) IsValid() bool {
	return o != nil
}

// XMLExtents 尺寸结构体
// 对应 XML: <a:ext cx="..." cy="..."/>
// 用于表示形状或对象的宽度/高度（EMU 单位）
type XMLExtents struct {
	Cx int64 `xml:"cx,attr"`
	Cy int64 `xml:"cy,attr"`
}

// IsZero 检查是否为零值（未设置或缺失属性）
func (e *XMLExtents) IsZero() bool {
	return e.Cx == 0 && e.Cy == 0
}

// IsValid 检查是否有效（OpenXML 规范要求 cx 和 cy 必须为正数）
// 零值或负值尺寸通常表示无效的形状
func (e *XMLExtents) IsValid() bool {
	return e != nil && e.Cx > 0 && e.Cy > 0
}

// XMLTransform 二维变换结构体
// 对应 XML: <a:xfrm>...</a:xfrm>
// 包含位置偏移和尺寸信息
// 注意：Go 的 xml 包会自动处理命名空间，所以使用本地名称即可
type XMLTransform struct {
	Off *XMLOffset  `xml:"off,omitempty"`
	Ext *XMLExtents `xml:"ext,omitempty"`
}

// XMLPlaceholder 占位符结构体
// 对应 XML: <p:ph type="..." idx="..."/>
// 用于标记母版/版式中的可填充区域类型
type XMLPlaceholder struct {
	Type string `xml:"type,attr,omitempty"`
	Idx  string `xml:"idx,attr,omitempty"`
}

// ============================================================================
// 中间包装层结构体
// ============================================================================

// XMLCNvPr 通用非视觉属性
// 对应 XML: <p:cNvPr id="..." name="..."/>
type XMLCNvPr struct {
	ID   int    `xml:"id,attr"`
	Name string `xml:"name,attr,omitempty"`
}

// XMLNvPr 非视觉属性
// 对应 XML: <p:nvPr>...</p:nvPr>
// 包含占位符定义（若存在）
type XMLNvPr struct {
	Ph *XMLPlaceholder `xml:"ph,omitempty"`
}

// XMLNvSpPr 非视觉形状属性
// 对应 XML: <p:nvSpPr>...</p:nvSpPr>
// 包含通用属性和非视觉属性
type XMLNvSpPr struct {
	CNvPr *XMLCNvPr `xml:"cNvPr,omitempty"`
	NvPr  *XMLNvPr  `xml:"nvPr,omitempty"`
}

// XMLSpPr 视觉形状属性
// 对应 XML: <p:spPr>...</p:spPr>
// 包含变换信息（位置和尺寸）
type XMLSpPr struct {
	Xfrm *XMLTransform `xml:"xfrm,omitempty"`
}

// XMLBackground 背景结构体
// 对应 XML: <p:bg>...</p:bg>
// 简单结构，包含背景属性
type XMLBackground struct {
	BgPr  *XMLBackgroundPr `xml:"bgPr,omitempty"`
	BgRef *XMLBackgroundRef `xml:"bgRef,omitempty"`
}

// XMLBackgroundRef 背景引用
// 对应 XML: <p:bgRef idx="..."><a:schemeClr val="..."/></p:bgRef>
type XMLBackgroundRef struct {
	Idx  string `xml:"idx,attr,omitempty"`
	Clr  *XMLSchemeColor `xml:"schemeClr,omitempty"`
}

// XMLBackgroundPr 背景属性
// 对应 XML: <p:bgPr>...</p:bgPr>
type XMLBackgroundPr struct {
	Fill *XMLFillProperties `xml:",any,omitempty"`
}

// XMLFillProperties 填充属性（联合类型）
// 对应 XML: <a:solidFill> / <a:gradFill> / <a:blipFill> 等
type XMLFillProperties struct {
	SolidFill *XMLSolidFill `xml:"a:solidFill,omitempty"`
	GradFill  *XMLGradFill  `xml:"a:gradFill,omitempty"`
	BlipFill  *XMLBlipFill  `xml:"a:blipFill,omitempty"`
	NoFill    *struct{}     `xml:"a:noFill,omitempty"`
}

// XMLSolidFill 纯色填充
// 对应 XML: <a:solidFill>...</a:solidFill>
type XMLSolidFill struct {
	SrgbClr   *XMLSRgbColor   `xml:"a:srgbClr,omitempty"`
	SchemeClr *XMLSchemeColor `xml:"a:schemeClr,omitempty"`
}

// XMLSRgbColor RGB 颜色
// 对应 XML: <a:srgbClr val="..."/>
type XMLSRgbColor struct {
	Val string `xml:"val,attr,omitempty"`
}

// XMLSchemeColor 主题颜色
// 对应 XML: <a:schemeClr val="..."/>
type XMLSchemeColor struct {
	Val string `xml:"val,attr,omitempty"`
}

// XMLGradFill 渐变填充
// 对应 XML: <a:gradFill>...</a:gradFill>
type XMLGradFill struct {
	GsLst *XMLGradientStopList `xml:"a:gsLst,omitempty"`
	Lin   *XMLLinearGradient   `xml:"a:lin,omitempty"`
}

// XMLGradientStopList 渐变色标列表
// 对应 XML: <a:gsLst>...</a:gsLst>
type XMLGradientStopList struct {
	Stops []XMLGradientStop `xml:"a:gs,omitempty"`
}

// XMLGradientStop 渐变色标
// 对应 XML: <a:gs pos="...">...</a:gs>
type XMLGradientStop struct {
	Pos       int64         `xml:"pos,attr,omitempty"`
	SolidFill *XMLSolidFill `xml:"a:solidFill,omitempty"`
}

// XMLLinearGradient 线性渐变
// 对应 XML: <a:lin ang="..." scaled="..."/>
type XMLLinearGradient struct {
	Ang    int64 `xml:"ang,attr,omitempty"`
	Scaled bool  `xml:"scaled,attr,omitempty"`
}

// XMLBlipFill 图片填充
// 对应 XML: <a:blipFill>...</a:blipFill>
type XMLBlipFill struct {
	Blip *XMLBlip `xml:"a:blip,omitempty"`
}

// XMLBlip 图片引用
// 对应 XML: <a:blip r:embed="..."/>
type XMLBlip struct {
	Embed string `xml:"r:embed,attr,omitempty"`
}

// ============================================================================
// 顶层结构体
// ============================================================================

// XMLShape 形状结构体
// 对应 XML: <p:sp>...</p:sp>
// 表示幻灯片中的单个形状元素
type XMLShape struct {
	NvSpPr *XMLNvSpPr `xml:"nvSpPr"`
	SpPr   *XMLSpPr   `xml:"spPr"`
}

// XMLShapeTree 形状树结构体
// 对应 XML: <p:spTree>...</p:spTree>
// 包含多个形状元素的集合
type XMLShapeTree struct {
	NvGrpSpPr   *XMLNvGrpSpPr   `xml:"nvGrpSpPr,omitempty"`
	GrpSpPr     *XMLGrpSpPr     `xml:"grpSpPr,omitempty"`
	Shapes      []XMLShape      `xml:"sp"`
	GroupShapes []XMLGroupShape `xml:"grpSp,omitempty"`
}

// XMLNvGrpSpPr 非视觉组属性
// 对应 XML: <p:nvGrpSpPr>...</p:nvGrpSpPr>
type XMLNvGrpSpPr struct {
	CNvPr      *XMLCNvPr      `xml:"cNvPr,omitempty"`
	CNvGrpSpPr *XMLCNvGrpSpPr `xml:"cNvGrpSpPr,omitempty"`
}

// XMLCNvGrpSpPr 组形状非视觉属性
// 对应 XML: <p:cNvGrpSpPr>...</p:cNvGrpSpPr>
type XMLCNvGrpSpPr struct {
	// 通常为空元素，预留扩展
}

// XMLGrpSpPr 组形状属性
// 对应 XML: <p:grpSpPr>...</p:grpSpPr>
type XMLGrpSpPr struct {
	Xfrm *XMLTransform `xml:"xfrm,omitempty"`
}

// XMLGroupShape 组形状
// 对应 XML: <p:grpSp>...</p:grpSp>
type XMLGroupShape struct {
	NvGrpSpPr *XMLNvGrpSpPr `xml:"nvGrpSpPr"`
	GrpSpPr   *XMLGrpSpPr   `xml:"grpSpPr"`
	Shapes    []XMLShape    `xml:"sp"`
}

// XMLCommonSlideData 通用幻灯片数据
// 对应 XML: <p:cSld>...</p:cSld>
// 包含背景和形状树
type XMLCommonSlideData struct {
	Bg      *XMLBackground  `xml:"bg,omitempty"`
	SpTree  *XMLShapeTree   `xml:"spTree"`
}

// XMLSlideLayout 幻灯片版式
// 对应 XML: <p:sldLayout>...</p:sldLayout>
// 根节点结构，定义单个版式
type XMLSlideLayout struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/presentationml/2006/main sldLayout"`

	// 命名空间声明
	XmlnsA string `xml:"xmlns:a,attr,omitempty"`
	XmlnsR string `xml:"xmlns:r,attr,omitempty"`
	XmlnsP string `xml:"xmlns:p,attr,omitempty"`

	CSld *XMLCommonSlideData `xml:"cSld"`
}

// XMLSlideMaster 幻灯片母版
// 对应 XML: <p:sldMaster>...</p:sldMaster>
// 根节点结构，定义母版
type XMLSlideMaster struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/presentationml/2006/main sldMaster"`

	// 命名空间声明
	XmlnsA string `xml:"xmlns:a,attr,omitempty"`
	XmlnsR string `xml:"xmlns:r,attr,omitempty"`
	XmlnsP string `xml:"xmlns:p,attr,omitempty"`

	CSld *XMLCommonSlideData `xml:"cSld"`
}
