package parts

import (
	"io"
	"sync"

	"github.com/Muprprpr/Go-pptx/opc"
)

// ============================================================================
// SlidePart 类型定义 - 对应 /ppt/slides/slideN.xml
// ============================================================================

type SlidePart struct {
	uri *opc.PackURI

	// 关联的布局 rId
	layoutRId string

	// 关联的母版 rId（用于快速访问）
	masterRId string

	// 幻灯片内容
	spTree *XSpTree // 形状树

	// 幻灯片属性
	colorMap *XClrMap // 颜色映射
	show     *bool   // 是否显示

	// 局部 ID 自增器（每个 Slide 内部独立，非并发安全，
	// 因为单个 Slide 由单 goroutine 负责生成）
	nextShapeID uint32

	// 页面级 Relationship 管理
	relMgr *SlideRelationships

	mu sync.RWMutex
}

// ============================================================================
// ShapeIDAllocator 形状 ID 分配器（单线程使用）
// ============================================================================

// ShapeIDAllocator 形状 ID 分配器（单线程使用）
// 用于管理单个 Slide 或 ShapeTree 内的 ID 分配
type ShapeIDAllocator struct {
	nextID     uint32          // 下一个可分配的 ID
	reservedID uint32          // 保留的起始 ID（通常为 1，spTree 自身使用）
	maxID      uint32          // 最大 ID 限制（0 表示无限制）
}

// ============================================================================
// ShapeIDAllocatorSync 线程安全的 ID 分配器
// ============================================================================

// ShapeIDAllocatorSync 线程安全的形状 ID 分配器
// 用于多 goroutine 环境下的 ID 分配
type ShapeIDAllocatorSync struct {
	nextID     uint32 // 下一个可分配的 ID
	reservedID uint32 // 保留的起始 ID
	maxID      uint32 // 最大 ID 限制
	mu         sync.Mutex
}

// ============================================================================
// SlideLayoutPart 类型定义 - 对应 /ppt/slideLayouts/slideLayoutN.xml
// ============================================================================

// SlideLayoutType 幻灯片布局类型
type SlideLayoutType int

const (
	SlideLayoutBlank       SlideLayoutType = iota // 空白布局
	SlideLayoutTitle                               // 标题布局
	SlideLayoutTitleAndContent                     // 标题和内容布局
	SlideLayoutTwoContent                          // 两栏内容布局
	SlideLayoutComparison                          // 比较布局
	SlideLayoutTitleOnly                          // 仅标题布局
	SlideLayoutBlankVertical                      // 空白垂直布局
	SlideLayoutObject                            // 对象布局
	SlideLayoutPictureAndCaption                  // 图片和标题布局
)

// SlideLayoutPart 对应 /ppt/slideLayouts/slideLayoutN.xml
type SlideLayoutPart struct {
	uri *opc.PackURI

	layoutType SlideLayoutType

	// 关联的母版 rId
	masterRId string

	// 布局内容
	spTree *XSpTree

	mu sync.RWMutex
}

// ============================================================================
// SlideRelationships 页面级 Relationship 管理
// ============================================================================

// SlideRelationships 页面级 Relationship 管理
// 维护本页面引用的图片、图表、布局等 rId 映射
type SlideRelationships struct {
	nextRId      int               // 下一个可分配的 rId（全局递增）
	imageRels    map[string]string // rId -> 目标 URI (图片)
	layoutRId    string            // 布局 rId
	masterRId    string            // 母版 rId
	mediaRels    map[string]string // rId -> 目标 URI (媒体)
	chartRels    map[string]string // rId -> 目标 URI (图表)
	tableRels    map[string]string // rId -> 目标 URI (表格)
}

// ============================================================================
// XML 结构类型定义
// ============================================================================

// SlideLayoutType 关系类型常量
const (
	RelationshipTypeImage        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
	RelationshipTypeMedia        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"
	RelationshipTypeChart        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
	RelationshipTypeSlideLayout  = "http://schemas.openxmlformats.org/presentationml/2006/relationships/slideLayout"
	RelationshipTypeSlideMaster  = "http://schemas.openxmlformats.org/presentationml/2006/relationships/slideMaster"
	RelationshipTypeTable        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
)

// XSlide 幻灯片 XML 结构
type XSlide struct {
	XMLName struct{} `xml:"sld"`

	XmlnsA string `xml:"xmlns:a,attr"`
	XmlnsR string `xml:"xmlns:r,attr"`
	XmlnsP string `xml:"xmlns:p,attr"`

	CSld      *XCSld      `xml:"cSld"`
	ClrMapOvr *XClrMapOvr `xml:"clrMapOvr"`
}

// XCSld 公共幻灯片数据
type XCSld struct {
	SpTree *XSpTree `xml:"spTree"`
}

// XClrMap 颜色映射
type XClrMap struct {
	BG1             string `xml:"bg1,attr,omitempty"`
	T1              string `xml:"t1,attr,omitempty"`
	BG2             string `xml:"bg2,attr,omitempty"`
	T2              string `xml:"t2,attr,omitempty"`
	Accent1         string `xml:"accent1,attr,omitempty"`
	Accent2         string `xml:"accent2,attr,omitempty"`
	Accent3         string `xml:"accent3,attr,omitempty"`
	Accent4         string `xml:"accent4,attr,omitempty"`
	Accent5         string `xml:"accent5,attr,omitempty"`
	Accent6         string `xml:"accent6,attr,omitempty"`
	HLink           string `xml:"hlink,attr,omitempty"`
	HLink1          string `xml:"hlink1,attr,omitempty"`
	HLink2          string `xml:"hlink2,attr,omitempty"`
 FollClr         string `xml:"follClr,attr,omitempty"`
	LastClr         string `xml:"lastClr,attr,omitempty"`
}

// XClrMapOvr 颜色映射覆盖
type XClrMapOvr struct {
	Accent1 string `xml:"accent1,attr,omitempty"`
}

// XSpTree 形状树
type XSpTree struct {
	XMLName struct{} `xml:"spTree"`

	NonVisual               nvGrpSpPr                 `xml:"nvGrpSpPr"`
	GroupShapeProperties    *XGroupShapeProperties    `xml:"grpSpPr,omitempty"`
	Children                []any                     `xml:"-"`
}

// nvGrpSpPr 非视觉组属性
type nvGrpSpPr struct {
	CNvPr      *XNvCxnSpPr `xml:"cNvPr"`
	CNvGrpSpPr *XNvGrpSpPr `xml:"cNvGrpSpPr"`
}

// XNvCxnSpPr 连接形状非视觉属性
type XNvCxnSpPr struct {
	ID   int    `xml:"id,attr"`
	Name string `xml:"name,attr,omitempty"`
}

// XNvGrpSpPr 组形状非视觉属性
type XNvGrpSpPr struct {
	CNvPr *XNvPr `xml:"cNvPr,omitempty"`
}

// XNvPr 非视觉属性
type XNvPr struct {
	ID   int    `xml:"id,attr,omitempty"`
	Name string `xml:"name,attr,omitempty"`
}

// XGroupShapeProperties 组形状属性
type XGroupShapeProperties struct {
	Xfrm *XTransform2D `xml:"xfrm,omitempty"`
}

// XTransform2D 二维变换
type XTransform2D struct {
	Offset    *XOv2DrOffset `xml:"off,omitempty"`
	Extent    *XOv2DrExtent `xml:"ext,omitempty"`
	Rotation  int           `xml:"rot,attr,omitempty"`
	FlipH     bool          `xml:"flipH,attr,omitempty"`
	FlipV     bool          `xml:"flipV,attr,omitempty"`
}

// XOv2DrOffset 偏移量
type XOv2DrOffset struct {
	X int `xml:"x,attr"`
	Y int `xml:"y,attr"`
}

// XOv2DrExtent 扩展尺寸
type XOv2DrExtent struct {
	Cx int `xml:"cx,attr"`
	Cy int `xml:"cy,attr"`
}

// XShapeProperties 形状属性
type XShapeProperties struct {
	XMLName          struct{}          `xml:"spPr"`
	Transform2D      *XTransform2D     `xml:"xfrm,omitempty"`
	PresetFill       *XPresetFill      `xml:"noFill|a:solidFill|a:gradFill|a:blipFill|a:pattFill|a:grpFill,omitempty"`
	Line             *XLineProperties   `xml:"ln,omitempty"`
}

// XPresetFill 预设填充
type XPresetFill struct {
	SrgbClr   *XSrgbClr   `xml:"srgbClr,omitempty"`
	SchemeClr *XSchemeClr `xml:"schemeClr,omitempty"`
}

// XSrgbClr RGB 颜色
type XSrgbClr struct {
	Val string `xml:"val,attr"`
}

// XSchemeClr 主题颜色
type XSchemeClr struct {
	Val string `xml:"val,attr"`
}

// XLineProperties 线条属性
type XLineProperties struct {
	Width     int           `xml:"w,attr,omitempty"`
	SolidFill *XPresetFill  `xml:"solidFill,omitempty"`
}

// XSp 形状
type XSp struct {
	XMLName struct{} `xml:"sp"`

	NonVisual        XNonVisualDrawingShape `xml:"nvSpPr"`
	ShapeProperties  *XShapeProperties     `xml:"spPr,omitempty"`
	ShapePreset      string                `xml:"-"`
	TextBody         *XTextBody            `xml:"txBody,omitempty"`
}

// XNonVisualDrawingShape 形状非视觉绘图属性
type XNonVisualDrawingShape struct {
	CNvPr  *XNvCxnSpPr `xml:"cNvPr"`
	CNvSpPr *XNvSpPr  `xml:"cNvSpPr"`
}

// XNvSpPr 形状非视觉属性
type XNvSpPr struct {
	CNvPr *XNvPr `xml:"cNvPr,omitempty"`
}

// XTextBody 文本主体
type XTextBody struct {
	XMLName   struct{}             `xml:"txBody"`
	BodyPr    *XBodyPr             `xml:"bodyPr,omitempty"`
	LstStyle  *XTextParagraphList `xml:"lstStyle,omitempty"`
	Paragraphs []XTextParagraph    `xml:"p"`
}

// XBodyPr 主体属性
type XBodyPr struct {
	Wrap      string `xml:"wrap,attr,omitempty"`
	Rotation  int    `xml:"rot,attr,omitempty"`
	Vertical  string `xml:"vert,attr,omitempty"`
	Anchor    string `xml:"anchor,attr,omitempty"`
	AnchorCtr bool   `xml:"anchorCtr,attr,omitempty"`
}

// XTextParagraphList 文本段落列表
type XTextParagraphList struct {
}

// XTextParagraph 文本段落
type XTextParagraph struct {
	Level      int        `xml:"lvl,attr,omitempty"`
	Indent     int        `xml:"indent,attr,omitempty"`
	Alignment  string     `xml:"algn,attr,omitempty"`
	TextRuns   []XTextRun `xml:"r"`
}

// XTextRun 文本片段
type XTextRun struct {
	Text            string            `xml:"t,omitempty"`
	TextProperties  *XTextProperties  `xml:"rPr,omitempty"`
}

// XTextProperties 文本属性
type XTextProperties struct {
	FontSize  int    `xml:"sz,attr,omitempty"`
	Bold      bool   `xml:"b,attr,omitempty"`
	Italic    bool   `xml:"i,attr,omitempty"`
	Underline string `xml:"u,attr,omitempty"`
	FontFace  string `xml:"typeface,attr,omitempty"`
	Color     string `xml:"solidFill,omitempty"`
}

// XPicture 图片
type XPicture struct {
	XMLName struct{} `xml:"pic"`

	NonVisual        XNonVisualDrawingPic `xml:"nvPicPr"`
	BlipFill        *XBlipFillProperties `xml:"blipFill,omitempty"`
	ShapeProperties *XShapeProperties    `xml:"spPr,omitempty"`
}

// XNonVisualDrawingPic 图片非视觉绘图属性
type XNonVisualDrawingPic struct {
	CNvPr   *XNvCxnSpPr `xml:"cNvPr"`
	CNvPicPr *XNvPicPr `xml:"cNvPicPr"`
}

// XNvPicPr 图片非视觉属性
type XNvPicPr struct {
	CNvPr *XNvPr `xml:"cNvPr,omitempty"`
}

// XBlipFillProperties 图片填充属性
type XBlipFillProperties struct {
	XMLName struct{}   `xml:"blipFill"`
	Blip    *XBlip             `xml:"blip,omitempty"`
	Stretch *XStretchProperties `xml:"stretch,omitempty"`
}

// XBlip 图片
type XBlip struct {
	Embed string `xml:"rembed,attr,omitempty"`
}

// XFillProperties 填充属性
type XFillProperties struct {
	XMLName struct{} `xml:"fillRect"`
}

// XStretchProperties 拉伸填充属性
type XStretchProperties struct {
	XMLName struct{} `xml:"stretch"`
	FillRect *XFillRectProperties `xml:"fillRect"`
}

// XFillRectProperties 填充矩形属性
type XFillRectProperties struct {
}

// XGraphicFrame 图形框架
type XGraphicFrame struct {
	XMLName struct{} `xml:"graphicFrame"`

	NonVisual   XNonVisualGraphicFrame  `xml:"nvGraphicFramePr"`
	Graphic     *XGraphic              `xml:"graphic,omitempty"`
	Transform2D *XTransform2D          `xml:"xfrm,omitempty"`
}

// XNonVisualGraphicFrame 图形框架非视觉属性
type XNonVisualGraphicFrame struct {
	CNvPr              *XNvCxnSpPr        `xml:"cNvPr"`
	CNvGraphicFramePr  *XNvGraphicFramePr `xml:"cNvGraphicFramePr"`
}

// XNvGraphicFramePr 图形框架非视觉属性
type XNvGraphicFramePr struct {
	CNvPr *XNvPr `xml:"cNvPr,omitempty"`
}

// XGraphic 图形
type XGraphic struct {
	Table *XTable `xml:"graphicData>a:tbl,omitempty"`
}

// XTableGrid 表格网格
type XTableGrid struct {
	XMLName struct{}       `xml:"tblGrid"`
	GridCols []XTableColumn `xml:"gridCol"`
}

// XTable 表格
type XTable struct {
	XMLName struct{}    `xml:"tbl"`
	Grid    *XTableGrid `xml:"tblGrid,omitempty"`
	Rows    []XTableRow `xml:"tr"`
}

// XTableColumn 表格列
type XTableColumn struct {
	W int `xml:"w,attr"`
}

// XTableRow 表格行
type XTableRow struct {
	GridSpan int           `xml:"gridSpan,attr,omitempty"`
	Cells    []XTableCell `xml:"tc"`
}

// XTableCell 表格单元格
type XTableCell struct {
	XMLName   struct{}    `xml:"tc"`
	GridSpan  int         `xml:"gridSpan,attr,omitempty"`
	RowSpan   int         `xml:"rowSpan,attr,omitempty"`
	Vertical  string      `xml:"anchor,attr,omitempty"`
	TextBody  *XTextBody  `xml:"txBody,omitempty"`
}

// XRelation 关系
type XRelation struct {
	ID     string `xml:"Id,attr"`
	Type   string `xml:"Type,attr"`
	Target string `xml:"Target,attr"`
}

// XSlideRelationships 幻灯片关系
type XSlideRelationships struct {
	XMLName struct{}     `xml:"Relationships"`
	Xmlns   string       `xml:"xmlns,attr"`
	Rels    []XRelation  `xml:"Relationship"`
}

// ============================================================================
// XMLWriter 类型定义
// ============================================================================

// XMLWriter XML 写入辅助结构，提供高效的流式 XML 生成
// 支持命名空间前缀缓存、属性缓冲和缩进控制
type XMLWriter struct {
	w          io.Writer
	buf        []byte
	indent     int
	indentStr  string
	useIndent  bool
	autoFlush  bool
	nsPrefixes map[string]string // 命名空间前缀缓存
}

// XMLWriterPool XMLWriter 对象池，用于减少内存分配
type XMLWriterPool struct {
	pool sync.Pool
}
