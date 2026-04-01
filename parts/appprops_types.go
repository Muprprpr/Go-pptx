package parts

// ============================================================================
// App Properties XML 结构类型定义 - 对应 /docProps/app.xml
// ============================================================================
//
// 应用程序属性基于 OpenXML 规范
// 命名空间: http://schemas.openxmlformats.org/officeDocument/2006/extended-properties
// 文件位置: /docProps/app.xml
//
// ============================================================================

import "encoding/xml"

// XMLAppProps 应用程序属性 XML 结构体
type XMLAppProps struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties Properties"`

	// 命名空间声明（用于序列化）
	XmlnsProp string `xml:"xmlns,attr,omitempty"`
	XmlnsVt   string `xml:"xmlns:vt,attr,omitempty"`

	// 应用程序信息
	Application string `xml:"Application,omitempty"` // 应用程序名称
	AppVersion string `xml:"AppVersion,omitempty"`  // 应用程序版本

	Security    string `xml:"DocSecurity,omitempty"`  // 文档安全级别

	// 文档统计信息
	TotalTime         *int   `xml:"TotalTime,omitempty"`   // 总编辑时间（分钟）
	Words             *int   `xml:"Words,omitempty"`      // 字数
	Characters        *int   `xml:"Characters,omitempty"` // 字符数
	Pages            *int   `xml:"Pages,omitempty"`       // 页数
	Paragraphs        *int   `xml:"Paragraphs,omitempty"` // 段落数
	Slides           *int   `xml:"Slides,omitempty"`      // 幻灯片数
	Notes            *int   `xml:"Notes,omitempty"`       // 备注数
	HiddenSlides     *int   `xml:"HiddenSlides,omitempty"` // 隐藏幻灯片数
	MMClips          *int   `xml:"MMClips,omitempty"`      // 多媒体剪辑数
	ScaleCrop        *bool  `xml:"ScaleCrop,omitempty"`    // 是否按比例裁剪

	// 组织信息
	Company          string `xml:"Company,omitempty"`   // 公司名称
	Manager          string `xml:"Manager,omitempty"`   // 管理者

	// 链接信息
	HyperlinkBase     string `xml:"HyperlinkBase,omitempty"`     // 超链接基础
	LinksUpToDate     *bool  `xml:"LinksUpToDate,omitempty"`     // 链接是否最新
	HyperlinksChanged *bool  `xml:"HyperlinksChanged,omitempty"` // 超链接是否更改
	SharedDoc         *bool  `xml:"SharedDoc,omitempty"`         // 是否共享文档

	// 标题对和部件标题（使用 InnerXML 保留原始结构）
	HeadingPairs     *XMLHeadingPairs `xml:"HeadingPairs,omitempty"`
	TitlesOfParts    *XMLTitlesOfParts  `xml:"TitlesOfParts,omitempty"`

	// 模板信息
	Template         string `xml:"Template,omitempty"` // 模板

	// 其他属性
	PresentationFormat string `xml:"PresentationFormat,omitempty"` // 演示文稿格式
	LineSketches      *bool  `xml:"LineSketches,omitempty"`      // 线条草图
}

// XMLHeadingPairs 标题对
// 注意：此结构较为复杂，通常使用 InnerXML 保留原始数据
// 命名空间：与父元素 Properties 相同（extended-properties）
type XMLHeadingPairs struct {
	XMLName  xml.Name `xml:"HeadingPairs"`
	InnerXML string   `xml:",innerxml"` // 保留原始 XML 内容
}

// XMLTitlesOfParts 部件标题
// 命名空间：与父元素 Properties 相同（extended-properties）
type XMLTitlesOfParts struct {
	XMLName  xml.Name `xml:"TitlesOfParts"`
	InnerXML string   `xml:",innerxml"` // 保留原始 XML 内容
}

// ============================================================================
// 娡板相关常量
// ============================================================================

// DefaultAppProps 默认应用属性模板
var DefaultAppProps = &XMLAppProps{
	Application:  "Microsoft Office PowerPoint",
	AppVersion: "15.0000",
	Company:     "",
	Manager:     "",
}

// ============================================================================
// 命名空间常量
// ============================================================================

const (
	// NamespaceExtendedProperties 扩展属性命名空间
	NamespaceExtendedProperties = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
	// NamespaceDocPropsVTypes 文档属性类型命名空间
	NamespaceDocPropsVTypes = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
)
