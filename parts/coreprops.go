package parts

import (
	"encoding/xml"
)

// ============================================================================
// Core Properties XML 结构体 - 对应 /docProps/core.xml
// ============================================================================
//
// OpenXML 核心属性基于 Dublin Core 元数据标准
// 命名空间:
//   - cp:  http://schemas.openxmlformats.org/package/2006/metadata/core-properties
//   - dc:  http://purl.org/dc/elements/1.1/
//   - dcterms: http://purl.org/dc/terms/
//   - xsi: http://www.w3.org/2001/XMLSchema-instance
// ============================================================================

// XMLCoreProperties 核心属性 XML 结构体
// 对应 XML: /docProps/core.xml
// 包含文档的元数据信息（标题、作者、创建/修改时间等）
type XMLCoreProperties struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/package/2006/metadata/core-properties coreProperties"`

	// 命名空间声明（用于序列化）
	XmlnsCp      string `xml:"xmlns:cp,attr,omitempty"`
	XmlnsDc      string `xml:"xmlns:dc,attr,omitempty"`
	XmlnsDcterms string `xml:"xmlns:dcterms,attr,omitempty"`
	XmlnsXsi     string `xml:"xmlns:xsi,attr,omitempty"`

	// Dublin Core 元素 (dc: 命名空间 -> http://purl.org/dc/elements/1.1/)
	Title       string `xml:"http://purl.org/dc/elements/1.1/ title,omitempty"`
	Creator     string `xml:"http://purl.org/dc/elements/1.1/ creator,omitempty"`
	Subject     string `xml:"http://purl.org/dc/elements/1.1/ subject,omitempty"`
	Description string `xml:"http://purl.org/dc/elements/1.1/ description,omitempty"`

	// Dublin Core Terms 元素 (dcterms: 命名空间 -> http://purl.org/dc/terms/)
	Created  *XMLW3CDTFDate `xml:"http://purl.org/dc/terms/ created,omitempty"`
	Modified *XMLW3CDTFDate `xml:"http://purl.org/dc/terms/ modified,omitempty"`

	// 核心属性扩展 (cp: 命名空间 -> http://schemas.openxmlformats.org/package/2006/metadata/core-properties)
	Keywords       string `xml:"http://schemas.openxmlformats.org/package/2006/metadata/core-properties keywords,omitempty"`
	LastModifiedBy string `xml:"http://schemas.openxmlformats.org/package/2006/metadata/core-properties lastModifiedBy,omitempty"`
	Revision       string `xml:"http://schemas.openxmlformats.org/package/2006/metadata/core-properties revision,omitempty"`
	Category       string `xml:"http://schemas.openxmlformats.org/package/2006/metadata/core-properties category,omitempty"`
	ContentType    string `xml:"http://schemas.openxmlformats.org/package/2006/metadata/core-properties contentType,omitempty"`
	Version        string `xml:"http://schemas.openxmlformats.org/package/2006/metadata/core-properties version,omitempty"`
	Identifier     string `xml:"http://schemas.openxmlformats.org/package/2006/metadata/core-properties identifier,omitempty"`
	Language       string `xml:"http://purl.org/dc/elements/1.1/ language,omitempty"`
}

// XMLW3CDTFDate W3CDTF 格式日期元素
// 对应 XML: <dcterms:created xsi:type="dcterms:W3CDTF">...</dcterms:created>
// W3CDTF 格式: YYYY-MM-DDThh:mm:ssZ
type XMLW3CDTFDate struct {
	Type  string `xml:"xsi:type,attr,omitempty"`
	Value string `xml:",chardata"`
}

// ============================================================================
// 常量定义
// ============================================================================

const (
	// 命名空间常量
	NamespaceCoreProperties = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
	NamespaceDublinCore     = "http://purl.org/dc/elements/1.1/"
	NamespaceDublinCoreTerms = "http://purl.org/dc/terms/"
	NamespaceXMLSchema      = "http://www.w3.org/2001/XMLSchema-instance"

	// W3CDTF 类型标识
	W3CDTFType = "dcterms:W3CDTF"
)

// XMLDeclaration OPC 包中所有 XML 文件的标准声明头
// 注意：此常量已在 xmlutils.go 中定义

// ============================================================================
// 构造函数
// ============================================================================

// NewXMLCoreProperties 创建带默认命名空间的核心属性结构体
func NewXMLCoreProperties() *XMLCoreProperties {
	return &XMLCoreProperties{
		XmlnsCp:      NamespaceCoreProperties,
		XmlnsDc:      NamespaceDublinCore,
		XmlnsDcterms: NamespaceDublinCoreTerms,
		XmlnsXsi:     NamespaceXMLSchema,
	}
}

// ============================================================================
// 辅助方法
// ============================================================================

// SetCreated 设置创建时间
func (cp *XMLCoreProperties) SetCreated(value string) {
	cp.Created = &XMLW3CDTFDate{
		Type:  W3CDTFType,
		Value: value,
	}
}

// SetModified 设置修改时间
func (cp *XMLCoreProperties) SetModified(value string) {
	cp.Modified = &XMLW3CDTFDate{
		Type:  W3CDTFType,
		Value: value,
	}
}

// GetCreated 获取创建时间值
func (cp *XMLCoreProperties) GetCreated() string {
	if cp.Created == nil {
		return ""
	}
	return cp.Created.Value
}

// GetModified 获取修改时间值
func (cp *XMLCoreProperties) GetModified() string {
	if cp.Modified == nil {
		return ""
	}
	return cp.Modified.Value
}

// ToXML 将核心属性序列化为 XML 字节
func (cp *XMLCoreProperties) ToXML() ([]byte, error) {
	output, err := xml.MarshalIndent(cp, "", "  ")
	if err != nil {
		return nil, err
	}
	return append([]byte(XMLDeclaration), output...), nil
}

// ParseCoreProperties 从 XML 字节解析核心属性
func ParseCoreProperties(data []byte) (*XMLCoreProperties, error) {
	var cp XMLCoreProperties
	if err := xml.Unmarshal(data, &cp); err != nil {
		return nil, err
	}
	return &cp, nil
}

// ParseCoreProps 是 ParseCoreProperties 的简写别名
// 提供更简洁的调用方式
func ParseCoreProps(data []byte) (*XMLCoreProperties, error) {
	return ParseCoreProperties(data)
}
