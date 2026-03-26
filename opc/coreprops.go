package opc

import (
	"encoding/xml"
	"fmt"
	"sync"
)

// CoreProperties 表示包的核心属性（Dublin Core 元数据）
type CoreProperties struct {
	title          string // 标题
	creator        string // 创建者
	subject        string // 主题
	description    string // 描述
	keywords       string // 关键词
	created        string // 创建时间
	modified       string // 修改时间
	lastModifiedBy string // 最后修改者
	revision       string // 版本号
	category       string // 类别
	contentType    string // 内容类型
	language       string // 语言
	mu             sync.RWMutex
}

// --- Getter 方法 ---

// Title 返回标题
func (cp *CoreProperties) Title() string {
	cp.mu.RLock()
	defer cp.mu.RUnlock()
	return cp.title
}

// Creator 返回创建者
func (cp *CoreProperties) Creator() string {
	cp.mu.RLock()
	defer cp.mu.RUnlock()
	return cp.creator
}

// Subject 返回主题
func (cp *CoreProperties) Subject() string {
	cp.mu.RLock()
	defer cp.mu.RUnlock()
	return cp.subject
}

// Description 返回描述
func (cp *CoreProperties) Description() string {
	cp.mu.RLock()
	defer cp.mu.RUnlock()
	return cp.description
}

// Keywords 返回关键词
func (cp *CoreProperties) Keywords() string {
	cp.mu.RLock()
	defer cp.mu.RUnlock()
	return cp.keywords
}

// Created 返回创建时间
func (cp *CoreProperties) Created() string {
	cp.mu.RLock()
	defer cp.mu.RUnlock()
	return cp.created
}

// Modified 返回修改时间
func (cp *CoreProperties) Modified() string {
	cp.mu.RLock()
	defer cp.mu.RUnlock()
	return cp.modified
}

// LastModifiedBy 返回最后修改者
func (cp *CoreProperties) LastModifiedBy() string {
	cp.mu.RLock()
	defer cp.mu.RUnlock()
	return cp.lastModifiedBy
}

// Revision 返回版本号
func (cp *CoreProperties) Revision() string {
	cp.mu.RLock()
	defer cp.mu.RUnlock()
	return cp.revision
}

// Category 返回类别
func (cp *CoreProperties) Category() string {
	cp.mu.RLock()
	defer cp.mu.RUnlock()
	return cp.category
}

// ContentType 返回内容类型
func (cp *CoreProperties) ContentType() string {
	cp.mu.RLock()
	defer cp.mu.RUnlock()
	return cp.contentType
}

// Language 返回语言
func (cp *CoreProperties) Language() string {
	cp.mu.RLock()
	defer cp.mu.RUnlock()
	return cp.language
}

// --- Setter 方法 ---

// SetTitle 设置标题
func (cp *CoreProperties) SetTitle(title string) {
	cp.mu.Lock()
	defer cp.mu.Unlock()
	cp.title = title
}

// SetCreator 设置创建者
func (cp *CoreProperties) SetCreator(creator string) {
	cp.mu.Lock()
	defer cp.mu.Unlock()
	cp.creator = creator
}

// SetSubject 设置主题
func (cp *CoreProperties) SetSubject(subject string) {
	cp.mu.Lock()
	defer cp.mu.Unlock()
	cp.subject = subject
}

// SetDescription 设置描述
func (cp *CoreProperties) SetDescription(description string) {
	cp.mu.Lock()
	defer cp.mu.Unlock()
	cp.description = description
}

// SetKeywords 设置关键词
func (cp *CoreProperties) SetKeywords(keywords string) {
	cp.mu.Lock()
	defer cp.mu.Unlock()
	cp.keywords = keywords
}

// SetCreated 设置创建时间
func (cp *CoreProperties) SetCreated(created string) {
	cp.mu.Lock()
	defer cp.mu.Unlock()
	cp.created = created
}

// SetModified 设置修改时间
func (cp *CoreProperties) SetModified(modified string) {
	cp.mu.Lock()
	defer cp.mu.Unlock()
	cp.modified = modified
}

// SetLastModifiedBy 设置最后修改者
func (cp *CoreProperties) SetLastModifiedBy(lastModifiedBy string) {
	cp.mu.Lock()
	defer cp.mu.Unlock()
	cp.lastModifiedBy = lastModifiedBy
}

// SetRevision 设置版本号
func (cp *CoreProperties) SetRevision(revision string) {
	cp.mu.Lock()
	defer cp.mu.Unlock()
	cp.revision = revision
}

// SetCategory 设置类别
func (cp *CoreProperties) SetCategory(category string) {
	cp.mu.Lock()
	defer cp.mu.Unlock()
	cp.category = category
}

// SetContentType 设置内容类型
func (cp *CoreProperties) SetContentType(contentType string) {
	cp.mu.Lock()
	defer cp.mu.Unlock()
	cp.contentType = contentType
}

// SetLanguage 设置语言
func (cp *CoreProperties) SetLanguage(language string) {
	cp.mu.Lock()
	defer cp.mu.Unlock()
	cp.language = language
}

// ===== XML 序列化 =====

// XCoreProperties XML 序列化的核心属性
type XCoreProperties struct {
	XMLName        xml.Name    `xml:"coreProperties"`
	XmlnsDc        string      `xml:"xmlns:dc,attr"`
	XmlnsDcterms   string      `xml:"xmlns:dcterms,attr"`
	XmlnsDcmitype  string      `xml:"xmlns:dcmitype,attr"`
	XmlnsXsi       string      `xml:"xmlns:xsi,attr"`
	XmlnsCore      string      `xml:"xmlns,attr"`
	Title          string      `xml:"dc:title"`
	Creator        string      `xml:"dc:creator"`
	Subject        string      `xml:"dc:subject"`
	Description    string      `xml:"dc:description"`
	Keywords       *XKeywords  `xml:"cp:keywords"`
	Created        *XDate      `xml:"dcterms:created"`
	Modified       *XDate      `xml:"dcterms:modified"`
	LastModifiedBy string      `xml:"cp:lastModifiedBy"`
	Revision       string      `xml:"cp:revision"`
	Category       string      `xml:"cp:category"`
	ContentType    string      `xml:"cp:contentType"`
}

// XKeywords 关键词元素
type XKeywords struct {
	Value string `xml:",chardata"`
}

// XDate 日期元素
type XDate struct {
	Type  string `xml:"xsi:type,attr"`
	Value string `xml:",chardata"`
}

// FromXML 从 XML 解析核心属性
func (cp *CoreProperties) FromXML(data []byte) error {
	cp.mu.Lock()
	defer cp.mu.Unlock()

	var xcp XCoreProperties
	if err := xml.Unmarshal(data, &xcp); err != nil {
		return fmt.Errorf("failed to unmarshal core properties: %w", err)
	}

	cp.title = xcp.Title
	cp.creator = xcp.Creator
	cp.subject = xcp.Subject
	cp.description = xcp.Description
	if xcp.Keywords != nil {
		cp.keywords = xcp.Keywords.Value
	}
	if xcp.Created != nil {
		cp.created = xcp.Created.Value
	}
	if xcp.Modified != nil {
		cp.modified = xcp.Modified.Value
	}
	cp.lastModifiedBy = xcp.LastModifiedBy
	cp.revision = xcp.Revision
	cp.category = xcp.Category
	cp.contentType = xcp.ContentType

	return nil
}

// ToXML 将核心属性序列化为 XML
func (cp *CoreProperties) ToXML() ([]byte, error) {
	cp.mu.RLock()
	defer cp.mu.RUnlock()

	xcp := XCoreProperties{
		XmlnsDc:        "http://purl.org/dc/elements/1.1/",
		XmlnsDcterms:   "http://purl.org/dc/terms/",
		XmlnsDcmitype:  "http://purl.org/dc/dcmitype/",
		XmlnsXsi:       "http://www.w3.org/2001/XMLSchema-instance",
		XmlnsCore:      "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
		Title:          cp.title,
		Creator:        cp.creator,
		Subject:        cp.subject,
		Description:    cp.description,
		LastModifiedBy: cp.lastModifiedBy,
		Revision:       cp.revision,
		Category:       cp.category,
		ContentType:    cp.contentType,
	}

	if cp.keywords != "" {
		xcp.Keywords = &XKeywords{Value: cp.keywords}
	}
	if cp.created != "" {
		xcp.Created = &XDate{Type: "dcterms:W3CDTF", Value: cp.created}
	}
	if cp.modified != "" {
		xcp.Modified = &XDate{Type: "dcterms:W3CDTF", Value: cp.modified}
	}

	output, err := xml.MarshalIndent(xcp, "", "  ")
	if err != nil {
		return nil, fmt.Errorf("failed to marshal core properties: %w", err)
	}

	return append([]byte(XMLDeclaration), output...), nil
}
