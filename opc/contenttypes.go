package opc

import (
	"encoding/xml"
	"fmt"
	"strings"
	"sync"
)

// ContentTypes 表示 [Content_Types].xml 的内容
// 定义包中所有部件的内容类型
type ContentTypes struct {
	defaults  map[string]string // 扩展名 -> 内容类型
	overrides map[string]string // URI -> 内容类型
	mu        sync.RWMutex
}

// NewContentTypes 创建新的内容类型定义
func NewContentTypes() *ContentTypes {
	ct := &ContentTypes{
		defaults:  make(map[string]string),
		overrides: make(map[string]string),
	}
	// 初始化默认内容类型
	for ext, ctType := range DefaultContentTypes {
		ct.defaults[ext] = ctType
	}
	return ct
}

// AddDefault 添加默认内容类型映射
func (ct *ContentTypes) AddDefault(extension, contentType string) {
	ct.mu.Lock()
	defer ct.mu.Unlock()
	ct.defaults[extension] = contentType
}

// AddOverride 添加特定 URI 的内容类型覆盖
func (ct *ContentTypes) AddOverride(uri *PackURI, contentType string) {
	ct.mu.Lock()
	defer ct.mu.Unlock()
	ct.overrides[uri.URI()] = contentType
}

// GetContentType 获取指定 URI 的内容类型
// 优先查找 overrides，然后查找 defaults
func (ct *ContentTypes) GetContentType(uri *PackURI) string {
	ct.mu.RLock()
	defer ct.mu.RUnlock()

	// 先查找 override
	if ct, ok := ct.overrides[uri.URI()]; ok {
		return ct
	}

	// 再查找 default
	ext := uri.Extension()
	if ct, ok := ct.defaults[ext]; ok {
		return ct
	}

	return ContentTypeDefault
}

// GetDefault 获取扩展名对应的默认内容类型
func (ct *ContentTypes) GetDefault(extension string) string {
	ct.mu.RLock()
	defer ct.mu.RUnlock()
	return ct.defaults[extension]
}

// GetOverride 获取 URI 对应的内容类型覆盖
func (ct *ContentTypes) GetOverride(uri *PackURI) string {
	ct.mu.RLock()
	defer ct.mu.RUnlock()
	return ct.overrides[uri.URI()]
}

// RemoveOverride 移除内容类型覆盖
func (ct *ContentTypes) RemoveOverride(uri *PackURI) {
	ct.mu.Lock()
	defer ct.mu.Unlock()
	delete(ct.overrides, uri.URI())
}

// Defaults 返回所有默认内容类型映射
func (ct *ContentTypes) Defaults() map[string]string {
	ct.mu.RLock()
	defer ct.mu.RUnlock()

	result := make(map[string]string, len(ct.defaults))
	for k, v := range ct.defaults {
		result[k] = v
	}
	return result
}

// Overrides 返回所有内容类型覆盖
func (ct *ContentTypes) Overrides() map[string]string {
	ct.mu.RLock()
	defer ct.mu.RUnlock()

	result := make(map[string]string, len(ct.overrides))
	for k, v := range ct.overrides {
		result[k] = v
	}
	return result
}

// ===== XML 序列化 =====

// XContentTypes XML 序列化的内容类型根元素
type XContentTypes struct {
	XMLName   xml.Name     `xml:"Types"`
	Xmlns     string       `xml:"xmlns,attr"`
	Defaults  []XDefault   `xml:"Default"`
	Overrides []XOverride  `xml:"Override"`
}

// XDefault XML 序列化的默认内容类型
type XDefault struct {
	Extension   string `xml:"Extension,attr"`
	ContentType string `xml:"ContentType,attr"`
}

// XOverride XML 序列化的内容类型覆盖
type XOverride struct {
	PartName    string `xml:"PartName,attr"`
	ContentType string `xml:"ContentType,attr"`
}

// FromXML 从 XML 解析内容类型
func (ct *ContentTypes) FromXML(data []byte) error {
	ct.mu.Lock()
	defer ct.mu.Unlock()

	var xct XContentTypes
	if err := xml.Unmarshal(data, &xct); err != nil {
		return fmt.Errorf("failed to unmarshal content types: %w", err)
	}

	ct.defaults = make(map[string]string)
	ct.overrides = make(map[string]string)

	for _, d := range xct.Defaults {
		ct.defaults[d.Extension] = d.ContentType
	}

	for _, o := range xct.Overrides {
		// PartName 通常是绝对路径如 /ppt/presentation.xml
		ct.overrides[o.PartName] = o.ContentType
	}

	return nil
}

// ToXML 将内容类型序列化为 XML
func (ct *ContentTypes) ToXML() ([]byte, error) {
	ct.mu.RLock()
	defer ct.mu.RUnlock()

	xct := XContentTypes{
		Xmlns: NamespaceOPCPackage,
	}

	for ext, ctType := range ct.defaults {
		xct.Defaults = append(xct.Defaults, XDefault{
			Extension:   strings.TrimPrefix(ext, "."),
			ContentType: ctType,
		})
	}

	for uri, ctType := range ct.overrides {
		xct.Overrides = append(xct.Overrides, XOverride{
			PartName:    uri,
			ContentType: ctType,
		})
	}

	output, err := xml.MarshalIndent(xct, "", "  ")
	if err != nil {
		return nil, fmt.Errorf("failed to marshal content types: %w", err)
	}

	return append([]byte(XMLDeclaration), output...), nil
}
