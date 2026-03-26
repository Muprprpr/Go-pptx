package parts

import (
	"encoding/xml"
)

// ============================================================================
// OpenXML Relationships XML 结构体 - 对应 *.rels 文件
// ============================================================================
//
// 关系文件位置示例：
//   - 包级别: /_rels/.rels
//   - 幻灯片: /ppt/slides/_rels/slide1.xml.rels
//   - 母版:   /ppt/slideMasters/_rels/slideMaster1.xml.rels
//
// 命名空间: http://schemas.openxmlformats.org/package/2006/relationships
// ============================================================================

// XMLRelationships 关系集合根节点
// 对应 XML: <Relationships xmlns="...">...</Relationships>
type XMLRelationships struct {
	XMLName       xml.Name          `xml:"Relationships"`
	Xmlns         string            `xml:"xmlns,attr,omitempty"`
	Relationships []XMLRelationship `xml:"Relationship"`
}

// XMLRelationship 单个关系
// 对应 XML: <Relationship Id="rId1" Type="..." Target="..."/>
type XMLRelationship struct {
	ID         string `xml:"Id,attr"`              // 关系 ID（如 rId1, rId2）
	Type       string `xml:"Type,attr"`            // 关系类型 URI
	Target     string `xml:"Target,attr"`          // 目标路径（相对或绝对）
	TargetMode string `xml:"TargetMode,attr,omitempty"` // Internal（默认）或 External
}

// ============================================================================
// 常量定义
// ============================================================================

const (
	// 关系命名空间
	NamespaceRelationships = "http://schemas.openxmlformats.org/package/2006/relationships"

	// 目标模式
	TargetModeInternal = "Internal"
	TargetModeExternal = "External"

	// 常用关系类型
	RelTypeImage       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
	RelTypeHyperlink   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
	RelTypeSlide       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
	RelTypeSlideLayout = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
	RelTypeSlideMaster = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
	RelTypeTheme       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
	RelTypeNotesSlide  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
	RelTypeComments    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
	RelTypeChart       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
	RelTypeTable       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
	RelTypeMedia       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"
	RelTypeAudio       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
)

// ============================================================================
// 构造函数
// ============================================================================

// NewXMLRelationships 创建带默认命名空间的关系集合
func NewXMLRelationships() *XMLRelationships {
	return &XMLRelationships{
		Xmlns:         NamespaceRelationships,
		Relationships: make([]XMLRelationship, 0),
	}
}

// NewXMLRelationship 创建单个关系
func NewXMLRelationship(id, relType, target string) XMLRelationship {
	return XMLRelationship{
		ID:     id,
		Type:   relType,
		Target: target,
	}
}

// NewXMLRelationshipExternal 创建外部关系
func NewXMLRelationshipExternal(id, relType, target string) XMLRelationship {
	return XMLRelationship{
		ID:         id,
		Type:       relType,
		Target:     target,
		TargetMode: TargetModeExternal,
	}
}

// ============================================================================
// 辅助方法
// ============================================================================

// Add 添加关系到集合
func (rs *XMLRelationships) Add(rel XMLRelationship) {
	rs.Relationships = append(rs.Relationships, rel)
}

// AddNew 创建并添加新关系
func (rs *XMLRelationships) AddNew(id, relType, target string) {
	rs.Add(NewXMLRelationship(id, relType, target))
}

// GetByID 根据 ID 获取关系
func (rs *XMLRelationships) GetByID(id string) *XMLRelationship {
	for i := range rs.Relationships {
		if rs.Relationships[i].ID == id {
			return &rs.Relationships[i]
		}
	}
	return nil
}

// GetByType 根据类型获取所有关系
func (rs *XMLRelationships) GetByType(relType string) []XMLRelationship {
	var result []XMLRelationship
	for _, rel := range rs.Relationships {
		if rel.Type == relType {
			result = append(result, rel)
		}
	}
	return result
}

// GetByTarget 根据目标路径获取关系
func (rs *XMLRelationships) GetByTarget(target string) *XMLRelationship {
	for i := range rs.Relationships {
		if rs.Relationships[i].Target == target {
			return &rs.Relationships[i]
		}
	}
	return nil
}

// Count 返回关系数量
func (rs *XMLRelationships) Count() int {
	return len(rs.Relationships)
}

// IsExternal 检查是否为外部关系
func (r *XMLRelationship) IsExternal() bool {
	return r.TargetMode == TargetModeExternal
}

// ============================================================================
// XML 序列化/反序列化
// ============================================================================

// ToXML 将关系集合序列化为 XML 字节
func (rs *XMLRelationships) ToXML() ([]byte, error) {
	output, err := xml.MarshalIndent(rs, "", "  ")
	if err != nil {
		return nil, err
	}
	return append([]byte(XMLDeclaration), output...), nil
}

// ParseRelationships 从 XML 字节解析关系集合
func ParseRelationships(data []byte) (*XMLRelationships, error) {
	var rs XMLRelationships
	if err := xml.Unmarshal(data, &rs); err != nil {
		return nil, err
	}
	return &rs, nil
}
