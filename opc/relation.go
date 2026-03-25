package opc

import (
	"encoding/xml"
	"fmt"
	"sort"
	"strings"
	"sync"
)

// Relationship 表示两个部件之间的关系
type Relationship struct {
	rID         string       // 关系ID (rId1, rId2, ...)
	relType     string       // 关系类型
	target      *PackURI     // 目标URI
	targetMode  string       // Internal 或 External
	isExternal  bool         // 是否为外部目标
	source      *PackURI     // 源部件URI（用于解析相对路径）
}

// NewRelationship 创建一个新的关系
func NewRelationship(rID, relType, targetURI string, isExternal bool, source *PackURI) *Relationship {
	rel := &Relationship{
		rID:        rID,
		relType:    relType,
		isExternal: isExternal,
		source:     source,
	}

	if isExternal {
		rel.targetMode = "External"
		rel.target = &PackURI{uri: targetURI}
	} else {
		rel.targetMode = "Internal"
		rel.target = NewPackURI(targetURI)
	}

	return rel
}

// RID 返回关系ID
func (r *Relationship) RID() string {
	return r.rID
}

// Type 返回关系类型
func (r *Relationship) Type() string {
	return r.relType
}

// TargetURI 返回目标URI
func (r *Relationship) TargetURI() *PackURI {
	return r.target
}

// TargetRef 返回目标引用（相对或绝对）
// 如果有源部件，返回从源部件到目标的相对路径
func (r *Relationship) TargetRef() string {
	if r.isExternal {
		return r.target.URI()
	}
	if r.source != nil {
		return r.target.RelPathFrom(r.source)
	}
	return r.target.URI()
}

// IsExternal 返回是否为外部关系
func (r *Relationship) IsExternal() bool {
	return r.isExternal
}

// TargetMode 返回目标模式
func (r *Relationship) TargetMode() string {
	return r.targetMode
}

// SourceURI 返回源URI
func (r *Relationship) SourceURI() *PackURI {
	return r.source
}

// SetSource 设置源URI
func (r *Relationship) SetSource(source *PackURI) {
	r.source = source
}

// Equals 比较两个关系是否相等
func (r *Relationship) Equals(other *Relationship) bool {
	if other == nil {
		return false
	}
	return r.rID == other.rID && r.relType == other.relType && r.target.Equals(other.target)
}

// Relationships 表示关系的集合
type Relationships struct {
	relationships map[string]*Relationship
	order        []string // 保持插入顺序
	mu           sync.RWMutex
	sourceURI    *PackURI // 关系所属的源部件
}

// NewRelationships 创建新的关系集合
func NewRelationships(sourceURI *PackURI) *Relationships {
	return &Relationships{
		relationships: make(map[string]*Relationship),
		order:        make([]string, 0),
		sourceURI:    sourceURI,
	}
}

// Add 添加一个关系
func (rs *Relationships) Add(rel *Relationship) error {
	rs.mu.Lock()
	defer rs.mu.Unlock()

	if _, exists := rs.relationships[rel.RID()]; exists {
		return fmt.Errorf("relationship with rID %s already exists", rel.RID())
	}

	// 设置源URI
	if rel.SourceURI() == nil && rs.sourceURI != nil {
		rel.SetSource(rs.sourceURI)
	}

	rs.relationships[rel.RID()] = rel
	rs.order = append(rs.order, rel.RID())
	return nil
}

// AddNew 创建并添加一个新关系
func (rs *Relationships) AddNew(relType, targetURI string, isExternal bool) (*Relationship, error) {
	rID := rs.NextRID()
	rel := NewRelationship(rID, relType, targetURI, isExternal, rs.sourceURI)
	err := rs.Add(rel)
	if err != nil {
		return nil, err
	}
	return rel, nil
}

// Get 根据rID获取关系
func (rs *Relationships) Get(rID string) *Relationship {
	rs.mu.RLock()
	defer rs.mu.RUnlock()
	return rs.relationships[rID]
}

// GetByType 根据关系类型获取所有关系
func (rs *Relationships) GetByType(relType string) []*Relationship {
	rs.mu.RLock()
	defer rs.mu.RUnlock()

	var result []*Relationship
	for _, rID := range rs.order {
		if rel := rs.relationships[rID]; rel.Type() == relType {
			result = append(result, rel)
		}
	}
	return result
}

// GetByTarget 根据目标URI获取关系
func (rs *Relationships) GetByTarget(targetURI *PackURI) *Relationship {
	rs.mu.RLock()
	defer rs.mu.RUnlock()

	for _, rel := range rs.relationships {
		if rel.TargetURI().Equals(targetURI) {
			return rel
		}
	}
	return nil
}

// Remove 删除一个关系
func (rs *Relationships) Remove(rID string) error {
	rs.mu.Lock()
	defer rs.mu.Unlock()

	if _, exists := rs.relationships[rID]; !exists {
		return fmt.Errorf("relationship with rID %s not found", rID)
	}

	delete(rs.relationships, rID)

	// 从order中移除
	for i, id := range rs.order {
		if id == rID {
			rs.order = append(rs.order[:i], rs.order[i+1:]...)
			break
		}
	}
	return nil
}

// Contains 检查是否包含指定rID的关系
func (rs *Relationships) Contains(rID string) bool {
	rs.mu.RLock()
	defer rs.mu.RUnlock()
	_, exists := rs.relationships[rID]
	return exists
}

// All 返回所有关系（按插入顺序）
func (rs *Relationships) All() []*Relationship {
	rs.mu.RLock()
	defer rs.mu.RUnlock()

	result := make([]*Relationship, 0, len(rs.order))
	for _, rID := range rs.order {
		result = append(result, rs.relationships[rID])
	}
	return result
}

// Count 返回关系数量
func (rs *Relationships) Count() int {
	rs.mu.RLock()
	defer rs.mu.RUnlock()
	return len(rs.relationships)
}

// NextRID 生成下一个关系ID
func (rs *Relationships) NextRID() string {
	rs.mu.RLock()
	defer rs.mu.RUnlock()

	// 找到最大的数字ID
	maxNum := 0
	for rID := range rs.relationships {
		if strings.HasPrefix(rID, "rId") {
			var num int
			_, err := fmt.Sscanf(rID, "rId%d", &num)
			if err == nil && num > maxNum {
				maxNum = num
			}
		}
	}
	return fmt.Sprintf("rId%d", maxNum+1)
}

// SetSourceURI 设置源URI
func (rs *Relationships) SetSourceURI(sourceURI *PackURI) {
	rs.mu.Lock()
	defer rs.mu.Unlock()
	rs.sourceURI = sourceURI
	// 更新所有关系的源
	for _, rel := range rs.relationships {
		rel.SetSource(sourceURI)
	}
}

// SourceURI 返回源URI
func (rs *Relationships) SourceURI() *PackURI {
	rs.mu.RLock()
	defer rs.mu.RUnlock()
	return rs.sourceURI
}

// Clone 克隆关系集合
func (rs *Relationships) Clone() *Relationships {
	rs.mu.RLock()
	defer rs.mu.RUnlock()

	newRs := NewRelationships(rs.sourceURI)
	for _, rID := range rs.order {
		rel := rs.relationships[rID]
		newRel := NewRelationship(rel.RID(), rel.Type(), rel.TargetURI().URI(), rel.IsExternal(), rel.SourceURI())
		newRs.relationships[rID] = newRel
		newRs.order = append(newRs.order, rID)
	}
	return newRs
}

// XML 结构体用于序列化关系

// XRelationships XML序列化的根元素
type XRelationships struct {
	XMLName      xml.Name       `xml:"Relationships"`
	Xmlns        string         `xml:"xmlns,attr"`
	Relationships []XRelationship `xml:"Relationship"`
}

// XRelationship XML序列化的关系元素
type XRelationship struct {
	ID         string `xml:"Id,attr"`
	Type       string `xml:"Type,attr"`
	Target     string `xml:"Target,attr"`
	TargetMode string `xml:"TargetMode,attr,omitempty"`
}

// MarshalXML 实现 xml.Marshaler 接口
func (rs *Relationships) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	rs.mu.RLock()
	defer rs.mu.RUnlock()

	xrels := XRelationships{
		Xmlns: NamespaceRelationships,
	}

	for _, rID := range rs.order {
		rel := rs.relationships[rID]
		xrel := XRelationship{
			ID:     rel.RID(),
			Type:   rel.Type(),
			Target: rel.TargetRef(),
		}
		if rel.IsExternal() {
			xrel.TargetMode = "External"
		}
		xrels.Relationships = append(xrels.Relationships, xrel)
	}

	return e.Encode(xrels)
}

// UnmarshalXML 实现 xml.Unmarshaler 接口
func (rs *Relationships) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	rs.mu.Lock()
	defer rs.mu.Unlock()

	var xrels XRelationships
	if err := d.DecodeElement(&xrels, &start); err != nil {
		return err
	}

	rs.relationships = make(map[string]*Relationship)
	rs.order = make([]string, 0)

	for _, xrel := range xrels.Relationships {
		isExternal := xrel.TargetMode == "External"
		rel := NewRelationship(xrel.ID, xrel.Type, xrel.Target, isExternal, rs.sourceURI)
		rs.relationships[xrel.ID] = rel
		rs.order = append(rs.order, xrel.ID)
	}

	return nil
}

// ToXML 将关系集合序列化为XML
func (rs *Relationships) ToXML() ([]byte, error) {
	return xml.Marshal(rs)
}

// FromXML 从XML解析关系集合
func (rs *Relationships) FromXML(data []byte) error {
	return xml.Unmarshal(data, rs)
}

// Relatable 可关联部件的接口
type Relatable interface {
	// PartURI 返回部件的URI
	PartURI() *PackURI
	// Relationships 返回部件的关系集合
	Relationships() *Relationships
	// AddRelationship 添加一个关系
	AddRelationship(relType, targetURI string, isExternal bool) (*Relationship, error)
}

// RelTypeCollection 关系类型集合（用于按类型分组查找）
type RelTypeCollection struct {
	types map[string][]*Relationship
}

// NewRelTypeCollection 创建新的关系类型集合
func NewRelTypeCollection() *RelTypeCollection {
	return &RelTypeCollection{
		types: make(map[string][]*Relationship),
	}
}

// Add 添加关系到类型集合
func (c *RelTypeCollection) Add(rel *Relationship) {
	c.types[rel.Type()] = append(c.types[rel.Type()], rel)
}

// GetByType 按类型获取关系
func (c *RelTypeCollection) GetByType(relType string) []*Relationship {
	return c.types[relType]
}

// Types 返回所有关系类型
func (c *RelTypeCollection) Types() []string {
	types := make([]string, 0, len(c.types))
	for t := range c.types {
		types = append(types, t)
	}
	sort.Strings(types)
	return types
}

// ParseRelationshipsFromXML 从XML数据解析关系
func ParseRelationshipsFromXML(data []byte, sourceURI *PackURI) (*Relationships, error) {
	rels := NewRelationships(sourceURI)
	if err := rels.FromXML(data); err != nil {
		return nil, fmt.Errorf("failed to parse relationships: %w", err)
	}
	return rels, nil
}
