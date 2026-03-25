package opc

import (
	"encoding/xml"
	"fmt"
	"io"
	"sync"
)

// Part 表示包中的一个部件
type Part struct {
	uri          *PackURI
	contentType  string
	blob         []byte
	relationships *Relationships
	dirty        bool // 是否被修改过
	mu           sync.RWMutex
}

// NewPart 创建一个新的部件
func NewPart(uri *PackURI, contentType string, blob []byte) *Part {
	return &Part{
		uri:          uri,
		contentType:  contentType,
		blob:         blob,
		relationships: NewRelationships(uri),
		dirty:        true,
	}
}

// NewPartFromReader 从Reader创建部件
func NewPartFromReader(uri *PackURI, contentType string, r io.Reader) (*Part, error) {
	blob, err := io.ReadAll(r)
	if err != nil {
		return nil, fmt.Errorf("failed to read part content: %w", err)
	}
	return NewPart(uri, contentType, blob), nil
}

// PartURI 返回部件URI
func (p *Part) PartURI() *PackURI {
	return p.uri
}

// ContentType 返回内容类型
func (p *Part) ContentType() string {
	return p.contentType
}

// SetContentType 设置内容类型
func (p *Part) SetContentType(ct string) {
	p.mu.Lock()
	defer p.mu.Unlock()
	p.contentType = ct
	p.dirty = true
}

// Blob 返回原始内容
func (p *Part) Blob() []byte {
	p.mu.RLock()
	defer p.mu.RUnlock()
	return p.blob
}

// SetBlob 设置内容
func (p *Part) SetBlob(blob []byte) {
	p.mu.Lock()
	defer p.mu.Unlock()
	p.blob = blob
	p.dirty = true
}

// SetBlobFromReader 从Reader设置内容
func (p *Part) SetBlobFromReader(r io.Reader) error {
	p.mu.Lock()
	defer p.mu.Unlock()

	blob, err := io.ReadAll(r)
	if err != nil {
		return fmt.Errorf("failed to read content: %w", err)
	}
	p.blob = blob
	p.dirty = true
	return nil
}

// Reader 返回内容的Reader
func (p *Part) Reader() io.Reader {
	p.mu.RLock()
	defer p.mu.RUnlock()
	return NewBytesReader(p.blob)
}

// Size 返回内容大小
func (p *Part) Size() int {
	p.mu.RLock()
	defer p.mu.RUnlock()
	return len(p.blob)
}

// Relationships 返回关系集合
func (p *Part) Relationships() *Relationships {
	return p.relationships
}

// AddRelationship 添加一个关系
func (p *Part) AddRelationship(relType, targetURI string, isExternal bool) (*Relationship, error) {
	rel, err := p.relationships.AddNew(relType, targetURI, isExternal)
	if err != nil {
		return nil, err
	}
	p.dirty = true
	return rel, nil
}

// RemoveRelationship 删除一个关系
func (p *Part) RemoveRelationship(rID string) error {
	err := p.relationships.Remove(rID)
	if err != nil {
		return err
	}
	p.dirty = true
	return nil
}

// GetRelatedPart 通过关系获取目标部件（需要Package上下文）
func (p *Part) GetRelatedPart(rID string) *PackURI {
	rel := p.relationships.Get(rID)
	if rel == nil {
		return nil
	}
	return rel.TargetURI()
}

// IsDirty 返回是否被修改
func (p *Part) IsDirty() bool {
	p.mu.RLock()
	defer p.mu.RUnlock()
	return p.dirty
}

// SetDirty 设置修改标记
func (p *Part) SetDirty(dirty bool) {
	p.mu.Lock()
	defer p.mu.Unlock()
	p.dirty = dirty
}

// LoadRelationships 从XML加载关系
func (p *Part) LoadRelationships(data []byte) error {
	return p.relationships.FromXML(data)
}

// RelationshipsBlob 返回关系的XML内容
func (p *Part) RelationshipsBlob() ([]byte, error) {
	if p.relationships.Count() == 0 {
		return nil, nil
	}
	return p.relationships.ToXML()
}

// HasRelationships 检查是否有关系
func (p *Part) HasRelationships() bool {
	return p.relationships.Count() > 0
}

// RelationshipsURI 返回关系文件的URI
func (p *Part) RelationshipsURI() *PackURI {
	return p.uri.RelationshipsURI()
}

// Clone 克隆部件
func (p *Part) Clone() *Part {
	p.mu.RLock()
	defer p.mu.RUnlock()

	blobCopy := make([]byte, len(p.blob))
	copy(blobCopy, p.blob)

	return &Part{
		uri:          p.uri.Clone(),
		contentType:  p.contentType,
		blob:         blobCopy,
		relationships: p.relationships.Clone(),
		dirty:        p.dirty,
	}
}

// UnmarshalBlob 从 blob 解析 XML 内容到 v
func (p *Part) UnmarshalBlob(v interface{}) error {
	p.mu.RLock()
	defer p.mu.RUnlock()

	return xml.Unmarshal(p.blob, v)
}

// MarshalToBlob 将 v 序列化为 XML 并存储到 blob
func (p *Part) MarshalToBlob(v interface{}) error {
	p.mu.Lock()
	defer p.mu.Unlock()

	data, err := xml.Marshal(v)
	if err != nil {
		return fmt.Errorf("failed to marshal XML: %w", err)
	}
	p.blob = data
	p.dirty = true
	return nil
}

// BytesReader 简单的bytes reader实现
type BytesReader struct {
	data []byte
	pos  int
}

// NewBytesReader 创建新的BytesReader
func NewBytesReader(data []byte) *BytesReader {
	return &BytesReader{data: data}
}

// Read 实现io.Reader接口
func (r *BytesReader) Read(p []byte) (n int, err error) {
	if r.pos >= len(r.data) {
		return 0, io.EOF
	}
	n = copy(p, r.data[r.pos:])
	r.pos += n
	return n, nil
}

// PartCollection 部件集合
type PartCollection struct {
	parts map[string]*Part
	order []string // 保持插入顺序
	mu    sync.RWMutex
}

// NewPartCollection 创建新的部件集合
func NewPartCollection() *PartCollection {
	return &PartCollection{
		parts: make(map[string]*Part),
		order: make([]string, 0),
	}
}

// Add 添加部件
func (c *PartCollection) Add(part *Part) error {
	c.mu.Lock()
	defer c.mu.Unlock()

	uri := part.PartURI().URI()
	if _, exists := c.parts[uri]; exists {
		return fmt.Errorf("part with URI %s already exists", uri)
	}

	c.parts[uri] = part
	c.order = append(c.order, uri)
	return nil
}

// Get 根据URI获取部件
func (c *PartCollection) Get(uri *PackURI) *Part {
	c.mu.RLock()
	defer c.mu.RUnlock()
	return c.parts[uri.URI()]
}

// GetByStr 根据字符串URI获取部件
func (c *PartCollection) GetByStr(uri string) *Part {
	return c.Get(NewPackURI(uri))
}

// Remove 删除部件
func (c *PartCollection) Remove(uri *PackURI) error {
	c.mu.Lock()
	defer c.mu.Unlock()

	key := uri.URI()
	if _, exists := c.parts[key]; !exists {
		return fmt.Errorf("part with URI %s not found", key)
	}

	delete(c.parts, key)

	// 从order中移除
	for i, u := range c.order {
		if u == key {
			c.order = append(c.order[:i], c.order[i+1:]...)
			break
		}
	}
	return nil
}

// Contains 检查是否包含指定部件
func (c *PartCollection) Contains(uri *PackURI) bool {
	c.mu.RLock()
	defer c.mu.RUnlock()
	_, exists := c.parts[uri.URI()]
	return exists
}

// All 返回所有部件（按插入顺序）
func (c *PartCollection) All() []*Part {
	c.mu.RLock()
	defer c.mu.RUnlock()

	result := make([]*Part, 0, len(c.order))
	for _, uri := range c.order {
		result = append(result, c.parts[uri])
	}
	return result
}

// Count 返回部件数量
func (c *PartCollection) Count() int {
	c.mu.RLock()
	defer c.mu.RUnlock()
	return len(c.parts)
}

// URIs 返回所有部件URI
func (c *PartCollection) URIs() []*PackURI {
	c.mu.RLock()
	defer c.mu.RUnlock()

	result := make([]*PackURI, 0, len(c.order))
	for _, uri := range c.order {
		result = append(result, NewPackURI(uri))
	}
	return result
}

// GetByType 根据内容类型获取部件
func (c *PartCollection) GetByType(contentType string) []*Part {
	c.mu.RLock()
	defer c.mu.RUnlock()

	var result []*Part
	for _, uri := range c.order {
		if part := c.parts[uri]; part.ContentType() == contentType {
			result = append(result, part)
		}
	}
	return result
}

// Clear 清空集合
func (c *PartCollection) Clear() {
	c.mu.Lock()
	defer c.mu.Unlock()
	c.parts = make(map[string]*Part)
	c.order = make([]string, 0)
}

// DirtyParts 返回所有被修改的部件
func (c *PartCollection) DirtyParts() []*Part {
	c.mu.RLock()
	defer c.mu.RUnlock()

	var result []*Part
	for _, uri := range c.order {
		if part := c.parts[uri]; part.IsDirty() {
			result = append(result, part)
		}
	}
	return result
}

// PartFactory 部件工厂接口
type PartFactory interface {
	// CreatePart 创建部件
	CreatePart(uri *PackURI, contentType string, blob []byte) (*Part, error)
}

// DefaultPartFactory 默认部件工厂
type DefaultPartFactory struct{}

// CreatePart 实现PartFactory接口
func (f *DefaultPartFactory) CreatePart(uri *PackURI, contentType string, blob []byte) (*Part, error) {
	return NewPart(uri, contentType, blob), nil
}
