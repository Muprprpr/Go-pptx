package parts

import (
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
	"strings"
	"sync"

	"github.com/Muprprpr/Go-pptx/opc"
	"github.com/Muprprpr/Go-pptx/utils"
)

// NewShapeIDAllocator 创建新的 ID 分配器
// reservedID: 保留的起始 ID，nextID 将从 reservedID + 1 开始
func NewShapeIDAllocator(reservedID uint32) *ShapeIDAllocator {
	return &ShapeIDAllocator{
		nextID:     reservedID + 1,
		reservedID: reservedID,
		maxID:      0, // 无限制
	}
}

// NewShapeIDAllocatorWithMax 创建带最大 ID 限制的分配器
func NewShapeIDAllocatorWithMax(reservedID, maxID uint32) *ShapeIDAllocator {
	return &ShapeIDAllocator{
		nextID:     reservedID + 1,
		reservedID: reservedID,
		maxID:      maxID,
	}
}

// Next 分配下一个 ID
// 返回新分配的 ID 值
func (a *ShapeIDAllocator) Next() uint32 {
	if a.maxID > 0 && a.nextID >= a.maxID {
		return 0 // 超出范围
	}
	id := a.nextID
	a.nextID++
	return id
}

// NextBatch 批量分配多个 ID
// count: 需要分配的 ID 数量
// 返回分配的 ID 数组（不包含 reservedID）
func (a *ShapeIDAllocator) NextBatch(count int) []uint32 {
	ids := make([]uint32, count)
	for i := 0; i < count; i++ {
		ids[i] = a.Next()
	}
	return ids
}

// Peek 返回下一个将被分配的 ID（不实际分配）
func (a *ShapeIDAllocator) Peek() uint32 {
	return a.nextID
}

// Current 返回当前 ID（最后分配的 ID）
func (a *ShapeIDAllocator) Current() uint32 {
	if a.nextID > a.reservedID+1 {
		return a.nextID - 1
	}
	return a.reservedID
}

// Reset 重置分配器，从 reservedID + 1 重新开始
func (a *ShapeIDAllocator) Reset() {
	a.nextID = a.reservedID + 1
}

// ResetFrom 从指定 ID 开始重新分配
func (a *ShapeIDAllocator) ResetFrom(startID uint32) {
	a.nextID = startID
}

// SetReserved 设置保留的起始 ID 并重置
func (a *ShapeIDAllocator) SetReserved(reservedID uint32) {
	a.reservedID = reservedID
	a.nextID = reservedID + 1
}

// Remaining 返回剩余可分配的 ID 数量
func (a *ShapeIDAllocator) Remaining() uint32 {
	if a.maxID == 0 {
		return ^uint32(0) // 最大值
	}
	return a.maxID - a.nextID + 1
}

// IsExhausted 检查 ID 是否已耗尽
func (a *ShapeIDAllocator) IsExhausted() bool {
	if a.maxID == 0 {
		return false
	}
	return a.nextID > a.maxID
}

// UsedCount 返回已使用的 ID 数量
func (a *ShapeIDAllocator) UsedCount() uint32 {
	return a.nextID - a.reservedID - 1
}

// ============================================================================
// NewShapeIDAllocatorSync 创建线程安全的 ID 分配器
func NewShapeIDAllocatorSync(reservedID uint32) *ShapeIDAllocatorSync {
	return &ShapeIDAllocatorSync{
		nextID:     reservedID + 1,
		reservedID: reservedID,
	}
}

// NewShapeIDAllocatorSyncWithMax 创建带最大 ID 限制的线程安全分配器
func NewShapeIDAllocatorSyncWithMax(reservedID, maxID uint32) *ShapeIDAllocatorSync {
	return &ShapeIDAllocatorSync{
		nextID:     reservedID + 1,
		reservedID: reservedID,
		maxID:      maxID,
	}
}

// Next 线程安全地分配下一个 ID
func (a *ShapeIDAllocatorSync) Next() uint32 {
	a.mu.Lock()
	defer a.mu.Unlock()

	if a.maxID > 0 && a.nextID >= a.maxID {
		return 0
	}
	id := a.nextID
	a.nextID++
	return id
}

// NextBatch 线程安全地批量分配多个 ID
func (a *ShapeIDAllocatorSync) NextBatch(count int) []uint32 {
	a.mu.Lock()
	defer a.mu.Unlock()

	ids := make([]uint32, count)
	for i := 0; i < count; i++ {
		if a.maxID > 0 && a.nextID >= a.maxID {
			ids[i] = 0
			continue
		}
		ids[i] = a.nextID
		a.nextID++
	}
	return ids
}

// TryNext 尝试分配 ID，失败时返回 false
func (a *ShapeIDAllocatorSync) TryNext() (uint32, bool) {
	a.mu.Lock()
	defer a.mu.Unlock()

	if a.maxID > 0 && a.nextID >= a.maxID {
		return 0, false
	}
	id := a.nextID
	a.nextID++
	return id, true
}

// Peek 线程安全地返回下一个将被分配的 ID
func (a *ShapeIDAllocatorSync) Peek() uint32 {
	a.mu.Lock()
	defer a.mu.Unlock()
	return a.nextID
}

// Reset 线程安全地重置分配器
func (a *ShapeIDAllocatorSync) Reset() {
	a.mu.Lock()
	defer a.mu.Unlock()
	a.nextID = a.reservedID + 1
}

// ResetFrom 线程安全地从指定 ID 开始重置
func (a *ShapeIDAllocatorSync) ResetFrom(startID uint32) {
	a.mu.Lock()
	defer a.mu.Unlock()
	a.nextID = startID
}

// ============================================================================
// SlidePart ID 管理方法
// ============================================================================

// Allocator 返回当前的 Shape ID 分配器（用于自定义管理）
// 注意：单个 Slide 由单 goroutine 负责生成，非并发安全
func (s *SlidePart) Allocator() *ShapeIDAllocator {
	return &ShapeIDAllocator{
		nextID:     s.nextShapeID,
		reservedID: 1,
	}
}

// NextShapeID 返回下一个可用的 Shape ID（不递增）
func (s *SlidePart) NextShapeID() uint32 {
	return s.nextShapeID
}

// AllocateShapeID 分配一个新的 shape ID
// 返回新分配的 ID 值
func (s *SlidePart) AllocateShapeID() uint32 {
	id := s.nextShapeID
	s.nextShapeID++
	return id
}

// AllocateShapeIDBatch 批量分配多个 shape ID
func (s *SlidePart) AllocateShapeIDBatch(count int) []uint32 {
	ids := make([]uint32, count)
	for i := 0; i < count; i++ {
		ids[i] = s.nextShapeID
		s.nextShapeID++
	}
	return ids
}

// AllocateShapeIDWithOffset 分配一个带偏移量的 shape ID
// offset: 在当前基础上增加的偏移量
func (s *SlidePart) AllocateShapeIDWithOffset(offset uint32) uint32 {
	id := s.nextShapeID + offset
	s.nextShapeID = id + 1
	return id
}

// PeekNextShapeID 查看下一个可用的 Shape ID（不递增）
func (s *SlidePart) PeekNextShapeID() uint32 {
	return s.nextShapeID
}

// CurrentShapeID 返回当前最后分配的 Shape ID
func (s *SlidePart) CurrentShapeID() uint32 {
	if s.nextShapeID > 2 {
		return s.nextShapeID - 1
	}
	return 1
}

// ResetShapeID 重置 Shape ID 分配器
func (s *SlidePart) ResetShapeID() {
	s.nextShapeID = 2
}

// SetShapeIDStart 设置 Shape ID 的起始值
func (s *SlidePart) SetShapeIDStart(startID uint32) {
	if startID < 2 {
		startID = 2
	}
	s.nextShapeID = startID
}

// ShapeIDCount 返回已分配的 Shape ID 数量
func (s *SlidePart) ShapeIDCount() uint32 {
	return s.nextShapeID - 2
}

// NewSlidePart 创建新的幻灯片部件
func NewSlidePart(id int) *SlidePart {
	uri := opc.NewPackURI(fmt.Sprintf("/ppt/slides/slide%d.xml", id))
	return &SlidePart{
		uri:         uri,
		spTree:      NewXSpTree(),
		nextShapeID: 2, // 1 已被 spTree 自身使用
		rels:        opc.NewRelationships(uri),
	}
}

// NewSlidePartWithURI 使用指定 URI 创建幻灯片部件
func NewSlidePartWithURI(uri *opc.PackURI) *SlidePart {
	return &SlidePart{
		uri:         uri,
		spTree:      NewXSpTree(),
		nextShapeID: 2,
		rels:        opc.NewRelationships(uri),
	}
}

// PartURI 返回部件 URI
func (s *SlidePart) PartURI() *opc.PackURI {
	return s.uri
}

// SetURI 设置部件 URI
func (s *SlidePart) SetURI(uri *opc.PackURI) {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.uri = uri
}

// LayoutRId 返回关联的布局 rId
func (s *SlidePart) LayoutRId() string {
	s.mu.RLock()
	defer s.mu.RUnlock()
	return s.layoutRId
}

// SetLayoutRId 设置关联的布局 rId
func (s *SlidePart) SetLayoutRId(rId string) {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.layoutRId = rId
}

// MasterRId 返回关联的母版 rId
func (s *SlidePart) MasterRId() string {
	s.mu.RLock()
	defer s.mu.RUnlock()
	return s.masterRId
}

// SetMasterRId 设置关联的母版 rId
func (s *SlidePart) SetMasterRId(rId string) {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.masterRId = rId
}

// Relationships 返回幻灯片关系管理器（使用 opc.Relationships）
func (s *SlidePart) Relationships() *opc.Relationships {
	s.mu.RLock()
	defer s.mu.RUnlock()
	return s.rels
}

// SpTree 返回形状树（用于高层构建器操作）
func (s *SlidePart) SpTree() *XSpTree {
	s.mu.RLock()
	defer s.mu.RUnlock()
	return s.spTree
}

// AppendShapeChild 向形状树添加子元素（用于高层构建器操作）
func (s *SlidePart) AppendShapeChild(child any) {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.spTree.Children = append(s.spTree.Children, child)
}

// NewXSpTree 创建新的形状树
func NewXSpTree() *XSpTree {
	return &XSpTree{
		NonVisual: nvGrpSpPr{
			CNvPr: &XNvCxnSpPr{
				ID:   1,
				Name: "Layout",
			},
		},
		Children: make([]any, 0),
	}
}

// ToXML 将 SlidePart 序列化为 XML
// 使用 XMLWriterPool 统一序列化底座
func (s *SlidePart) ToXML() ([]byte, error) {
	s.mu.RLock()
	defer s.mu.RUnlock()

	xs := XSlide{
		XmlnsA: "http://schemas.openxmlformats.org/drawingml/2006/main",
		XmlnsR: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		XmlnsP: "http://schemas.openxmlformats.org/presentationml/2006/main",
		ClrMapOvr: &XClrMapOvr{
			Accent1: "accent1",
		},
		CSld: &XCSld{
			SpTree: s.spTree,
		},
	}

	// 使用 XMLWriterPool 统一序列化
	xw := globalXMLWriterPool.Get()
	defer globalXMLWriterPool.Put(xw)

	xw.ResetBuffer()
	if err := xw.Declaration(); err != nil {
		return nil, err
	}
	if err := xs.WriteXML(xw); err != nil {
		return nil, err
	}
	return xw.Bytes(), nil
}

// FromXML 从 XML 反序列化为 SlidePart
func (s *SlidePart) FromXML(data []byte) error {
	// 去除命名空间前缀以兼容 Go 的 xml.Unmarshal
	cleanData, err := StripNamespacePrefixes(data)
	if err != nil {
		return fmt.Errorf("failed to clean XML: %w", err)
	}

	var xs XSlide
	if err := xml.Unmarshal(cleanData, &xs); err != nil {
		return fmt.Errorf("failed to unmarshal slide XML: %w", err)
	}

	s.mu.Lock()
	defer s.mu.Unlock()

	if xs.CSld != nil && xs.CSld.SpTree != nil {
		s.spTree = xs.CSld.SpTree
	}

	return nil
}

// NewSlideLayoutPart 创建新的幻灯片布局部件
func NewSlideLayoutPart(id int) *SlideLayoutPart {
	return &SlideLayoutPart{
		uri:       opc.NewPackURI(fmt.Sprintf("/ppt/slideLayouts/slideLayout%d.xml", id)),
		spTree:    NewXSpTree(),
		layoutType: SlideLayoutBlank,
	}
}

// PartURI 返回部件 URI
func (s *SlideLayoutPart) PartURI() *opc.PackURI {
	return s.uri
}

// LayoutType 返回布局类型
func (s *SlideLayoutPart) LayoutType() SlideLayoutType {
	s.mu.RLock()
	defer s.mu.RUnlock()
	return s.layoutType
}

// SetLayoutType 设置布局类型
func (s *SlideLayoutPart) SetLayoutType(t SlideLayoutType) {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.layoutType = t
}

// MasterRId 返回关联的母版 rId
func (s *SlideLayoutPart) MasterRId() string {
	s.mu.RLock()
	defer s.mu.RUnlock()
	return s.masterRId
}

// SetMasterRId 设置关联的母版 rId
func (s *SlideLayoutPart) SetMasterRId(rId string) {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.masterRId = rId
}

// AddImageRel 添加图片关系（带去重）
// 返回分配的 rId
func (s *SlidePart) AddImageRel(targetURI string) string {
	s.mu.Lock()
	defer s.mu.Unlock()

	target := opc.NewPackURI(targetURI)

	// 去重检查：如果已存在相同目标的关系，复用其 rId
	if existing := s.rels.GetByTarget(target); existing != nil {
		return existing.RID()
	}

	// 创建新关系
	rel, err := s.rels.AddNew(RelTypeImage, targetURI, false)
	if err != nil {
		return ""
	}
	return rel.RID()
}

// AddMediaRel 添加媒体关系（带去重）
// 返回分配的 rId
func (s *SlidePart) AddMediaRel(targetURI string) string {
	s.mu.Lock()
	defer s.mu.Unlock()

	target := opc.NewPackURI(targetURI)

	// 去重检查
	if existing := s.rels.GetByTarget(target); existing != nil {
		return existing.RID()
	}

	rel, err := s.rels.AddNew(RelTypeMedia, targetURI, false)
	if err != nil {
		return ""
	}
	return rel.RID()
}

// AddChartRel 添加图表关系（带去重）
// 返回分配的 rId
func (s *SlidePart) AddChartRel(targetURI string) string {
	s.mu.Lock()
	defer s.mu.Unlock()

	target := opc.NewPackURI(targetURI)

	// 去重检查
	if existing := s.rels.GetByTarget(target); existing != nil {
		return existing.RID()
	}

	rel, err := s.rels.AddNew(RelTypeChart, targetURI, false)
	if err != nil {
		return ""
	}
	return rel.RID()
}

// AddTableRel 添加表格关系（带去重）
// 返回分配的 rId
func (s *SlidePart) AddTableRel(targetURI string) string {
	s.mu.Lock()
	defer s.mu.Unlock()

	target := opc.NewPackURI(targetURI)

	// 去重检查
	if existing := s.rels.GetByTarget(target); existing != nil {
		return existing.RID()
	}

	rel, err := s.rels.AddNew(RelTypeTable, targetURI, false)
	if err != nil {
		return ""
	}
	return rel.RID()
}

// NewXMLWriter 创建新的 XMLWriter
func NewXMLWriter(w io.Writer) *XMLWriter {
	return &XMLWriter{
		w:          w,
		buf:        make([]byte, 0, 4096),
		indent:     0,
		indentStr:  "  ",
		useIndent:  true,
		autoFlush:  false,
		nsPrefixes: make(map[string]string),
	}
}

// NewXMLWriterWithIndent 创建带缩进的 XMLWriter
func NewXMLWriterWithIndent(w io.Writer, indentStr string) *XMLWriter {
	return &XMLWriter{
		w:          w,
		buf:        make([]byte, 0, 4096),
		indent:     0,
		indentStr:  indentStr,
		useIndent:  true,
		autoFlush:  false,
		nsPrefixes: make(map[string]string),
	}
}

// NewXMLWriterBuffered 创建使用缓冲区的 XMLWriter
func NewXMLWriterBuffered(cap int) *XMLWriter {
	return &XMLWriter{
		w:          nil,
		buf:        make([]byte, 0, cap),
		indent:     0,
		indentStr:  "  ",
		useIndent:  true,
		autoFlush:  false,
		nsPrefixes: make(map[string]string),
	}
}

// SetAutoFlush 设置是否自动刷新到 writer
func (xw *XMLWriter) SetAutoFlush(enable bool) {
	xw.autoFlush = enable
}

// SetIndent 设置缩进字符串
func (xw *XMLWriter) SetIndent(indentStr string) {
	xw.indentStr = indentStr
}

// SetUseIndent 设置是否使用缩进
func (xw *XMLWriter) SetUseIndent(use bool) {
	xw.useIndent = use
}

// writeRaw 直接写入字节
func (xw *XMLWriter) writeRaw(data []byte) error {
	xw.buf = append(xw.buf, data...)
	if xw.autoFlush && xw.w != nil {
		_, err := xw.w.Write(xw.buf)
		xw.buf = xw.buf[:0]
		return err
	}
	return nil
}

// writeIndent 写入缩进
func (xw *XMLWriter) writeIndent() error {
	if xw.useIndent {
		for i := 0; i < xw.indent; i++ {
			if err := xw.writeRaw([]byte(xw.indentStr)); err != nil {
				return err
			}
		}
	}
	return nil
}

// Declaration 写入 XML 声明
func (xw *XMLWriter) Declaration() error {
	return xw.writeRaw([]byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`))
}

// DeclarationWithEncoding 写入带编码的 XML 声明
func (xw *XMLWriter) DeclarationWithEncoding(encoding string) error {
	return xw.writeRaw([]byte(`<?xml version="1.0" encoding="` + encoding + `" standalone="yes"?>`))
}

// StartElement 写入起始元素标签
func (xw *XMLWriter) StartElement(prefix, localName string) error {
	if prefix != "" {
		return xw.writeRaw([]byte("<" + prefix + ":" + localName + ">"))
	}
	return xw.writeRaw([]byte("<" + localName + ">"))
}

// StartElementNS 写入带命名空间的起始元素
func (xw *XMLWriter) StartElementNS(prefix, localName, ns string) error {
	xw.nsPrefixes[ns] = prefix
	if prefix != "" {
		return xw.writeRaw([]byte("<" + prefix + ":" + localName + " xmlns:" + prefix + "=\"" + ns + "\">"))
	}
	return xw.writeRaw([]byte("<" + localName + " xmlns=\"" + ns + "\">"))
}

// StartElementWithAttrs 写入带属性的起始元素
func (xw *XMLWriter) StartElementWithAttrs(prefix, localName string, attrs ...string) error {
	tag := "<"
	if prefix != "" {
		tag += prefix + ":" + localName
	} else {
		tag += localName
	}
	for i := 0; i < len(attrs); i += 2 {
		if i+1 < len(attrs) {
			tag += " " + attrs[i] + "=\"" + xw.escapeAttr(attrs[i+1]) + "\""
		}
	}
	tag += ">"
	return xw.writeRaw([]byte(tag))
}

// StartElementNSWithAttrs 写入带命名空间和属性的起始元素
func (xw *XMLWriter) StartElementNSWithAttrs(prefix, localName, ns string, attrs ...string) error {
	xw.nsPrefixes[ns] = prefix
	tag := "<"
	if prefix != "" {
		tag += prefix + ":" + localName + " xmlns:" + prefix + "=\"" + ns + "\""
	} else {
		tag += localName + " xmlns=\"" + ns + "\""
	}
	for i := 0; i < len(attrs); i += 2 {
		if i+1 < len(attrs) {
			tag += " " + attrs[i] + "=\"" + xw.escapeAttr(attrs[i+1]) + "\""
		}
	}
	tag += ">"
	return xw.writeRaw([]byte(tag))
}

// StartElementRaw 写入起始元素（不转义属性值）
func (xw *XMLWriter) StartElementRaw(prefix, localName string, attrs ...string) error {
	tag := "<"
	if prefix != "" {
		tag += prefix + ":" + localName
	} else {
		tag += localName
	}
	for i := 0; i < len(attrs); i += 2 {
		if i+1 < len(attrs) {
			tag += " " + attrs[i] + "=\"" + attrs[i+1] + "\""
		}
	}
	tag += ">"
	return xw.writeRaw([]byte(tag))
}

// EndElement 写入结束元素标签
func (xw *XMLWriter) EndElement(prefix, localName string) error {
	if prefix != "" {
		return xw.writeRaw([]byte("</" + prefix + ":" + localName + ">"))
	}
	return xw.writeRaw([]byte("</" + localName + ">"))
}

// EmptyElement 写入空元素标签
func (xw *XMLWriter) EmptyElement(prefix, localName string) error {
	if prefix != "" {
		return xw.writeRaw([]byte("<" + prefix + ":" + localName + "/>"))
	}
	return xw.writeRaw([]byte("<" + localName + "/>"))
}

// EmptyElementWithAttrs 写入带属性的空元素
func (xw *XMLWriter) EmptyElementWithAttrs(prefix, localName string, attrs ...string) error {
	tag := "<"
	if prefix != "" {
		tag += prefix + ":" + localName
	} else {
		tag += localName
	}
	for i := 0; i < len(attrs); i += 2 {
		if i+1 < len(attrs) {
			tag += " " + attrs[i] + "=\"" + xw.escapeAttr(attrs[i+1]) + "\""
		}
	}
	tag += "/>"
	return xw.writeRaw([]byte(tag))
}

// Text 写入文本内容
func (xw *XMLWriter) Text(content string) error {
	return xw.writeRaw([]byte(xw.escapeText(content)))
}

// TextRaw 写入原始文本（不转义）
func (xw *XMLWriter) TextRaw(content string) error {
	return xw.writeRaw([]byte(content))
}

// CharData 写入字符数据
func (xw *XMLWriter) CharData(data []byte) error {
	return xw.writeRaw([]byte(xw.escapeText(string(data))))
}

// Comment 写入注释
func (xw *XMLWriter) Comment(content string) error {
	if err := xw.writeRaw([]byte("<!--")); err != nil {
		return err
	}
	if err := xw.writeRaw([]byte(content)); err != nil {
		return err
	}
	return xw.writeRaw([]byte("-->"))
}

// CData 写入 CDATA 节
func (xw *XMLWriter) CData(content string) error {
	if err := xw.writeRaw([]byte("<![CDATA[")); err != nil {
		return err
	}
	if err := xw.writeRaw([]byte(content)); err != nil {
		return err
	}
	return xw.writeRaw([]byte("]]>"))
}

// ProcessingInstruction 写入处理指令
func (xw *XMLWriter) ProcessingInstruction(target, data string) error {
	if data != "" {
		return xw.writeRaw([]byte("<?" + target + " " + data + "?>"))
	}
	return xw.writeRaw([]byte("<?" + target + "?>"))
}

// Indent 增加缩进级别
func (xw *XMLWriter) Indent() {
	xw.indent++
}

// Dedent 减少缩进级别
func (xw *XMLWriter) Dedent() {
	if xw.indent > 0 {
		xw.indent--
	}
}

// WithIndent 临时增加缩进，执行后恢复
func (xw *XMLWriter) WithIndent(fn func()) {
	xw.Indent()
	fn()
	xw.Dedent()
}

// Newline 写入换行符
func (xw *XMLWriter) Newline() error {
	return xw.writeRaw([]byte{'\n'})
}

// Raw 写入原始内容
func (xw *XMLWriter) Raw(content string) error {
	return xw.writeRaw([]byte(content))
}

// Flush 将缓冲区内容刷新到 writer
func (xw *XMLWriter) Flush() error {
	if xw.w == nil {
		return fmt.Errorf("XMLWriter: writer is nil, no io.Writer configured")
	}
	_, err := xw.w.Write(xw.buf)
	xw.buf = xw.buf[:0]
	return err
}

// Bytes 返回缓冲区内容的副本（并发安全）
// 注意：必须返回副本，因为 XMLWriter 会被池化复用
// 如果返回引用，当 writer 被放回池并重新使用时会导致数据竞争
func (xw *XMLWriter) Bytes() []byte {
	// 创建副本以避免竞态条件
	result := make([]byte, len(xw.buf))
	copy(result, xw.buf)
	return result
}

// String 返回缓冲区内容的字符串形式
func (xw *XMLWriter) String() string {
	return string(xw.buf)
}

// Reset 重置 writer
func (xw *XMLWriter) Reset(w io.Writer) {
	xw.w = w
	xw.buf = xw.buf[:0]
	xw.indent = 0
	xw.nsPrefixes = make(map[string]string)
}

// ResetBuffer 重置缓冲区
func (xw *XMLWriter) ResetBuffer() {
	xw.buf = xw.buf[:0]
}

// Size 返回缓冲区当前大小
func (xw *XMLWriter) Size() int {
	return len(xw.buf)
}

// Capacity 返回缓冲区容量
func (xw *XMLWriter) Capacity() int {
	return cap(xw.buf)
}

// escapeText 转义 XML 文本内容
func (xw *XMLWriter) escapeText(s string) string {
	// 预检查是否需要转义
	if !strings.ContainsAny(s, "<>&\"'") {
		return s
	}

	var sb strings.Builder
	sb.Grow(len(s) + 16)
	for _, r := range s {
		switch r {
		case '<':
			sb.WriteString("&lt;")
		case '>':
			sb.WriteString("&gt;")
		case '&':
			sb.WriteString("&amp;")
		case '"':
			sb.WriteString("&quot;")
		case '\'':
			sb.WriteString("&apos;")
		default:
			sb.WriteRune(r)
		}
	}
	return sb.String()
}

// escapeAttr 转义 XML 属性值
func (xw *XMLWriter) escapeAttr(s string) string {
	// 预检查是否需要转义
	if !strings.ContainsAny(s, "<>&\"'") {
		return s
	}

	var sb strings.Builder
	sb.Grow(len(s) + 16)
	for _, r := range s {
		switch r {
		case '<':
			sb.WriteString("&lt;")
		case '>':
			sb.WriteString("&gt;")
		case '&':
			sb.WriteString("&amp;")
		case '"':
			sb.WriteString("&quot;")
		case '\'':
			sb.WriteString("&apos;")
		default:
			sb.WriteRune(r)
		}
	}
	return sb.String()
}

// WriteInt 写入整数值
func (xw *XMLWriter) WriteInt(val int) error {
	return xw.writeRaw([]byte(strconv.Itoa(val)))
}

// WriteInt64 写入 64 位整数值
func (xw *XMLWriter) WriteInt64(val int64) error {
	return xw.writeRaw([]byte(strconv.FormatInt(val, 10)))
}

// WriteUint64 写入无符号 64 位整数值
func (xw *XMLWriter) WriteUint64(val uint64) error {
	return xw.writeRaw([]byte(strconv.FormatUint(val, 10)))
}

// WriteFloat64 写入浮点数值
func (xw *XMLWriter) WriteFloat64(val float64, prec int) error {
	return xw.writeRaw([]byte(strconv.FormatFloat(val, 'f', prec, 64)))
}

// WriteBool 写入布尔值
func (xw *XMLWriter) WriteBool(val bool) error {
	if val {
		return xw.writeRaw([]byte("1"))
	}
	return xw.writeRaw([]byte("0"))
}

// WriteBoolStr 写入布尔值字符串（true/false）
func (xw *XMLWriter) WriteBoolStr(val bool) error {
	if val {
		return xw.writeRaw([]byte("true"))
	}
	return xw.writeRaw([]byte("false"))
}

// WriteEMUs 写入 EMU 单位值（用于 PowerPoint 尺寸）
func (xw *XMLWriter) WriteEMUs(val int64) error {
	return xw.writeRaw([]byte(utils.WriteEMUAttr(val)))
}

// WriteEMUsWithUnit 写入带单位的 EMU 值
func (xw *XMLWriter) WriteEMUsWithUnit(val int64) error {
	return xw.writeRaw([]byte(utils.WriteEMUAttr(val)))
}

// WriteEMUsF 写入浮点 EMU 值（基于常用单位转换）
func (xw *XMLWriter) WriteEMUsF(val float64) error {
	return xw.writeRaw([]byte(utils.WriteEMUAttr(int64(val))))
}

// WriteInchesAsEMU 将英寸转换为 EMU 并写入
func (xw *XMLWriter) WriteInchesAsEMU(inches float64) error {
	return xw.writeRaw([]byte(utils.WriteEMUAttr(utils.InchesToEMU(inches))))
}

// WriteCentimetersAsEMU 将厘米转换为 EMU 并写入
func (xw *XMLWriter) WriteCentimetersAsEMU(cm float64) error {
	return xw.writeRaw([]byte(utils.WriteEMUAttr(utils.CentimetersToEMU(cm))))
}

// WriteMillimetersAsEMU 将毫米转换为 EMU 并写入
func (xw *XMLWriter) WriteMillimetersAsEMU(mm float64) error {
	return xw.writeRaw([]byte(utils.WriteEMUAttr(utils.MillimetersToEMU(mm))))
}

// WritePointsAsEMU 将磅转换为 EMU 并写入
func (xw *XMLWriter) WritePointsAsEMU(points float64) error {
	return xw.writeRaw([]byte(utils.WriteEMUAttr(utils.PointsToEMU(points))))
}

// WritePixelsAsEMU 将像素转换为 EMU 并写入
func (xw *XMLWriter) WritePixelsAsEMU(pixels float64) error {
	return xw.writeRaw([]byte(utils.WriteEMUAttr(utils.PixelsToEMU(pixels))))
}

// WritePercentage 写入百分比值
func (xw *XMLWriter) WritePercentage(val int) error {
	return xw.writeRaw([]byte(strconv.Itoa(val)))
}

// NewXMLWriterPool 创建新的 XMLWriterPool
func NewXMLWriterPool() *XMLWriterPool {
	return &XMLWriterPool{
		pool: sync.Pool{
			New: func() any {
				return &XMLWriter{
					buf:        make([]byte, 0, 4096),
					indent:     0,
					indentStr:  "  ",
					useIndent:  true,
					autoFlush:  false,
					nsPrefixes: make(map[string]string),
				}
			},
		},
	}
}

// globalXMLWriterPool 全局 XMLWriter 池，用于统一序列化底座
var globalXMLWriterPool = NewXMLWriterPool()

// Get 获取或创建一个 XMLWriter
func (p *XMLWriterPool) Get() *XMLWriter {
	xw := p.pool.Get().(*XMLWriter)
	xw.buf = xw.buf[:0]
	xw.indent = 0
	xw.nsPrefixes = make(map[string]string)
	return xw
}

// Put 回收 XMLWriter 到池中
func (p *XMLWriterPool) Put(xw *XMLWriter) {
	xw.w = nil
	xw.buf = xw.buf[:0]
	xw.autoFlush = false
	xw.nsPrefixes = make(map[string]string)
	p.pool.Put(xw)
}

// GetWithWriter 获取配置了 writer 的 XMLWriter
func (p *XMLWriterPool) GetWithWriter(w io.Writer) *XMLWriter {
	xw := p.Get()
	xw.w = w
	return xw
}

// GetBuffered 获取使用缓冲区的 XMLWriter
func (p *XMLWriterPool) GetBuffered() *XMLWriter {
	return p.Get()
}

// ============================================================================
// XML 结构 WriteXML 方法
// ============================================================================

// WriteXML 将 XSlide 序列化为 XML
func (xs *XSlide) WriteXML(xw *XMLWriter) error {
	// 写入 p:sld 起始标签及命名空间
	if err := xw.StartElementWithAttrs("p", "sld",
		"xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main",
		"xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		"xmlns:p", "http://schemas.openxmlformats.org/presentationml/2006/main",
	); err != nil {
		return err
	}

	// 写入颜色映射覆盖
	if err := xw.StartElement("p", "clrMapOvr"); err != nil {
		return err
	}
	if err := xw.EmptyElementWithAttrs("a", "defRgbClrModel", "val", "bg1"); err != nil {
		return err
	}
	if err := xw.EndElement("p", "clrMapOvr"); err != nil {
		return err
	}

	// 写入形状树
	if xs.CSld != nil && xs.CSld.SpTree != nil {
		if err := xs.CSld.SpTree.WriteXML(xw); err != nil {
			return err
		}
	}

	return xw.EndElement("p", "sld")
}

// WriteXML 将 XSpTree 序列化为 XML
func (xst *XSpTree) WriteXML(xw *XMLWriter) error {
	if err := xw.StartElement("p", "spTree"); err != nil {
		return err
	}

	// 写入非视觉组属性
	if err := xw.StartElement("p", "nvGrpSpPr"); err != nil {
		return err
	}
	if err := xw.StartElement("p", "cNvPr"); err != nil {
		return err
	}
	if err := xw.WriteInt(xst.NonVisual.CNvPr.ID); err != nil {
		return err
	}
	if xst.NonVisual.CNvPr.Name != "" {
		if err := xw.Raw(" name=\"" + xst.NonVisual.CNvPr.Name + "\""); err != nil {
			return err
		}
	}
	if err := xw.EndElement("p", "cNvPr"); err != nil {
		return err
	}
	if err := xw.StartElement("p", "cNvGrpSpPr"); err != nil {
		return err
	}
	if err := xw.EmptyElement("p", "cNvPr"); err != nil {
		return err
	}
	if err := xw.EndElement("p", "cNvGrpSpPr"); err != nil {
		return err
	}
	if err := xw.StartElement("p", "cNvPr"); err != nil {
		return err
	}
	if err := xw.EndElement("p", "cNvPr"); err != nil {
		return err
	}
	if err := xw.EndElement("p", "nvGrpSpPr"); err != nil {
		return err
	}

	// 写入组形状属性
	if xst.GroupShapeProperties != nil && xst.GroupShapeProperties.Xfrm != nil {
		if err := xst.GroupShapeProperties.Xfrm.WriteXML(xw); err != nil {
			return err
		}
	}

	// 写入子元素
	for _, child := range xst.Children {
		switch v := child.(type) {
		case *XSp:
			if err := v.WriteXML(xw); err != nil {
				return err
			}
		case *XPicture:
			if err := v.WriteXML(xw); err != nil {
				return err
			}
		case *XGraphicFrame:
			if err := v.WriteXML(xw); err != nil {
				return err
			}
		}
	}

	return xw.EndElement("p", "spTree")
}

// WriteXML 将 XTransform2D 序列化为 XML
func (xt *XTransform2D) WriteXML(xw *XMLWriter) error {
	if err := xw.StartElement("a", "xfrm"); err != nil {
		return err
	}

	// 写入偏移量
	if xt.Offset != nil {
		if err := xw.EmptyElementWithAttrs("a", "off", "x", strconv.Itoa(xt.Offset.X), "y", strconv.Itoa(xt.Offset.Y)); err != nil {
			return err
		}
	}

	// 写入扩展尺寸
	if xt.Extent != nil {
		if err := xw.EmptyElementWithAttrs("a", "ext", "cx", strconv.Itoa(xt.Extent.Cx), "cy", strconv.Itoa(xt.Extent.Cy)); err != nil {
			return err
		}
	}

	// 写入旋转角度
	if xt.Rotation != 0 {
		if err := xw.Raw(" rot=\"" + strconv.Itoa(xt.Rotation) + "\""); err != nil {
			return err
		}
	}

	return xw.EndElement("a", "xfrm")
}

// WriteXML 将 XShapeProperties 序列化为 XML
func (xsp *XShapeProperties) WriteXML(xw *XMLWriter) error {
	if err := xw.StartElement("p", "spPr"); err != nil {
		return err
	}

	// 写入变换
	if xsp.Transform2D != nil {
		if err := xsp.Transform2D.WriteXML(xw); err != nil {
			return err
		}
	}

	// 写入填充
	if xsp.PresetFill != nil {
		if err := xw.StartElement("a", "solidFill"); err != nil {
			return err
		}
		if xsp.PresetFill.SrgbClr != nil {
			if err := xw.EmptyElementWithAttrs("a", "srgbClr", "val", xsp.PresetFill.SrgbClr.Val); err != nil {
				return err
			}
		} else if xsp.PresetFill.SchemeClr != nil {
			if err := xw.EmptyElementWithAttrs("a", "schemeClr", "val", xsp.PresetFill.SchemeClr.Val); err != nil {
				return err
			}
		}
		if err := xw.EndElement("a", "solidFill"); err != nil {
			return err
		}
	}

	// 写入线条
	if xsp.Line != nil {
		if err := xw.StartElement("a", "ln"); err != nil {
			return err
		}
		if xsp.Line.Width != 0 {
			if err := xw.Raw(" w=\"" + strconv.Itoa(xsp.Line.Width) + "\""); err != nil {
				return err
			}
		}
		if xsp.Line.SolidFill != nil {
			if err := xw.StartElement("a", "solidFill"); err != nil {
				return err
			}
			if xsp.Line.SolidFill.SrgbClr != nil {
				if err := xw.EmptyElementWithAttrs("a", "srgbClr", "val", xsp.Line.SolidFill.SrgbClr.Val); err != nil {
					return err
				}
			}
			if err := xw.EndElement("a", "solidFill"); err != nil {
				return err
			}
		}
		if err := xw.EndElement("a", "ln"); err != nil {
			return err
		}
	}

	return xw.EndElement("p", "spPr")
}

// WriteXML 将 XSp 序列化为 XML
func (xs *XSp) WriteXML(xw *XMLWriter) error {
	if err := xw.StartElement("p", "sp"); err != nil {
		return err
	}

	// 写入非视觉属性
	if err := xw.StartElement("p", "nvSpPr"); err != nil {
		return err
	}
	if err := xw.StartElement("p", "cNvPr"); err != nil {
		return err
	}
	if err := xw.WriteInt(xs.NonVisual.CNvPr.ID); err != nil {
		return err
	}
	if xs.NonVisual.CNvPr.Name != "" {
		if err := xw.Raw(" name=\"" + xs.NonVisual.CNvPr.Name + "\""); err != nil {
			return err
		}
	}
	if err := xw.EndElement("p", "cNvPr"); err != nil {
		return err
	}
	if err := xw.StartElement("p", "cNvSpPr"); err != nil {
		return err
	}
	if err := xw.EmptyElement("p", "cNvPr"); err != nil {
		return err
	}
	if err := xw.EndElement("p", "cNvSpPr"); err != nil {
		return err
	}
	if err := xw.EndElement("p", "nvSpPr"); err != nil {
		return err
	}

	// 写入形状属性
	if xs.ShapeProperties != nil {
		if err := xs.ShapeProperties.WriteXML(xw); err != nil {
			return err
		}
	}

	// 写入文本内容
	if xs.TextBody != nil {
		if err := xs.TextBody.WriteXML(xw); err != nil {
			return err
		}
	}

	return xw.EndElement("p", "sp")
}

// WriteXML 将 XTextBody 序列化为 XML
func (xtb *XTextBody) WriteXML(xw *XMLWriter) error {
	if err := xw.StartElement("p", "txBody"); err != nil {
		return err
	}

	// 写入 body 属性
	if err := xw.StartElement("a", "bodyPr"); err != nil {
		return err
	}
	if xtb.BodyPr != nil {
		if xtb.BodyPr.Wrap != "" {
			if err := xw.Raw(" wrap=\"" + xtb.BodyPr.Wrap + "\""); err != nil {
				return err
			}
		}
		if xtb.BodyPr.Rotation != 0 {
			if err := xw.Raw(" rot=\"" + strconv.Itoa(xtb.BodyPr.Rotation) + "\""); err != nil {
				return err
			}
		}
		if xtb.BodyPr.Vertical != "" {
			if err := xw.Raw(" vert=\"" + xtb.BodyPr.Vertical + "\""); err != nil {
				return err
			}
		}
		if xtb.BodyPr.Anchor != "" {
			if err := xw.Raw(" anchor=\"" + xtb.BodyPr.Anchor + "\""); err != nil {
				return err
			}
		}
	}
	if err := xw.EndElement("a", "bodyPr"); err != nil {
		return err
	}

	// 写入列表样式
	if xtb.LstStyle != nil {
		if err := xw.EmptyElement("a", "lstStyle"); err != nil {
			return err
		}
	}

	// 写入段落
	for _, para := range xtb.Paragraphs {
		if err := para.WriteXML(xw); err != nil {
			return err
		}
	}

	return xw.EndElement("p", "txBody")
}

// WriteXML 将 XTextParagraph 序列化为 XML
func (xtp *XTextParagraph) WriteXML(xw *XMLWriter) error {
	if err := xw.StartElement("a", "p"); err != nil {
		return err
	}

	// 写入段落属性
	if xtp.Level != 0 {
		if err := xw.Raw(" lvl=\"" + strconv.Itoa(xtp.Level) + "\""); err != nil {
			return err
		}
	}
	if xtp.Indent != 0 {
		if err := xw.Raw(" indent=\"" + strconv.Itoa(xtp.Indent) + "\""); err != nil {
			return err
		}
	}
	if xtp.Alignment != "" {
		if err := xw.Raw(" algn=\"" + xtp.Alignment + "\""); err != nil {
			return err
		}
	}

	// 写入文本片段
	for _, run := range xtp.TextRuns {
		if err := run.WriteXML(xw); err != nil {
			return err
		}
	}

	return xw.EndElement("a", "p")
}

// WriteXML 将 XTextRun 序列化为 XML
func (xtr *XTextRun) WriteXML(xw *XMLWriter) error {
	if err := xw.StartElement("a", "r"); err != nil {
		return err
	}

	// 写入文本属性
	if xtr.TextProperties != nil {
		if err := xw.StartElement("a", "rPr"); err != nil {
			return err
		}
		props := xtr.TextProperties
		if props.FontSize != 0 {
			if err := xw.Raw(" sz=\"" + strconv.Itoa(props.FontSize) + "\""); err != nil {
				return err
			}
		}
		if props.Bold {
			if err := xw.Raw(" b=\"1\""); err != nil {
				return err
			}
		}
		if props.Italic {
			if err := xw.Raw(" i=\"1\""); err != nil {
				return err
			}
		}
		if props.Underline != "" {
			if err := xw.Raw(" u=\"" + props.Underline + "\""); err != nil {
				return err
			}
		}
		if props.FontFace != "" {
			if err := xw.Raw(" i=\"1\" lang=\"zh-CN\">"); err != nil {
				return err
			}
			if err := xw.EmptyElementWithAttrs("a", "latin", "typeface", props.FontFace); err != nil {
				return err
			}
		}
		if props.Color != "" {
			if err := xw.StartElement("a", "solidFill"); err != nil {
				return err
			}
			if err := xw.EmptyElementWithAttrs("a", "srgbClr", "val", props.Color); err != nil {
				return err
			}
			if err := xw.EndElement("a", "solidFill"); err != nil {
				return err
			}
		}
		if err := xw.EndElement("a", "rPr"); err != nil {
			return err
		}
	}

	// 写入文本内容
	if err := xw.StartElement("a", "t"); err != nil {
		return err
	}
	if err := xw.Text(xtr.Text); err != nil {
		return err
	}
	if err := xw.EndElement("a", "t"); err != nil {
		return err
	}

	return xw.EndElement("a", "r")
}

// WriteXML 将 XPicture 序列化为 XML
func (xp *XPicture) WriteXML(xw *XMLWriter) error {
	if err := xw.StartElement("p", "pic"); err != nil {
		return err
	}

	// 写入非视觉属性
	if err := xw.StartElement("p", "nvPicPr"); err != nil {
		return err
	}
	if err := xw.StartElement("p", "cNvPr"); err != nil {
		return err
	}
	if err := xw.WriteInt(xp.NonVisual.CNvPr.ID); err != nil {
		return err
	}
	if xp.NonVisual.CNvPr.Name != "" {
		if err := xw.Raw(" name=\"" + xp.NonVisual.CNvPr.Name + "\""); err != nil {
			return err
		}
	}
	if err := xw.EndElement("p", "cNvPr"); err != nil {
		return err
	}
	if err := xw.StartElement("p", "cNvPicPr"); err != nil {
		return err
	}
	if err := xw.EmptyElement("p", "cNvPr"); err != nil {
		return err
	}
	if err := xw.EndElement("p", "cNvPicPr"); err != nil {
		return err
	}
	if err := xw.EndElement("p", "nvPicPr"); err != nil {
		return err
	}

	// 写入图片填充
	if xp.BlipFill != nil {
		if err := xp.BlipFill.WriteXML(xw); err != nil {
			return err
		}
	}

	// 写入形状属性
	if xp.ShapeProperties != nil {
		if err := xp.ShapeProperties.WriteXML(xw); err != nil {
			return err
		}
	}

	return xw.EndElement("p", "pic")
}

// WriteXML 将 XBlipFillProperties 序列化为 XML
func (xbfp *XBlipFillProperties) WriteXML(xw *XMLWriter) error {
	if err := xw.StartElement("p", "blipFill"); err != nil {
		return err
	}

	// 写入 blip
	if xbfp.Blip != nil {
		if err := xw.StartElement("a", "blip"); err != nil {
			return err
		}
		if xbfp.Blip.Embed != "" {
			if err := xw.Raw(" r:embed=\"" + xbfp.Blip.Embed + "\""); err != nil {
				return err
			}
		}
		if err := xw.EndElement("a", "blip"); err != nil {
			return err
		}
	}

	// 写入拉伸填充
	if xbfp.Stretch != nil {
		if err := xw.StartElement("a", "stretch"); err != nil {
			return err
		}
		if err := xw.EmptyElement("a", "fillRect"); err != nil {
			return err
		}
		if err := xw.EndElement("a", "stretch"); err != nil {
			return err
		}
	}

	return xw.EndElement("p", "blipFill")
}

// WriteXML 将 XGraphicFrame 序列化为 XML
func (xgf *XGraphicFrame) WriteXML(xw *XMLWriter) error {
	if err := xw.StartElement("p", "graphicFrame"); err != nil {
		return err
	}

	// 写入非视觉属性
	if err := xw.StartElement("p", "nvGraphicFramePr"); err != nil {
		return err
	}
	if err := xw.StartElement("p", "cNvPr"); err != nil {
		return err
	}
	if err := xw.WriteInt(xgf.NonVisual.CNvPr.ID); err != nil {
		return err
	}
	if xgf.NonVisual.CNvPr.Name != "" {
		if err := xw.Raw(" name=\"" + xgf.NonVisual.CNvPr.Name + "\""); err != nil {
			return err
		}
	}
	if err := xw.EndElement("p", "cNvPr"); err != nil {
		return err
	}
	if err := xw.StartElement("p", "cNvGraphicFramePr"); err != nil {
		return err
	}
	if err := xw.EmptyElement("p", "cNvPr"); err != nil {
		return err
	}
	if err := xw.EndElement("p", "cNvGraphicFramePr"); err != nil {
		return err
	}
	if err := xw.EndElement("p", "nvGraphicFramePr"); err != nil {
		return err
	}

	// 写入变换
	if xgf.Transform2D != nil {
		if err := xgf.Transform2D.WriteXML(xw); err != nil {
			return err
		}
	}

	// 写入图形
	if xgf.Graphic != nil && xgf.Graphic.Table != nil {
		if err := xw.StartElement("a", "graphic"); err != nil {
			return err
		}
		if err := xgf.Graphic.Table.WriteXML(xw); err != nil {
			return err
		}
		if err := xw.EndElement("a", "graphic"); err != nil {
			return err
		}
	}

	return xw.EndElement("p", "graphicFrame")
}

// WriteXML 将 XTable 序列化为 XML
func (xt *XTable) WriteXML(xw *XMLWriter) error {
	if err := xw.StartElement("a", "graphicData"); err != nil {
		return err
	}
	if err := xw.Raw(" uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\""); err != nil {
		return err
	}

	if err := xw.StartElement("a", "tbl"); err != nil {
		return err
	}

	// 写入表格网格
	if err := xw.StartElement("a", "tblGrid"); err != nil {
		return err
	}
	for _, col := range xt.Grid.GridCols {
		if err := xw.EmptyElementWithAttrs("a", "gridCol", "w", strconv.Itoa(col.W)); err != nil {
			return err
		}
	}
	if err := xw.EndElement("a", "tblGrid"); err != nil {
		return err
	}

	// 写入行
	for _, row := range xt.Rows {
		if err := row.WriteXML(xw); err != nil {
			return err
		}
	}

	if err := xw.EndElement("a", "tbl"); err != nil {
		return err
	}
	if err := xw.EndElement("a", "graphicData"); err != nil {
		return err
	}

	return nil
}

// WriteXML 将 XTableRow 序列化为 XML
func (xtr *XTableRow) WriteXML(xw *XMLWriter) error {
	if err := xw.StartElement("a", "tr"); err != nil {
		return err
	}
	if xtr.GridSpan > 1 {
		if err := xw.Raw(" gridSpan=\"" + strconv.Itoa(xtr.GridSpan) + "\""); err != nil {
			return err
		}
	}

	// 写入单元格
	for _, cell := range xtr.Cells {
		if err := cell.WriteXML(xw); err != nil {
			return err
		}
	}

	return xw.EndElement("a", "tr")
}

// WriteXML 将 XTableCell 序列化为 XML
func (xtc *XTableCell) WriteXML(xw *XMLWriter) error {
	if err := xw.StartElement("a", "tc"); err != nil {
		return err
	}

	if xtc.GridSpan > 1 {
		if err := xw.Raw(" gridSpan=\"" + strconv.Itoa(xtc.GridSpan) + "\""); err != nil {
			return err
		}
	}
	if xtc.RowSpan > 1 {
		if err := xw.Raw(" rowSpan=\"" + strconv.Itoa(xtc.RowSpan) + "\""); err != nil {
			return err
		}
	}
	if xtc.Vertical != "" {
		if err := xw.Raw(" anchor=\"" + xtc.Vertical + "\""); err != nil {
			return err
		}
	}

	// 写入文本内容
	if xtc.TextBody != nil {
		if err := xtc.TextBody.WriteXML(xw); err != nil {
			return err
		}
	}

	return xw.EndElement("a", "tc")
}

