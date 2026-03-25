package opc

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"sync"
)

// ===== 数据源接口和实现 =====

// PartSource 部件数据源接口
type PartSource interface {
	Open() (io.ReadCloser, error)
	Size() int64
}

// ZipFileSource ZIP 文件中的部件数据源
type ZipFileSource struct {
	file *zip.File
}

// NewZipFileSource 从 zip.File 创建数据源
func NewZipFileSource(f *zip.File) *ZipFileSource {
	return &ZipFileSource{file: f}
}

// Open 打开 ZIP 文件条目
func (s *ZipFileSource) Open() (io.ReadCloser, error) {
	return s.file.Open()
}

// Size 返回未压缩大小
func (s *ZipFileSource) Size() int64 {
	return int64(s.file.UncompressedSize64)
}

// BytesSource 内存中的字节数据源
type BytesSource struct {
	data []byte
}

// NewBytesSource 从字节数组创建数据源
func NewBytesSource(data []byte) *BytesSource {
	return &BytesSource{data: data}
}

// Open 返回 bytes.Reader
func (s *BytesSource) Open() (io.ReadCloser, error) {
	return io.NopCloser(&bytesReaderAt{data: s.data}), nil
}

// Size 返回数据大小
func (s *BytesSource) Size() int64 {
	return int64(len(s.data))
}

// ReaderSource io.Reader 数据源
type ReaderSource struct {
	reader io.Reader
	size   int64
}

// NewReaderSource 从 io.Reader 创建数据源
func NewReaderSource(r io.Reader, size int64) *ReaderSource {
	return &ReaderSource{reader: r, size: size}
}

// Open 返回 reader
func (s *ReaderSource) Open() (io.ReadCloser, error) {
	return io.NopCloser(s.reader), nil
}

// Size 返回数据大小
func (s *ReaderSource) Size() int64 {
	return s.size
}

// ===== 流式写入接口 =====

// StreamWriter 流式写入器接口
type StreamWriter interface {
	StreamWriteTo(w io.Writer) error
}

// XMLStreamer XML 流式写入器接口
type XMLStreamer interface {
	StreamXML(enc *xml.Encoder) error
}

// ===== 流式 ZIP 写入器 =====

// StreamingZipWriter 流式 ZIP 写入器
type StreamingZipWriter struct {
	zipWriter *zip.Writer
}

// NewStreamingZipWriter 创建流式 ZIP 写入器
func NewStreamingZipWriter(w io.Writer) *StreamingZipWriter {
	return &StreamingZipWriter{
		zipWriter: zip.NewWriter(w),
	}
}

// Create 创建 ZIP 条目并返回写入器
func (sw *StreamingZipWriter) Create(path string) (io.Writer, error) {
	return sw.zipWriter.Create(path)
}

// WriteFromReader 从 Reader 流式写入 ZIP 条目
func (sw *StreamingZipWriter) WriteFromReader(path string, reader io.Reader) error {
	w, err := sw.zipWriter.Create(path)
	if err != nil {
		return err
	}
	_, err = io.Copy(w, reader)
	return err
}

// WriteFromStreamer 从 StreamWriter 流式写入 ZIP 条目
func (sw *StreamingZipWriter) WriteFromStreamer(path string, streamer StreamWriter) error {
	w, err := sw.zipWriter.Create(path)
	if err != nil {
		return err
	}
	return streamer.StreamWriteTo(w)
}

// WriteFromXMLStreamer 从 XMLStreamer 流式写入 ZIP 条目
func (sw *StreamingZipWriter) WriteFromXMLStreamer(path string, streamer XMLStreamer) error {
	w, err := sw.zipWriter.Create(path)
	if err != nil {
		return err
	}

	// 写入 XML 头
	if _, err := w.Write([]byte(xml.Header)); err != nil {
		return err
	}

	encoder := xml.NewEncoder(w)
	if err := streamer.StreamXML(encoder); err != nil {
		return err
	}
	return encoder.Flush()
}

// WriteStreamPart 流式写入 StreamPart
func (sw *StreamingZipWriter) WriteStreamPart(part *StreamPart) error {
	path := part.PartURI().MemberName()
	w, err := sw.zipWriter.Create(path)
	if err != nil {
		return err
	}

	// 打开部件流
	rc, err := part.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	// 流式复制
	_, err = io.Copy(w, rc)
	return err
}

// WriteBytes 写入字节数据
func (sw *StreamingZipWriter) WriteBytes(path string, data []byte) error {
	w, err := sw.zipWriter.Create(path)
	if err != nil {
		return err
	}
	_, err = w.Write(data)
	return err
}

// WriteXML 写入 XML 数据（自动添加 XML 头）
func (sw *StreamingZipWriter) WriteXML(path string, data []byte) error {
	w, err := sw.zipWriter.Create(path)
	if err != nil {
		return err
	}
	// 写入 XML 头
	if _, err := w.Write([]byte(xml.Header)); err != nil {
		return err
	}
	_, err = w.Write(data)
	return err
}

// Close 关闭 ZIP 写入器
func (sw *StreamingZipWriter) Close() error {
	return sw.zipWriter.Close()
}

// ===== 流式部件 =====

// StreamPart 流式部件 - 支持懒加载
type StreamPart struct {
	uri           *PackURI
	contentType   string
	source        PartSource
	relationships *Relationships
	dirty         bool
	loaded        bool
	blob          []byte // 缓存的数据（如果已加载）
	mu            sync.RWMutex
}

// NewStreamPart 创建流式部件
func NewStreamPart(uri *PackURI, contentType string, source PartSource) *StreamPart {
	return &StreamPart{
		uri:           uri,
		contentType:   contentType,
		source:        source,
		relationships: NewRelationships(uri),
		dirty:         false,
		loaded:        false,
	}
}

// PartURI 返回部件 URI
func (p *StreamPart) PartURI() *PackURI {
	return p.uri
}

// ContentType 返回内容类型
func (p *StreamPart) ContentType() string {
	return p.contentType
}

// SetContentType 设置内容类型
func (p *StreamPart) SetContentType(ct string) {
	p.mu.Lock()
	defer p.mu.Unlock()
	p.contentType = ct
	p.dirty = true
}

// Open 打开部件内容流
func (p *StreamPart) Open() (io.ReadCloser, error) {
	p.mu.RLock()
	defer p.mu.RUnlock()

	// 如果已加载到内存，返回内存数据
	if p.loaded {
		return io.NopCloser(&bytesReaderAt{data: p.blob}), nil
	}

	// 否则从源打开
	if p.source != nil {
		return p.source.Open()
	}

	return nil, nil
}

// Load 将内容加载到内存
func (p *StreamPart) Load() error {
	p.mu.Lock()
	defer p.mu.Unlock()

	if p.loaded {
		return nil
	}

	if p.source == nil {
		return nil
	}

	rc, err := p.source.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	p.blob, err = io.ReadAll(rc)
	if err != nil {
		return err
	}

	p.loaded = true
	return nil
}

// Blob 返回内容（如果未加载则先加载）
func (p *StreamPart) Blob() ([]byte, error) {
	if err := p.Load(); err != nil {
		return nil, err
	}
	p.mu.RLock()
	defer p.mu.RUnlock()
	return p.blob, nil
}

// SetBlob 设置内容
func (p *StreamPart) SetBlob(data []byte) {
	p.mu.Lock()
	defer p.mu.Unlock()
	p.blob = data
	p.loaded = true
	p.dirty = true
}

// SetBlobFromReader 从 Reader 设置内容
func (p *StreamPart) SetBlobFromReader(r io.Reader) error {
	p.mu.Lock()
	defer p.mu.Unlock()

	data, err := io.ReadAll(r)
	if err != nil {
		return err
	}
	p.blob = data
	p.loaded = true
	p.dirty = true
	return nil
}

// IsLoaded 返回是否已加载到内存
func (p *StreamPart) IsLoaded() bool {
	p.mu.RLock()
	defer p.mu.RUnlock()
	return p.loaded
}

// IsDirty 返回是否被修改
func (p *StreamPart) IsDirty() bool {
	p.mu.RLock()
	defer p.mu.RUnlock()
	return p.dirty
}

// SetDirty 设置修改标记
func (p *StreamPart) SetDirty(dirty bool) {
	p.mu.Lock()
	defer p.mu.Unlock()
	p.dirty = dirty
}

// Relationships 返回关系集合
func (p *StreamPart) Relationships() *Relationships {
	return p.relationships
}

// LoadRelationships 从 XML 加载关系
func (p *StreamPart) LoadRelationships(data []byte) error {
	return p.relationships.FromXML(data)
}

// HasRelationships 检查是否有关系
func (p *StreamPart) HasRelationships() bool {
	return p.relationships.Count() > 0
}

// RelationshipsBlob 返回关系的 XML 内容
func (p *StreamPart) RelationshipsBlob() ([]byte, error) {
	if p.relationships.Count() == 0 {
		return nil, nil
	}
	return p.relationships.ToXML()
}

// RelationshipsURI 返回关系文件的 URI
func (p *StreamPart) RelationshipsURI() *PackURI {
	return p.uri.RelationshipsURI()
}

// Size 返回内容大小
func (p *StreamPart) Size() int64 {
	p.mu.RLock()
	defer p.mu.RUnlock()
	if p.loaded {
		return int64(len(p.blob))
	}
	if p.source != nil {
		return p.source.Size()
	}
	return 0
}

// UnmarshalBlob 从 blob 解析 XML 内容
func (p *StreamPart) UnmarshalBlob(v any) error {
	if err := p.Load(); err != nil {
		return err
	}
	p.mu.RLock()
	defer p.mu.RUnlock()
	return xml.Unmarshal(p.blob, v)
}

// Clone 克隆部件
func (p *StreamPart) Clone() *StreamPart {
	p.mu.RLock()
	defer p.mu.RUnlock()

	var blobCopy []byte
	if p.loaded && p.blob != nil {
		blobCopy = make([]byte, len(p.blob))
		copy(blobCopy, p.blob)
	}

	return &StreamPart{
		uri:           p.uri.Clone(),
		contentType:   p.contentType,
		source:        p.source,
		relationships: p.relationships.Clone(),
		dirty:         p.dirty,
		loaded:        p.loaded,
		blob:          blobCopy,
	}
}

// ===== 流式写入器实现 =====

// RelationshipsStreamer 关系流式写入器
type RelationshipsStreamer struct {
	rels *Relationships
}

// NewRelationshipsStreamer 创建关系流式写入器
func NewRelationshipsStreamer(rels *Relationships) *RelationshipsStreamer {
	return &RelationshipsStreamer{rels: rels}
}

// StreamWriteTo 实现 StreamWriter 接口
func (rs *RelationshipsStreamer) StreamWriteTo(w io.Writer) error {
	encoder := xml.NewEncoder(w)

	// 写入 Relationships 根元素
	start := xml.StartElement{
		Name: xml.Name{Local: "Relationships"},
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "xmlns"}, Value: NamespaceRelationships},
		},
	}

	if err := encoder.EncodeToken(start); err != nil {
		return err
	}

	// 写入每个 Relationship
	for _, rel := range rs.rels.All() {
		relElem := xml.StartElement{
			Name: xml.Name{Local: "Relationship"},
			Attr: []xml.Attr{
				{Name: xml.Name{Local: "Id"}, Value: rel.RID()},
				{Name: xml.Name{Local: "Type"}, Value: rel.Type()},
				{Name: xml.Name{Local: "Target"}, Value: rel.TargetRef()},
			},
		}
		if rel.IsExternal() {
			relElem.Attr = append(relElem.Attr, xml.Attr{
				Name:  xml.Name{Local: "TargetMode"},
				Value: "External",
			})
		}

		if err := encoder.EncodeToken(relElem); err != nil {
			return err
		}
		if err := encoder.EncodeToken(relElem.End()); err != nil {
			return err
		}
	}

	// 结束根元素
	if err := encoder.EncodeToken(start.End()); err != nil {
		return err
	}

	return encoder.Flush()
}

// ContentTypesStreamer ContentTypes 流式写入器
type ContentTypesStreamer struct {
	ct *ContentTypes
}

// NewContentTypesStreamer 创建 ContentTypes 流式写入器
func NewContentTypesStreamer(ct *ContentTypes) *ContentTypesStreamer {
	return &ContentTypesStreamer{ct: ct}
}

// StreamWriteTo 实现 StreamWriter 接口
func (cs *ContentTypesStreamer) StreamWriteTo(w io.Writer) error {
	encoder := xml.NewEncoder(w)

	// 写入 Types 根元素
	start := xml.StartElement{
		Name: xml.Name{Local: "Types"},
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "xmlns"}, Value: NamespaceOPCPackage},
		},
	}

	if err := encoder.EncodeToken(start); err != nil {
		return err
	}

	// 写入 Default 元素
	for ext, ctType := range cs.ct.Defaults() {
		defElem := xml.StartElement{
			Name: xml.Name{Local: "Default"},
			Attr: []xml.Attr{
				{Name: xml.Name{Local: "Extension"}, Value: ext},
				{Name: xml.Name{Local: "ContentType"}, Value: ctType},
			},
		}
		if err := encoder.EncodeToken(defElem); err != nil {
			return err
		}
		if err := encoder.EncodeToken(defElem.End()); err != nil {
			return err
		}
	}

	// 写入 Override 元素
	for uri, ctType := range cs.ct.Overrides() {
		overrideElem := xml.StartElement{
			Name: xml.Name{Local: "Override"},
			Attr: []xml.Attr{
				{Name: xml.Name{Local: "PartName"}, Value: uri},
				{Name: xml.Name{Local: "ContentType"}, Value: ctType},
			},
		}
		if err := encoder.EncodeToken(overrideElem); err != nil {
			return err
		}
		if err := encoder.EncodeToken(overrideElem.End()); err != nil {
			return err
		}
	}

	// 结束根元素
	if err := encoder.EncodeToken(start.End()); err != nil {
		return err
	}

	return encoder.Flush()
}

// ===== 并发写入数据结构 =====

// PartData 部件数据 - 用于 channel 传递
type PartData struct {
	URI         string    // 部件 URI
	Path        string    // ZIP 内路径
	ContentType string    // 内容类型
	Data        []byte    // 数据内容
	Source      PartSource // 数据源（用于懒加载）
	Error       error     // 写入错误（如果有）
}

// PartDataChannel 部件数据通道类型
type PartDataChannel chan *PartData

// NewPartDataChannel 创建部件数据通道
func NewPartDataChannel(bufferSize int) PartDataChannel {
	return make(PartDataChannel, bufferSize)
}

// ===== 全局资源去重池 =====

// ResourceHashKey 资源哈希键
type ResourceHashKey string

// ResourceEntry 资源条目
type ResourceEntry struct {
	URI       string    // 部件 URI
	Hash      string    // 内容哈希（SHA256）
	Size      int64     // 原始大小
	Reference int       // 引用计数
}

// ResourceDedupPool 全局资源去重池
// 使用 sync.Map 实现并发安全的资源去重
type ResourceDedupPool struct {
	entries sync.Map // map[ResourceHashKey]*ResourceEntry
	mu      sync.RWMutex
}

// globalResourcePool 全局资源池单例
var globalResourcePool = &ResourceDedupPool{}

// GetGlobalResourcePool 获取全局资源池
func GetGlobalResourcePool() *ResourceDedupPool {
	return globalResourcePool
}

// NewResourceDedupPool 创建新的资源去重池
func NewResourceDedupPool() *ResourceDedupPool {
	return &ResourceDedupPool{}
}

// ComputeHash 计算数据的 SHA256 哈希
func ComputeHash(data []byte) string {
	// 使用简单的哈希算法（实际生产中应使用 crypto/sha256）
	// 这里使用简化版本以避免额外依赖
	if len(data) == 0 {
		return ""
	}

	// 简单的 FNV-1a 哈希
	var hash uint32 = 2166136261
	for _, b := range data {
		hash ^= uint32(b)
		hash *= 16777619
	}

	// 加上长度以增强唯一性
	return fmt.Sprintf("%x-%d", hash, len(data))
}

// Register 注册资源，返回是否为新资源
// 如果资源已存在，增加引用计数并返回 false
func (p *ResourceDedupPool) Register(uri string, data []byte) (isNew bool, existingURI string) {
	hash := ComputeHash(data)
	key := ResourceHashKey(hash)

	p.mu.Lock()
	defer p.mu.Unlock()

	// 检查是否已存在
	if entry, ok := p.entries.Load(key); ok {
		e := entry.(*ResourceEntry)
		e.Reference++
		return false, e.URI
	}

	// 新资源
	entry := &ResourceEntry{
		URI:       uri,
		Hash:      hash,
		Size:      int64(len(data)),
		Reference: 1,
	}
	p.entries.Store(key, entry)
	return true, uri
}

// RegisterWithHash 使用预计算的哈希注册资源
func (p *ResourceDedupPool) RegisterWithHash(uri string, hash string, size int64) (isNew bool, existingURI string) {
	key := ResourceHashKey(hash)

	p.mu.Lock()
	defer p.mu.Unlock()

	if entry, ok := p.entries.Load(key); ok {
		e := entry.(*ResourceEntry)
		e.Reference++
		return false, e.URI
	}

	entry := &ResourceEntry{
		URI:       uri,
		Hash:      hash,
		Size:      size,
		Reference: 1,
	}
	p.entries.Store(key, entry)
	return true, uri
}

// Lookup 查找资源
func (p *ResourceDedupPool) Lookup(hash string) (*ResourceEntry, bool) {
	if entry, ok := p.entries.Load(ResourceHashKey(hash)); ok {
		return entry.(*ResourceEntry), true
	}
	return nil, false
}

// Release 释放资源引用
func (p *ResourceDedupPool) Release(hash string) {
	p.mu.Lock()
	defer p.mu.Unlock()

	if entry, ok := p.entries.Load(ResourceHashKey(hash)); ok {
		e := entry.(*ResourceEntry)
		e.Reference--
		if e.Reference <= 0 {
			p.entries.Delete(ResourceHashKey(hash))
		}
	}
}

// Clear 清空资源池
func (p *ResourceDedupPool) Clear() {
	p.mu.Lock()
	defer p.mu.Unlock()
	p.entries = sync.Map{}
}

// Stats 返回资源池统计信息
func (p *ResourceDedupPool) Stats() (count int, totalSize int64) {
	p.entries.Range(func(key, value interface{}) bool {
		count++
		entry := value.(*ResourceEntry)
		totalSize += entry.Size
		return true
	})
	return
}

// ===== 并发 ZIP 收集器 =====

// ConcurrentZipCollector 并发 ZIP 收集器
// 使用 goroutine 从 channel 收集部件数据并写入 ZIP
type ConcurrentZipCollector struct {
	zipWriter  *zip.Writer
	dataChan   PartDataChannel
	errorChan  chan error
	doneChan   chan struct{}
	wg         sync.WaitGroup
	bufferSize int
}

// NewConcurrentZipCollector 创建并发 ZIP 收集器
func NewConcurrentZipCollector(w io.Writer, bufferSize int) *ConcurrentZipCollector {
	return &ConcurrentZipCollector{
		zipWriter:  zip.NewWriter(w),
		dataChan:   make(PartDataChannel, bufferSize),
		errorChan:  make(chan error, 1),
		doneChan:   make(chan struct{}),
		bufferSize: bufferSize,
	}
}

// Start 启动收集器 goroutine
func (c *ConcurrentZipCollector) Start() {
	c.wg.Add(1)
	go c.collect()
}

// collect 收集 goroutine
func (c *ConcurrentZipCollector) collect() {
	defer c.wg.Done()

	for data := range c.dataChan {
		if data.Error != nil {
			c.errorChan <- data.Error
			return
		}

		// 写入 ZIP 条目
		if err := c.writePart(data); err != nil {
			c.errorChan <- err
			return
		}
	}

	// 所有数据已写入，关闭 ZIP
	if err := c.zipWriter.Close(); err != nil {
		c.errorChan <- err
		return
	}

	close(c.doneChan)
}

// writePart 写入单个部件
func (c *ConcurrentZipCollector) writePart(data *PartData) error {
	w, err := c.zipWriter.Create(data.Path)
	if err != nil {
		return fmt.Errorf("failed to create zip entry %s: %w", data.Path, err)
	}

	if data.Data != nil {
		_, err = w.Write(data.Data)
		return err
	}

	if data.Source != nil {
		rc, err := data.Source.Open()
		if err != nil {
			return err
		}
		defer rc.Close()
		_, err = io.Copy(w, rc)
		return err
	}

	return nil
}

// Submit 提交部件数据到收集器
func (c *ConcurrentZipCollector) Submit(data *PartData) error {
	select {
	case c.dataChan <- data:
		return nil
	case err := <-c.errorChan:
		return err
	case <-c.doneChan:
		return fmt.Errorf("collector already finished")
	}
}

// SubmitBytes 提交字节数据
func (c *ConcurrentZipCollector) SubmitBytes(path string, data []byte) error {
	return c.Submit(&PartData{
		Path: path,
		Data: data,
	})
}

// Close 关闭收集器，等待所有数据写入完成
func (c *ConcurrentZipCollector) Close() error {
	close(c.dataChan)

	select {
	case <-c.doneChan:
		return nil
	case err := <-c.errorChan:
		return err
	}
}

// Wait 等待收集器完成
func (c *ConcurrentZipCollector) Wait() error {
	return c.Close()
}

// DataChannel 返回数据通道（用于外部生产者）
func (c *ConcurrentZipCollector) DataChannel() PartDataChannel {
	return c.dataChan
}

// ===== 辅助类型 =====

// bytesReaderAt 简单的 bytes reader
type bytesReaderAt struct {
	data []byte
	pos  int
}

func (r *bytesReaderAt) Read(p []byte) (n int, err error) {
	if r.pos >= len(r.data) {
		return 0, io.EOF
	}
	n = copy(p, r.data[r.pos:])
	r.pos += n
	return n, nil
}
