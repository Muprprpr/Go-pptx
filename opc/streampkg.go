package opc

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"path"
	"strings"
	"sync"
)

// StreamPackage 流式 OPC 包 - 支持懒加载和流式写入
type StreamPackage struct {
	parts          map[string]*StreamPart // URI -> StreamPart
	partOrder      []string               // 保持插入顺序
	relationships  *Relationships
	contentTypes   *ContentTypes
	coreProperties *CoreProperties
	zipReader      *zip.Reader
	zipFile        *os.File // 保持文件打开以便懒加载
	mu             sync.RWMutex
}

// NewStreamPackage 创建新的流式包
func NewStreamPackage() *StreamPackage {
	return &StreamPackage{
		parts:         make(map[string]*StreamPart),
		partOrder:     make([]string, 0),
		relationships: NewRelationships(RootURI()),
		contentTypes:  NewContentTypes(),
	}
}

// OpenStream 流式打开 OPC 包
// 文件句柄会保持打开状态以支持懒加载
func OpenStream(path string) (*StreamPackage, error) {
	file, err := os.Open(path)
	if err != nil {
		return nil, fmt.Errorf("failed to open file: %w", err)
	}

	stat, err := file.Stat()
	if err != nil {
		file.Close()
		return nil, fmt.Errorf("failed to get file info: %w", err)
	}

	zipReader, err := zip.NewReader(file, stat.Size())
	if err != nil {
		file.Close()
		return nil, fmt.Errorf("failed to open zip: %w", err)
	}

	pkg := &StreamPackage{
		parts:         make(map[string]*StreamPart),
		partOrder:     make([]string, 0),
		relationships: NewRelationships(RootURI()),
		contentTypes:  NewContentTypes(),
		zipReader:     zipReader,
		zipFile:       file,
	}

	// 只加载元数据，不加载实际内容
	if err := pkg.loadContentTypes(); err != nil {
		file.Close()
		return nil, fmt.Errorf("failed to load content types: %w", err)
	}

	if err := pkg.loadPartMetadata(); err != nil {
		file.Close()
		return nil, fmt.Errorf("failed to load part metadata: %w", err)
	}

	if err := pkg.loadRelationships(); err != nil {
		file.Close()
		return nil, fmt.Errorf("failed to load relationships: %w", err)
	}

	return pkg, nil
}

// OpenStreamFromReader 从 io.ReaderAt 流式打开
// 注意：调用者负责保持 ReaderAt 的有效性
func OpenStreamFromReader(r io.ReaderAt, size int64) (*StreamPackage, error) {
	zipReader, err := zip.NewReader(r, size)
	if err != nil {
		return nil, fmt.Errorf("failed to open zip: %w", err)
	}

	pkg := &StreamPackage{
		parts:         make(map[string]*StreamPart),
		partOrder:     make([]string, 0),
		relationships: NewRelationships(RootURI()),
		contentTypes:  NewContentTypes(),
		zipReader:     zipReader,
	}

	if err := pkg.loadContentTypes(); err != nil {
		return nil, fmt.Errorf("failed to load content types: %w", err)
	}

	if err := pkg.loadPartMetadata(); err != nil {
		return nil, fmt.Errorf("failed to load part metadata: %w", err)
	}

	if err := pkg.loadRelationships(); err != nil {
		return nil, fmt.Errorf("failed to load relationships: %w", err)
	}

	return pkg, nil
}

// loadContentTypes 加载内容类型（必须立即加载）
func (p *StreamPackage) loadContentTypes() error {
	for _, f := range p.zipReader.File {
		if f.Name == PathContentTypes {
			rc, err := f.Open()
			if err != nil {
				return err
			}
			data, err := io.ReadAll(rc)
			rc.Close()
			if err != nil {
				return err
			}
			return p.contentTypes.FromXML(data)
		}
	}
	return fmt.Errorf("[Content_Types].xml not found")
}

// loadPartMetadata 只加载部件元数据，不加载内容
func (p *StreamPackage) loadPartMetadata() error {
	for _, f := range p.zipReader.File {
		// 跳过特殊文件
		if f.Name == PathContentTypes {
			continue
		}
		if strings.Contains(f.Name, PathRelsDir+"/") && strings.HasSuffix(f.Name, ".rels") {
			continue
		}
		if strings.HasSuffix(f.Name, "/") {
			continue
		}

		uri := NewPackURI("/" + f.Name)
		contentType := p.contentTypes.GetContentType(uri)

		// 创建流式部件，使用 ZipFileSource 实现懒加载
		part := NewStreamPart(uri, contentType, NewZipFileSource(f))

		p.parts[uri.URI()] = part
		p.partOrder = append(p.partOrder, uri.URI())
	}

	return nil
}

// loadRelationships 加载所有关系
func (p *StreamPackage) loadRelationships() error {
	for _, f := range p.zipReader.File {
		if !strings.Contains(f.Name, PathRelsDir+"/") || !strings.HasSuffix(f.Name, ".rels") {
			continue
		}

		rc, err := f.Open()
		if err != nil {
			return err
		}
		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			return err
		}

		relURI := NewPackURI("/" + f.Name)
		sourceURI := relURI.SourceURI()

		rels := NewRelationships(sourceURI)
		if err := rels.FromXML(data); err != nil {
			return err
		}

		if sourceURI.IsPackageRels() {
			p.relationships = rels
		} else {
			part := p.parts[sourceURI.URI()]
			if part != nil {
				part.LoadRelationships(data)
			}
		}
	}

	return nil
}

// ===== 部件访问 =====

// GetPart 获取部件（内容按需加载）
func (p *StreamPackage) GetPart(uri *PackURI) *StreamPart {
	p.mu.RLock()
	defer p.mu.RUnlock()
	return p.parts[uri.URI()]
}

// GetPartByStr 根据字符串 URI 获取部件
func (p *StreamPackage) GetPartByStr(uri string) *StreamPart {
	return p.GetPart(NewPackURI(uri))
}

// ContainsPart 检查部件是否存在
func (p *StreamPackage) ContainsPart(uri *PackURI) bool {
	p.mu.RLock()
	defer p.mu.RUnlock()
	_, exists := p.parts[uri.URI()]
	return exists
}

// AllParts 返回所有部件
func (p *StreamPackage) AllParts() []*StreamPart {
	p.mu.RLock()
	defer p.mu.RUnlock()

	result := make([]*StreamPart, 0, len(p.partOrder))
	for _, uri := range p.partOrder {
		result = append(result, p.parts[uri])
	}
	return result
}

// PartCount 返回部件数量
func (p *StreamPackage) PartCount() int {
	p.mu.RLock()
	defer p.mu.RUnlock()
	return len(p.parts)
}

// PartURIs 返回所有部件 URI
func (p *StreamPackage) PartURIs() []*PackURI {
	p.mu.RLock()
	defer p.mu.RUnlock()

	result := make([]*PackURI, 0, len(p.partOrder))
	for _, uri := range p.partOrder {
		result = append(result, NewPackURI(uri))
	}
	return result
}

// GetPartsByType 根据内容类型获取部件
func (p *StreamPackage) GetPartsByType(contentType string) []*StreamPart {
	p.mu.RLock()
	defer p.mu.RUnlock()

	var result []*StreamPart
	for _, uri := range p.partOrder {
		if part := p.parts[uri]; part.ContentType() == contentType {
			result = append(result, part)
		}
	}
	return result
}

// ===== 部件管理 =====

// AddPart 添加流式部件
func (p *StreamPackage) AddPart(part *StreamPart) error {
	p.mu.Lock()
	defer p.mu.Unlock()

	uri := part.PartURI().URI()
	if _, exists := p.parts[uri]; exists {
		return fmt.Errorf("part with URI %s already exists", uri)
	}

	p.parts[uri] = part
	p.partOrder = append(p.partOrder, uri)
	return nil
}

// CreateStreamPart 创建并添加流式部件
func (p *StreamPackage) CreateStreamPart(uri *PackURI, contentType string, source PartSource) (*StreamPart, error) {
	part := NewStreamPart(uri, contentType, source)
	if err := p.AddPart(part); err != nil {
		return nil, err
	}
	return part, nil
}

// CreatePartFromBytes 从字节创建部件（立即加载到内存）
func (p *StreamPackage) CreatePartFromBytes(uri *PackURI, contentType string, data []byte) (*StreamPart, error) {
	part := NewStreamPart(uri, contentType, NewBytesSource(data))
	part.SetBlob(data) // 立即加载到内存
	if err := p.AddPart(part); err != nil {
		return nil, err
	}
	return part, nil
}

// RemovePart 移除部件
func (p *StreamPackage) RemovePart(uri *PackURI) error {
	p.mu.Lock()
	defer p.mu.Unlock()

	key := uri.URI()
	if _, exists := p.parts[key]; !exists {
		return fmt.Errorf("part with URI %s not found", key)
	}

	delete(p.parts, key)

	// 从 order 中移除
	for i, u := range p.partOrder {
		if u == key {
			p.partOrder = append(p.partOrder[:i], p.partOrder[i+1:]...)
			break
		}
	}

	return nil
}

// ===== 关系管理 =====

// Relationships 返回包级别关系
func (p *StreamPackage) Relationships() *Relationships {
	return p.relationships
}

// AddRelationship 添加包级别关系
func (p *StreamPackage) AddRelationship(relType, targetURI string, isExternal bool) (*Relationship, error) {
	return p.relationships.AddNew(relType, targetURI, isExternal)
}

// GetPartByRelType 通过关系类型获取目标部件
func (p *StreamPackage) GetPartByRelType(relType string) *StreamPart {
	rels := p.relationships.GetByType(relType)
	if len(rels) == 0 {
		return nil
	}
	return p.parts[rels[0].TargetURI().URI()]
}

// ===== 流式保存 =====

// StreamSave 流式保存到 io.Writer
func (p *StreamPackage) StreamSave(w io.Writer) error {
	sw := NewStreamingZipWriter(w)

	// 1. 写入 [Content_Types].xml（流式）
	if err := p.streamWriteContentTypes(sw); err != nil {
		return fmt.Errorf("failed to write content types: %w", err)
	}

	// 2. 写入包级别关系（流式）
	if err := p.streamWritePackageRelationships(sw); err != nil {
		return fmt.Errorf("failed to write package relationships: %w", err)
	}

	// 3. 流式写入所有部件
	if err := p.streamWriteParts(sw); err != nil {
		return fmt.Errorf("failed to write parts: %w", err)
	}

	// 4. 写入核心属性（如果有）
	if p.coreProperties != nil {
		if err := p.streamWriteCoreProperties(sw); err != nil {
			return fmt.Errorf("failed to write core properties: %w", err)
		}
	}

	return sw.Close()
}

// StreamSaveFile 流式保存到文件
func (p *StreamPackage) StreamSaveFile(path string) error {
	file, err := os.Create(path)
	if err != nil {
		return fmt.Errorf("failed to create file: %w", err)
	}
	defer file.Close()

	return p.StreamSave(file)
}

func (p *StreamPackage) streamWriteContentTypes(sw *StreamingZipWriter) error {
	p.updateContentTypes()
	streamer := NewContentTypesStreamer(p.contentTypes)
	return sw.WriteFromStreamer(PathContentTypes, streamer)
}

func (p *StreamPackage) streamWritePackageRelationships(sw *StreamingZipWriter) error {
	if p.relationships.Count() == 0 {
		return nil
	}

	streamer := NewRelationshipsStreamer(p.relationships)
	relPath := PathRelsDir + "/" + PathRelsFile
	return sw.WriteFromStreamer(relPath, streamer)
}

func (p *StreamPackage) streamWriteParts(sw *StreamingZipWriter) error {
	p.mu.RLock()
	defer p.mu.RUnlock()

	for _, uri := range p.partOrder {
		part := p.parts[uri]

		// 流式写入部件内容
		if err := sw.WriteStreamPart(part); err != nil {
			return err
		}

		// 流式写入关系（如果有）
		if part.HasRelationships() {
			relPath := p.relFilePath(part.PartURI())
			streamer := NewRelationshipsStreamer(part.Relationships())
			if err := sw.WriteFromStreamer(relPath, streamer); err != nil {
				return err
			}
		}
	}

	return nil
}

func (p *StreamPackage) streamWriteCoreProperties(sw *StreamingZipWriter) error {
	data, err := p.coreProperties.ToXML()
	if err != nil {
		return err
	}
	return sw.WriteFromReader("docProps/core.xml", &bytesReaderAt{data: data})
}

func (p *StreamPackage) relFilePath(uri *PackURI) string {
	dir := path.Dir(strings.TrimPrefix(uri.URI(), "/"))
	filename := path.Base(uri.URI())
	return path.Join(dir, PathRelsDir, filename+".rels")
}

func (p *StreamPackage) updateContentTypes() {
	for _, uri := range p.partOrder {
		part := p.parts[uri]
		contentType := part.ContentType()
		packURI := part.PartURI()

		if packURI.IsRelationshipsPart() {
			contentType = ContentTypeRelationships
		}

		ext := packURI.Extension()
		defaultCT := p.contentTypes.GetDefault(ext)

		if contentType != "" && contentType != ContentTypeDefault {
			if defaultCT == "" || defaultCT == ContentTypeDefault || defaultCT != contentType {
				p.contentTypes.AddOverride(packURI, contentType)
			}
		}
	}
}

// ===== 其他方法 =====

// ContentTypes 返回内容类型定义
func (p *StreamPackage) ContentTypes() *ContentTypes {
	return p.contentTypes
}

// CoreProperties 返回核心属性
func (p *StreamPackage) CoreProperties() *CoreProperties {
	p.mu.RLock()
	defer p.mu.RUnlock()
	return p.coreProperties
}

// SetCoreProperties 设置核心属性
func (p *StreamPackage) SetCoreProperties(props *CoreProperties) {
	p.mu.Lock()
	defer p.mu.Unlock()
	p.coreProperties = props
}

// Close 关闭包，释放资源
func (p *StreamPackage) Close() error {
	p.mu.Lock()
	defer p.mu.Unlock()

	// 关闭底层文件
	if p.zipFile != nil {
		err := p.zipFile.Close()
		p.zipFile = nil
		return err
	}

	return nil
}

// ===== 懒加载迭代器 =====

// PartIterator 部件迭代器
type PartIterator struct {
	pkg    *StreamPackage
	index  int
	filter func(*StreamPart) bool
}

// NewPartIterator 创建部件迭代器
func (p *StreamPackage) NewPartIterator() *PartIterator {
	return &PartIterator{
		pkg:    p,
		index:  0,
		filter: nil,
	}
}

// FilterByType 按内容类型过滤
func (it *PartIterator) FilterByType(contentType string) *PartIterator {
	it.filter = func(part *StreamPart) bool {
		return part.ContentType() == contentType
	}
	return it
}

// Next 移动到下一个部件
func (it *PartIterator) Next() bool {
	it.pkg.mu.RLock()
	defer it.pkg.mu.RUnlock()

	for it.index < len(it.pkg.partOrder) {
		uri := it.pkg.partOrder[it.index]
		it.index++
		part := it.pkg.parts[uri]
		if it.filter == nil || it.filter(part) {
			return true
		}
	}
	return false
}

// Part 返回当前部件
func (it *PartIterator) Part() *StreamPart {
	if it.index <= 0 || it.index > len(it.pkg.partOrder) {
		return nil
	}
	uri := it.pkg.partOrder[it.index-1]
	return it.pkg.parts[uri]
}

// Open 打开当前部件的内容流
func (it *PartIterator) Open() (io.ReadCloser, error) {
	part := it.Part()
	if part == nil {
		return nil, fmt.Errorf("no current part")
	}
	return part.Open()
}
