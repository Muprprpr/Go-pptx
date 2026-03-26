package opc

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path"
	"strings"
	"sync"
)

// Package 表示一个 OPC 包（如 PPTX 文件）
type Package struct {
	parts          *PartCollection
	relationships  *Relationships
	contentTypes   *ContentTypes
	coreProperties *CoreProperties
	mu             sync.RWMutex
}

// NewPackage 创建一个新的空 OPC 包
func NewPackage() *Package {
	return &Package{
		parts:         NewPartCollection(),
		relationships: NewRelationships(RootURI()),
		contentTypes:  NewContentTypes(),
	}
}

// Parts 返回部件集合
func (p *Package) Parts() *PartCollection {
	return p.parts
}

// Relationships 返回包级别关系
func (p *Package) Relationships() *Relationships {
	return p.relationships
}

// ContentTypes 返回内容类型定义
func (p *Package) ContentTypes() *ContentTypes {
	return p.contentTypes
}

// CoreProperties 返回核心属性
func (p *Package) CoreProperties() *CoreProperties {
	p.mu.RLock()
	defer p.mu.RUnlock()
	return p.coreProperties
}

// SetCoreProperties 设置核心属性
func (p *Package) SetCoreProperties(props *CoreProperties) {
	p.mu.Lock()
	defer p.mu.Unlock()
	p.coreProperties = props
}

// ===== 包的打开 =====

// Open 从 ZIP 流打开一个 OPC 包
func Open(r io.ReaderAt, size int64) (*Package, error) {
	zipReader, err := zip.NewReader(r, size)
	if err != nil {
		return nil, fmt.Errorf("failed to open zip: %w", err)
	}

	pkg := NewPackage()

	if err := pkg.loadContentTypes(zipReader); err != nil {
		return nil, fmt.Errorf("failed to load content types: %w", err)
	}

	if err := pkg.loadParts(zipReader); err != nil {
		return nil, fmt.Errorf("failed to load parts: %w", err)
	}

	if err := pkg.loadRelationships(zipReader); err != nil {
		return nil, fmt.Errorf("failed to load relationships: %w", err)
	}

	for _, part := range pkg.parts.All() {
		part.SetDirty(false)
	}

	return pkg, nil
}

// OpenFile 从文件路径打开一个 OPC 包
func OpenFile(path string) (*Package, error) {
	file, err := os.Open(path)
	if err != nil {
		return nil, fmt.Errorf("failed to open file: %w", err)
	}
	defer file.Close()

	stat, err := file.Stat()
	if err != nil {
		return nil, fmt.Errorf("failed to get file info: %w", err)
	}

	return Open(file, stat.Size())
}

func (p *Package) loadContentTypes(zipReader *zip.Reader) error {
	var ctFile *zip.File
	for _, f := range zipReader.File {
		if f.Name == PathContentTypes {
			ctFile = f
			break
		}
	}

	if ctFile == nil {
		return fmt.Errorf("[Content_Types].xml not found")
	}

	rc, err := ctFile.Open()
	if err != nil {
		return fmt.Errorf("failed to open [Content_Types].xml: %w", err)
	}
	defer rc.Close()

	data, err := io.ReadAll(rc)
	if err != nil {
		return fmt.Errorf("failed to read [Content_Types].xml: %w", err)
	}

	return p.contentTypes.FromXML(data)
}

func (p *Package) loadParts(zipReader *zip.Reader) error {
	for _, f := range zipReader.File {
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

		rc, err := f.Open()
		if err != nil {
			return fmt.Errorf("failed to open %s: %w", f.Name, err)
		}

		blob, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			return fmt.Errorf("failed to read %s: %w", f.Name, err)
		}

		contentType := p.contentTypes.GetContentType(uri)
		part := NewPart(uri, contentType, blob)
		part.SetDirty(false)

		if err := p.parts.Add(part); err != nil {
			return fmt.Errorf("failed to add part %s: %w", uri.URI(), err)
		}
	}

	return nil
}

func (p *Package) loadRelationships(zipReader *zip.Reader) error {
	for _, f := range zipReader.File {
		if !strings.Contains(f.Name, PathRelsDir+"/") || !strings.HasSuffix(f.Name, ".rels") {
			continue
		}

		rc, err := f.Open()
		if err != nil {
			return fmt.Errorf("failed to open %s: %w", f.Name, err)
		}

		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			return fmt.Errorf("failed to read %s: %w", f.Name, err)
		}

		relURI := NewPackURI("/" + f.Name)
		sourceURI := relURI.SourceURI()

		rels := NewRelationships(sourceURI)
		if err := rels.FromXML(data); err != nil {
			return fmt.Errorf("failed to parse relationships %s: %w", f.Name, err)
		}

		if relURI.IsPackageRels() {
			p.relationships = rels
		} else {
			part := p.parts.Get(sourceURI)
			if part != nil {
				part.LoadRelationships(data)
			}
		}
	}

	return nil
}

// ===== 部件管理 =====

// AddPart 添加部件到包
func (p *Package) AddPart(part *Part) error {
	return p.parts.Add(part)
}

// CreatePart 创建并添加新部件
func (p *Package) CreatePart(uri *PackURI, contentType string, blob []byte) (*Part, error) {
	part := NewPart(uri, contentType, blob)
	if err := p.parts.Add(part); err != nil {
		return nil, err
	}
	return part, nil
}

// CreatePartFromReader 从 Reader 创建并添加部件
func (p *Package) CreatePartFromReader(uri *PackURI, contentType string, r io.Reader) (*Part, error) {
	part, err := NewPartFromReader(uri, contentType, r)
	if err != nil {
		return nil, err
	}
	if err := p.parts.Add(part); err != nil {
		return nil, err
	}
	return part, nil
}

// CreatePartFromXML 从 XML 结构创建并添加部件
func (p *Package) CreatePartFromXML(uri *PackURI, contentType string, v interface{}) (*Part, error) {
	data, err := xml.Marshal(v)
	if err != nil {
		return nil, fmt.Errorf("failed to marshal XML: %w", err)
	}
	data = append([]byte(XMLDeclaration), data...)
	return p.CreatePart(uri, contentType, data)
}

// GetPart 根据 URI 获取部件
func (p *Package) GetPart(uri *PackURI) *Part {
	return p.parts.Get(uri)
}

// GetPartByStr 根据字符串 URI 获取部件
func (p *Package) GetPartByStr(uri string) *Part {
	return p.parts.GetByStr(uri)
}

// GetPartsByType 根据内容类型获取所有部件
func (p *Package) GetPartsByType(contentType string) []*Part {
	return p.parts.GetByType(contentType)
}

// ContainsPart 检查部件是否存在
func (p *Package) ContainsPart(uri *PackURI) bool {
	return p.parts.Contains(uri)
}

// RemovePart 从包中移除部件
func (p *Package) RemovePart(uri *PackURI) error {
	return p.parts.Remove(uri)
}

// PartCount 返回部件数量
func (p *Package) PartCount() int {
	return p.parts.Count()
}

// AllParts 返回所有部件
func (p *Package) AllParts() []*Part {
	return p.parts.All()
}

// PartURIs 返回所有部件 URI
func (p *Package) PartURIs() []*PackURI {
	return p.parts.URIs()
}

// DirtyParts 返回所有被修改的部件
func (p *Package) DirtyParts() []*Part {
	return p.parts.DirtyParts()
}

// ===== 关系管理 =====

// AddRelationship 添加包级别关系
func (p *Package) AddRelationship(relType, targetURI string, isExternal bool) (*Relationship, error) {
	return p.relationships.AddNew(relType, targetURI, isExternal)
}

// GetRelationship 根据 rID 获取包级别关系
func (p *Package) GetRelationship(rID string) *Relationship {
	return p.relationships.Get(rID)
}

// GetRelationshipsByType 根据类型获取包级别关系
func (p *Package) GetRelationshipsByType(relType string) []*Relationship {
	return p.relationships.GetByType(relType)
}

// GetPartByRelType 通过关系类型获取目标部件
func (p *Package) GetPartByRelType(relType string) *Part {
	rels := p.relationships.GetByType(relType)
	if len(rels) == 0 {
		return nil
	}
	return p.parts.Get(rels[0].TargetURI())
}

// ResolveRelationship 解析部件间关系获取目标部件
func (p *Package) ResolveRelationship(source *Part, relType string) *Part {
	rels := source.Relationships().GetByType(relType)
	if len(rels) == 0 {
		return nil
	}
	return p.parts.Get(rels[0].TargetURI())
}

// ===== 包的保存 =====

// Save 将包保存为 ZIP 格式写入到 w
func (p *Package) Save(w io.Writer) error {
	zipWriter := zip.NewWriter(w)
	defer zipWriter.Close()

	if err := p.writeContentTypes(zipWriter); err != nil {
		return fmt.Errorf("failed to write content types: %w", err)
	}

	if err := p.writePackageRelationships(zipWriter); err != nil {
		return fmt.Errorf("failed to write package relationships: %w", err)
	}

	if err := p.writeParts(zipWriter); err != nil {
		return fmt.Errorf("failed to write parts: %w", err)
	}

	if p.coreProperties != nil {
		if err := p.writeCoreProperties(zipWriter); err != nil {
			return fmt.Errorf("failed to write core properties: %w", err)
		}
	}

	return nil
}

// SaveFile 将包保存到文件
func (p *Package) SaveFile(path string) error {
	file, err := os.Create(path)
	if err != nil {
		return fmt.Errorf("failed to create file: %w", err)
	}
	defer file.Close()

	return p.Save(file)
}

// SaveToBytes 将包保存到字节数组
func (p *Package) SaveToBytes() ([]byte, error) {
	buf := &bytesBuffer{}
	if err := p.Save(buf); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

func (p *Package) writeContentTypes(zipWriter *zip.Writer) error {
	p.updateContentTypes()

	data, err := p.contentTypes.ToXML()
	if err != nil {
		return err
	}

	return p.writeZipEntry(zipWriter, PathContentTypes, data)
}

func (p *Package) writePackageRelationships(zipWriter *zip.Writer) error {
	if p.relationships.Count() == 0 {
		return nil
	}

	data, err := p.relationships.ToXML()
	if err != nil {
		return err
	}

	relPath := PathRelsDir + "/" + PathRelsFile
	return p.writeZipEntry(zipWriter, relPath, data)
}

func (p *Package) writeParts(zipWriter *zip.Writer) error {
	for _, part := range p.parts.All() {
		filePath := strings.TrimPrefix(part.PartURI().URI(), "/")
		if err := p.writeZipEntry(zipWriter, filePath, part.Blob()); err != nil {
			return fmt.Errorf("failed to write part %s: %w", filePath, err)
		}

		if part.HasRelationships() {
			relPath := p.relFilePath(part.PartURI())
			relData, err := part.RelationshipsBlob()
			if err != nil {
				return fmt.Errorf("failed to serialize relationships for %s: %w", filePath, err)
			}
			if relData != nil {
				if err := p.writeZipEntry(zipWriter, relPath, relData); err != nil {
					return fmt.Errorf("failed to write relationships for %s: %w", filePath, err)
				}
			}
		}
	}

	return nil
}

func (p *Package) writeCoreProperties(zipWriter *zip.Writer) error {
	data, err := p.coreProperties.ToXML()
	if err != nil {
		return err
	}

	return p.writeZipEntry(zipWriter, "docProps/core.xml", data)
}

func (p *Package) writeZipEntry(zipWriter *zip.Writer, path string, data []byte) error {
	// 剥离前导斜杠，确保符合 ZIP 规范（Windows 和其他系统都要求）
	path = strings.TrimPrefix(path, "/")

	writer, err := zipWriter.Create(path)
	if err != nil {
		return fmt.Errorf("failed to create zip entry %s: %w", path, err)
	}

	_, err = writer.Write(data)
	if err != nil {
		return fmt.Errorf("failed to write zip entry %s: %w", path, err)
	}

	return nil
}

func (p *Package) relFilePath(uri *PackURI) string {
	dir := path.Dir(strings.TrimPrefix(uri.URI(), "/"))
	filename := path.Base(uri.URI())
	return path.Join(dir, PathRelsDir, filename+".rels")
}

func (p *Package) updateContentTypes() {
	for _, part := range p.parts.All() {
		uri := part.PartURI()
		contentType := part.ContentType()

		if uri.IsRelationshipsPart() {
			contentType = ContentTypeRelationships
		}

		ext := uri.Extension()
		defaultCT := p.contentTypes.GetDefault(ext)

		if contentType != "" && contentType != ContentTypeDefault {
			if defaultCT == "" || defaultCT == ContentTypeDefault || defaultCT != contentType {
				p.contentTypes.AddOverride(uri, contentType)
			}
		}
	}
}

// ===== 其他方法 =====

// Clone 克隆整个包
func (p *Package) Clone() *Package {
	p.mu.RLock()
	defer p.mu.RUnlock()

	newPkg := NewPackage()

	for _, part := range p.parts.All() {
		newPart := part.Clone()
		_ = newPkg.parts.Add(newPart)
	}

	newPkg.relationships = p.relationships.Clone()

	newPkg.contentTypes = &ContentTypes{
		defaults:  make(map[string]string),
		overrides: make(map[string]string),
	}
	for k, v := range p.contentTypes.Defaults() {
		newPkg.contentTypes.AddDefault(k, v)
	}
	for k, v := range p.contentTypes.Overrides() {
		newPkg.contentTypes.AddOverride(NewPackURI(k), v)
	}

	if p.coreProperties != nil {
		newPkg.coreProperties = &CoreProperties{}
		newPkg.coreProperties.SetTitle(p.coreProperties.Title())
		newPkg.coreProperties.SetCreator(p.coreProperties.Creator())
		newPkg.coreProperties.SetSubject(p.coreProperties.Subject())
		newPkg.coreProperties.SetDescription(p.coreProperties.Description())
		newPkg.coreProperties.SetKeywords(p.coreProperties.Keywords())
		newPkg.coreProperties.SetCreated(p.coreProperties.Created())
		newPkg.coreProperties.SetModified(p.coreProperties.Modified())
		newPkg.coreProperties.SetLastModifiedBy(p.coreProperties.LastModifiedBy())
		newPkg.coreProperties.SetRevision(p.coreProperties.Revision())
		newPkg.coreProperties.SetCategory(p.coreProperties.Category())
		newPkg.coreProperties.SetContentType(p.coreProperties.ContentType())
		newPkg.coreProperties.SetLanguage(p.coreProperties.Language())
	}

	return newPkg
}

// Close 关闭包，释放资源
func (p *Package) Close() error {
	p.mu.Lock()
	defer p.mu.Unlock()

	p.parts.Clear()
	p.relationships = nil

	return nil
}

// bytesBuffer 简单的 bytes buffer，实现 io.Writer
type bytesBuffer struct {
	data []byte
}

func (b *bytesBuffer) Write(p []byte) (n int, err error) {
	b.data = append(b.data, p...)
	return len(p), nil
}

func (b *bytesBuffer) Bytes() []byte {
	return b.data
}
