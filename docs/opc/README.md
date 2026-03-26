package opc // import "github.com/Muprprpr/Go-pptx/opc"

Package opc 提供 OOXML Open Packaging Convention (OPC) 的 Go 实现 用于处理 PPTX 等 Office
Open XML 文件格式

CONSTANTS

const (
	// OPC 关系类型
	ContentTypeRelationships = "application/vnd.openxmlformats-package.relationships+xml"

	// PPTX 核心内容类型
	ContentTypePresentation   = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
	ContentTypeSlide          = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
	ContentTypeSlideLayout    = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
	ContentTypeSlideMaster    = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
	ContentTypeNotesSlide     = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
	ContentTypeHandoutMaster  = "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml"
	ContentTypeNotesMaster    = "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"
	ContentTypePresentationML = "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"

	// 主题和样式
	ContentTypeTheme         = "application/vnd.openxmlformats-officedocument.theme+xml"
	ContentTypeThemeOverride = "application/vnd.openxmlformats-officedocument.themeOverride+xml"
	ContentTypeStyles        = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"

	// 核心属性
	ContentTypeCoreProperties = "application/vnd.openxmlformats-package.core-properties+xml"

	// 扩展属性
	ContentTypeExtendedProperties = "application/vnd.openxmlformats-officedocument.extended-properties+xml"

	// 自定义属性
	ContentTypeCustomProperties = "application/vnd.openxmlformats-officedocument.custom-properties+xml"

	// 图片内容类型
	ContentTypePNG  = "image/png"
	ContentTypeJPEG = "image/jpeg"
	ContentTypeGIF  = "image/gif"
	ContentTypeBMP  = "image/bmp"
	ContentTypeTIFF = "image/tiff"
	ContentTypeWMF  = "image/x-wmf"
	ContentTypeEMF  = "image/x-emf"
	ContentTypeSVG  = "image/svg+xml"

	// 音频内容类型
	ContentTypeWAV  = "audio/wav"
	ContentTypeMP3  = "audio/mpeg"
	ContentTypeMIDI = "audio/midi"

	// 视频内容类型
	ContentTypeMP4 = "video/mp4"
	ContentTypeAVI = "video/x-msvideo"
	ContentTypeWMV = "video/x-ms-wmv"

	// 其他
	ContentTypeXML  = "application/xml"
	ContentTypeFont = "application/x-font"

	// 默认内容类型映射（基于扩展名）
	ContentTypeDefault = "application/octet-stream"
)
    内容类型常量 (Content Types)

const (
	// OPC 核心关系
	RelTypeCoreProperties = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"

	// Office 文档关系
	RelTypeOfficeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"

	// 扩展属性
	RelTypeExtendedProperties = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
	RelTypeCustomProperties   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"

	// 幻灯片关系
	RelTypeSlide         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
	RelTypeSlideLayout   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
	RelTypeSlideMaster   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
	RelTypeNotesSlide    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
	RelTypeNotesMaster   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster"
	RelTypeHandoutMaster = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster"

	// 主题关系
	RelTypeTheme         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
	RelTypeThemeOverride = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride"

	// 媒体关系
	RelTypeImage = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
	RelTypeAudio = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
	RelTypeVideo = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"
	RelTypeMedia = "http://schemas.microsoft.com/office/2007/relationships/media"

	// 超链接
	RelTypeHyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

	// 字体
	RelTypeFont = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"

	// OLE 对象
	RelTypeOLEObject = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"

	// 缩略图
	RelTypeThumbnail = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"

	// 样式
	RelTypeStyles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
)
    关系类型常量 (Relationship Types)

const (
	NamespaceOPCPackage      = "http://schemas.openxmlformats.org/package/2006/content-types"
	NamespaceRelationships   = "http://schemas.openxmlformats.org/package/2006/relationships"
	NamespaceRelationshipsNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)
    OPC 命名空间

const (
	PathContentTypes = "[Content_Types].xml"
	PathRelsDir      = "_rels"
	PathRelsFile     = ".rels"
)
    OPC 默认路径


VARIABLES

var ContentTypeToExtension = map[string]string{
	ContentTypePNG:           ".png",
	ContentTypeJPEG:          ".jpg",
	ContentTypeGIF:           ".gif",
	ContentTypeBMP:           ".bmp",
	ContentTypeTIFF:          ".tiff",
	ContentTypeWMF:           ".wmf",
	ContentTypeEMF:           ".emf",
	ContentTypeSVG:           ".svg",
	ContentTypeWAV:           ".wav",
	ContentTypeMP3:           ".mp3",
	ContentTypeMIDI:          ".mid",
	ContentTypeMP4:           ".mp4",
	ContentTypeAVI:           ".avi",
	ContentTypeWMV:           ".wmv",
	ContentTypeRelationships: ".rels",
	ContentTypeXML:           ".xml",
}
    内容类型到扩展名的反向映射

var DefaultContentTypes = map[string]string{
	".xml":   ContentTypeXML,
	".rels":  ContentTypeRelationships,
	".png":   ContentTypePNG,
	".jpg":   ContentTypeJPEG,
	".jpeg":  ContentTypeJPEG,
	".gif":   ContentTypeGIF,
	".bmp":   ContentTypeBMP,
	".tiff":  ContentTypeTIFF,
	".tif":   ContentTypeTIFF,
	".wmf":   ContentTypeWMF,
	".emf":   ContentTypeEMF,
	".svg":   ContentTypeSVG,
	".wav":   ContentTypeWAV,
	".mp3":   ContentTypeMP3,
	".mid":   ContentTypeMIDI,
	".midi":  ContentTypeMIDI,
	".mp4":   ContentTypeMP4,
	".avi":   ContentTypeAVI,
	".wmv":   ContentTypeWMV,
	".font":  ContentTypeFont,
	".odttf": ContentTypeFont,
}
    默认内容类型映射（扩展名 -> 内容类型）


FUNCTIONS

func ComputeHash(data []byte) string
    ComputeHash 计算数据的 SHA256 哈希

func GetContentTypeByExtension(ext string) string
    GetContentTypeByExtension 根据文件扩展名获取内容类型

func GetExtensionByContentType(contentType string) string
    GetExtensionByContentType 根据内容类型获取文件扩展名

func IsValidPackURI(uri string) bool
    IsValidPackURI 检查 URI 是否有效

func NormalizeURI(uri string) string
    NormalizeURI 规范化 URI


TYPES

type BytesReader struct {
	// Has unexported fields.
}
    BytesReader 简单的bytes reader实现

func NewBytesReader(data []byte) *BytesReader
    NewBytesReader 创建新的BytesReader

func (r *BytesReader) Read(p []byte) (n int, err error)
    Read 实现io.Reader接口

type BytesSource struct {
	// Has unexported fields.
}
    BytesSource 内存中的字节数据源

func NewBytesSource(data []byte) *BytesSource
    NewBytesSource 从字节数组创建数据源

func (s *BytesSource) Open() (io.ReadCloser, error)
    Open 返回 bytes.Reader

func (s *BytesSource) Size() int64
    Size 返回数据大小

type ConcurrentZipCollector struct {
	// Has unexported fields.
}
    ConcurrentZipCollector 并发 ZIP 收集器 使用 goroutine 从 channel 收集部件数据并写入 ZIP

func NewConcurrentZipCollector(w io.Writer, bufferSize int) *ConcurrentZipCollector
    NewConcurrentZipCollector 创建并发 ZIP 收集器

func (c *ConcurrentZipCollector) Close() error
    Close 关闭收集器，等待所有数据写入完成

func (c *ConcurrentZipCollector) DataChannel() PartDataChannel
    DataChannel 返回数据通道（用于外部生产者）

func (c *ConcurrentZipCollector) Start()
    Start 启动收集器 goroutine

func (c *ConcurrentZipCollector) Submit(data *PartData) error
    Submit 提交部件数据到收集器

func (c *ConcurrentZipCollector) SubmitBytes(path string, data []byte) error
    SubmitBytes 提交字节数据

func (c *ConcurrentZipCollector) Wait() error
    Wait 等待收集器完成

type ContentTypes struct {
	// Has unexported fields.
}
    ContentTypes 表示 [Content_Types].xml 的内容 定义包中所有部件的内容类型

func NewContentTypes() *ContentTypes
    NewContentTypes 创建新的内容类型定义

func (ct *ContentTypes) AddDefault(extension, contentType string)
    AddDefault 添加默认内容类型映射

func (ct *ContentTypes) AddOverride(uri *PackURI, contentType string)
    AddOverride 添加特定 URI 的内容类型覆盖

func (ct *ContentTypes) Defaults() map[string]string
    Defaults 返回所有默认内容类型映射

func (ct *ContentTypes) FromXML(data []byte) error
    FromXML 从 XML 解析内容类型

func (ct *ContentTypes) GetContentType(uri *PackURI) string
    GetContentType 获取指定 URI 的内容类型 优先查找 overrides，然后查找 defaults

func (ct *ContentTypes) GetDefault(extension string) string
    GetDefault 获取扩展名对应的默认内容类型

func (ct *ContentTypes) GetOverride(uri *PackURI) string
    GetOverride 获取 URI 对应的内容类型覆盖

func (ct *ContentTypes) Overrides() map[string]string
    Overrides 返回所有内容类型覆盖

func (ct *ContentTypes) RemoveOverride(uri *PackURI)
    RemoveOverride 移除内容类型覆盖

func (ct *ContentTypes) ToXML() ([]byte, error)
    ToXML 将内容类型序列化为 XML

type ContentTypesStreamer struct {
	// Has unexported fields.
}
    ContentTypesStreamer ContentTypes 流式写入器

func NewContentTypesStreamer(ct *ContentTypes) *ContentTypesStreamer
    NewContentTypesStreamer 创建 ContentTypes 流式写入器

func (cs *ContentTypesStreamer) StreamWriteTo(w io.Writer) error
    StreamWriteTo 实现 StreamWriter 接口

type CoreProperties struct {
	// Has unexported fields.
}
    CoreProperties 表示包的核心属性（Dublin Core 元数据）

func (cp *CoreProperties) Category() string
    Category 返回类别

func (cp *CoreProperties) ContentType() string
    ContentType 返回内容类型

func (cp *CoreProperties) Created() string
    Created 返回创建时间

func (cp *CoreProperties) Creator() string
    Creator 返回创建者

func (cp *CoreProperties) Description() string
    Description 返回描述

func (cp *CoreProperties) FromXML(data []byte) error
    FromXML 从 XML 解析核心属性

func (cp *CoreProperties) Keywords() string
    Keywords 返回关键词

func (cp *CoreProperties) Language() string
    Language 返回语言

func (cp *CoreProperties) LastModifiedBy() string
    LastModifiedBy 返回最后修改者

func (cp *CoreProperties) Modified() string
    Modified 返回修改时间

func (cp *CoreProperties) Revision() string
    Revision 返回版本号

func (cp *CoreProperties) SetCategory(category string)
    SetCategory 设置类别

func (cp *CoreProperties) SetContentType(contentType string)
    SetContentType 设置内容类型

func (cp *CoreProperties) SetCreated(created string)
    SetCreated 设置创建时间

func (cp *CoreProperties) SetCreator(creator string)
    SetCreator 设置创建者

func (cp *CoreProperties) SetDescription(description string)
    SetDescription 设置描述

func (cp *CoreProperties) SetKeywords(keywords string)
    SetKeywords 设置关键词

func (cp *CoreProperties) SetLanguage(language string)
    SetLanguage 设置语言

func (cp *CoreProperties) SetLastModifiedBy(lastModifiedBy string)
    SetLastModifiedBy 设置最后修改者

func (cp *CoreProperties) SetModified(modified string)
    SetModified 设置修改时间

func (cp *CoreProperties) SetRevision(revision string)
    SetRevision 设置版本号

func (cp *CoreProperties) SetSubject(subject string)
    SetSubject 设置主题

func (cp *CoreProperties) SetTitle(title string)
    SetTitle 设置标题

func (cp *CoreProperties) Subject() string
    Subject 返回主题

func (cp *CoreProperties) Title() string
    Title 返回标题

func (cp *CoreProperties) ToXML() ([]byte, error)
    ToXML 将核心属性序列化为 XML

type DefaultPartFactory struct{}
    DefaultPartFactory 默认部件工厂

func (f *DefaultPartFactory) CreatePart(uri *PackURI, contentType string, blob []byte) (*Part, error)
    CreatePart 实现PartFactory接口

type PackURI struct {
	// Has unexported fields.
}
    PackURI 表示包内的URI（如 /ppt/slides/slide1.xml） 遵循 OPC 规范中的 URI 规则

func ContentTypesURI() *PackURI
    ContentTypesURI 返回 [Content_Types].xml 的 URI

func NewPackURI(uri string) *PackURI
    NewPackURI 创建一个新的 PackURI

func PackageRelsURI() *PackURI
    PackageRelsURI 返回包级别关系文件的 URI

func RootURI() *PackURI
    RootURI 返回包根目录 URI

func (p *PackURI) BaseName() string
    BaseName 返回 URI 的基本名称（不含扩展名） 例如: /ppt/slides/slide1.xml -> slide1

func (p *PackURI) Clone() *PackURI
    Clone 创建 PackURI 的副本

func (p *PackURI) DirName() string
    DirName 返回父目录路径 例如: /ppt/slides/slide1.xml -> /ppt/slides

func (p *PackURI) DirSegments() []string
    DirSegments 返回目录段的切片（不含文件名）

func (p *PackURI) Equals(other *PackURI) bool
    Equals 比较两个 PackURI 是否相等

func (p *PackURI) EqualsStr(uri string) bool
    EqualsStr 比较与字符串 URI 是否相等

func (p *PackURI) Extension() string
    Extension 返回文件扩展名 例如: /ppt/slides/slide1.xml -> .xml

func (p *PackURI) FileName() string
    FileName 返回文件名（含扩展名） 例如: /ppt/slides/slide1.xml -> slide1.xml

func (p *PackURI) IsAbsolute() bool
    IsAbsolute 检查是否为绝对路径

func (p *PackURI) IsPackageRels() bool
    IsPackageRels 检查是否为包级别关系文件

func (p *PackURI) IsRelationshipsPart() bool
    IsRelationshipsPart 检查是否为关系部件

func (p *PackURI) Join(relativePath string) *PackURI
    Join 连接相对路径并返回新的 PackURI

func (p PackURI) MarshalText() ([]byte, error)
    MarshalText 实现 encoding.TextMarshaler 接口

func (p *PackURI) MemberName() string
    MemberName 返回相对于包根目录的成员名称（用于 ZIP） 去掉开头的 /

func (p *PackURI) RelPath() string
    RelPath 返回相对路径（用于关系目标）

func (p *PackURI) RelPathFrom(other *PackURI) string
    RelPathFrom 计算从另一个 URI 到此 URI 的相对路径

func (p *PackURI) RelationshipsURI() *PackURI
    RelationshipsURI 返回此部件对应的关系文件 URI 例如: /ppt/slides/slide1.xml ->
    /ppt/slides/_rels/slide1.xml.rels

func (p *PackURI) Segments() []string
    Segments 返回 URI 的各个段

func (p *PackURI) SourceURI() *PackURI
    SourceURI 从关系文件路径获取源部件 URI 例如: /ppt/slides/_rels/slide1.xml.rels ->
    /ppt/slides/slide1.xml

func (p *PackURI) String() string
    String 返回 URI 字符串表示

func (p *PackURI) URI() string
    URI 返回原始 URI 字符串

func (p *PackURI) UnmarshalText(data []byte) error
    UnmarshalText 实现 encoding.TextUnmarshaler 接口

type Package struct {
	// Has unexported fields.
}
    Package 表示一个 OPC 包（如 PPTX 文件）

func NewPackage() *Package
    NewPackage 创建一个新的空 OPC 包

func Open(r io.ReaderAt, size int64) (*Package, error)
    Open 从 ZIP 流打开一个 OPC 包

func OpenFile(path string) (*Package, error)
    OpenFile 从文件路径打开一个 OPC 包

func (p *Package) AddPart(part *Part) error
    AddPart 添加部件到包

func (p *Package) AddRelationship(relType, targetURI string, isExternal bool) (*Relationship, error)
    AddRelationship 添加包级别关系

func (p *Package) AllParts() []*Part
    AllParts 返回所有部件

func (p *Package) Clone() *Package
    Clone 克隆整个包

func (p *Package) Close() error
    Close 关闭包，释放资源

func (p *Package) ContainsPart(uri *PackURI) bool
    ContainsPart 检查部件是否存在

func (p *Package) ContentTypes() *ContentTypes
    ContentTypes 返回内容类型定义

func (p *Package) CoreProperties() *CoreProperties
    CoreProperties 返回核心属性

func (p *Package) CreatePart(uri *PackURI, contentType string, blob []byte) (*Part, error)
    CreatePart 创建并添加新部件

func (p *Package) CreatePartFromReader(uri *PackURI, contentType string, r io.Reader) (*Part, error)
    CreatePartFromReader 从 Reader 创建并添加部件

func (p *Package) CreatePartFromXML(uri *PackURI, contentType string, v interface{}) (*Part, error)
    CreatePartFromXML 从 XML 结构创建并添加部件

func (p *Package) DirtyParts() []*Part
    DirtyParts 返回所有被修改的部件

func (p *Package) GetPart(uri *PackURI) *Part
    GetPart 根据 URI 获取部件

func (p *Package) GetPartByRelType(relType string) *Part
    GetPartByRelType 通过关系类型获取目标部件

func (p *Package) GetPartByStr(uri string) *Part
    GetPartByStr 根据字符串 URI 获取部件

func (p *Package) GetPartsByType(contentType string) []*Part
    GetPartsByType 根据内容类型获取所有部件

func (p *Package) GetRelationship(rID string) *Relationship
    GetRelationship 根据 rID 获取包级别关系

func (p *Package) GetRelationshipsByType(relType string) []*Relationship
    GetRelationshipsByType 根据类型获取包级别关系

func (p *Package) PartCount() int
    PartCount 返回部件数量

func (p *Package) PartURIs() []*PackURI
    PartURIs 返回所有部件 URI

func (p *Package) Parts() *PartCollection
    Parts 返回部件集合

func (p *Package) Relationships() *Relationships
    Relationships 返回包级别关系

func (p *Package) RemovePart(uri *PackURI) error
    RemovePart 从包中移除部件

func (p *Package) ResolveRelationship(source *Part, relType string) *Part
    ResolveRelationship 解析部件间关系获取目标部件

func (p *Package) Save(w io.Writer) error
    Save 将包保存为 ZIP 格式写入到 w

func (p *Package) SaveFile(path string) error
    SaveFile 将包保存到文件

func (p *Package) SaveToBytes() ([]byte, error)
    SaveToBytes 将包保存到字节数组

func (p *Package) SetCoreProperties(props *CoreProperties)
    SetCoreProperties 设置核心属性

type Part struct {
	// Has unexported fields.
}
    Part 表示包中的一个部件

func NewPart(uri *PackURI, contentType string, blob []byte) *Part
    NewPart 创建一个新的部件

func NewPartFromReader(uri *PackURI, contentType string, r io.Reader) (*Part, error)
    NewPartFromReader 从Reader创建部件

func (p *Part) AddRelationship(relType, targetURI string, isExternal bool) (*Relationship, error)
    AddRelationship 添加一个关系

func (p *Part) Blob() []byte
    Blob 返回原始内容

func (p *Part) Clone() *Part
    Clone 克隆部件

func (p *Part) ContentType() string
    ContentType 返回内容类型

func (p *Part) GetRelatedPart(rID string) *PackURI
    GetRelatedPart 通过关系获取目标部件（需要Package上下文）

func (p *Part) HasRelationships() bool
    HasRelationships 检查是否有关系

func (p *Part) IsDirty() bool
    IsDirty 返回是否被修改

func (p *Part) LoadRelationships(data []byte) error
    LoadRelationships 从XML加载关系

func (p *Part) MarshalToBlob(v interface{}) error
    MarshalToBlob 将 v 序列化为 XML 并存储到 blob

func (p *Part) PartURI() *PackURI
    PartURI 返回部件URI

func (p *Part) Reader() io.Reader
    Reader 返回内容的Reader

func (p *Part) Relationships() *Relationships
    Relationships 返回关系集合

func (p *Part) RelationshipsBlob() ([]byte, error)
    RelationshipsBlob 返回关系的XML内容

func (p *Part) RelationshipsURI() *PackURI
    RelationshipsURI 返回关系文件的URI

func (p *Part) RemoveRelationship(rID string) error
    RemoveRelationship 删除一个关系

func (p *Part) SetBlob(blob []byte)
    SetBlob 设置内容

func (p *Part) SetBlobFromReader(r io.Reader) error
    SetBlobFromReader 从Reader设置内容

func (p *Part) SetContentType(ct string)
    SetContentType 设置内容类型

func (p *Part) SetDirty(dirty bool)
    SetDirty 设置修改标记

func (p *Part) Size() int
    Size 返回内容大小

func (p *Part) UnmarshalBlob(v interface{}) error
    UnmarshalBlob 从 blob 解析 XML 内容到 v

type PartCollection struct {
	// Has unexported fields.
}
    PartCollection 部件集合

func NewPartCollection() *PartCollection
    NewPartCollection 创建新的部件集合

func (c *PartCollection) Add(part *Part) error
    Add 添加部件

func (c *PartCollection) All() []*Part
    All 返回所有部件（按插入顺序）

func (c *PartCollection) Clear()
    Clear 清空集合

func (c *PartCollection) Contains(uri *PackURI) bool
    Contains 检查是否包含指定部件

func (c *PartCollection) Count() int
    Count 返回部件数量

func (c *PartCollection) DirtyParts() []*Part
    DirtyParts 返回所有被修改的部件

func (c *PartCollection) Get(uri *PackURI) *Part
    Get 根据URI获取部件

func (c *PartCollection) GetByStr(uri string) *Part
    GetByStr 根据字符串URI获取部件

func (c *PartCollection) GetByType(contentType string) []*Part
    GetByType 根据内容类型获取部件

func (c *PartCollection) Remove(uri *PackURI) error
    Remove 删除部件

func (c *PartCollection) URIs() []*PackURI
    URIs 返回所有部件URI

type PartData struct {
	URI         string     // 部件 URI
	Path        string     // ZIP 内路径
	ContentType string     // 内容类型
	Data        []byte     // 数据内容
	Source      PartSource // 数据源（用于懒加载）
	Error       error      // 写入错误（如果有）
}
    PartData 部件数据 - 用于 channel 传递

type PartDataChannel chan *PartData
    PartDataChannel 部件数据通道类型

func NewPartDataChannel(bufferSize int) PartDataChannel
    NewPartDataChannel 创建部件数据通道

type PartFactory interface {
	// CreatePart 创建部件
	CreatePart(uri *PackURI, contentType string, blob []byte) (*Part, error)
}
    PartFactory 部件工厂接口

type PartIterator struct {
	// Has unexported fields.
}
    PartIterator 部件迭代器

func (it *PartIterator) FilterByType(contentType string) *PartIterator
    FilterByType 按内容类型过滤

func (it *PartIterator) Next() bool
    Next 移动到下一个部件

func (it *PartIterator) Open() (io.ReadCloser, error)
    Open 打开当前部件的内容流

func (it *PartIterator) Part() *StreamPart
    Part 返回当前部件

type PartSource interface {
	Open() (io.ReadCloser, error)
	Size() int64
}
    PartSource 部件数据源接口

type ReaderSource struct {
	// Has unexported fields.
}
    ReaderSource io.Reader 数据源

func NewReaderSource(r io.Reader, size int64) *ReaderSource
    NewReaderSource 从 io.Reader 创建数据源

func (s *ReaderSource) Open() (io.ReadCloser, error)
    Open 返回 reader

func (s *ReaderSource) Size() int64
    Size 返回数据大小

type RelTypeCollection struct {
	// Has unexported fields.
}
    RelTypeCollection 关系类型集合（用于按类型分组查找）

func NewRelTypeCollection() *RelTypeCollection
    NewRelTypeCollection 创建新的关系类型集合

func (c *RelTypeCollection) Add(rel *Relationship)
    Add 添加关系到类型集合

func (c *RelTypeCollection) GetByType(relType string) []*Relationship
    GetByType 按类型获取关系

func (c *RelTypeCollection) Types() []string
    Types 返回所有关系类型

type Relatable interface {
	// PartURI 返回部件的URI
	PartURI() *PackURI
	// Relationships 返回部件的关系集合
	Relationships() *Relationships
	// AddRelationship 添加一个关系
	AddRelationship(relType, targetURI string, isExternal bool) (*Relationship, error)
}
    Relatable 可关联部件的接口

type Relationship struct {
	// Has unexported fields.
}
    Relationship 表示两个部件之间的关系

func NewRelationship(rID, relType, targetURI string, isExternal bool, source *PackURI) *Relationship
    NewRelationship 创建一个新的关系

func (r *Relationship) Equals(other *Relationship) bool
    Equals 比较两个关系是否相等

func (r *Relationship) IsExternal() bool
    IsExternal 返回是否为外部关系

func (r *Relationship) RID() string
    RID 返回关系ID

func (r *Relationship) SetSource(source *PackURI)
    SetSource 设置源URI

func (r *Relationship) SourceURI() *PackURI
    SourceURI 返回源URI

func (r *Relationship) TargetMode() string
    TargetMode 返回目标模式

func (r *Relationship) TargetRef() string
    TargetRef 返回目标引用（相对或绝对） 如果有源部件，返回从源部件到目标的相对路径

func (r *Relationship) TargetURI() *PackURI
    TargetURI 返回目标URI

func (r *Relationship) Type() string
    Type 返回关系类型

type Relationships struct {
	// Has unexported fields.
}
    Relationships 表示关系的集合

func NewRelationships(sourceURI *PackURI) *Relationships
    NewRelationships 创建新的关系集合

func ParseRelationshipsFromXML(data []byte, sourceURI *PackURI) (*Relationships, error)
    ParseRelationshipsFromXML 从XML数据解析关系

func (rs *Relationships) Add(rel *Relationship) error
    Add 添加一个关系

func (rs *Relationships) AddNew(relType, targetURI string, isExternal bool) (*Relationship, error)
    AddNew 创建并添加一个新关系（使用原子操作分配 ID）

func (rs *Relationships) All() []*Relationship
    All 返回所有关系（按插入顺序）

func (rs *Relationships) Clone() *Relationships
    Clone 克隆关系集合

func (rs *Relationships) Contains(rID string) bool
    Contains 检查是否包含指定rID的关系

func (rs *Relationships) Count() int
    Count 返回关系数量

func (rs *Relationships) FromXML(data []byte) error
    FromXML 从XML解析关系集合

func (rs *Relationships) Get(rID string) *Relationship
    Get 根据rID获取关系

func (rs *Relationships) GetByTarget(targetURI *PackURI) *Relationship
    GetByTarget 根据目标URI获取关系

func (rs *Relationships) GetByType(relType string) []*Relationship
    GetByType 根据关系类型获取所有关系

func (rs *Relationships) InitRIDCounter()
    InitRIDCounter 初始化 rID 计数器（从现有关系中找到最大值） 在从 XML 加载关系后调用，确保新分配的 ID 不会与现有的冲突

func (rs *Relationships) MarshalXML(e *xml.Encoder, start xml.StartElement) error
    MarshalXML 实现 xml.Marshaler 接口

func (rs *Relationships) NextRID() string
    NextRID 返回下一个关系ID（预览，不消耗） 此方法是幂等的，多次调用返回相同的值，直到 AddNew 实际使用

func (rs *Relationships) Remove(rID string) error
    Remove 删除一个关系

func (rs *Relationships) SetSourceURI(sourceURI *PackURI)
    SetSourceURI 设置源URI

func (rs *Relationships) SourceURI() *PackURI
    SourceURI 返回源URI

func (rs *Relationships) ToXML() ([]byte, error)
    ToXML 将关系集合序列化为XML

func (rs *Relationships) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error
    UnmarshalXML 实现 xml.Unmarshaler 接口

type RelationshipsStreamer struct {
	// Has unexported fields.
}
    RelationshipsStreamer 关系流式写入器

func NewRelationshipsStreamer(rels *Relationships) *RelationshipsStreamer
    NewRelationshipsStreamer 创建关系流式写入器

func (rs *RelationshipsStreamer) StreamWriteTo(w io.Writer) error
    StreamWriteTo 实现 StreamWriter 接口

type ResourceDedupPool struct {
	// Has unexported fields.
}
    ResourceDedupPool 全局资源去重池 使用 sync.Map 实现并发安全的资源去重

func GetGlobalResourcePool() *ResourceDedupPool
    GetGlobalResourcePool 获取全局资源池

func NewResourceDedupPool() *ResourceDedupPool
    NewResourceDedupPool 创建新的资源去重池

func (p *ResourceDedupPool) Clear()
    Clear 清空资源池

func (p *ResourceDedupPool) Lookup(hash string) (*ResourceEntry, bool)
    Lookup 查找资源

func (p *ResourceDedupPool) Register(uri string, data []byte) (isNew bool, existingURI string)
    Register 注册资源，返回是否为新资源 如果资源已存在，增加引用计数并返回 false

func (p *ResourceDedupPool) RegisterWithHash(uri string, hash string, size int64) (isNew bool, existingURI string)
    RegisterWithHash 使用预计算的哈希注册资源

func (p *ResourceDedupPool) Release(hash string)
    Release 释放资源引用

func (p *ResourceDedupPool) Stats() (count int, totalSize int64)
    Stats 返回资源池统计信息

type ResourceEntry struct {
	URI       string // 部件 URI
	Hash      string // 内容哈希（SHA256）
	Size      int64  // 原始大小
	Reference int    // 引用计数
}
    ResourceEntry 资源条目

type ResourceHashKey string
    ResourceHashKey 资源哈希键

type StreamPackage struct {
	// Has unexported fields.
}
    StreamPackage 流式 OPC 包 - 支持懒加载和流式写入

func NewStreamPackage() *StreamPackage
    NewStreamPackage 创建新的流式包

func OpenStream(path string) (*StreamPackage, error)
    OpenStream 流式打开 OPC 包 文件句柄会保持打开状态以支持懒加载

func OpenStreamFromReader(r io.ReaderAt, size int64) (*StreamPackage, error)
    OpenStreamFromReader 从 io.ReaderAt 流式打开 注意：调用者负责保持 ReaderAt 的有效性

func (p *StreamPackage) AddMediaPartWithDedup(uri *PackURI, contentType string, data []byte) (actualURI *PackURI, isNew bool, err error)
    AddMediaPartWithDedup 添加媒体部件并进行去重 如果资源已存在，返回已存在部件的 URI 而不创建新部件

func (p *StreamPackage) AddPart(part *StreamPart) error
    AddPart 添加流式部件

func (p *StreamPackage) AddRelationship(relType, targetURI string, isExternal bool) (*Relationship, error)
    AddRelationship 添加包级别关系

func (p *StreamPackage) AllParts() []*StreamPart
    AllParts 返回所有部件

func (p *StreamPackage) ClearMediaDedupPool()
    ClearMediaDedupPool 清空媒体资源去重池

func (p *StreamPackage) Close() error
    Close 关闭包，释放资源

func (p *StreamPackage) ConcurrentStreamSave(w io.Writer, workerCount, bufferSize int) error
    ConcurrentStreamSave 使用 goroutine 收集器并发保存 workerCount: 并发工作 goroutine 数量
    bufferSize: channel 缓冲区大小

func (p *StreamPackage) ConcurrentStreamSaveFile(path string, workerCount, bufferSize int) error
    ConcurrentStreamSaveFile 使用并发方式保存到文件

func (p *StreamPackage) ContainsPart(uri *PackURI) bool
    ContainsPart 检查部件是否存在

func (p *StreamPackage) ContentTypes() *ContentTypes
    ContentTypes 返回内容类型定义

func (p *StreamPackage) CoreProperties() *CoreProperties
    CoreProperties 返回核心属性

func (p *StreamPackage) CreatePartFromBytes(uri *PackURI, contentType string, data []byte) (*StreamPart, error)
    CreatePartFromBytes 从字节创建部件（立即加载到内存）

func (p *StreamPackage) CreateStreamPart(uri *PackURI, contentType string, source PartSource) (*StreamPart, error)
    CreateStreamPart 创建并添加流式部件

func (p *StreamPackage) GetMediaDedupStats() (count int, totalSize int64)
    GetMediaDedupStats 获取媒体资源去重统计

func (p *StreamPackage) GetPart(uri *PackURI) *StreamPart
    GetPart 获取部件（内容按需加载）

func (p *StreamPackage) GetPartByRelType(relType string) *StreamPart
    GetPartByRelType 通过关系类型获取目标部件

func (p *StreamPackage) GetPartByStr(uri string) *StreamPart
    GetPartByStr 根据字符串 URI 获取部件

func (p *StreamPackage) GetPartsByType(contentType string) []*StreamPart
    GetPartsByType 根据内容类型获取部件

func (p *StreamPackage) NewPartIterator() *PartIterator
    NewPartIterator 创建部件迭代器

func (p *StreamPackage) PartCount() int
    PartCount 返回部件数量

func (p *StreamPackage) PartURIs() []*PackURI
    PartURIs 返回所有部件 URI

func (p *StreamPackage) RegisterMediaWithDedup(uri string, data []byte) (isNew bool, existingURI string)
    RegisterMediaWithDedup 注册媒体资源并进行去重 返回：isNew 表示是否为新资源，existingURI 表示已存在资源的
    URI

func (p *StreamPackage) Relationships() *Relationships
    Relationships 返回包级别关系

func (p *StreamPackage) RemovePart(uri *PackURI) error
    RemovePart 移除部件

func (p *StreamPackage) SetCoreProperties(props *CoreProperties)
    SetCoreProperties 设置核心属性

func (p *StreamPackage) StreamSave(w io.Writer) error
    StreamSave 流式保存到 io.Writer

func (p *StreamPackage) StreamSaveFile(path string) error
    StreamSaveFile 流式保存到文件

type StreamPart struct {
	// Has unexported fields.
}
    StreamPart 流式部件 - 支持懒加载

func NewStreamPart(uri *PackURI, contentType string, source PartSource) *StreamPart
    NewStreamPart 创建流式部件

func (p *StreamPart) Blob() ([]byte, error)
    Blob 返回内容（如果未加载则先加载）

func (p *StreamPart) Clone() *StreamPart
    Clone 克隆部件

func (p *StreamPart) ContentType() string
    ContentType 返回内容类型

func (p *StreamPart) HasRelationships() bool
    HasRelationships 检查是否有关系

func (p *StreamPart) IsDirty() bool
    IsDirty 返回是否被修改

func (p *StreamPart) IsLoaded() bool
    IsLoaded 返回是否已加载到内存

func (p *StreamPart) Load() error
    Load 将内容加载到内存

func (p *StreamPart) LoadRelationships(data []byte) error
    LoadRelationships 从 XML 加载关系

func (p *StreamPart) Open() (io.ReadCloser, error)
    Open 打开部件内容流

func (p *StreamPart) PartURI() *PackURI
    PartURI 返回部件 URI

func (p *StreamPart) Relationships() *Relationships
    Relationships 返回关系集合

func (p *StreamPart) RelationshipsBlob() ([]byte, error)
    RelationshipsBlob 返回关系的 XML 内容

func (p *StreamPart) RelationshipsURI() *PackURI
    RelationshipsURI 返回关系文件的 URI

func (p *StreamPart) SetBlob(data []byte)
    SetBlob 设置内容

func (p *StreamPart) SetBlobFromReader(r io.Reader) error
    SetBlobFromReader 从 Reader 设置内容

func (p *StreamPart) SetContentType(ct string)
    SetContentType 设置内容类型

func (p *StreamPart) SetDirty(dirty bool)
    SetDirty 设置修改标记

func (p *StreamPart) Size() int64
    Size 返回内容大小

func (p *StreamPart) UnmarshalBlob(v any) error
    UnmarshalBlob 从 blob 解析 XML 内容

type StreamWriter interface {
	StreamWriteTo(w io.Writer) error
}
    StreamWriter 流式写入器接口

type StreamingZipWriter struct {
	// Has unexported fields.
}
    StreamingZipWriter 流式 ZIP 写入器

func NewStreamingZipWriter(w io.Writer) *StreamingZipWriter
    NewStreamingZipWriter 创建流式 ZIP 写入器

func (sw *StreamingZipWriter) Close() error
    Close 关闭 ZIP 写入器

func (sw *StreamingZipWriter) Create(path string) (io.Writer, error)
    Create 创建 ZIP 条目并返回写入器

func (sw *StreamingZipWriter) WriteBytes(path string, data []byte) error
    WriteBytes 写入字节数据

func (sw *StreamingZipWriter) WriteFromReader(path string, reader io.Reader) error
    WriteFromReader 从 Reader 流式写入 ZIP 条目

func (sw *StreamingZipWriter) WriteFromStreamer(path string, streamer StreamWriter) error
    WriteFromStreamer 从 StreamWriter 流式写入 ZIP 条目

func (sw *StreamingZipWriter) WriteFromXMLStreamer(path string, streamer XMLStreamer) error
    WriteFromXMLStreamer 从 XMLStreamer 流式写入 ZIP 条目

func (sw *StreamingZipWriter) WriteStreamPart(part *StreamPart) error
    WriteStreamPart 流式写入 StreamPart

func (sw *StreamingZipWriter) WriteXML(path string, data []byte) error
    WriteXML 写入 XML 数据（自动添加 XML 头）

type XContentTypes struct {
	XMLName   xml.Name    `xml:"Types"`
	Xmlns     string      `xml:"xmlns,attr"`
	Defaults  []XDefault  `xml:"Default"`
	Overrides []XOverride `xml:"Override"`
}
    XContentTypes XML 序列化的内容类型根元素

type XCoreProperties struct {
	XMLName        xml.Name   `xml:"coreProperties"`
	XmlnsDc        string     `xml:"xmlns:dc,attr"`
	XmlnsDcterms   string     `xml:"xmlns:dcterms,attr"`
	XmlnsDcmitype  string     `xml:"xmlns:dcmitype,attr"`
	XmlnsXsi       string     `xml:"xmlns:xsi,attr"`
	XmlnsCore      string     `xml:"xmlns,attr"`
	Title          string     `xml:"dc:title"`
	Creator        string     `xml:"dc:creator"`
	Subject        string     `xml:"dc:subject"`
	Description    string     `xml:"dc:description"`
	Keywords       *XKeywords `xml:"cp:keywords"`
	Created        *XDate     `xml:"dcterms:created"`
	Modified       *XDate     `xml:"dcterms:modified"`
	LastModifiedBy string     `xml:"cp:lastModifiedBy"`
	Revision       string     `xml:"cp:revision"`
	Category       string     `xml:"cp:category"`
	ContentType    string     `xml:"cp:contentType"`
}
    XCoreProperties XML 序列化的核心属性

type XDate struct {
	Type  string `xml:"xsi:type,attr"`
	Value string `xml:",chardata"`
}
    XDate 日期元素

type XDefault struct {
	Extension   string `xml:"Extension,attr"`
	ContentType string `xml:"ContentType,attr"`
}
    XDefault XML 序列化的默认内容类型

type XKeywords struct {
	Value string `xml:",chardata"`
}
    XKeywords 关键词元素

type XMLStreamer interface {
	StreamXML(enc *xml.Encoder) error
}
    XMLStreamer XML 流式写入器接口

type XOverride struct {
	PartName    string `xml:"PartName,attr"`
	ContentType string `xml:"ContentType,attr"`
}
    XOverride XML 序列化的内容类型覆盖

type XRelationship struct {
	ID         string `xml:"Id,attr"`
	Type       string `xml:"Type,attr"`
	Target     string `xml:"Target,attr"`
	TargetMode string `xml:"TargetMode,attr,omitempty"`
}
    XRelationship XML序列化的关系元素

type XRelationships struct {
	XMLName       xml.Name        `xml:"Relationships"`
	Xmlns         string          `xml:"xmlns,attr"`
	Relationships []XRelationship `xml:"Relationship"`
}
    XRelationships XML序列化的根元素

type ZipFileSource struct {
	// Has unexported fields.
}
    ZipFileSource ZIP 文件中的部件数据源

func NewZipFileSource(f *zip.File) *ZipFileSource
    NewZipFileSource 从 zip.File 创建数据源

func (s *ZipFileSource) Open() (io.ReadCloser, error)
    Open 打开 ZIP 文件条目

func (s *ZipFileSource) Size() int64
    Size 返回未压缩大小

