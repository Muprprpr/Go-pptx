package parts

import (
	"io"
)

// ============================================================================
// 媒体资源结构体 - 统一处理图片、音频、视频
// ============================================================================
//
// 设计原则：
// 1. 所有字段均为只读（小写字段通过构造函数初始化）
// 2. 支持小文件直接存储 []byte，大文件使用 io.Reader
// 3. 针对高并发读取优化，无需加锁即可安全读取
// ============================================================================

// MediaType 媒体类型枚举
type MediaType int8

const (
	MediaTypeUnknown MediaType = iota
	MediaTypeImage   // 图片
	MediaTypeAudio   // 音频
	MediaTypeVideo   // 视频
)

// MediaResource 媒体资源结构体（只读）
// 统一处理 PPTX 中的图片、音频、视频等媒体文件
type MediaResource struct {
	// 基础属性
	fileName     string      // 文件名（如 image1.png, audio1.mp3）
	contentType  string      // MIME 类型（如 image/png, audio/mpeg）
	mediaType    MediaType   // 媒体类型枚举
	target       string      // 在 ZIP 中的完整路径（如 ppt/media/image1.png）

	// 数据存储（二选一）
	data     []byte   // 小文件：直接存储字节数据
	dataSize int64    // 数据大小（字节）
	reader   io.Reader // 大文件：延迟加载的 Reader

	// 关联信息
	rId       string // 关系 ID（在 slide/slideLayout/slideMaster 中的引用）
	extension string // 文件扩展名（如 .png, .mp3）
	hash      string // 内容 Hash（MD5，用于去重）
}

// ============================================================================
// 构造函数
// ============================================================================

// NewMediaResourceFromBytes 从字节数据创建媒体资源
// 适用于小文件（如小图片）
func NewMediaResourceFromBytes(fileName, contentType, target string, data []byte) *MediaResource {
	return &MediaResource{
		fileName:    fileName,
		contentType: contentType,
		mediaType:   detectMediaType(contentType),
		target:      target,
		data:        data,
		dataSize:    int64(len(data)),
		extension:   extractExtension(fileName),
	}
}

// NewMediaResourceFromReader 从 Reader 创建媒体资源
// 适用于大文件（如视频、大图片）
func NewMediaResourceFromReader(fileName, contentType, target string, reader io.Reader, size int64) *MediaResource {
	return &MediaResource{
		fileName:    fileName,
		contentType: contentType,
		mediaType:   detectMediaType(contentType),
		target:      target,
		reader:      reader,
		dataSize:    size,
		extension:   extractExtension(fileName),
	}
}

// ============================================================================
// Getter 方法
// ============================================================================

// FileName 返回文件名（如 image1.png）
func (m *MediaResource) FileName() string { return m.fileName }

// ContentType 返回 MIME 类型（如 image/png）
func (m *MediaResource) ContentType() string { return m.contentType }

// MediaType 返回媒体类型枚举
func (m *MediaResource) MediaType() MediaType { return m.mediaType }

// Target 返回在 ZIP 中的完整路径（如 ppt/media/image1.png）
func (m *MediaResource) Target() string { return m.target }

// Data 返回字节数据（如果存在）
// 对于大文件（使用 Reader 创建），返回 nil
func (m *MediaResource) Data() []byte { return m.data }

// DataSize 返回数据大小（字节）
func (m *MediaResource) DataSize() int64 { return m.dataSize }

// Reader 返回数据 Reader
// 对于小文件（使用 Bytes 创建），返回 nil
func (m *MediaResource) Reader() io.Reader { return m.reader }

// RID 返回关系 ID
func (m *MediaResource) RID() string { return m.rId }

// Extension 返回文件扩展名（如 .png）
func (m *MediaResource) Extension() string { return m.extension }

// HasData 检查是否有字节数据
func (m *MediaResource) HasData() bool { return m.data != nil }

// HasReader 检查是否有 Reader
func (m *MediaResource) HasReader() bool { return m.reader != nil }

// IsImage 检查是否为图片类型
func (m *MediaResource) IsImage() bool { return m.mediaType == MediaTypeImage }

// IsAudio 检查是否为音频类型
func (m *MediaResource) IsAudio() bool { return m.mediaType == MediaTypeAudio }

// IsVideo 检查是否为视频类型
func (m *MediaResource) IsVideo() bool { return m.mediaType == MediaTypeVideo }

// ============================================================================
// Setter 方法（仅用于初始化阶段）
// ============================================================================

// SetRID 设置关系 ID
func (m *MediaResource) SetRID(rId string) {
	m.rId = rId
}

// SetHash 设置内容 Hash
func (m *MediaResource) SetHash(hash string) {
	m.hash = hash
}

// Hash 返回内容 Hash
func (m *MediaResource) Hash() string { return m.hash }

// ============================================================================
// 辅助函数
// ============================================================================

// detectMediaType 根据 MIME 类型检测媒体类型
func detectMediaType(contentType string) MediaType {
	switch {
	case isImageContentType(contentType):
		return MediaTypeImage
	case isAudioContentType(contentType):
		return MediaTypeAudio
	case isVideoContentType(contentType):
		return MediaTypeVideo
	default:
		return MediaTypeUnknown
	}
}

// isImageContentType 检查是否为图片 MIME 类型
func isImageContentType(ct string) bool {
	prefixes := []string{
		"image/png",
		"image/jpeg",
		"image/gif",
		"image/bmp",
		"image/tiff",
		"image/svg+xml",
		"image/webp",
		"image/x-emf",
		"image/x-wmf",
	}
	for _, p := range prefixes {
		if ct == p {
			return true
		}
	}
	return false
}

// isAudioContentType 检查是否为音频 MIME 类型
func isAudioContentType(ct string) bool {
	prefixes := []string{
		"audio/mpeg",
		"audio/wav",
		"audio/ogg",
		"audio/aac",
		"audio/mp4",
	}
	for _, p := range prefixes {
		if ct == p {
			return true
		}
	}
	return false
}

// isVideoContentType 检查是否为视频 MIME 类型
func isVideoContentType(ct string) bool {
	prefixes := []string{
		"video/mp4",
		"video/webm",
		"video/ogg",
		"video/quicktime",
		"video/x-msvideo",
		"video/x-ms-wmv",
	}
	for _, p := range prefixes {
		if ct == p {
			return true
		}
	}
	return false
}

// extractExtension 从文件名提取扩展名
func extractExtension(fileName string) string {
	for i := len(fileName) - 1; i >= 0; i-- {
		if fileName[i] == '.' {
			return fileName[i:]
		}
	}
	return ""
}

// ============================================================================
// MediaType String 方法
// ============================================================================

// String 返回媒体类型的字符串表示
func (mt MediaType) String() string {
	switch mt {
	case MediaTypeImage:
		return "image"
	case MediaTypeAudio:
		return "audio"
	case MediaTypeVideo:
		return "video"
	default:
		return "unknown"
	}
}
