// Package opc 提供 OOXML Open Packaging Convention (OPC) 的 Go 实现
// 用于处理 PPTX 等 Office Open XML 文件格式
package opc

// 内容类型常量 (Content Types)
const (
	// OPC 关系类型
	ContentTypeRelationships = "application/vnd.openxmlformats-package.relationships+xml"

	// PPTX 核心内容类型
	ContentTypePresentation    = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
	ContentTypeSlide           = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
	ContentTypeSlideLayout     = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
	ContentTypeSlideMaster     = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
	ContentTypeNotesSlide      = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
	ContentTypeHandoutMaster   = "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml"
	ContentTypeNotesMaster     = "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"
	ContentTypePresentationML  = "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"

	// 主题和样式
	ContentTypeTheme          = "application/vnd.openxmlformats-officedocument.theme+xml"
	ContentTypeThemeOverride  = "application/vnd.openxmlformats-officedocument.themeOverride+xml"
	ContentTypeStyles         = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"

	// 图表
	ContentTypeChart          = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
	ContentTypeChartEx        = "application/vnd.ms-office.chartex+xml"

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
	ContentTypeMP4  = "video/mp4"
	ContentTypeAVI  = "video/x-msvideo"
	ContentTypeWMV  = "video/x-ms-wmv"

	// 其他
	ContentTypeXML  = "application/xml"
	ContentTypeFont = "application/x-font"

	// 默认内容类型映射（基于扩展名）
	ContentTypeDefault = "application/octet-stream"
)

// 关系类型常量 (Relationship Types)
const (
	// OPC 核心关系
	RelTypeCoreProperties = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"

	// Office 文档关系
	RelTypeOfficeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"

	// 扩展属性
	RelTypeExtendedProperties = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
	RelTypeCustomProperties   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"

	// 幻灯片关系
	RelTypeSlide        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
	RelTypeSlideLayout  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
	RelTypeSlideMaster  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
	RelTypeNotesSlide   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
	RelTypeNotesMaster  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster"
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

// 默认内容类型映射（扩展名 -> 内容类型）
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

// 内容类型到扩展名的反向映射
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

// GetContentTypeByExtension 根据文件扩展名获取内容类型
func GetContentTypeByExtension(ext string) string {
	if ct, ok := DefaultContentTypes[ext]; ok {
		return ct
	}
	return ContentTypeDefault
}

// GetExtensionByContentType 根据内容类型获取文件扩展名
func GetExtensionByContentType(contentType string) string {
	if ext, ok := ContentTypeToExtension[contentType]; ok {
		return ext
	}
	return ".bin"
}

// OPC 命名空间
const (
	NamespaceOPCPackage      = "http://schemas.openxmlformats.org/package/2006/content-types"
	NamespaceRelationships   = "http://schemas.openxmlformats.org/package/2006/relationships"
	NamespaceRelationshipsNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)

// OPC 默认路径
const (
	PathContentTypes = "[Content_Types].xml"
	PathRelsDir      = "_rels"
	PathRelsFile     = ".rels"
)

// XMLDeclaration OPC 包中所有 XML 文件的标准声明头
const XMLDeclaration = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`

// IsImmutableContentType 判断内容类型是否为不可变资源
// 不可变资源可以使用 zero-copy 共享，无需深拷贝
func IsImmutableContentType(contentType string) bool {
	switch contentType {
	// 图片类型 - 二进制数据，只读
	case ContentTypePNG, ContentTypeJPEG, ContentTypeGIF,
		ContentTypeBMP, ContentTypeTIFF, ContentTypeWMF,
		ContentTypeEMF, ContentTypeSVG:
		return true

	// 音视频类型 - 二进制数据，只读
	case ContentTypeWAV, ContentTypeMP3, ContentTypeMIDI,
		ContentTypeMP4, ContentTypeAVI, ContentTypeWMV:
		return true

	// 主题和母版 - 模板文件，通常不变
	case ContentTypeTheme, ContentTypeThemeOverride,
		ContentTypeSlideMaster, ContentTypeSlideLayout:
		return true

	// 字体文件 - 只读
	case ContentTypeFont:
		return true

	default:
		return false
	}
}

// IsLargeBinaryContentType 判断是否为大块二进制内容
// 用于判断是否值得使用 zero-copy 优化
func IsLargeBinaryContentType(contentType string) bool {
	switch contentType {
	case ContentTypePNG, ContentTypeJPEG, ContentTypeGIF,
		ContentTypeBMP, ContentTypeTIFF, ContentTypeWMF,
		ContentTypeEMF, ContentTypeSVG,
		ContentTypeWAV, ContentTypeMP3, ContentTypeMIDI,
		ContentTypeMP4, ContentTypeAVI, ContentTypeWMV,
		ContentTypeFont:
		return true
	default:
		return false
	}
}

// IsImageContentType 判断是否为图片内容类型
func IsImageContentType(contentType string) bool {
	switch contentType {
	case ContentTypePNG, ContentTypeJPEG, ContentTypeGIF,
		ContentTypeBMP, ContentTypeTIFF, ContentTypeWMF,
		ContentTypeEMF, ContentTypeSVG:
		return true
	default:
		return false
	}
}

// IsMediaContentType 判断是否为音视频内容类型
func IsMediaContentType(contentType string) bool {
	switch contentType {
	case ContentTypeWAV, ContentTypeMP3, ContentTypeMIDI,
		ContentTypeMP4, ContentTypeAVI, ContentTypeWMV:
		return true
	default:
		return false
	}
}
