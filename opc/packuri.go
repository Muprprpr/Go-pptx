package opc

import (
	"path"
	"path/filepath"
	"strings"
)

// PackURI 表示包内的URI（如 /ppt/slides/slide1.xml）
// 遵循 OPC 规范中的 URI 规则
type PackURI struct {
	uri      string
	segments []string
}

// NewPackURI 创建一个新的 PackURI
func NewPackURI(uri string) *PackURI {
	// 规范化 URI：确保以 / 开头
	if !strings.HasPrefix(uri, "/") {
		uri = "/" + uri
	}
	// 移除重复的斜杠
	for strings.Contains(uri, "//") {
		uri = strings.ReplaceAll(uri, "//", "/")
	}

	p := &PackURI{
		uri: uri,
	}
	p.segments = p.parseSegments()
	return p
}

// parseSegments 解析 URI 的各个段
func (p *PackURI) parseSegments() []string {
	trimmed := strings.Trim(p.uri, "/")
	if trimmed == "" {
		return []string{}
	}
	return strings.Split(trimmed, "/")
}

// URI 返回原始 URI 字符串
func (p *PackURI) URI() string {
	return p.uri
}

// String 返回 URI 字符串表示
func (p *PackURI) String() string {
	return p.uri
}

// BaseName 返回 URI 的基本名称（不含扩展名）
// 例如: /ppt/slides/slide1.xml -> slide1
func (p *PackURI) BaseName() string {
	filename := p.FileName()
	ext := path.Ext(filename)
	return strings.TrimSuffix(filename, ext)
}

// FileName 返回文件名（含扩展名）
// 例如: /ppt/slides/slide1.xml -> slide1.xml
func (p *PackURI) FileName() string {
	return path.Base(p.uri)
}

// Extension 返回文件扩展名
// 例如: /ppt/slides/slide1.xml -> .xml
func (p *PackURI) Extension() string {
	return path.Ext(p.uri)
}

// DirName 返回父目录路径
// 例如: /ppt/slides/slide1.xml -> /ppt/slides
func (p *PackURI) DirName() string {
	return path.Dir(p.uri)
}

// Segments 返回 URI 的各个段
func (p *PackURI) Segments() []string {
	return p.segments
}

// MemberName 返回相对于包根目录的成员名称（用于 ZIP）
// 去掉开头的 /
func (p *PackURI) MemberName() string {
	return strings.TrimPrefix(p.uri, "/")
}

// RelPath 返回相对路径（用于关系目标）
func (p *PackURI) RelPath() string {
	return p.MemberName()
}

// Join 连接相对路径并返回新的 PackURI
func (p *PackURI) Join(relativePath string) *PackURI {
	// 处理绝对路径
	if strings.HasPrefix(relativePath, "/") {
		return NewPackURI(relativePath)
	}

	// 解析相对路径
	result := p.DirName()
	parts := strings.Split(relativePath, "/")

	for _, part := range parts {
		if part == ".." {
			// 向上一级目录
			if result != "/" && result != "" {
				result = path.Dir(result)
			}
		} else if part != "." && part != "" {
			// 添加路径段
			result = path.Join(result, part)
		}
	}

	return NewPackURI(result)
}

// RelPathFrom 计算从另一个 URI 到此 URI 的相对路径
func (p *PackURI) RelPathFrom(other *PackURI) string {
	fromSegs := other.DirSegments()
	toSegs := p.DirSegments()

	// 找到公共前缀
	commonLen := 0
	minLen := len(fromSegs)
	if len(toSegs) < minLen {
		minLen = len(toSegs)
	}

	for i := 0; i < minLen; i++ {
		if fromSegs[i] == toSegs[i] {
			commonLen++
		} else {
			break
		}
	}

	// 构建相对路径
	var upLevels []string
	for i := commonLen; i < len(fromSegs); i++ {
		upLevels = append(upLevels, "..")
	}

	var downPath []string
	for i := commonLen; i < len(toSegs); i++ {
		downPath = append(downPath, toSegs[i])
	}
	downPath = append(downPath, p.FileName())

	relativeParts := append(upLevels, downPath...)
	if len(relativeParts) == 0 {
		return p.FileName()
	}
	return strings.Join(relativeParts, "/")
}

// DirSegments 返回目录段的切片（不含文件名）
func (p *PackURI) DirSegments() []string {
	dir := p.DirName()
	if dir == "/" || dir == "." {
		return []string{}
	}
	return strings.Split(strings.Trim(dir, "/"), "/")
}

// IsRelationshipsPart 检查是否为关系部件
func (p *PackURI) IsRelationshipsPart() bool {
	return strings.HasSuffix(p.FileName(), ".rels")
}

// RelationshipsURI 返回此部件对应的关系文件 URI
// 例如: /ppt/slides/slide1.xml -> /ppt/slides/_rels/slide1.xml.rels
func (p *PackURI) RelationshipsURI() *PackURI {
	if p.IsRelationshipsPart() {
		return p
	}

	dir := p.DirName()
	filename := p.FileName()

	// 构建关系文件路径
	relPath := path.Join(dir, PathRelsDir, filename+".rels")
	return NewPackURI(relPath)
}

// SourceURI 从关系文件路径获取源部件 URI
// 例如: /ppt/slides/_rels/slide1.xml.rels -> /ppt/slides/slide1.xml
func (p *PackURI) SourceURI() *PackURI {
	if !p.IsRelationshipsPart() {
		return p
	}

	// 解析关系文件路径
	dir := p.DirName()
	filename := p.FileName()

	// 检查是否在 _rels 目录中
	if path.Base(dir) != PathRelsDir {
		return p
	}

	// 移除 _rels 目录
	parentDir := path.Dir(dir)
	// 移除 .rels 扩展名
	sourceFilename := strings.TrimSuffix(filename, ".rels")

	return NewPackURI(path.Join(parentDir, sourceFilename))
}

// Equals 比较两个 PackURI 是否相等
func (p *PackURI) Equals(other *PackURI) bool {
	if other == nil {
		return false
	}
	return p.uri == other.uri
}

// EqualsStr 比较与字符串 URI 是否相等
func (p *PackURI) EqualsStr(uri string) bool {
	other := NewPackURI(uri)
	return p.Equals(other)
}

// IsAbsolute 检查是否为绝对路径
func (p *PackURI) IsAbsolute() bool {
	return strings.HasPrefix(p.uri, "/")
}

// Clone 创建 PackURI 的副本
func (p *PackURI) Clone() *PackURI {
	return NewPackURI(p.uri)
}

// MarshalText 实现 encoding.TextMarshaler 接口
func (p PackURI) MarshalText() ([]byte, error) {
	return []byte(p.uri), nil
}

// UnmarshalText 实现 encoding.TextUnmarshaler 接口
func (p *PackURI) UnmarshalText(data []byte) error {
	p.uri = string(data)
	if !strings.HasPrefix(p.uri, "/") {
		p.uri = "/" + p.uri
	}
	p.segments = p.parseSegments()
	return nil
}

// RootURI 返回包根目录 URI
func RootURI() *PackURI {
	return NewPackURI("/")
}

// ContentTypesURI 返回 [Content_Types].xml 的 URI
func ContentTypesURI() *PackURI {
	return NewPackURI("/" + PathContentTypes)
}

// PackageRelsURI 返回包级别关系文件的 URI
func PackageRelsURI() *PackURI {
	return NewPackURI("/" + PathRelsDir + "/" + PathRelsFile)
}

// IsPackageRels 检查是否为包级别关系文件
func (p *PackURI) IsPackageRels() bool {
	return p.uri == "/"+PathRelsDir+"/"+PathRelsFile
}

// IsValidPackURI 检查 URI 是否有效
func IsValidPackURI(uri string) bool {
	// 基本验证
	if uri == "" {
		return false
	}

	// 必须以 / 开头（绝对路径）
	if !strings.HasPrefix(uri, "/") {
		return false
	}

	// 检查非法字符
	illegalChars := []string{"\\", ":", "*", "?", "\"", "<", ">", "|"}
	for _, char := range illegalChars {
		if strings.Contains(uri, char) {
			return false
		}
	}

	return true
}

// NormalizeURI 规范化 URI
func NormalizeURI(uri string) string {
	// 将反斜杠转换为正斜杠
	uri = filepath.ToSlash(uri)

	// 确保以 / 开头
	if !strings.HasPrefix(uri, "/") {
		uri = "/" + uri
	}

	// 移除重复的斜杠
	for strings.Contains(uri, "//") {
		uri = strings.ReplaceAll(uri, "//", "/")
	}

	// 移除结尾的斜杠（除非是根目录）
	if len(uri) > 1 && strings.HasSuffix(uri, "/") {
		uri = strings.TrimSuffix(uri, "/")
	}

	return uri
}
