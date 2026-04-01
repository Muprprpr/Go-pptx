package parts

import (
	"encoding/xml"
	"fmt"
	"sync"

	"github.com/Muprprpr/Go-pptx/opc"
)

// ============================================================================
// ThemePart - 主题部件
// ============================================================================
//
// 对应 /ppt/theme/themeN.xml
// 定义演示文稿的颜色方案、字体方案和格式方案
//
// ============================================================================

// ThemePart 主题部件
type ThemePart struct {
	uri *opc.PackURI

	// 主题数据
	theme *XTheme

	// 缓存的颜色查找表
	colorCache map[ColorRole]*XColorVariant

	mu sync.RWMutex
}

// NewThemePart 创建新的主题部件
func NewThemePart(id int) *ThemePart {
	return &ThemePart{
		uri:        opc.NewPackURI(fmt.Sprintf("/ppt/theme/theme%d.xml", id)),
		colorCache: make(map[ColorRole]*XColorVariant),
	}
}

// NewThemePartWithURI 使用指定 URI 创建主题部件
func NewThemePartWithURI(uri *opc.PackURI) *ThemePart {
	return &ThemePart{
		uri:        uri,
		colorCache: make(map[ColorRole]*XColorVariant),
	}
}

// PartURI 返回部件 URI
func (t *ThemePart) PartURI() *opc.PackURI {
	return t.uri
}

// Theme 返回主题数据
func (t *ThemePart) Theme() *XTheme {
	t.mu.RLock()
	defer t.mu.RUnlock()
	return t.theme
}

// SetThemeData 设置主题数据（用于设置克隆后的主题）
func (t *ThemePart) SetThemeData(theme *XTheme) {
	t.mu.Lock()
	defer t.mu.Unlock()
	t.theme = theme
	// 清空缓存
	t.colorCache = make(map[ColorRole]*XColorVariant)
}

// ColorScheme 返回颜色方案
func (t *ThemePart) ColorScheme() *XColorScheme {
	t.mu.RLock()
	defer t.mu.RUnlock()
	if t.theme == nil || t.theme.ThemeElements == nil {
		return nil
	}
	return t.theme.ThemeElements.ColorScheme
}

// FontScheme 返回字体方案
func (t *ThemePart) FontScheme() *XFontScheme {
	t.mu.RLock()
	defer t.mu.RUnlock()
	if t.theme == nil || t.theme.ThemeElements == nil {
		return nil
	}
	return t.theme.ThemeElements.FontScheme
}

// ============================================================================
// 颜色访问方法
// ============================================================================

// GetThemeColor 获取指定角色的颜色
func (t *ThemePart) GetThemeColor(role ColorRole) *XColorVariant {
	t.mu.RLock()
	defer t.mu.RUnlock()

	// 检查缓存
	if c, ok := t.colorCache[role]; ok {
		return c
	}

	// 从颜色方案中获取
	scheme := t.ColorScheme()
	if scheme == nil {
		return nil
	}

	var color *XColorVariant
	switch role {
	case ColorRoleDark1:
		color = scheme.Dark1
	case ColorRoleLight1:
		color = scheme.Light1
	case ColorRoleDark2:
		color = scheme.Dark2
	case ColorRoleLight2:
		color = scheme.Light2
	case ColorRoleAccent1:
		color = scheme.Accent1
	case ColorRoleAccent2:
		color = scheme.Accent2
	case ColorRoleAccent3:
		color = scheme.Accent3
	case ColorRoleAccent4:
		color = scheme.Accent4
	case ColorRoleAccent5:
		color = scheme.Accent5
	case ColorRoleAccent6:
		color = scheme.Accent6
	case ColorRoleHyperlink:
		color = scheme.Hyperlink
	case ColorRoleFollowedHyperlink:
		color = scheme.FollowedHyperlink
	}

	// 缓存结果
	if color != nil {
		t.colorCache[role] = color
	}

	return color
}

// GetThemeColorRGB 获取指定角色的 RGB 颜色值
// 返回 6 位十六进制字符串（如 "FF0000"），如果无法获取则返回空字符串
func (t *ThemePart) GetThemeColorRGB(role ColorRole) string {
	color := t.GetThemeColor(role)
	if color == nil {
		return ""
	}

	if color.SRGBColor != nil {
		return color.SRGBColor.Val
	}

	if color.SysColor != nil && color.SysColor.LastClr != "" {
		return color.SysColor.LastClr
	}

	return ""
}

// GetThemeColorType 获取指定角色的颜色类型
func (t *ThemePart) GetThemeColorType(role ColorRole) ColorType {
	color := t.GetThemeColor(role)
	if color == nil {
		return ColorTypeUnknown
	}

	if color.SRGBColor != nil {
		return ColorTypeRGB
	}

	if color.SysColor != nil {
		return ColorTypeSystem
	}

	return ColorTypeUnknown
}

// ============================================================================
// 颜色设置方法
// ============================================================================

// SetThemeColorRGB 设置指定角色的 RGB 颜色值
// rgb 为 6 位十六进制字符串（如 "FF0000"）
func (t *ThemePart) SetThemeColorRGB(role ColorRole, rgb string) {
	t.mu.Lock()
	defer t.mu.Unlock()

	// 确保主题结构存在
	t.ensureThemeStructure()

	scheme := t.theme.ThemeElements.ColorScheme

	// 创建颜色变体
	color := &XColorVariant{
		SRGBColor: &XSRGBColor{Val: rgb},
	}

	// 设置对应角色的颜色
	switch role {
	case ColorRoleDark1:
		scheme.Dark1 = color
	case ColorRoleLight1:
		scheme.Light1 = color
	case ColorRoleDark2:
		scheme.Dark2 = color
	case ColorRoleLight2:
		scheme.Light2 = color
	case ColorRoleAccent1:
		scheme.Accent1 = color
	case ColorRoleAccent2:
		scheme.Accent2 = color
	case ColorRoleAccent3:
		scheme.Accent3 = color
	case ColorRoleAccent4:
		scheme.Accent4 = color
	case ColorRoleAccent5:
		scheme.Accent5 = color
	case ColorRoleAccent6:
		scheme.Accent6 = color
	case ColorRoleHyperlink:
		scheme.Hyperlink = color
	case ColorRoleFollowedHyperlink:
		scheme.FollowedHyperlink = color
	}

	// 更新缓存
	t.colorCache[role] = color
}

// SetThemeColorSystem 设置指定角色的系统颜色
// sysColorName 为系统颜色名称（如 "windowText", "window"）
// lastClr 为回退 RGB 颜色值（6 位十六进制）
func (t *ThemePart) SetThemeColorSystem(role ColorRole, sysColorName, lastClr string) {
	t.mu.Lock()
	defer t.mu.Unlock()

	// 确保主题结构存在
	t.ensureThemeStructure()

	scheme := t.theme.ThemeElements.ColorScheme

	// 创建颜色变体
	color := &XColorVariant{
		SysColor: &XSysColor{
			Val:     sysColorName,
			LastClr: lastClr,
		},
	}

	// 设置对应角色的颜色
	switch role {
	case ColorRoleDark1:
		scheme.Dark1 = color
	case ColorRoleLight1:
		scheme.Light1 = color
	case ColorRoleDark2:
		scheme.Dark2 = color
	case ColorRoleLight2:
		scheme.Light2 = color
	case ColorRoleAccent1:
		scheme.Accent1 = color
	case ColorRoleAccent2:
		scheme.Accent2 = color
	case ColorRoleAccent3:
		scheme.Accent3 = color
	case ColorRoleAccent4:
		scheme.Accent4 = color
	case ColorRoleAccent5:
		scheme.Accent5 = color
	case ColorRoleAccent6:
		scheme.Accent6 = color
	case ColorRoleHyperlink:
		scheme.Hyperlink = color
	case ColorRoleFollowedHyperlink:
		scheme.FollowedHyperlink = color
	}

	// 更新缓存
	t.colorCache[role] = color
}

// ensureThemeStructure 确保主题结构存在（调用时需持有锁）
func (t *ThemePart) ensureThemeStructure() {
	if t.theme == nil {
		t.theme = &XTheme{
			XmlnsA: "http://schemas.openxmlformats.org/drawingml/2006/main",
		}
	}
	if t.theme.ThemeElements == nil {
		t.theme.ThemeElements = &XThemeElements{}
	}
	if t.theme.ThemeElements.ColorScheme == nil {
		t.theme.ThemeElements.ColorScheme = &XColorScheme{}
	}
	if t.theme.ThemeElements.FontScheme == nil {
		t.theme.ThemeElements.FontScheme = &XFontScheme{}
	}
}

// ============================================================================
// 字体设置方法
// ============================================================================

// SetThemeMajorFont 设置标题字体
func (t *ThemePart) SetThemeMajorFont(latin, eastAsia, complex string) {
	t.mu.Lock()
	defer t.mu.Unlock()

	t.ensureThemeStructure()

	fontScheme := t.theme.ThemeElements.FontScheme
	if fontScheme.MajorFont == nil {
		fontScheme.MajorFont = &XFontCollection{}
	}

	fontScheme.MajorFont.Latin = latin
	fontScheme.MajorFont.EastAsia = eastAsia
	fontScheme.MajorFont.Complex = complex
}

// SetThemeMinorFont 设置正文字体
func (t *ThemePart) SetThemeMinorFont(latin, eastAsia, complex string) {
	t.mu.Lock()
	defer t.mu.Unlock()

	t.ensureThemeStructure()

	fontScheme := t.theme.ThemeElements.FontScheme
	if fontScheme.MinorFont == nil {
		fontScheme.MinorFont = &XFontCollection{}
	}

	fontScheme.MinorFont.Latin = latin
	fontScheme.MinorFont.EastAsia = eastAsia
	fontScheme.MinorFont.Complex = complex
}

// SetThemeScriptFont 设置脚本特定字体
// isMajor: true 为标题字体，false 为正文字体
func (t *ThemePart) SetThemeScriptFont(isMajor bool, script, typeface string) {
	t.mu.Lock()
	defer t.mu.Unlock()

	t.ensureThemeStructure()

	fontScheme := t.theme.ThemeElements.FontScheme
	var fontColl *XFontCollection
	if isMajor {
		if fontScheme.MajorFont == nil {
			fontScheme.MajorFont = &XFontCollection{}
		}
		fontColl = fontScheme.MajorFont
	} else {
		if fontScheme.MinorFont == nil {
			fontScheme.MinorFont = &XFontCollection{}
		}
		fontColl = fontScheme.MinorFont
	}

	// 查找是否已存在该脚本的字体
	for i, f := range fontColl.Fonts {
		if f.Script == script {
			fontColl.Fonts[i].Typeface = typeface
			return
		}
	}

	// 添加新字体
	fontColl.Fonts = append(fontColl.Fonts, XScriptFont{
		Script:   script,
		Typeface: typeface,
	})
}

// ============================================================================
// XML 序列化/反序列化
// ============================================================================

// ToXML 将主题序列化为 XML
func (t *ThemePart) ToXML() ([]byte, error) {
	t.mu.RLock()
	defer t.mu.RUnlock()

	if t.theme == nil {
		return nil, fmt.Errorf("theme data is nil")
	}

	output, err := xml.MarshalIndent(t.theme, "", "  ")
	if err != nil {
		return nil, err
	}

	return append([]byte(XMLDeclaration), output...), nil
}

// FromXML 从 XML 反序列化主题
func (t *ThemePart) FromXML(data []byte) error {
	// 去除命名空间前缀以兼容 Go 的 xml.Unmarshal
	cleanData, err := StripNamespacePrefixes(data)
	if err != nil {
		return fmt.Errorf("failed to clean XML: %w", err)
	}

	var theme XTheme
	if err := xml.Unmarshal(cleanData, &theme); err != nil {
		return fmt.Errorf("failed to unmarshal theme XML: %w", err)
	}

	t.mu.Lock()
	defer t.mu.Unlock()

	t.theme = &theme

	// 清空缓存
	t.colorCache = make(map[ColorRole]*XColorVariant)

	return nil
}

// ParseTheme 从 XML 字节解析主题
func ParseTheme(data []byte) (*XTheme, error) {
	cleanData, err := StripNamespacePrefixes(data)
	if err != nil {
		return nil, fmt.Errorf("failed to clean XML: %w", err)
	}

	var theme XTheme
	if err := xml.Unmarshal(cleanData, &theme); err != nil {
		return nil, fmt.Errorf("failed to parse theme: %w", err)
	}

	return &theme, nil
}

// ============================================================================
// XColorVariant 辅助方法
// ============================================================================

// Type 返回颜色类型
func (c *XColorVariant) Type() ColorType {
	if c == nil {
		return ColorTypeUnknown
	}
	if c.SRGBColor != nil {
		return ColorTypeRGB
	}
	if c.SysColor != nil {
		return ColorTypeSystem
	}
	return ColorTypeUnknown
}

// RGB 返回 RGB 颜色值
// 对于 RGB 颜色，直接返回值
// 对于系统颜色，返回 LastClr（回退颜色）
func (c *XColorVariant) RGB() string {
	if c == nil {
		return ""
	}

	if c.SRGBColor != nil {
		return c.SRGBColor.Val
	}

	if c.SysColor != nil {
		return c.SysColor.LastClr
	}

	return ""
}

// IsRGB 判断是否为 RGB 颜色
func (c *XColorVariant) IsRGB() bool {
	return c != nil && c.SRGBColor != nil
}

// IsSystem 判断是否为系统颜色
func (c *XColorVariant) IsSystem() bool {
	return c != nil && c.SysColor != nil
}

// SystemColorName 返回系统颜色名称
func (c *XColorVariant) SystemColorName() string {
	if c == nil || c.SysColor == nil {
		return ""
	}
	return c.SysColor.Val
}
