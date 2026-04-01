package parts

// ============================================================================
// Theme XML 结构类型定义 - 对应 /ppt/theme/themeN.xml
// ============================================================================
//
// 主题文件定义了演示文稿的颜色方案、字体方案和格式方案。
// 命名空间: http://schemas.openxmlformats.org/drawingml/2006/main
//
// 文件位置示例：
//   - /ppt/theme/theme1.xml
//   - /ppt/theme/theme2.xml
//
// ============================================================================

// XTheme 主题 XML 根元素
type XTheme struct {
	XMLName       struct{}          `xml:"theme"`
	XmlnsA        string            `xml:"xmlns:a,attr"`
	Name          string            `xml:"name,attr,omitempty"`
	ThemeElements *XThemeElements  `xml:"themeElements"`
}

// XThemeElements 主题元素集合
type XThemeElements struct {
	ColorScheme *XColorScheme `xml:"clrScheme"`
	FontScheme  *XFontScheme  `xml:"fontScheme,omitempty"`
	FmtScheme   *XFmtScheme   `xml:"fmtScheme,omitempty"`
}

// ============================================================================
// 颜色方案 (Color Scheme)
// ============================================================================

// XColorScheme 颜色方案
// 定义演示文稿中使用的 12 种标准颜色
type XColorScheme struct {
	XMLName   struct{}        `xml:"clrScheme"`
	XmlnsA    string          `xml:"xmlns:a,attr,omitempty"`
	Name      string          `xml:"name,attr,omitempty"`
	Dark1     *XColorVariant  `xml:"dk1"`      // 深色 1
	Light1    *XColorVariant  `xml:"lt1"`      // 浅色 1
	Dark2     *XColorVariant  `xml:"dk2"`      // 深色 2
	Light2    *XColorVariant  `xml:"lt2"`      // 浅色 2
	Accent1   *XColorVariant  `xml:"accent1"`  // 强调色 1
	Accent2   *XColorVariant  `xml:"accent2"`  // 强调色 2
	Accent3   *XColorVariant  `xml:"accent3"`  // 强调色 3
	Accent4   *XColorVariant  `xml:"accent4"`  // 强调色 4
	Accent5   *XColorVariant  `xml:"accent5"`  // 强调色 5
	Accent6   *XColorVariant  `xml:"accent6"`  // 强调色 6
	Hyperlink *XColorVariant  `xml:"hlink"`    // 超链接
	FollowedHyperlink *XColorVariant `xml:"folHlink"` // 访问过的超链接
}

// XColorVariant 颜色变体（可以是 RGB 或系统颜色）
type XColorVariant struct {
	SRGBColor *XSRGBColor `xml:"srgbClr,omitempty"`
	SysColor  *XSysColor  `xml:"sysClr,omitempty"`
}

// XSRGBColor RGB 颜色
type XSRGBColor struct {
	Val string `xml:"val,attr"` // 6 位十六进制 RGB 值（如 "FF0000"）
}

// XSysColor 系统颜色
type XSysColor struct {
	Val     string `xml:"val,attr"`           // 系统颜色名称
	LastClr string `xml:"lastClr,attr,omitempty"` // 最后使用的 RGB 值（回退颜色）
}

// ColorType 颜色类型枚举
type ColorType int

const (
	ColorTypeUnknown ColorType = iota
	ColorTypeRGB               // RGB 颜色
	ColorTypeSystem            // 系统颜色
)

// ColorRole 颜色角色枚举
// 用于标识颜色方案中的各个颜色
type ColorRole int

const (
	ColorRoleDark1 ColorRole = iota // 深色 1（通常是黑色）
	ColorRoleLight1                  // 浅色 1（通常是白色）
	ColorRoleDark2                   // 深色 2
	ColorRoleLight2                  // 浅色 2
	ColorRoleAccent1                 // 强调色 1
	ColorRoleAccent2                 // 强调色 2
	ColorRoleAccent3                 // 强调色 3
	ColorRoleAccent4                 // 强调色 4
	ColorRoleAccent5                 // 强调色 5
	ColorRoleAccent6                 // 强调色 6
	ColorRoleHyperlink               // 超链接
	ColorRoleFollowedHyperlink       // 访问过的超链接
)

// String 返回颜色角色的名称
func (r ColorRole) String() string {
	names := []string{
		"Dark1",
		"Light1",
		"Dark2",
		"Light2",
		"Accent1",
		"Accent2",
		"Accent3",
		"Accent4",
		"Accent5",
		"Accent6",
		"Hyperlink",
		"FollowedHyperlink",
	}
	if int(r) < len(names) {
		return names[r]
	}
	return "Unknown"
}

// ============================================================================
// 字体方案 (Font Scheme)
// ============================================================================

// XFontScheme 字体方案
type XFontScheme struct {
	XMLName   struct{}      `xml:"fontScheme"`
	XmlnsA    string        `xml:"xmlns:a,attr,omitempty"`
	Name      string        `xml:"name,attr,omitempty"`
	MajorFont *XFontCollection `xml:"majorFont,omitempty"` // 标题字体
	MinorFont *XFontCollection `xml:"minorFont,omitempty"` // 正文字体
}

// XFontCollection 字体集合
type XFontCollection struct {
	Latin    string          `xml:"latin typeface,attr,omitempty"` // 拉丁字体
	EastAsia string          `xml:"ea typeface,attr,omitempty"`    // 东亚字体
	Complex  string          `xml:"cs typeface,attr,omitempty"`    // 复杂脚本字体
	Fonts    []XScriptFont   `xml:"font"`                          // 脚本特定字体
}

// XScriptFont 脚本特定字体
type XScriptFont struct {
	Script   string `xml:"script,attr"`   // 脚本代码（如 "Jpan", "Hans"）
	Typeface string `xml:"typeface,attr"` // 字体名称
}

// ============================================================================
// 格式方案 (Format Scheme)
// ============================================================================

// XFmtScheme 格式方案
// 包含填充、线条、效果和背景填充的样式列表
type XFmtScheme struct {
	XMLName        struct{}        `xml:"fmtScheme"`
	XmlnsA         string          `xml:"xmlns:a,attr,omitempty"`
	Name           string          `xml:"name,attr,omitempty"`
	FillStyleLst   *XFillStyleList   `xml:"fillStyleLst,omitempty"`
	LnStyleLst     *XLineStyleList   `xml:"lnStyleLst,omitempty"`
	EffectStyleLst *XEffectStyleList `xml:"effectStyleLst,omitempty"`
	BgFillStyleLst *XFillStyleList   `xml:"bgFillStyleLst,omitempty"`
}

// XFillStyleList 填充样式列表
type XFillStyleList struct {
	InnerXML string `xml:",innerxml"` // 保留原始 XML 内容，避免丢失数据
}

// XLineStyleList 线条样式列表
type XLineStyleList struct {
	InnerXML string `xml:",innerxml"` // 保留原始 XML 内容
}

// XEffectStyleList 效果样式列表
type XEffectStyleList struct {
	InnerXML string `xml:",innerxml"` // 保留原始 XML 内容
}
