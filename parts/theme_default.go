package parts

// ============================================================================
// 默认主题模板 - Office 主题
// ============================================================================
//
// 基于标准 Office 主题，在程序启动时只解析一次
// 用于创建新主题时的快速克隆
//
// ============================================================================

import (
	"sync"
)

// DefaultThemeXML 默认 Office 主题 XML（完整版）
// 直接复制自标准 Office 主题，确保内容不缺漏
const DefaultThemeXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office 主题">
    <a:themeElements>
        <a:clrScheme name="Office">
            <a:dk1>
                <a:sysClr val="windowText" lastClr="000000"/>
            </a:dk1>
            <a:lt1>
                <a:sysClr val="window" lastClr="FFFFFF"/>
            </a:lt1>
            <a:dk2>
                <a:srgbClr val="44546A"/>
            </a:dk2>
            <a:lt2>
                <a:srgbClr val="E7E6E6"/>
            </a:lt2>
            <a:accent1>
                <a:srgbClr val="5B9BD5"/>
            </a:accent1>
            <a:accent2>
                <a:srgbClr val="ED7D31"/>
            </a:accent2>
            <a:accent3>
                <a:srgbClr val="A5A5A5"/>
            </a:accent3>
            <a:accent4>
                <a:srgbClr val="FFC000"/>
            </a:accent4>
            <a:accent5>
                <a:srgbClr val="4472C4"/>
            </a:accent5>
            <a:accent6>
                <a:srgbClr val="70AD47"/>
            </a:accent6>
            <a:hlink>
                <a:srgbClr val="0563C1"/>
            </a:hlink>
            <a:folHlink>
                <a:srgbClr val="954F72"/>
            </a:folHlink>
        </a:clrScheme>
        <a:fontScheme name="Office">
            <a:majorFont>
                <a:latin typeface="Calibri Light" panose="020F0302020204030204"/>
                <a:ea typeface=""/>
                <a:cs typeface=""/>
                <a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>
                <a:font script="Hang" typeface="맑은 고딕"/>
                <a:font script="Hans" typeface="宋体"/>
                <a:font script="Hant" typeface="新細明體"/>
                <a:font script="Arab" typeface="Times New Roman"/>
                <a:font script="Hebr" typeface="Times New Roman"/>
                <a:font script="Thai" typeface="Angsana New"/>
                <a:font script="Ethi" typeface="Nyala"/>
                <a:font script="Beng" typeface="Vrinda"/>
                <a:font script="Gujr" typeface="Shruti"/>
                <a:font script="Khmr" typeface="MoolBoran"/>
                <a:font script="Knda" typeface="Tunga"/>
                <a:font script="Guru" typeface="Raavi"/>
                <a:font script="Cans" typeface="Euphemia"/>
                <a:font script="Cher" typeface="Plantagenet Cherokee"/>
                <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
                <a:font script="Tibt" typeface="Microsoft Himalaya"/>
                <a:font script="Thaa" typeface="MV Boli"/>
                <a:font script="Deva" typeface="Mangal"/>
                <a:font script="Telu" typeface="Gautami"/>
                <a:font script="Taml" typeface="Latha"/>
                <a:font script="Syrc" typeface="Estrangelo Edessa"/>
                <a:font script="Orya" typeface="Kalinga"/>
                <a:font script="Mlym" typeface="Kartika"/>
                <a:font script="Laoo" typeface="DokChampa"/>
                <a:font script="Sinh" typeface="Iskoola Pota"/>
                <a:font script="Mong" typeface="Mongolian Baiti"/>
                <a:font script="Viet" typeface="Times New Roman"/>
                <a:font script="Uigh" typeface="Microsoft Uighur"/>
                <a:font script="Geor" typeface="Sylfaen"/>
            </a:majorFont>
            <a:minorFont>
                <a:latin typeface="Calibri" panose="020F0502020204030204"/>
                <a:ea typeface=""/>
                <a:cs typeface=""/>
                <a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>
                <a:font script="Hang" typeface="맑은 고딕"/>
                <a:font script="Hans" typeface="宋体"/>
                <a:font script="Hant" typeface="新細明體"/>
                <a:font script="Arab" typeface="Arial"/>
                <a:font script="Hebr" typeface="Arial"/>
                <a:font script="Thai" typeface="Cordia New"/>
                <a:font script="Ethi" typeface="Nyala"/>
                <a:font script="Beng" typeface="Vrinda"/>
                <a:font script="Gujr" typeface="Shruti"/>
                <a:font script="Khmr" typeface="DaunPenh"/>
                <a:font script="Knda" typeface="Tunga"/>
                <a:font script="Guru" typeface="Raavi"/>
                <a:font script="Cans" typeface="Euphemia"/>
                <a:font script="Cher" typeface="Plantagenet Cherokee"/>
                <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
                <a:font script="Tibt" typeface="Microsoft Himalaya"/>
                <a:font script="Thaa" typeface="MV Boli"/>
                <a:font script="Deva" typeface="Mangal"/>
                <a:font script="Telu" typeface="Gautami"/>
                <a:font script="Taml" typeface="Latha"/>
                <a:font script="Syrc" typeface="Estrangelo Edessa"/>
                <a:font script="Orya" typeface="Kalinga"/>
                <a:font script="Mlym" typeface="Kartika"/>
                <a:font script="Laoo" typeface="DokChampa"/>
                <a:font script="Sinh" typeface="Iskoola Pota"/>
                <a:font script="Mong" typeface="Mongolian Baiti"/>
                <a:font script="Viet" typeface="Arial"/>
                <a:font script="Uigh" typeface="Microsoft Uighur"/>
                <a:font script="Geor" typeface="Sylfaen"/>
            </a:minorFont>
        </a:fontScheme>
        <a:fmtScheme name="Office">
            <a:fillStyleLst>
                <a:solidFill>
                    <a:schemeClr val="phClr"/>
                </a:solidFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0">
                            <a:schemeClr val="phClr">
                                <a:lumMod val="110000"/>
                                <a:satMod val="105000"/>
                                <a:tint val="67000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="50000">
                            <a:schemeClr val="phClr">
                                <a:lumMod val="105000"/>
                                <a:satMod val="103000"/>
                                <a:tint val="73000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="100000">
                            <a:schemeClr val="phClr">
                                <a:lumMod val="105000"/>
                                <a:satMod val="109000"/>
                                <a:tint val="81000"/>
                            </a:schemeClr>
                        </a:gs>
                    </a:gsLst>
                    <a:lin ang="5400000" scaled="0"/>
                </a:gradFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0">
                            <a:schemeClr val="phClr">
                                <a:satMod val="103000"/>
                                <a:lumMod val="102000"/>
                                <a:tint val="94000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="50000">
                            <a:schemeClr val="phClr">
                                <a:satMod val="110000"/>
                                <a:lumMod val="100000"/>
                                <a:shade val="100000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="100000">
                            <a:schemeClr val="phClr">
                                <a:lumMod val="99000"/>
                                <a:satMod val="120000"/>
                                <a:shade val="78000"/>
                            </a:schemeClr>
                        </a:gs>
                    </a:gsLst>
                    <a:lin ang="5400000" scaled="0"/>
                </a:gradFill>
            </a:fillStyleLst>
            <a:lnStyleLst>
                <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill>
                        <a:schemeClr val="phClr"/>
                    </a:solidFill>
                    <a:prstDash val="solid"/>
                    <a:miter lim="800000"/>
                </a:ln>
                <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill>
                        <a:schemeClr val="phClr"/>
                    </a:solidFill>
                    <a:prstDash val="solid"/>
                    <a:miter lim="800000"/>
                </a:ln>
                <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill>
                        <a:schemeClr val="phClr"/>
                    </a:solidFill>
                    <a:prstDash val="solid"/>
                    <a:miter lim="800000"/>
                </a:ln>
            </a:lnStyleLst>
            <a:effectStyleLst>
                <a:effectStyle>
                    <a:effectLst/>
                </a:effectStyle>
                <a:effectStyle>
                    <a:effectLst/>
                </a:effectStyle>
                <a:effectStyle>
                    <a:effectLst>
                        <a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
                            <a:srgbClr val="000000">
                                <a:alpha val="63000"/>
                            </a:srgbClr>
                        </a:outerShdw>
                    </a:effectLst>
                </a:effectStyle>
            </a:effectStyleLst>
            <a:bgFillStyleLst>
                <a:solidFill>
                    <a:schemeClr val="phClr"/>
                </a:solidFill>
                <a:solidFill>
                    <a:schemeClr val="phClr">
                        <a:tint val="95000"/>
                        <a:satMod val="170000"/>
                    </a:schemeClr>
                </a:solidFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0">
                            <a:schemeClr val="phClr">
                                <a:tint val="93000"/>
                                <a:satMod val="150000"/>
                                <a:shade val="98000"/>
                                <a:lumMod val="102000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="50000">
                            <a:schemeClr val="phClr">
                                <a:tint val="98000"/>
                                <a:satMod val="130000"/>
                                <a:shade val="90000"/>
                                <a:lumMod val="103000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="100000">
                            <a:schemeClr val="phClr">
                                <a:shade val="63000"/>
                                <a:satMod val="120000"/>
                            </a:schemeClr>
                        </a:gs>
                    </a:gsLst>
                    <a:lin ang="5400000" scaled="0"/>
                </a:gradFill>
            </a:bgFillStyleLst>
        </a:fmtScheme>
    </a:themeElements>
    <a:objectDefaults/>
    <a:extraClrSchemeLst/>
    <a:extLst>
        <a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}">
            <thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/>
        </a:ext>
    </a:extLst>
</a:theme>`

// defaultThemeOnce 单例：解析后的默认主题
// 程序启动时通过 sync.Once 解析一次
var defaultThemeOnce struct {
	theme *XTheme
	once sync.Once
}

// DefaultTheme 返回默认主题（懒加载解析）
func DefaultTheme() *XTheme {
	if defaultThemeOnce.theme == nil {
		theme, err := ParseTheme([]byte(DefaultThemeXML))
		if err != nil {
			panic("failed to parse default theme: " + err.Error())
		}
		defaultThemeOnce.theme = theme
	}
	return defaultThemeOnce.theme
}

// CloneTheme 克隆主题（深拷贝）
func CloneTheme() *XTheme {
	src := DefaultTheme()
	if src == nil {
		return nil
	}

	// 深拷贝
	dst := &XTheme{
		XmlnsA: src.XmlnsA,
		Name:   src.Name,
	}

	if src.ThemeElements != nil {
		dst.ThemeElements = &XThemeElements{
			ColorScheme: cloneColorScheme(src.ThemeElements.ColorScheme),
			FontScheme:  cloneFontScheme(src.ThemeElements.FontScheme),
			FmtScheme:   cloneFmtScheme(src.ThemeElements.FmtScheme),
		}
	}

	return dst
}

// cloneColorScheme 深拷贝颜色方案
func cloneColorScheme(src *XColorScheme) *XColorScheme {
	if src == nil {
		return nil
	}
	dst := &XColorScheme{
		Name:            src.Name,
		Dark1:            cloneColorVariant(src.Dark1),
		Light1:           cloneColorVariant(src.Light1),
		Dark2:            cloneColorVariant(src.Dark2),
		Light2:           cloneColorVariant(src.Light2),
		Accent1:          cloneColorVariant(src.Accent1),
		Accent2:          cloneColorVariant(src.Accent2),
		Accent3:          cloneColorVariant(src.Accent3),
		Accent4:          cloneColorVariant(src.Accent4),
		Accent5:          cloneColorVariant(src.Accent5),
		Accent6:          cloneColorVariant(src.Accent6),
		Hyperlink:        cloneColorVariant(src.Hyperlink),
		FollowedHyperlink: cloneColorVariant(src.FollowedHyperlink),
	}
	return dst
}

// cloneColorVariant 深拷贝颜色变体
func cloneColorVariant(src *XColorVariant) *XColorVariant {
	if src == nil {
		return nil
	}
	dst := &XColorVariant{}
	if src.SRGBColor != nil {
		dst.SRGBColor = &XSRGBColor{Val: src.SRGBColor.Val}
	}
	if src.SysColor != nil {
		dst.SysColor = &XSysColor{
			Val:     src.SysColor.Val,
			LastClr: src.SysColor.LastClr,
		}
	}
	return dst
}
// cloneFontScheme 深拷贝字体方案
func cloneFontScheme(src *XFontScheme) *XFontScheme {
	if src == nil {
		return nil
	}
	dst := &XFontScheme{
		Name: src.Name,
	}
	if src.MajorFont != nil {
		dst.MajorFont = cloneFontCollection(src.MajorFont)
	}
	if src.MinorFont != nil {
		dst.MinorFont = cloneFontCollection(src.MinorFont)
	}
	return dst
}
// cloneFontCollection 深拷贝字体集合
func cloneFontCollection(src *XFontCollection) *XFontCollection {
	if src == nil {
		return nil
	}
	dst := &XFontCollection{
		Latin:    src.Latin,
		EastAsia: src.EastAsia,
		Complex:  src.Complex,
	}
	if len(src.Fonts) > 0 {
		dst.Fonts = make([]XScriptFont, len(src.Fonts))
		copy(dst.Fonts, src.Fonts)
	}
	return dst
}
// cloneFmtScheme 深拷贝格式方案
func cloneFmtScheme(src *XFmtScheme) *XFmtScheme {
	if src == nil {
		return nil
	}
	dst := &XFmtScheme{
		Name: src.Name,
	}
	if src.FillStyleLst != nil {
		dst.FillStyleLst = &XFillStyleList{InnerXML: src.FillStyleLst.InnerXML}
	}
	if src.LnStyleLst != nil {
		dst.LnStyleLst = &XLineStyleList{InnerXML: src.LnStyleLst.InnerXML}
	}
	if src.EffectStyleLst != nil {
		dst.EffectStyleLst = &XEffectStyleList{InnerXML: src.EffectStyleLst.InnerXML}
	}
	if src.BgFillStyleLst != nil {
		dst.BgFillStyleLst = &XFillStyleList{InnerXML: src.BgFillStyleLst.InnerXML}
	}
	return dst
}
