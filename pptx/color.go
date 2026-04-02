// Package pptx 提供 PPTX 文件的高级操作接口
package pptx

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

// ============================================================================
// 颜色系统 - 颜色映射与验证
// ============================================================================
//
// OOXML 颜色规范：
// 1. 输出格式必须为 6 位十六进制，无 # 前缀（如 "FF0000"）
// 2. 透明度使用独立标签 <a:alpha val="50000"/>，范围 0-100000
// 3. 统一输入格式：#RRGGBBAA（8位，最后两位为 alpha）
//    - #FF0000FF = 100% 不透明红色
//    - #FF000000 = 0% 完全透明红色
//
// PowerPoint 支持两种颜色表示方式：
// 1. RGB 颜色 (srgbClr) - 直接指定，如 "FF0000" 表示红色
// 2. 主题颜色 (schemeClr) - 引用主题中的颜色，如 "accent1"
//
// ============================================================================

// Alpha 常量（OOXML 范围：0-100000）
const (
	// AlphaOpaque 完全不透明
	AlphaOpaque = 100000
	// AlphaTransparent 完全透明
	AlphaTransparent = 0
	// AlphaDefault 默认透明度（100%）
	AlphaDefault = 100000
)

// ColorType 颜色类型
type ColorType int

const (
	// ColorTypeRGB RGB 颜色
	ColorTypeRGB ColorType = iota
	// ColorTypeScheme 主题颜色
	ColorTypeScheme
	// ColorTypeInvalid 无效颜色
	ColorTypeInvalid
)

// Color 颜色结构体
type Color struct {
	// Type 颜色类型
	Type ColorType
	// RGB 十六进制值 (6位，如 "FF0000"，无 # 前缀)
	RGB string
	// Scheme 主题颜色名称 (如 "accent1")
	Scheme string
	// Alpha 透明度 (0-100000，OOXML 标准)
	// 100000 = 100% 不透明，0 = 完全透明
	Alpha int
	// IsValid 是否有效
	IsValid bool
}

// ============================================================================
// 预设颜色常量
// ============================================================================

// 预设颜色常量（6位十六进制，无 # 前缀）
var (
	// 基础颜色
	ColorBlack   = RGBColor("000000")
	ColorWhite   = RGBColor("FFFFFF")
	ColorRed     = RGBColor("FF0000")
	ColorGreen   = RGBColor("00FF00")
	ColorBlue    = RGBColor("0000FF")
	ColorYellow  = RGBColor("FFFF00")
	ColorCyan    = RGBColor("00FFFF")
	ColorMagenta = RGBColor("FF00FF")

	// 透明色
	ColorTransparent = Color{Type: ColorTypeRGB, RGB: "000000", Alpha: AlphaTransparent, IsValid: true}

	// 常用 UI 颜色
	ColorGray      = RGBColor("808080")
	ColorLightGray = RGBColor("C0C0C0")
	ColorDarkGray  = RGBColor("404040")
	ColorOrange    = RGBColor("FFA500")
	ColorPurple    = RGBColor("800080")
	ColorPink      = RGBColor("FFC0CB")
	ColorBrown     = RGBColor("A52A2A")
	ColorNavy      = RGBColor("000080")
	ColorTeal      = RGBColor("008080")
	ColorOlive     = RGBColor("808000")
	ColorMaroon    = RGBColor("800000")
	ColorLime      = RGBColor("00FF00")
	ColorAqua      = RGBColor("00FFFF")
	ColorSilver    = RGBColor("C0C0C0")
	ColorGold      = RGBColor("FFD700")
)

// ============================================================================
// 主题颜色常量
// ============================================================================

// 主题颜色名称常量
const (
	// 背景色
	SchemeBg1 = "bg1" // 背景色 1 (通常是白色或浅色)
	SchemeBg2 = "bg2" // 背景色 2
	SchemeFg1 = "fg1" // 前景色/文本色 1 (通常是黑色或深色)
	SchemeFg2 = "fg2" // 前景色/文本色 2

	// 强调色
	SchemeAccent1 = "accent1" // 强调色 1
	SchemeAccent2 = "accent2" // 强调色 2
	SchemeAccent3 = "accent3" // 强调色 3
	SchemeAccent4 = "accent4" // 强调色 4
	SchemeAccent5 = "accent5" // 强调色 5
	SchemeAccent6 = "accent6" // 强调色 6

	// 超链接色
	SchemeHlink    = "hlink"    // 超链接颜色
	SchemeFolHlink = "folHlink" // 已访问超链接颜色

	// 特殊色
	SchemePhClr = "phClr" // 幻灯片标题颜色
	SchemeTx1   = "tx1"   // 文本色 1
	SchemeTx2   = "tx2"   // 文本色 2
)

// 主题颜色列表
var SchemeColors = []string{
	SchemeBg1, SchemeBg2, SchemeFg1, SchemeFg2,
	SchemeAccent1, SchemeAccent2, SchemeAccent3, SchemeAccent4, SchemeAccent5, SchemeAccent6,
	SchemeHlink, SchemeFolHlink,
	SchemeTx1, SchemeTx2,
}

// ============================================================================
// 颜色解析函数
// ============================================================================

// hexColorRegex 十六进制颜色正则（支持 6 位和 8 位）
var hexColorRegex = regexp.MustCompile(`^#?([0-9A-Fa-f]{6})([0-9A-Fa-f]{2})?$`)

// hexColor3Regex 3位十六进制颜色正则
var hexColor3Regex = regexp.MustCompile(`^#?([0-9A-Fa-f]{3})$`)

// rgbColorRegex RGB 颜色正则
var rgbColorRegex = regexp.MustCompile(`^rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$`)

// rgbaColorRegex RGBA 颜色正则
var rgbaColorRegex = regexp.MustCompile(`^rgba\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*([\d.]+)\s*\)$`)

// ParseColor 解析颜色字符串
// 支持格式：
// - "#FF0000" 或 "FF0000" (6位十六进制)
// - "#FF0000FF" (8位十六进制，最后两位为 alpha)
// - "rgb(255, 0, 0)" (RGB)
// - "rgba(255, 0, 0, 0.5)" (RGBA，alpha 0-1)
// - "accent1", "bg1" 等主题颜色名称
func ParseColor(s string) Color {
	s = strings.TrimSpace(s)
	if s == "" {
		return Color{Type: ColorTypeInvalid, IsValid: false}
	}

	// 尝试解析 8 位十六进制颜色（#RRGGBBAA）
	if matches := hexColorRegex.FindStringSubmatch(s); matches != nil {
		hex := strings.ToUpper(matches[1])
		alpha := AlphaDefault
		if matches[2] != "" {
			alpha = hexToAlpha(matches[2])
		}
		return Color{
			Type:    ColorTypeRGB,
			RGB:     hex,
			Alpha:   alpha,
			IsValid: true,
		}
	}

	// 尝试解析 3 位十六进制颜色
	if matches := hexColor3Regex.FindStringSubmatch(s); matches != nil {
		hex3 := matches[1]
		hex := strings.ToUpper(string(hex3[0]) + string(hex3[0]) + string(hex3[1]) + string(hex3[1]) + string(hex3[2]) + string(hex3[2]))
		return Color{
			Type:    ColorTypeRGB,
			RGB:     hex,
			Alpha:   AlphaDefault,
			IsValid: true,
		}
	}

	// 尝试解析 RGBA 颜色
	if matches := rgbaColorRegex.FindStringSubmatch(s); matches != nil {
		r, _ := strconv.Atoi(matches[1])
		g, _ := strconv.Atoi(matches[2])
		b, _ := strconv.Atoi(matches[3])
		a, _ := strconv.ParseFloat(matches[4], 64)
		alpha := int(a * float64(AlphaOpaque))
		return ColorFromRGBA(r, g, b, alpha)
	}

	// 尝试解析 RGB 颜色
	if matches := rgbColorRegex.FindStringSubmatch(s); matches != nil {
		r, _ := strconv.Atoi(matches[1])
		g, _ := strconv.Atoi(matches[2])
		b, _ := strconv.Atoi(matches[3])
		return ColorFromRGB(r, g, b)
	}

	// 尝试解析主题颜色
	if IsSchemeColor(s) {
		return SchemeColor(s)
	}

	return Color{Type: ColorTypeInvalid, IsValid: false}
}

// hexToAlpha 将 2 位十六进制转换为 OOXML alpha 值 (0-100000)
func hexToAlpha(hex string) int {
	val, _ := strconv.ParseInt(hex, 16, 64)
	// 0x00 -> 0 (透明), 0xFF -> 100000 (不透明)
	return int(float64(val) / 255.0 * float64(AlphaOpaque))
}

// alphaToHex 将 OOXML alpha 值转换为 2 位十六进制
func alphaToHex(alpha int) string {
	if alpha < 0 {
		alpha = 0
	}
	if alpha > AlphaOpaque {
		alpha = AlphaOpaque
	}
	val := int(float64(alpha) / float64(AlphaOpaque) * 255.0)
	return fmt.Sprintf("%02X", val)
}

// ============================================================================
// 颜色创建函数
// ============================================================================

// RGBColor 创建 RGB 颜色（无 alpha）
func RGBColor(hex string) Color {
	hex = strings.ToUpper(strings.TrimSpace(hex))
	hex = strings.TrimPrefix(hex, "#")

	// 验证十六进制格式
	if len(hex) != 6 {
		return Color{Type: ColorTypeInvalid, IsValid: false}
	}
	if _, err := strconv.ParseInt(hex, 16, 64); err != nil {
		return Color{Type: ColorTypeInvalid, IsValid: false}
	}

	return Color{
		Type:    ColorTypeRGB,
		RGB:     hex,
		Alpha:   AlphaDefault,
		IsValid: true,
	}
}

// RGBAColor 创建带 alpha 的 RGB 颜色
// hex: 6位十六进制 RGB
// alpha: 0-100000 (OOXML 标准)
func RGBAColor(hex string, alpha int) Color {
	c := RGBColor(hex)
	if !c.IsValid {
		return c
	}
	c.Alpha = alpha
	if c.Alpha < 0 {
		c.Alpha = 0
	}
	if c.Alpha > AlphaOpaque {
		c.Alpha = AlphaOpaque
	}
	return c
}

// ColorFromRGB 从 RGB 值创建颜色
func ColorFromRGB(r, g, b int) Color {
	if r < 0 || r > 255 || g < 0 || g > 255 || b < 0 || b > 255 {
		return Color{Type: ColorTypeInvalid, IsValid: false}
	}
	hex := fmt.Sprintf("%02X%02X%02X", r, g, b)
	return RGBColor(hex)
}

// ColorFromRGBA 从 RGBA 值创建颜色
// r, g, b: 0-255
// alpha: 0-100000 (OOXML 标准)
func ColorFromRGBA(r, g, b, alpha int) Color {
	c := ColorFromRGB(r, g, b)
	if !c.IsValid {
		return c
	}
	c.Alpha = alpha
	if c.Alpha < 0 {
		c.Alpha = 0
	}
	if c.Alpha > AlphaOpaque {
		c.Alpha = AlphaOpaque
	}
	return c
}

// SchemeColor 创建主题颜色
func SchemeColor(name string) Color {
	name = strings.ToLower(strings.TrimSpace(name))
	if !IsSchemeColor(name) {
		return Color{Type: ColorTypeInvalid, IsValid: false}
	}
	return Color{
		Type:    ColorTypeScheme,
		Scheme:  name,
		Alpha:   AlphaDefault,
		IsValid: true,
	}
}

// IsSchemeColor 检查是否为有效的主题颜色名称
func IsSchemeColor(name string) bool {
	name = strings.ToLower(strings.TrimSpace(name))
	for _, scheme := range SchemeColors {
		if scheme == name {
			return true
		}
	}
	return false
}

// ============================================================================
// 颜色输出函数 - 符合 OOXML 规范
// ============================================================================

// ToRGB 转换为 RGB 十六进制字符串（6位，无 # 前缀）
// 符合 OOXML 规范：<a:srgbClr val="FF0000"/>
func (c Color) ToRGB() string {
	if c.Type == ColorTypeRGB {
		return c.RGB
	}
	return ""
}

// ToHex 转换为带 # 前缀的 6 位十六进制字符串（用于显示）
func (c Color) ToHex() string {
	if c.Type == ColorTypeRGB {
		return "#" + c.RGB
	}
	return ""
}

// ToHexA 转换为带 # 前缀的 8 位十六进制字符串（包含 alpha）
func (c Color) ToHexA() string {
	if c.Type == ColorTypeRGB {
		return "#" + c.RGB + alphaToHex(c.Alpha)
	}
	return ""
}

// ToScheme 转换为主题颜色名称
func (c Color) ToScheme() string {
	if c.Type == ColorTypeScheme {
		return c.Scheme
	}
	return ""
}

// AlphaValue 返回 OOXML 格式的 alpha 值 (0-100000)
// 用于 <a:alpha val="50000"/>
func (c Color) AlphaValue() int {
	return c.Alpha
}

// AlphaPercent 返回百分比形式的 alpha (0-100)
func (c Color) AlphaPercent() float64 {
	return float64(c.Alpha) / float64(AlphaOpaque) * 100
}

// String 返回颜色的字符串表示
func (c Color) String() string {
	switch c.Type {
	case ColorTypeRGB:
		if c.Alpha != AlphaDefault {
			return "#" + c.RGB + alphaToHex(c.Alpha)
		}
		return "#" + c.RGB
	case ColorTypeScheme:
		return c.Scheme
	default:
		return "invalid"
	}
}

// RGBComponents 返回 RGB 分量 (r, g, b)
func (c Color) RGBComponents() (r, g, b int, ok bool) {
	if c.Type != ColorTypeRGB || len(c.RGB) != 6 {
		return 0, 0, 0, false
	}
	ri, _ := strconv.ParseInt(c.RGB[0:2], 16, 64)
	gi, _ := strconv.ParseInt(c.RGB[2:4], 16, 64)
	bi, _ := strconv.ParseInt(c.RGB[4:6], 16, 64)
	return int(ri), int(gi), int(bi), true
}

// WithAlpha 设置透明度并返回新颜色
// alpha: 0-100000 (OOXML 标准)
func (c Color) WithAlpha(alpha int) Color {
	if alpha < 0 {
		alpha = 0
	}
	if alpha > AlphaOpaque {
		alpha = AlphaOpaque
	}
	c.Alpha = alpha
	return c
}

// WithAlphaPercent 设置透明度（百分比）并返回新颜色
// percent: 0-100
func (c Color) WithAlphaPercent(percent float64) Color {
	if percent < 0 {
		percent = 0
	}
	if percent > 100 {
		percent = 100
	}
	c.Alpha = int(percent / 100.0 * float64(AlphaOpaque))
	return c
}

// ============================================================================
// 颜色验证函数
// ============================================================================

// ColorValidationResult 颜色验证结果
type ColorValidationResult struct {
	// IsValid 是否有效
	IsValid bool
	// Color 解析后的颜色
	Color Color
	// Original 原始输入
	Original string
	// Message 验证消息
	Message string
}

// ValidateColor 验证颜色
func ValidateColor(s string) ColorValidationResult {
	color := ParseColor(s)
	return ColorValidationResult{
		IsValid:  color.IsValid,
		Color:    color,
		Original: s,
		Message:  color.validateMessage(),
	}
}

// validateMessage 生成验证消息
func (c Color) validateMessage() string {
	if c.IsValid {
		switch c.Type {
		case ColorTypeRGB:
			if c.Alpha != AlphaDefault {
				return fmt.Sprintf("有效的 RGB 颜色: #%s (alpha: %d/100000)", c.RGB, c.Alpha)
			}
			return fmt.Sprintf("有效的 RGB 颜色: #%s", c.RGB)
		case ColorTypeScheme:
			return fmt.Sprintf("有效的主题颜色: %s", c.Scheme)
		}
	}
	return "无效的颜色格式"
}

// ============================================================================
// 颜色映射表
// ============================================================================

// ColorMap 颜色映射表
// 用于将颜色名称映射到实际颜色值
type ColorMap struct {
	colors map[string]Color
}

// NewColorMap 创建颜色映射表
func NewColorMap() *ColorMap {
	return &ColorMap{
		colors: make(map[string]Color),
	}
}

// DefaultColorMap 默认颜色映射表
func DefaultColorMap() *ColorMap {
	cm := NewColorMap()
	// 添加常用颜色
	cm.Set("black", ColorBlack)
	cm.Set("white", ColorWhite)
	cm.Set("red", ColorRed)
	cm.Set("green", ColorGreen)
	cm.Set("blue", ColorBlue)
	cm.Set("yellow", ColorYellow)
	cm.Set("cyan", ColorCyan)
	cm.Set("magenta", ColorMagenta)
	cm.Set("gray", ColorGray)
	cm.Set("grey", ColorGray)
	cm.Set("lightgray", ColorLightGray)
	cm.Set("lightgrey", ColorLightGray)
	cm.Set("darkgray", ColorDarkGray)
	cm.Set("darkgrey", ColorDarkGray)
	cm.Set("orange", ColorOrange)
	cm.Set("purple", ColorPurple)
	cm.Set("pink", ColorPink)
	cm.Set("brown", ColorBrown)
	cm.Set("navy", ColorNavy)
	cm.Set("teal", ColorTeal)
	cm.Set("olive", ColorOlive)
	cm.Set("maroon", ColorMaroon)
	cm.Set("lime", ColorLime)
	cm.Set("aqua", ColorAqua)
	cm.Set("silver", ColorSilver)
	cm.Set("gold", ColorGold)
	cm.Set("transparent", ColorTransparent)
	return cm
}

// Set 设置颜色映射
func (cm *ColorMap) Set(name string, color Color) {
	cm.colors[strings.ToLower(name)] = color
}

// Get 获取颜色映射
func (cm *ColorMap) Get(name string) (Color, bool) {
	color, ok := cm.colors[strings.ToLower(name)]
	return color, ok
}

// Resolve 解析颜色（支持名称、十六进制、RGB、主题色）
func (cm *ColorMap) Resolve(s string) Color {
	// 先尝试从映射表查找
	if color, ok := cm.Get(s); ok {
		return color
	}
	// 再尝试解析
	return ParseColor(s)
}

// All 返回所有颜色映射
func (cm *ColorMap) All() map[string]Color {
	result := make(map[string]Color, len(cm.colors))
	for k, v := range cm.colors {
		result[k] = v
	}
	return result
}
