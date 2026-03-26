package utils

// ============================================================================
// PowerPoint 单位转换工具
// ============================================================================
//
// PowerPoint 内部使用 EMU (English Metric Units) 作为基本单位。
// 1 英寸 = 914400 EMU
// 1 厘米 = 360000 EMU
//
// 参考: ECMA-376 Office Open XML File Formats
// ============================================================================

import (
	"math"
	"strconv"
)

// EMU (English Metric Unit) 常量定义
const (
	// EMUsPerInch: 1 英寸对应的 EMU 数量
	EMUsPerInch int64 = 914400

	// EMUsPerCentimeter: 1 厘米对应的 EMU 数量
	EMUsPerCentimeter int64 = 360000

	// EMUsPerMillimeter: 1 毫米对应的 EMU 数量
	EMUsPerMillimeter int64 = 36000

	// EMUsPerPoint: 1 磅对应的 EMU 数量 (1 磅 = 1/72 英寸)
	EMUsPerPoint int64 = EMUsPerInch / 72

	// EMUsPerPixel: 默认像素分辨率下的 EMU 数量 (96 DPI)
	EMUsPerPixel int64 = EMUsPerInch / 96
)

// EMU 浮点常量（用于乘法运算）
var (
	// EMUsPerInchF: 1 英寸对应的 EMU 数量（浮点）
	EMUsPerInchF float64 = 914400.0

	// EMUsPerCentimeterF: 1 厘米对应的 EMU 数量（浮点）
	EMUsPerCentimeterF float64 = 360000.0

	// EMUsPerMillimeterF: 1 毫米对应的 EMU 数量（浮点）
	EMUsPerMillimeterF float64 = 36000.0

	// EMUsPerPointF: 1 磅对应的 EMU 数量（浮点）
	EMUsPerPointF float64 = 914400.0 / 72.0

	// EMUsPerPixelF: 默认像素分辨率下的 EMU 数量（浮点）
	EMUsPerPixelF float64 = 914400.0 / 96.0
)

// ============================================================================
// 类型别名和结构体
// ============================================================================

// EMU int64 类型的别名，用于语义化表达
type EMU int64

// Unit 类型，支持多种单位
type Unit struct {
	Value float64
	Unit  UnitType
}

// UnitType 单位类型
type UnitType int

const (
	UnitTypeEMU       UnitType = iota // EMU
	UnitTypeInch                      // 英寸
	UnitTypeCentimeter                // 厘米
	UnitTypeMillimeter                // 毫米
	UnitTypePoint                     // 磅/点
	UnitTypePixel                     // 像素
)

// ============================================================================
// EMU 单位转换函数
// ============================================================================

// InchesToEMU 将英寸转换为 EMU
func InchesToEMU(inches float64) int64 {
	return int64(math.Round(inches * EMUsPerInchF))
}

// EMUToInches 将 EMU 转换为英寸
func EMUToInches(emu int64) float64 {
	return float64(emu) / EMUsPerInchF
}

// CentimetersToEMU 将厘米转换为 EMU
func CentimetersToEMU(cm float64) int64 {
	return int64(math.Round(cm * EMUsPerCentimeterF))
}

// EMUToCentimeters 将 EMU 转换为厘米
func EMUToCentimeters(emu int64) float64 {
	return float64(emu) / EMUsPerCentimeterF
}

// MillimetersToEMU 将毫米转换为 EMU
func MillimetersToEMU(mm float64) int64 {
	return int64(math.Round(mm * EMUsPerMillimeterF))
}

// EMUToMillimeters 将 EMU 转换为毫米
func EMUToMillimeters(emu int64) float64 {
	return float64(emu) / EMUsPerMillimeterF
}

// PointsToEMU 将磅转换为 EMU
func PointsToEMU(points float64) int64 {
	return int64(math.Round(points * EMUsPerPointF))
}

// EMUToPoints 将 EMU 转换为磅
func EMUToPoints(emu int64) float64 {
	return float64(emu) / EMUsPerPointF
}

// PixelsToEMU 将像素转换为 EMU（基于 96 DPI）
func PixelsToEMU(pixels float64) int64 {
	return int64(math.Round(pixels * EMUsPerPixelF))
}

// EMUToPixels 将 EMU 转换为像素（基于 96 DPI）
func EMUToPixels(emu int64) float64 {
	return float64(emu) / EMUsPerPixelF
}

// ============================================================================
// EMU 方法
// ============================================================================

// NewEMU 创建 EMU 值
func NewEMU(value int64) EMU {
	return EMU(value)
}

// Inches 英寸转 EMU
func (e EMU) Inches() float64 {
	return EMUToInches(int64(e))
}

// Centimeters EMU 转厘米
func (e EMU) Centimeters() float64 {
	return EMUToCentimeters(int64(e))
}

// Millimeters EMU 转毫米
func (e EMU) Millimeters() float64 {
	return EMUToMillimeters(int64(e))
}

// Points EMU 转磅
func (e EMU) Points() float64 {
	return EMUToPoints(int64(e))
}

// Pixels EMU 转像素
func (e EMU) Pixels() float64 {
	return EMUToPixels(int64(e))
}

// ============================================================================
// 单位转换器
// ============================================================================

// Converter 单位转换器
type Converter struct {
	dpi            int
	emusPerPixelF  float64
}

// NewConverter 创建新的单位转换器
func NewConverter() *Converter {
	return &Converter{
		dpi:           96,
		emusPerPixelF: EMUsPerInchF / 96.0,
	}
}

// NewConverterWithDPI 创建指定 DPI 的单位转换器
func NewConverterWithDPI(dpi int) *Converter {
	return &Converter{
		dpi:           dpi,
		emusPerPixelF: EMUsPerInchF / float64(dpi),
	}
}

// EMUsPerPixelForDPI 根据 DPI 计算每像素 EMU 数
func (c *Converter) EMUsPerPixelForDPI() int64 {
	return int64(math.Round(c.emusPerPixelF))
}

// PixelsToEMU 将像素转换为 EMU（基于当前 DPI）
func (c *Converter) PixelsToEMU(pixels float64) int64 {
	return int64(math.Round(pixels * c.emusPerPixelF))
}

// EMUToPixels 将 EMU 转换为像素（基于当前 DPI）
func (c *Converter) EMUToPixels(emu int64) float64 {
	return float64(emu) / c.emusPerPixelF
}

// SetDPI 设置 DPI
func (c *Converter) SetDPI(dpi int) {
	c.dpi = dpi
	c.emusPerPixelF = EMUsPerInchF / float64(dpi)
}

// DPI 返回当前 DPI
func (c *Converter) DPI() int {
	return c.dpi
}

// ============================================================================
// 常用尺寸常量
// ============================================================================

// SlideWidth 标准幻灯片宽度（宽屏 16:9）
const SlideWidth EMU = EMU(12192000) // 13.333 英寸

// SlideHeight 标准幻灯片高度（宽屏 16:9）
const SlideHeight EMU = EMU(6858000) // 7.5 英寸

// StandardSlideWidth 标准幻灯片宽度（4:3）
const StandardSlideWidth EMU = EMU(9144000) // 10 英寸

// StandardSlideHeight 标准幻灯片高度（4:3）
const StandardSlideHeight EMU = EMU(6858000) // 7.5 英寸

// ============================================================================
// 便捷函数
// ============================================================================

// MakeEMUPair 创建 EMU 坐标对
func MakeEMUPair(x, y int64) (int64, int64) {
	return x, y
}

// MakeSizeEMU 创建 EMU 尺寸
func MakeSizeEMU(cx, cy int64) (int64, int64) {
	return cx, cy
}

// RectEMU EMU 矩形
type RectEMU struct {
	X  int64
	Y  int64
	Cx int64
	Cy int64
}

// NewRectEMU 创建 EMU 矩形
func NewRectEMU(x, y, cx, cy int64) RectEMU {
	return RectEMU{X: x, Y: y, Cx: cx, Cy: cy}
}

// FromInches 从英寸创建 EMU 矩形
func (r RectEMU) FromInches(x, y, width, height float64) RectEMU {
	return RectEMU{
		X:  InchesToEMU(x),
		Y:  InchesToEMU(y),
		Cx: InchesToEMU(width),
		Cy: InchesToEMU(height),
	}
}

// FromCentimeters 从厘米创建 EMU 矩形
func (r RectEMU) FromCentimeters(x, y, width, height float64) RectEMU {
	return RectEMU{
		X:  CentimetersToEMU(x),
		Y:  CentimetersToEMU(y),
		Cx: CentimetersToEMU(width),
		Cy: CentimetersToEMU(height),
	}
}

// ============================================================================
// XML 属性写入辅助
// ============================================================================

// WriteEMUAttr 写入 EMU 属性值到字符串
func WriteEMUAttr(value int64) string {
	return strconv.FormatInt(value, 10)
}

// WriteEMUAttrs 写入多个 EMU 属性
func WriteEMUAttrs(values ...int64) []string {
	attrs := make([]string, len(values))
	for i, v := range values {
		attrs[i] = strconv.FormatInt(v, 10)
	}
	return attrs
}
