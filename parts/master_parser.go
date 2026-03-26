package parts

import (
	"encoding/xml"
	"fmt"
	"strconv"

	"github.com/Muprprpr/Go-pptx/utils"
)

// ============================================================================
// 母版/版式 XML 解析器
// ============================================================================
//
// 将嵌套的 XML 结构体转换为干净的只读领域模型
// ============================================================================

// ParseLayout 解析幻灯片版式 XML
// 传入 xml 字节，解析并提取出该版式中所有的占位符坐标和类型
func ParseLayout(xmlData []byte) (*SlideLayoutData, error) {
	var xmlLayout XMLSlideLayout
	if err := xml.Unmarshal(xmlData, &xmlLayout); err != nil {
		return nil, fmt.Errorf("解析版式 XML 失败: %w", err)
	}

	if xmlLayout.CSld == nil || xmlLayout.CSld.SpTree == nil {
		return nil, fmt.Errorf("版式缺少必要节点 p:cSld 或 p:spTree")
	}

	// 提取占位符
	placeholders := extractPlaceholders(xmlLayout.CSld.SpTree)

	// 提取背景
	var background *Background
	if xmlLayout.CSld.Bg != nil {
		background = parseBackground(xmlLayout.CSld.Bg)
	}

	return &SlideLayoutData{
		id:           generateLayoutID(),
		name:         "",
		background:   background,
		placeholders: placeholders,
	}, nil
}

// ParseMaster 解析幻灯片母版 XML
// 传入 xml 字节，解析并提取出母版中的占位符定义
func ParseMaster(xmlData []byte) (*SlideMasterData, error) {
	var xmlMaster XMLSlideMaster
	if err := xml.Unmarshal(xmlData, &xmlMaster); err != nil {
		return nil, fmt.Errorf("解析母版 XML 失败: %w", err)
	}

	if xmlMaster.CSld == nil || xmlMaster.CSld.SpTree == nil {
		return nil, fmt.Errorf("母版缺少必要节点 p:cSld 或 p:spTree")
	}

	// 提取占位符
	placeholders := extractPlaceholders(xmlMaster.CSld.SpTree)

	// 提取背景
	var background *Background
	if xmlMaster.CSld.Bg != nil {
		background = parseBackground(xmlMaster.CSld.Bg)
	}

	return &SlideMasterData{
		id:           generateMasterID(),
		name:         "",
		background:   background,
		placeholders: placeholders,
	}, nil
}

// ============================================================================
// 占位符提取
// ============================================================================

// extractPlaceholders 从形状树中提取所有占位符
func extractPlaceholders(spTree *XMLShapeTree) map[string]*Placeholder {
	placeholders := make(map[string]*Placeholder)

	// 提取普通形状中的占位符
	for _, shape := range spTree.Shapes {
		ph := extractPlaceholderFromShape(shape)
		if ph != nil {
			placeholders[ph.id] = ph
		}
	}

	// 递归提取组形状中的占位符
	for _, grpSp := range spTree.GroupShapes {
		extractPlaceholdersFromGroup(&grpSp, placeholders)
	}

	return placeholders
}

// extractPlaceholdersFromGroup 从组形状中递归提取占位符
func extractPlaceholdersFromGroup(grpSp *XMLGroupShape, placeholders map[string]*Placeholder) {
	for _, shape := range grpSp.Shapes {
		ph := extractPlaceholderFromShape(shape)
		if ph != nil {
			placeholders[ph.id] = ph
		}
	}
}

// extractPlaceholderFromShape 从单个形状中提取占位符信息
func extractPlaceholderFromShape(shape XMLShape) *Placeholder {
	// 检查是否有占位符定义
	if shape.NvSpPr == nil || shape.NvSpPr.NvPr == nil || shape.NvSpPr.NvPr.Ph == nil {
		return nil
	}

	xmlPh := shape.NvSpPr.NvPr.Ph

	// 解析占位符类型
	phType := parsePlaceholderType(xmlPh.Type)

	// 生成占位符 ID
	phID := generatePlaceholderID(xmlPh.Idx, shape.NvSpPr.CNvPr)

	// 提取坐标和尺寸
	var x, y, cx, cy int64
	if shape.SpPr != nil && shape.SpPr.Xfrm != nil {
		if shape.SpPr.Xfrm.Off != nil {
			x = shape.SpPr.Xfrm.Off.X
			y = shape.SpPr.Xfrm.Off.Y
		}
		if shape.SpPr.Xfrm.Ext != nil {
			cx = shape.SpPr.Xfrm.Ext.Cx
			cy = shape.SpPr.Xfrm.Ext.Cy
		}
	}

	return &Placeholder{
		id:              phID,
		placeholderType: phType,
		x:               x,
		y:               y,
		cx:              cx,
		cy:              cy,
	}
}

// ============================================================================
// 背景解析
// ============================================================================

// parseBackground 解析背景 XML 结构
func parseBackground(xmlBg *XMLBackground) *Background {
	// 处理 <p:bgRef> 背景引用（主题色引用）
	if xmlBg.BgRef != nil {
		return &Background{
			backgroundType: BackgroundTypeThemeColor,
			solidColorRGB: xmlBg.BgRef.Clr.Val, // 主题色名称，如 "bg1"
			opacity:        1.0,
		}
	}

	// 处理 <p:bgPr> 背景属性
	if xmlBg.BgPr == nil || xmlBg.BgPr.Fill == nil {
		return nil
	}

	fill := xmlBg.BgPr.Fill

	// 纯色填充
	if fill.SolidFill != nil {
		rgb := extractColorFromSolidFill(fill.SolidFill)
		if rgb != "" {
			return &Background{
				backgroundType: BackgroundTypeSolidColor,
				solidColorRGB:  rgb,
				opacity:        1.0,
			}
		}
	}

	// 渐变填充
	if fill.GradFill != nil {
		return parseGradientBackground(fill.GradFill)
	}

	// 图片填充
	if fill.BlipFill != nil && fill.BlipFill.Blip != nil {
		return &Background{
			backgroundType: BackgroundTypePicture,
			pictureRId:     fill.BlipFill.Blip.Embed,
			opacity:        1.0,
		}
	}

	// 无填充
	if fill.NoFill != nil {
		return &Background{
			backgroundType: BackgroundTypeNone,
			opacity:        1.0,
		}
	}

	return nil
}

// parseGradientBackground 解析渐变背景
func parseGradientBackground(gradFill *XMLGradFill) *Background {
	bg := &Background{
		backgroundType: BackgroundTypeGradient,
		opacity:        1.0,
	}

	// 解析渐变角度
	if gradFill.Lin != nil {
		bg.gradientAngle = int32(gradFill.Lin.Ang / 60000) // 转换为度
	}

	// 解析渐变色标
	if gradFill.GsLst != nil {
		bg.gradientColors = make([]GradientStop, 0, len(gradFill.GsLst.Stops))
		for _, stop := range gradFill.GsLst.Stops {
			rgb := ""
			if stop.SolidFill != nil {
				rgb = extractColorFromSolidFill(stop.SolidFill)
			}
			if rgb != "" {
				bg.gradientColors = append(bg.gradientColors, GradientStop{
					position: float32(stop.Pos) / 100000.0, // 转换为 0.0-1.0
					colorRGB: rgb,
				})
			}
		}
	}

	return bg
}

// extractColorFromSolidFill 从纯色填充中提取颜色
func extractColorFromSolidFill(solidFill *XMLSolidFill) string {
	if solidFill.SrgbClr != nil {
		return solidFill.SrgbClr.Val
	}
	if solidFill.SchemeClr != nil {
		// 主题色需要查表，这里暂时返回空
		// 后续可以扩展为主题色名称
		return ""
	}
	return ""
}

// ============================================================================
// 占位符类型转换
// ============================================================================

// parsePlaceholderType 将 XML 字符串类型转换为枚举
func parsePlaceholderType(typeStr string) PlaceholderType {
	switch typeStr {
	case "title":
		return PlaceholderTypeTitle
	case "body":
		return PlaceholderTypeBody
	case "ctrTitle":
		return PlaceholderTypeCenterTitle
	case "subTitle":
		return PlaceholderTypeSubTitle
	case "dt":
		return PlaceholderTypeDateTime
	case "sldNum":
		return PlaceholderTypeSlideNumber
	case "ftr":
		return PlaceholderTypeFooter
	case "hdr":
		return PlaceholderTypeHeader
	case "obj":
		return PlaceholderTypeObject
	case "chart":
		return PlaceholderTypeChart
	case "tbl":
		return PlaceholderTypeTable
	case "clipArt":
		return PlaceholderTypeClipArt
	case "dgm":
		return PlaceholderTypeOrgChart
	case "media":
		return PlaceholderTypeMedia
	case "sldImg":
		return PlaceholderTypeSlideImage
	case "pic":
		return PlaceholderTypePicture
	default:
		return PlaceholderTypeNone
	}
}

// ============================================================================
// ID 生成
// ============================================================================

var (
	layoutIDCounter  int64
	masterIDCounter  int64
)

// generateLayoutID 生成版式 ID
func generateLayoutID() string {
	layoutIDCounter++
	return "layout_" + strconv.FormatInt(layoutIDCounter, 10)
}

// generateMasterID 生成母版 ID
func generateMasterID() string {
	masterIDCounter++
	return "master_" + strconv.FormatInt(masterIDCounter, 10)
}

// generatePlaceholderID 生成占位符 ID
// 优先使用 XML 中的 idx，否则使用 cNvPr 中的 id
func generatePlaceholderID(xmlIdx string, cnvPr *XMLCNvPr) string {
	if xmlIdx != "" {
		return "ph_" + xmlIdx
	}
	if cnvPr != nil && cnvPr.ID > 0 {
		return "ph_" + strconv.Itoa(cnvPr.ID)
	}
	// 生成随机 ID
	return "ph_unknown"
}

// ============================================================================
// EMU 单位转换辅助（便捷方法）
// ============================================================================

// EMUToPixels 将 EMU 转换为像素（96 DPI）
var EMUToPixels = utils.EMUToPixels

// EMUToPoints 将 EMU 转换为磅
var EMUToPoints = utils.EMUToPoints

// EMUToInches 将 EMU 转换为英寸
var EMUToInches = utils.EMUToInches

// EMUToCentimeters 将 EMU 转换为厘米
var EMUToCentimeters = utils.EMUToCentimeters
