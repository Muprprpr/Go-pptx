// Package pptx 提供 PPTX 文件的高级操作接口
package pptx

import (
	"fmt"

	"github.com/Muprprpr/Go-pptx/opc"
	"github.com/Muprprpr/Go-pptx/parts"
)

// ============================================================================
// SlideBuilder - 高层幻灯片构建器
// ============================================================================
//
// 提供便捷的 API 来构建幻灯片内容，包括：
// - 添加文本框
// - 添加自动形状
// - 添加图片
// - 添加表格
// - 管理媒体关系
// ============================================================================

// SlideBuilder 幻灯片构建器
type SlideBuilder struct {
	slide *parts.SlidePart
}

// NewSlideBuilder 创建幻灯片构建器
func NewSlideBuilder(slide *parts.SlidePart) *SlideBuilder {
	return &SlideBuilder{
		slide: slide,
	}
}

// Slide 返回底层 SlidePart
func (b *SlideBuilder) Slide() *parts.SlidePart {
	return b.slide
}

// ============================================================================
// 形状添加方法
// ============================================================================

// AddTextBox 添加文本框到幻灯片
// x, y, cx, cy: 位置和尺寸（EMU 单位）
// text: 文本内容
func (b *SlideBuilder) AddTextBox(x, y, cx, cy int, text string) *parts.XSp {
	sp := &parts.XSp{
		NonVisual: parts.XNonVisualDrawingShape{
			CNvPr: &parts.XNvCxnSpPr{
				ID:   int(b.slide.AllocateShapeID()),
				Name: fmt.Sprintf("TextBox %d", b.slide.CurrentShapeID()),
			},
			CNvSpPr: &parts.XNvSpPr{},
		},
		ShapeProperties: &parts.XShapeProperties{
			Transform2D: &parts.XTransform2D{
				Offset: &parts.XOv2DrOffset{X: x, Y: y},
				Extent: &parts.XOv2DrExtent{Cx: cx, Cy: cy},
			},
		},
		TextBody: &parts.XTextBody{
			BodyPr: &parts.XBodyPr{},
			LstStyle: &parts.XTextParagraphList{},
			Paragraphs: []parts.XTextParagraph{
				{
					TextRuns: []parts.XTextRun{
						{Text: text},
					},
				},
			},
		},
	}

	b.slide.AppendShapeChild(sp)
	return sp
}

// AddAutoShape 添加自动形状到幻灯片
// x, y, cx, cy: 位置和尺寸（EMU 单位）
// presetID: 预设形状类型 (如 "rectangle", "ellipse", "roundRect" 等)
func (b *SlideBuilder) AddAutoShape(x, y, cx, cy int, presetID string) *parts.XSp {
	sp := &parts.XSp{
		NonVisual: parts.XNonVisualDrawingShape{
			CNvPr: &parts.XNvCxnSpPr{
				ID:   int(b.slide.AllocateShapeID()),
				Name: fmt.Sprintf("%s %d", presetID, b.slide.CurrentShapeID()),
			},
			CNvSpPr: &parts.XNvSpPr{},
		},
		ShapeProperties: &parts.XShapeProperties{
			Transform2D: &parts.XTransform2D{
				Offset: &parts.XOv2DrOffset{X: x, Y: y},
				Extent: &parts.XOv2DrExtent{Cx: cx, Cy: cy},
			},
		},
		ShapePreset: presetID,
	}

	b.slide.AppendShapeChild(sp)
	return sp
}

// AddPicture 添加图片到幻灯片
// x, y, cx, cy: 位置和尺寸（EMU 单位）
// imageRId: 图片关系 ID
func (b *SlideBuilder) AddPicture(x, y, cx, cy int, imageRId string) *parts.XPicture {
	pic := &parts.XPicture{
		NonVisual: parts.XNonVisualDrawingPic{
			CNvPr: &parts.XNvCxnSpPr{
				ID:   int(b.slide.AllocateShapeID()),
				Name: fmt.Sprintf("Picture %d", b.slide.CurrentShapeID()),
			},
			CNvPicPr: &parts.XNvPicPr{},
		},
		BlipFill: &parts.XBlipFillProperties{
			Blip: &parts.XBlip{
				Embed: imageRId,
			},
			Stretch: &parts.XStretchProperties{},
		},
		ShapeProperties: &parts.XShapeProperties{
			Transform2D: &parts.XTransform2D{
				Offset: &parts.XOv2DrOffset{X: x, Y: y},
				Extent: &parts.XOv2DrExtent{Cx: cx, Cy: cy},
			},
		},
	}

	b.slide.AppendShapeChild(pic)
	return pic
}

// AddTable 添加表格到幻灯片
// x, y, cx, cy: 位置和尺寸（EMU 单位）
// rows, cols: 行列数
func (b *SlideBuilder) AddTable(x, y, cx, cy, rows, cols int) *parts.XGraphicFrame {
	// 计算单元格尺寸
	cellW := cx / cols

	// 构建表格网格
	gridCols := make([]parts.XTableColumn, cols)
	for i := range gridCols {
		gridCols[i] = parts.XTableColumn{W: cellW}
	}

	// 构建行
	tableRows := make([]parts.XTableRow, rows)
	for r := range tableRows {
		cells := make([]parts.XTableCell, cols)
		for c := range cells {
			cells[c] = parts.XTableCell{
				TextBody: &parts.XTextBody{
					BodyPr:   &parts.XBodyPr{},
					LstStyle: &parts.XTextParagraphList{},
					Paragraphs: []parts.XTextParagraph{
						{TextRuns: []parts.XTextRun{{Text: ""}}},
					},
				},
			}
		}
		tableRows[r] = parts.XTableRow{GridSpan: 1, Cells: cells}
	}

	table := parts.XTable{
		Grid: &parts.XTableGrid{GridCols: gridCols},
		Rows: tableRows,
	}

	gf := &parts.XGraphicFrame{
		NonVisual: parts.XNonVisualGraphicFrame{
			CNvPr: &parts.XNvCxnSpPr{
				ID:   int(b.slide.AllocateShapeID()),
				Name: fmt.Sprintf("Table %d", b.slide.CurrentShapeID()),
			},
			CNvGraphicFramePr: &parts.XNvGraphicFramePr{},
		},
		Graphic: &parts.XGraphic{
			Table: &table,
		},
		Transform2D: &parts.XTransform2D{
			Offset: &parts.XOv2DrOffset{X: x, Y: y},
			Extent: &parts.XOv2DrExtent{Cx: cx, Cy: cy},
		},
	}

	b.slide.AppendShapeChild(gf)
	return gf
}

// SetTableCellText 设置表格单元格文本
func (b *SlideBuilder) SetTableCellText(gf *parts.XGraphicFrame, row, col int, text string) {
	if gf == nil || gf.Graphic == nil || gf.Graphic.Table == nil {
		return
	}
	table := gf.Graphic.Table
	if row < 0 || row >= len(table.Rows) || col < 0 || col >= len(table.Rows[row].Cells) {
		return
	}
	table.Rows[row].Cells[col].TextBody.Paragraphs[0].TextRuns[0].Text = text
}

// GetOrAddPicture 添加图片到幻灯片并返回 XPicture
// 自动处理图片关系 ID
func (b *SlideBuilder) GetOrAddPicture(x, y, cx, cy int, imageURI string) *parts.XPicture {
	rId := b.GetImageRId(imageURI)
	return b.AddPicture(x, y, cx, cy, rId)
}

// ============================================================================
// 关系管理方法
// ============================================================================

// AddImage 添加图片关系并返回 rId
func (b *SlideBuilder) AddImage(targetURI string) string {
	return b.slide.AddImageRel(targetURI)
}

// AddMedia 添加媒体关系并返回 rId
func (b *SlideBuilder) AddMedia(targetURI string) string {
	return b.slide.AddMediaRel(targetURI)
}

// AddChart 添加图表关系并返回 rId
func (b *SlideBuilder) AddChart(targetURI string) string {
	return b.slide.AddChartRel(targetURI)
}

// HasImage 判断是否已存在某图片关系
func (b *SlideBuilder) HasImage(targetURI string) bool {
	rels := b.slide.Relationships()
	target := opc.NewPackURI(targetURI)
	return rels.GetByTarget(target) != nil
}

// HasMedia 判断是否已存在某媒体关系
func (b *SlideBuilder) HasMedia(targetURI string) bool {
	rels := b.slide.Relationships()
	target := opc.NewPackURI(targetURI)
	return rels.GetByTarget(target) != nil
}

// GetImageRId 获取图片 rId，不存在则添加
func (b *SlideBuilder) GetImageRId(targetURI string) string {
	return b.slide.AddImageRel(targetURI)
}

// GetMediaRId 获取媒体 rId，不存在则添加
func (b *SlideBuilder) GetMediaRId(targetURI string) string {
	return b.slide.AddMediaRel(targetURI)
}

// GetChartRId 获取图表 rId，不存在则添加
func (b *SlideBuilder) GetChartRId(targetURI string) string {
	return b.slide.AddChartRel(targetURI)
}

// GetRelationshipURI 根据 rId 获取目标 URI
func (b *SlideBuilder) GetRelationshipURI(rId string) string {
	rels := b.slide.Relationships()
	rel := rels.Get(rId)
	if rel != nil {
		return rel.TargetURI().URI()
	}
	return ""
}
