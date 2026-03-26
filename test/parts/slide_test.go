package parts_test

import (
	"encoding/xml"
	"strings"
	"testing"

	"github.com/Muprprpr/Go-pptx/parts"
)

// TestSlideBuilder_AddText 测试 Slide Builder API 的文本添加逻辑
// 不涉及 XML 序列化，直接断言 Go 结构体
func TestSlideBuilder_AddText(t *testing.T) {
	// 实例化一个空白的 SlidePart 对象
	slide := parts.NewSlidePart(1)
	if slide == nil {
		t.Fatal("NewSlidePart 返回 nil")
	}

	// 调用 AddTextBox 方法添加文本
	testText := "测试标题"
	x, y, cx, cy := 914400, 457200, 9144000, 1143000 // EMU 单位
	sp := slide.AddTextBox(x, y, cx, cy, testText)

	// 断言返回的形状不为 nil（证明 ShapeTree 成功增加了一个形状）
	if sp == nil {
		t.Fatal("AddTextBox 返回 nil")
	}

	// 检查该形状内部的 TextBody -> Paragraph -> Run -> Text 的值
	textBody := sp.TextBody
	if textBody == nil {
		t.Fatal("形状的 TextBody 为 nil")
	}

	if len(textBody.Paragraphs) == 0 {
		t.Fatal("TextBody.Paragraphs 长度为 0")
	}

	para := textBody.Paragraphs[0]
	if len(para.TextRuns) == 0 {
		t.Fatal("Paragraph.TextRuns 长度为 0")
	}

	run := para.TextRuns[0]
	if run.Text != testText {
		t.Errorf("Text = %q, want %q", run.Text, testText)
	}

	t.Logf("成功验证文本内容: %q", run.Text)
}

// TestSlideBuilder_AddTextBox_Multiple 测试添加多个文本框
func TestSlideBuilder_AddTextBox_Multiple(t *testing.T) {
	slide := parts.NewSlidePart(1)

	// 添加多个文本框并收集返回的形状指针
	texts := []string{"标题文本", "正文段落1", "正文段落2"}
	var shapes []*parts.XSp
	for i, text := range texts {
		y := 457200 + i*914400 // 递增 Y 坐标
		sp := slide.AddTextBox(914400, y, 9144000, 457200, text)
		if sp == nil {
			t.Fatalf("第 %d 个 AddTextBox 返回 nil", i+1)
		}
		shapes = append(shapes, sp)
	}

	// 验证返回的形状数量
	if len(shapes) != len(texts) {
		t.Errorf("形状数量 = %d, want %d", len(shapes), len(texts))
	}

	// 验证每个形状的文本内容
	for i, sp := range shapes {
		if sp.TextBody == nil || len(sp.TextBody.Paragraphs) == 0 {
			t.Errorf("第 %d 个形状的 TextBody 结构不完整", i+1)
			continue
		}
		para := sp.TextBody.Paragraphs[0]
		if len(para.TextRuns) == 0 {
			t.Errorf("第 %d 个形状没有文本片段", i+1)
			continue
		}
		if para.TextRuns[0].Text != texts[i] {
			t.Errorf("第 %d 个形状文本 = %q, want %q", i+1, para.TextRuns[0].Text, texts[i])
		}
	}

	t.Logf("成功添加 %d 个文本框", len(shapes))
}

// TestSlideBuilder_AddTextBox_VerifyStructure 测试 AddTextBox 创建的完整结构
func TestSlideBuilder_AddTextBox_VerifyStructure(t *testing.T) {
	slide := parts.NewSlidePart(1)

	testText := "完整结构测试"
	x, y, cx, cy := 1000000, 2000000, 3000000, 4000000
	sp := slide.AddTextBox(x, y, cx, cy, testText)

	// 验证形状的非视觉属性
	if sp.NonVisual.CNvPr == nil {
		t.Fatal("NonVisual.CNvPr 为 nil")
	}
	if sp.NonVisual.CNvPr.ID == 0 {
		t.Error("形状 ID 为 0")
	}
	t.Logf("形状 ID: %d, Name: %s", sp.NonVisual.CNvPr.ID, sp.NonVisual.CNvPr.Name)

	// 验证形状属性中的位置和尺寸
	if sp.ShapeProperties == nil {
		t.Fatal("ShapeProperties 为 nil")
	}
	if sp.ShapeProperties.Transform2D == nil {
		t.Fatal("Transform2D 为 nil")
	}
	if sp.ShapeProperties.Transform2D.Offset == nil {
		t.Fatal("Offset 为 nil")
	}
	if sp.ShapeProperties.Transform2D.Extent == nil {
		t.Fatal("Extent 为 nil")
	}

	// 验证坐标
	offset := sp.ShapeProperties.Transform2D.Offset
	if offset.X != x || offset.Y != y {
		t.Errorf("Offset = (%d, %d), want (%d, %d)", offset.X, offset.Y, x, y)
	}

	extent := sp.ShapeProperties.Transform2D.Extent
	if extent.Cx != cx || extent.Cy != cy {
		t.Errorf("Extent = (%d, %d), want (%d, %d)", extent.Cx, extent.Cy, cx, cy)
	}

	t.Logf("位置验证通过: offset=(%d, %d), extent=(%d, %d)", offset.X, offset.Y, extent.Cx, extent.Cy)
}

// TestSlide_MarshalComponents 测试底层组件的 XML 序列化
// 验证 omitempty 标签是否生效，确保不会生成多余的空标签
func TestSlide_MarshalComponents(t *testing.T) {
	// 测试 XTextParagraph 序列化
	// 注意：XTextParagraph 没有 XMLName 字段，所以根标签是结构体名
	t.Run("XTextParagraph_OmitEmpty", func(t *testing.T) {
		// 手动构造一个最简单的 XTextParagraph，只包含一个写着 "Hello" 的 XTextRun
		// 不给任何非必填的属性赋值
		para := parts.XTextParagraph{
			TextRuns: []parts.XTextRun{
				{Text: "Hello"},
			},
		}

		// 使用 xml.Marshal 将其序列化
		data, err := xml.Marshal(&para)
		if err != nil {
			t.Fatalf("xml.Marshal 失败: %v", err)
		}

		xmlStr := string(data)
		t.Logf("生成的 XML: %s", xmlStr)

		// 核心断言：绝对不包含空的属性标签
		// Level, Indent, Alignment 都是 omitempty，值为零时不应该出现
		forbiddenPatterns := []string{
			`lvl="0"`,
			`indent="0"`,
			`algn=""`,
		}

		for _, pattern := range forbiddenPatterns {
			if strings.Contains(xmlStr, pattern) {
				t.Errorf("XML 包含零值属性: %s（应该被 omitempty 省略）", pattern)
			}
		}

		// 验证文本内容正确序列化
		if !strings.Contains(xmlStr, "Hello") {
			t.Error("应包含文本 'Hello'")
		}
	})

	// 测试 XTextRun 序列化 - 验证 TextProperties 的 omitempty
	t.Run("XTextRun_OmitEmpty", func(t *testing.T) {
		// 不设置 TextProperties
		run := parts.XTextRun{
			Text: "World",
		}

		data, err := xml.Marshal(&run)
		if err != nil {
			t.Fatalf("xml.Marshal 失败: %v", err)
		}

		xmlStr := string(data)
		t.Logf("生成的 XML: %s", xmlStr)

		// 不应该包含 rPr 标签（因为 TextProperties 为 nil 且 omitempty）
		if strings.Contains(xmlStr, "<a:rPr>") || strings.Contains(xmlStr, "<a:rPr/>") {
			t.Error("XTextRun 不应包含空的 <a:rPr> 标签（TextProperties 为 nil）")
		}

		// 验证文本内容
		if !strings.Contains(xmlStr, "World") {
			t.Error("应包含文本 'World'")
		}
	})

	// 测试 XTextRun 带属性 - 验证属性级的 omitempty
	t.Run("XTextRun_WithProperties", func(t *testing.T) {
		run := parts.XTextRun{
			Text: "Styled",
			TextProperties: &parts.XTextProperties{
				FontSize: 2400, // 24pt
				Bold:     true,
			},
		}

		data, err := xml.Marshal(&run)
		if err != nil {
			t.Fatalf("xml.Marshal 失败: %v", err)
		}

		xmlStr := string(data)
		t.Logf("生成的 XML: %s", xmlStr)

		// 应该包含设置的属性
		if !strings.Contains(xmlStr, `sz="2400"`) {
			t.Error("应包含 sz=\"2400\" 属性")
		}
		// 注意：Go xml 包将 bool 序列化为 "true"/"false"
		if !strings.Contains(xmlStr, `b="true"`) {
			t.Error("应包含 b=\"true\" 属性")
		}

		// 不应包含未设置的属性（Italic, Underline, FontFace, Color）
		unexpectedAttrs := []string{"i=", "u=", "typeface=", "solidFill"}
		for _, attr := range unexpectedAttrs {
			if strings.Contains(xmlStr, attr) {
				t.Errorf("不应包含未设置的属性: %s", attr)
			}
		}
	})

	// 测试带属性的 XTextParagraph（验证 omitempty 对属性也生效）
	t.Run("XTextParagraph_WithAttributes", func(t *testing.T) {
		para := parts.XTextParagraph{
			Level:     1,       // 设置属性
			Alignment: "ctr",   // 居中对齐
			TextRuns: []parts.XTextRun{
				{Text: "Centered"},
			},
		}

		data, err := xml.Marshal(&para)
		if err != nil {
			t.Fatalf("xml.Marshal 失败: %v", err)
		}

		xmlStr := string(data)
		t.Logf("生成的 XML: %s", xmlStr)

		// 应该包含 lvl 和 algn 属性
		if !strings.Contains(xmlStr, `lvl="1"`) {
			t.Error("应包含 lvl=\"1\" 属性")
		}
		if !strings.Contains(xmlStr, `algn="ctr"`) {
			t.Error("应包含 algn=\"ctr\" 属性")
		}

		// 不应包含 indent 属性（因为未设置，零值被 omitempty）
		if strings.Contains(xmlStr, "indent") {
			t.Error("不应包含 indent 属性（值为零且 omitempty）")
		}
	})

	// 测试 XShapeProperties 的 omitempty
	t.Run("XShapeProperties_OmitEmpty", func(t *testing.T) {
		sp := parts.XSp{
			NonVisual: parts.XNonVisualDrawingShape{
				CNvPr: &parts.XNvCxnSpPr{
					ID:   1,
					Name: "Test",
				},
				CNvSpPr: &parts.XNvSpPr{},
			},
			ShapeProperties: &parts.XShapeProperties{
				Transform2D: &parts.XTransform2D{
					Offset: &parts.XOv2DrOffset{X: 0, Y: 0},
					Extent: &parts.XOv2DrExtent{Cx: 100, Cy: 100},
				},
			},
			TextBody: &parts.XTextBody{
				Paragraphs: []parts.XTextParagraph{
					{TextRuns: []parts.XTextRun{{Text: "Test"}}},
				},
			},
		}

		data, err := xml.Marshal(&sp)
		if err != nil {
			t.Fatalf("xml.Marshal 失败: %v", err)
		}

		xmlStr := string(data)
		t.Logf("生成的 XML (前200字符): %.200s...", xmlStr)

		// 验证 Rotation, FlipH, FlipV 等 omitempty 属性未出现
		rotationPatterns := []string{`rot="0"`, `flipH="false"`, `flipV="false"`}
		for _, pattern := range rotationPatterns {
			if strings.Contains(xmlStr, pattern) {
				t.Errorf("不应包含零值属性: %s", pattern)
			}
		}
	})
}

// TestSlide_MarshalFullPage 整页幻灯片的序列化冒烟测试
// 验证命名空间声明是否正确，这是 PPT 能否打开的关键
func TestSlide_MarshalFullPage(t *testing.T) {
	// 直接构造 XSlide 结构体，验证命名空间序列化
	xslide := parts.XSlide{
		XmlnsA: "http://schemas.openxmlformats.org/drawingml/2006/main",
		XmlnsR: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		XmlnsP: "http://schemas.openxmlformats.org/presentationml/2006/main",
		ClrMapOvr: &parts.XClrMapOvr{Accent1: "accent1"},
		CSld: &parts.XCSld{
			SpTree: parts.NewXSpTree(),
		},
	}

	// 使用 ToXMLFast 进行序列化（生成带命名空间前缀的 OOXML 格式）
	data, err := xslide.ToXMLFast()
	if err != nil {
		t.Fatalf("ToXMLFast 失败: %v", err)
	}

	xmlStr := string(data)
	t.Logf("生成的 XML:\n%s", xmlStr)

	// 关键断言：检查必需的命名空间声明
	requiredNamespaces := []string{
		`xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`,
		`xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"`,
		`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`,
	}

	for _, ns := range requiredNamespaces {
		if !strings.Contains(xmlStr, ns) {
			t.Errorf("缺少必需的命名空间声明: %s", ns)
		}
	}

	// 检查根节点 <p:sld>
	if !strings.Contains(xmlStr, "<p:sld") {
		t.Error("缺少根节点 <p:sld>")
	}
	if !strings.Contains(xmlStr, "</p:sld>") {
		t.Error("缺少根节点结束标签 </p:sld>")
	}

	// 检查形状树 <p:spTree>
	if !strings.Contains(xmlStr, "<p:spTree>") {
		t.Error("缺少形状树节点 <p:spTree>")
	}
	if !strings.Contains(xmlStr, "</p:spTree>") {
		t.Error("缺少形状树结束标签 </p:spTree>")
	}

	t.Logf("命名空间和根节点验证通过")
}

// TestSlide_MarshalFullPage_WithContent 测试带内容的完整幻灯片序列化
func TestSlide_MarshalFullPage_WithContent(t *testing.T) {
	// 构造包含文本的 XTextBody
	textBody := parts.XTextBody{
		Paragraphs: []parts.XTextParagraph{
			{
				TextRuns: []parts.XTextRun{
					{Text: "演示标题"},
				},
			},
		},
	}

	// 使用 ToXMLFast 进行序列化（生成带命名空间前缀的 OOXML 格式）
	data, err := textBody.ToXMLFast()
	if err != nil {
		t.Fatalf("ToXMLFast 失败: %v", err)
	}

	xmlStr := string(data)
	t.Logf("生成的 XML:\n%s", xmlStr)

	// 检查文本内容
	if !strings.Contains(xmlStr, "演示标题") {
		t.Error("缺少文本内容: 演示标题")
	}

	// 检查 <a:t> 标签
	if !strings.Contains(xmlStr, "<a:t>") {
		t.Error("缺少文本标签 <a:t>")
	}

	t.Logf("文本内容序列化验证通过")
}
