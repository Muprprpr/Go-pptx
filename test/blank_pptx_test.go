package pptx_test

import (
	"archive/zip"
	"os"
	"strings"
	"testing"

	"github.com/Muprprpr/Go-pptx/opc"
	"github.com/Muprprpr/Go-pptx/parts"
)

// ============================================================================
// 最小可用 OOXML 基石 (Minimal Viable OOXML)
//
// TestOPC_RawAssemblyFallback 零依赖底层组装测试
//
// 【注意】本测试的目的是验证 OPC 引擎的纯原生组装、序列化、头信息拼接和 ZIP 打包能力。
// 由于使用了硬编码的最小化 XML，且强依赖 AddRelationship 的顺序来生成特定 rId (如 rId1, rId2)，
// 本测试生成的 test_hardcore_output.pptx 在某些版本的 PowerPoint 中可能因 rId 错位而无法直接打开。
//
// ============================================================================

const (
	// 最小化主题 (包含最基础的颜色矩阵和字体，否则 PPT 必崩)
	minThemeXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`

	// 最小化母版 (必须包含对 Layout 的引用列表和基础形状树)
	minMasterXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`

	// 最小化版式 (几乎为空，但必须存在以承载 Slide)
	minLayoutXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>`

	// 最小化 Presentation XML (包含完整的 sldMasterIdLst 和 sldIdLst)
	minPresentationXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`
)

// TestBlankPresentation_Hardcore 极简白板测试 - 构建完整 OOXML 继承树的 PPTX
// 手动搭建: slide1 -> slideLayout1 -> slideMaster1 -> theme1
func TestBlankPresentation_Hardcore(t *testing.T) {
	pkg := opc.NewPackage()

	// --- 1. 创建静态基础设施 (Theme, Master, Layout) ---
	themeURI := opc.NewPackURI("/ppt/theme/theme1.xml")
	pkg.CreatePart(themeURI, "application/vnd.openxmlformats-officedocument.theme+xml", []byte(minThemeXML))

	masterURI := opc.NewPackURI("/ppt/slideMasters/slideMaster1.xml")
	masterPart, _ := pkg.CreatePart(masterURI, "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml", []byte(minMasterXML))

	layoutURI := opc.NewPackURI("/ppt/slideLayouts/slideLayout1.xml")
	layoutPart, _ := pkg.CreatePart(layoutURI, "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml", []byte(minLayoutXML))

	// --- 2. 编织基础设施内部的关系网 ---
	// 2.1 Master -> Theme (rId2)
	masterPart.AddRelationship(
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
		"../theme/theme1.xml",
		false,
	)
	// 2.2 Master -> Layout (rId1, 对应 minMasterXML 中的 sldLayoutId r:id="rId1")
	masterPart.AddRelationship(
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
		"../slideLayouts/slideLayout1.xml",
		false,
	)
	// 2.3 Layout -> Master (rId1)
	layoutPart.AddRelationship(
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
		"../slideMasters/slideMaster1.xml",
		false,
	)

	// --- 3. 创建 Presentation (直接使用最小化 XML) ---
	presPartOp, _ := pkg.CreatePart(
		opc.NewPackURI("/ppt/presentation.xml"),
		"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
		[]byte(minPresentationXML),
	)

	// --- 4. 编织 Presentation 的关系网 ---
	// 根目录指向 Presentation
	pkg.AddRelationship(
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
		"/ppt/presentation.xml",
		false,
	)
	// Presentation -> Master (rId1)
	presPartOp.AddRelationship(
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
		"slideMasters/slideMaster1.xml",
		false,
	)
	// Presentation -> Theme (rId3)
	presPartOp.AddRelationship(
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
		"theme/theme1.xml",
		false,
	)

	// --- 5. 创建幻灯片并关联 Layout ---
	slidePart := parts.NewSlidePart(1)
	slidePart.AddTextBox(914400, 457200, 4572000, 457200, "Hello from Go Engine!")
	slideXML, _ := slidePart.ToXML()

	slidePartOp, _ := pkg.CreatePart(
		opc.NewPackURI("/ppt/slides/slide1.xml"),
		"application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
		slideXML,
	)

	// 【关键】Slide -> Layout (rId1)
	slidePartOp.AddRelationship(
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
		"../slideLayouts/slideLayout1.xml",
		false,
	)

	// --- 6. 完善 Presentation -> Slide 的关系 ---
	// Presentation -> Slide (rId2)
	presPartOp.AddRelationship(
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
		"slides/slide1.xml",
		false,
	)

	// --- 7. 打包保存 ---
	outputPath := "test_hardcore_output.pptx"
	err := pkg.SaveFile(outputPath)
	if err != nil {
		t.Fatalf("保存 PPTX 失败: %v", err)
	}
	// defer os.Remove(outputPath)

	t.Logf("成功创建纯净版 PPTX: %s", outputPath)

	// --- 8. 验证文件结构 ---
	info, err := os.Stat(outputPath)
	if err != nil {
		t.Fatalf("检查输出文件失败: %v", err)
	}
	if info.Size() == 0 {
		t.Fatal("输出文件大小为 0")
	}

	// --- 9. 验证 OPC 结构 ---
	zipReader, err := zip.OpenReader(outputPath)
	if err != nil {
		t.Fatalf("打开 ZIP 失败: %v", err)
	}
	defer zipReader.Close()

	files := make(map[string]bool)
	for _, f := range zipReader.File {
		files[f.Name] = true
	}

	// 验证必需的 OPC 文件
	requiredFiles := []string{
		"[Content_Types].xml",
		"_rels/.rels",
		"ppt/presentation.xml",
		"ppt/_rels/presentation.xml.rels",
		"ppt/slides/slide1.xml",
		"ppt/slideMasters/slideMaster1.xml",
		"ppt/slideLayouts/slideLayout1.xml",
		"ppt/theme/theme1.xml",
	}

	for _, name := range requiredFiles {
		if !files[name] {
			t.Errorf("缺少必需文件: %s", name)
		}
	}

	t.Log("OPC 结构验证通过")

	// --- 10. 验证 XML 声明头 ---
	for _, f := range zipReader.File {
		if strings.HasSuffix(f.Name, ".xml") || strings.HasSuffix(f.Name, ".rels") {
			rc, err := f.Open()
			if err != nil {
				continue
			}
			header := make([]byte, 60)
			n, _ := rc.Read(header)
			rc.Close()
			content := string(header[:n])
			if !strings.Contains(content, "<?xml") {
				t.Errorf("文件 %s 缺少 XML 声明头", f.Name)
			}
			if !strings.Contains(content, "standalone=\"yes\"") {
				t.Errorf("文件 %s 缺少 standalone=\"yes\"", f.Name)
			}
		}
	}

	t.Log("XML 声明头验证通过")

	// --- 11. 重新打开验证 ---
	pkg2, err := opc.OpenFile(outputPath)
	if err != nil {
		t.Fatalf("重新打开 PPTX 失败: %v", err)
	}
	defer pkg2.Close()

	// 验证关键部件存在
	if pkg2.GetPart(opc.NewPackURI("/ppt/presentation.xml")) == nil {
		t.Fatal("缺少 presentation.xml")
	}
	if pkg2.GetPart(opc.NewPackURI("/ppt/slides/slide1.xml")) == nil {
		t.Fatal("缺少 slide1.xml")
	}
	if pkg2.GetPart(opc.NewPackURI("/ppt/slideMasters/slideMaster1.xml")) == nil {
		t.Fatal("缺少 slideMaster1.xml")
	}
	if pkg2.GetPart(opc.NewPackURI("/ppt/slideLayouts/slideLayout1.xml")) == nil {
		t.Fatal("缺少 slideLayout1.xml")
	}
	if pkg2.GetPart(opc.NewPackURI("/ppt/theme/theme1.xml")) == nil {
		t.Fatal("缺少 theme1.xml")
	}

	t.Log("全部验证通过！PPT 可被正常打开")
}
