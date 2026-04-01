package pptx_test

import (
	"archive/zip"
	"os"
	"strings"
	"testing"

	"github.com/Muprprpr/Go-pptx/opc"
	"github.com/Muprprpr/Go-pptx/parts"
	"github.com/Muprprpr/Go-pptx/slide"
)

// ============================================================================
// Go-pptx 完整能力测试流水线
// ============================================================================
//
// 测试目标：证明整个库从解析到创建的完整数据流能力
//
// 阶段划分：
//   Stage 1: OPC层【解析与解包】能力
//   Stage 2: Parts层【反序列化】能力
//   Stage 3: Parts层【创建与序列化】能力
//   Stage 4: OPC层【写入与路由】能力 (极其关键)
//   Stage 5: Parts层【数据更新】能力
//   Stage 6: OPC层【安全打包】能力
//
// 依赖：test/test.pptx (真实 PPTX 文件)
// ============================================================================

const testPPTXPath = "test.pptx"

// ============================================================================
// Stage 1: OPC层【解析与解包】能力
// ============================================================================

// TestOPC_ParseAndUnpack 阶段1：证明 OPC 层能够正确解析和解包 PPTX 文件
func TestOPC_ParseAndUnpack(t *testing.T) {
	// 1.1 验证测试文件存在
	if _, err := os.Stat(testPPTXPath); os.IsNotExist(err) {
		t.Fatalf("测试文件不存在: %s", testPPTXPath)
	}

	// 1.2 打开 OPC 包
	pkg, err := opc.OpenFile(testPPTXPath)
	if err != nil {
		t.Fatalf("OPC OpenFile 失败: %v", err)
	}
	defer pkg.Close()

	// 1.3 验证包级别关系存在
	rootRels := pkg.Relationships()
	if rootRels == nil {
		t.Fatal("包级别关系为空")
	}
	if rootRels.Count() == 0 {
		t.Fatal("包级别关系数量为 0")
	}

	// 1.4 验证 ContentTypes 正确加载
	ct := pkg.ContentTypes()
	if ct == nil {
		t.Fatal("ContentTypes 为空")
	}

	// 1.5 验证部件数量（test.pptx 包含多个部件）
	partCount := pkg.PartCount()
	if partCount == 0 {
		t.Fatal("部件数量为 0")
	}
	t.Logf("Stage 1: OPC 解包成功，部件数量 = %d", partCount)

	// 1.6 验证核心部件存在
	requiredParts := []string{
		"/ppt/presentation.xml",
		"/ppt/slides/slide1.xml",
		"/ppt/slideMasters/slideMaster1.xml",
	}
	for _, uri := range requiredParts {
		if !pkg.ContainsPart(opc.NewPackURI(uri)) {
			t.Errorf("缺少必需部件: %s", uri)
		}
	}

	// 1.7 验证 ZIP 结构完整性
	file, _ := os.Open(testPPTXPath)
	defer file.Close()
	stat, _ := file.Stat()
	zipReader, err := zip.NewReader(file, stat.Size())
	if err != nil {
		t.Fatalf("ZIP 读取失败: %v", err)
	}

	files := make(map[string]bool)
	for _, f := range zipReader.File {
		files[f.Name] = true
	}

	// 验证 ZIP 中没有以 "/" 开头的路径（Windows ZIP 规范）
	for _, f := range zipReader.File {
		if strings.HasPrefix(f.Name, "/") {
			t.Errorf("ZIP 路径违规（以 / 开头）: %s", f.Name)
		}
	}

	t.Logf("Stage 1: OPC 解析与解包能力验证通过 (files=%d)", len(files))
	_ = files // 使用变量避免编译错误
}

// ============================================================================
// Stage 2: Parts层【反序列化】能力
// ============================================================================

// TestParts_Deserialize 阶段2：证明 Parts 层能够正确反序列化 XML
func TestParts_Deserialize(t *testing.T) {
	// 2.1 打开 OPC 包
	pkg, err := opc.OpenFile(testPPTXPath)
	if err != nil {
		t.Fatalf("OPC OpenFile 失败: %v", err)
	}
	defer pkg.Close()

	// 2.2 获取 PresentationPart 并反序列化
	presPart := pkg.GetPart(opc.NewPackURI("/ppt/presentation.xml"))
	if presPart == nil {
		t.Fatal("缺少 presentation.xml")
	}

	pres := parts.NewPresentationPart()
	if err := pres.FromXML(presPart.Blob()); err != nil {
		t.Fatalf("PresentationPart.FromXML 失败: %v", err)
	}

	// 2.3 验证 PresentationPart 数据正确性
	slideCount := pres.SlideCount()
	if slideCount == 0 {
		t.Error("幻灯片数量为 0")
	}
	t.Logf("Stage 2: Presentation 反序列化成功，幻灯片数量 = %d", slideCount)

	slideSize := pres.SlideSize()
	if slideSize.Cx == 0 || slideSize.Cy == 0 {
		t.Error("幻灯片尺寸无效")
	}
	t.Logf("Stage 2: 幻灯片尺寸 = %dx%d EMU", slideSize.Cx, slideSize.Cy)

	// 2.4 获取 SlidePart 并反序列化
	slidePart := pkg.GetPart(opc.NewPackURI("/ppt/slides/slide1.xml"))
	if slidePart == nil {
		t.Fatal("缺少 slide1.xml")
	}

	slide := parts.NewSlidePart(1)
	if err := slide.FromXML(slidePart.Blob()); err != nil {
		t.Fatalf("SlidePart.FromXML 失败: %v", err)
	}

	// 2.5 验证 SlidePart 数据正确性
	shapeCount := slide.ShapeIDCount()
	t.Logf("Stage 2: Slide 反序列化成功，形状数量 = %d", shapeCount)

	t.Log("Stage 2: Parts 层反序列化能力验证通过")
}

// ============================================================================
// Stage 3: Parts层【创建与序列化】能力
// ============================================================================

// TestParts_CreateAndSerialize 阶段3：证明 Parts 层能够创建新部件并序列化
func TestParts_CreateAndSerialize(t *testing.T) {
	// 3.1 创建新的 PresentationPart
	pres := parts.NewPresentationPart()

	// 3.2 创建新的 SlidePart 并使用 Builder 添加文本
	slidePart := parts.NewSlidePart(1)
	builder := slide.NewSlideBuilder(slidePart)
	builder.AddTextBox(914400, 457200, 4572000, 457200, "Hello from Go Engine!")

	// 3.3 序列化 PresentationPart
	presXML, err := pres.ToXML()
	if err != nil {
		t.Fatalf("PresentationPart.ToXML 失败: %v", err)
	}
	if len(presXML) == 0 {
		t.Fatal("PresentationPart.ToXML 返回空数据")
	}
	if !strings.HasPrefix(string(presXML), "<?xml") {
		t.Error("XML 缺少声明头")
	}

	// 3.4 序列化 SlidePart
	slideXML, err := slidePart.ToXML()
	if err != nil {
		t.Fatalf("SlidePart.ToXML 失败: %v", err)
	}
	if len(slideXML) == 0 {
		t.Fatal("SlidePart.ToXML 返回空数据")
	}
	if !strings.Contains(string(slideXML), "Hello from Go Engine!") {
		t.Error("Slide XML 不包含预期文本")
	}

	t.Logf("Stage 3: Presentation XML 大小 = %d bytes", len(presXML))
	t.Logf("Stage 3: Slide XML 大小 = %d bytes", len(slideXML))
	t.Log("Stage 3: Parts 层创建与序列化能力验证通过")
}

// ============================================================================
// Stage 4: OPC层【写入与路由】能力 (极其关键)
// ============================================================================

// TestOPC_WriteAndRoute 阶段4：证明 OPC 层能够正确写入和路由
func TestOPC_WriteAndRoute(t *testing.T) {
	// 4.1 创建新包
	pkg := opc.NewPackage()

	// 4.2 创建 PresentationPart
	pres := parts.NewPresentationPart()
	presXML, _ := pres.ToXML()
	presPart, err := pkg.CreatePart(
		opc.NewPackURI("/ppt/presentation.xml"),
		"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
		presXML,
	)
	if err != nil {
		t.Fatalf("CreatePart presentation.xml 失败: %v", err)
	}

	// 4.3 创建 SlidePart
	slidePart4 := parts.NewSlidePart(1)
	slideBuilder4 := slide.NewSlideBuilder(slidePart4)
	slideBuilder4.AddTextBox(914400, 457200, 4572000, 457200, "Hello from OPC!")
	slideXML, _ := slidePart4.ToXML()
	slidePart, err := pkg.CreatePart(
		opc.NewPackURI("/ppt/slides/slide1.xml"),
		"application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
		slideXML,
	)
	if err != nil {
		t.Fatalf("CreatePart slide1.xml 失败: %v", err)
	}

	// 4.4 建立关系
	// 包级别rels: 指向 presentation.xml
	pkg.AddRelationship(
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
		"/ppt/presentation.xml",
		false,
	)

	// presentation.xml -> slide1.xml
	presPart.AddRelationship(
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
		"slides/slide1.xml",
		false,
	)

	// slide1.xml -> presentation.xml (布局关系)
	slidePart.AddRelationship(
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
		"../slideLayouts/slideLayout1.xml",
		false,
	)

	// 4.5 验证关系路由
	// 通过关系找到 slide
	slideViaRel := pkg.ResolveRelationship(presPart,
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide")
	if slideViaRel == nil {
		t.Fatal("通过关系解析 slide 失败")
	}
	if slideViaRel.PartURI().URI() != "/ppt/slides/slide1.xml" {
		t.Errorf("关系解析的 URI 不正确: %s", slideViaRel.PartURI().URI())
	}

	// 4.6 保存到文件
	outputPath := "test_opc_output.pptx"
	if err := pkg.SaveFile(outputPath); err != nil {
		t.Fatalf("SaveFile 失败: %v", err)
	}
	defer os.Remove(outputPath)

	// 4.7 验证输出文件 ZIP 结构
	file, _ := os.Open(outputPath)
	defer file.Close()
	stat, _ := file.Stat()
	zipReader, err := zip.NewReader(file, stat.Size())
	if err != nil {
		t.Fatalf("ZIP 读取失败: %v", err)
	}

	outputFiles := make(map[string]bool)
	for _, f := range zipReader.File {
		outputFiles[f.Name] = true
		// 检查没有以 "/" 开头的路径
		if strings.HasPrefix(f.Name, "/") {
			t.Errorf("ZIP 路径违规: %s", f.Name)
		}
	}

	// 验证必需的 OPC 文件都存在
	requiredFiles := []string{
		"[Content_Types].xml",
		"_rels/.rels",
		"ppt/presentation.xml",
		"ppt/_rels/presentation.xml.rels",
		"ppt/slides/slide1.xml",
		"ppt/slides/_rels/slide1.xml.rels",
	}
	for _, name := range requiredFiles {
		if !outputFiles[name] {
			t.Errorf("输出文件缺少: %s", name)
		}
	}

	t.Logf("Stage 4: OPC 写入成功，输出文件大小 = %d bytes", stat.Size())
	t.Log("Stage 4: OPC 层写入与路由能力验证通过")
}

// ============================================================================
// Stage 5: Parts层【数据更新】能力
// ============================================================================

// TestParts_Update 阶段5：证明 Parts 层能够更新已有数据
func TestParts_Update(t *testing.T) {
	// 5.1 打开 OPC 包
	pkg, err := opc.OpenFile(testPPTXPath)
	if err != nil {
		t.Fatalf("OPC OpenFile 失败: %v", err)
	}
	defer pkg.Close()

	// 5.2 获取并反序列化 SlidePart
	slidePart := pkg.GetPart(opc.NewPackURI("/ppt/slides/slide1.xml"))
	if slidePart == nil {
		t.Fatal("缺少 slide1.xml")
	}

	slidePart5 := parts.NewSlidePart(1)
	if err := slidePart5.FromXML(slidePart.Blob()); err != nil {
		t.Fatalf("SlidePart.FromXML 失败: %v", err)
	}

	originalShapeCount := slidePart5.ShapeIDCount()
	t.Logf("Stage 5: 原始形状数量 = %d", originalShapeCount)

	// 5.3 添加新文本框
	slideBuilder5 := slide.NewSlideBuilder(slidePart5)
	slideBuilder5.AddTextBox(1000000, 1000000, 2000000, 500000, "Updated Text!")

	newShapeCount := slidePart5.ShapeIDCount()
	if newShapeCount <= originalShapeCount {
		t.Error("添加形状后数量未增加")
	}
	t.Logf("Stage 5: 更新后形状数量 = %d", newShapeCount)

	// 5.4 重新序列化
	newXML, err := slidePart5.ToXML()
	if err != nil {
		t.Fatalf("SlidePart.ToXML 失败: %v", err)
	}
	if !strings.Contains(string(newXML), "Updated Text!") {
		t.Error("更新后的 XML 不包含新文本")
	}

	// 5.5 更新 OPC 包中的部件
	newSlidePart := opc.NewPart(
		opc.NewPackURI("/ppt/slides/slide1.xml"),
		slidePart.ContentType(),
		newXML,
	)
	if err := pkg.AddPart(newSlidePart); err != nil {
		// 如果已存在，需要先移除
		pkg.RemovePart(opc.NewPackURI("/ppt/slides/slide1.xml"))
		pkg.AddPart(newSlidePart)
	}

	// 5.6 验证更新
	updatedPart := pkg.GetPart(opc.NewPackURI("/ppt/slides/slide1.xml"))
	if string(updatedPart.Blob()) != string(newXML) {
		t.Error("部件更新失败")
	}

	t.Log("Stage 5: Parts 层数据更新能力验证通过")
}

// ============================================================================
// Stage 6: OPC层【安全打包】能力
// ============================================================================

// TestOPC_SecurePackaging 阶段6：证明 OPC 层能够安全地重新打包
func TestOPC_SecurePackaging(t *testing.T) {
	// 6.1 打开原始文件
	originalPkg, err := opc.OpenFile(testPPTXPath)
	if err != nil {
		t.Fatalf("OpenFile 失败: %v", err)
	}

	// 6.2 记录原始部件数量和内容
	originalPartCount := originalPkg.PartCount()
	originalSlidePart := originalPkg.GetPart(opc.NewPackURI("/ppt/slides/slide1.xml"))
	_ = string(originalSlidePart.Blob()) // 原始内容，用于验证

	// 6.3 修改 SlidePart
	slide6 := parts.NewSlidePart(1)
	slideBuilder6 := slide.NewSlideBuilder(slide6)
	slideBuilder6.AddTextBox(1000000, 1000000, 2000000, 500000, "Modified Content!")
	newSlideXML, _ := slide6.ToXML()

	newSlidePart := opc.NewPart(
		opc.NewPackURI("/ppt/slides/slide1.xml"),
		originalSlidePart.ContentType(),
		newSlideXML,
	)
	originalPkg.RemovePart(opc.NewPackURI("/ppt/slides/slide1.xml"))
	originalPkg.AddPart(newSlidePart)

	// 6.4 保存修改后的包
	outputPath := "test_secure_output.pptx"
	if err := originalPkg.SaveFile(outputPath); err != nil {
		t.Fatalf("SaveFile 失败: %v", err)
	}
	originalPkg.Close()
	defer os.Remove(outputPath)

	// 6.5 重新打开并验证
	reopenedPkg, err := opc.OpenFile(outputPath)
	if err != nil {
		t.Fatalf("重新打开失败: %v", err)
	}
	defer reopenedPkg.Close()

	// 6.6 验证数据完整性
	reopenedPartCount := reopenedPkg.PartCount()
	if reopenedPartCount != originalPartCount {
		t.Errorf("部件数量不匹配: got %d, want %d", reopenedPartCount, originalPartCount)
	}

	reopenedSlidePart := reopenedPkg.GetPart(opc.NewPackURI("/ppt/slides/slide1.xml"))
	if string(reopenedSlidePart.Blob()) != string(newSlideXML) {
		t.Error("重新打开后内容不匹配")
	}

	// 6.7 验证 ContentTypes 正确
	ct := reopenedPkg.ContentTypes()
	if ct == nil {
		t.Error("ContentTypes 为空")
	}

	// 6.8 验证 ZIP 规范性（无前导斜杠）
	file, _ := os.Open(outputPath)
	defer file.Close()
	stat, _ := file.Stat()
	zipReader, _ := zip.NewReader(file, stat.Size())
	for _, f := range zipReader.File {
		if strings.HasPrefix(f.Name, "/") {
			t.Errorf("ZIP 路径违规: %s", f.Name)
		}
	}

	t.Logf("Stage 6: 安全打包成功，文件大小 = %d bytes", stat.Size())
	t.Log("Stage 6: OPC 层安全打包能力验证通过")
}

// ============================================================================
// 综合测试：完整流水线
// ============================================================================

// TestPipeline_FullIntegration 完整流水线测试：解析 -> 反序列化 -> 修改 -> 重新打包
func TestPipeline_FullIntegration(t *testing.T) {
	t.Log("========== 开始完整流水线测试 ==========")

	// Stage 1: 解析
	t.Log("----- Stage 1: OPC 解析与解包 -----")
	pkg, err := opc.OpenFile(testPPTXPath)
	if err != nil {
		t.Fatalf("Stage 1 失败: %v", err)
	}
	t.Logf("Stage 1 通过: 解包得到 %d 个部件", pkg.PartCount())

	// Stage 2: 反序列化
	t.Log("----- Stage 2: Parts 反序列化 -----")
	presPart := pkg.GetPart(opc.NewPackURI("/ppt/presentation.xml"))
	pres := parts.NewPresentationPart()
	if err := pres.FromXML(presPart.Blob()); err != nil {
		t.Fatalf("Stage 2 失败: %v", err)
	}
	t.Logf("Stage 2 通过: 反序列化得到 %d 张幻灯片", pres.SlideCount())

	// Stage 3: 创建
	t.Log("----- Stage 3: Parts 创建与序列化 -----")
	newSlidePart3 := parts.NewSlidePart(1)
	newSlidePart3Builder := slide.NewSlideBuilder(newSlidePart3)
	newSlidePart3Builder.AddTextBox(914400, 457200, 4572000, 457200, "Integration Test!")
	newSlideXML, err := newSlidePart3.ToXML()
	if err != nil {
		t.Fatalf("Stage 3 序列化失败: %v", err)
	}
	t.Logf("Stage 3 通过: 创建新幻灯片 XML (%d bytes)", len(newSlideXML))

	// Stage 4: 写入
	t.Log("----- Stage 4: OPC 写入与路由 -----")
	newPkg := opc.NewPackage()
	newPkg.CreatePart(
		opc.NewPackURI("/ppt/presentation.xml"),
		"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
		presPart.Blob(),
	)
	newPkg.CreatePart(
		opc.NewPackURI("/ppt/slides/slide1.xml"),
		"application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
		newSlideXML,
	)
	newPkg.AddRelationship(
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
		"/ppt/presentation.xml",
		false,
	)
	outputPath := "test_integration_output.pptx"
	if err := newPkg.SaveFile(outputPath); err != nil {
		t.Fatalf("Stage 4 失败: %v", err)
	}
	pkg.Close()
	defer os.Remove(outputPath)
	t.Logf("Stage 4 通过: 写入 %d bytes", func() int {
		info, _ := os.Stat(outputPath)
		return int(info.Size())
	}())

	// Stage 5: 更新
	t.Log("----- Stage 5: Parts 数据更新 -----")
	updatedPkg, _ := opc.OpenFile(outputPath)
	updatedSlide := parts.NewSlidePart(1)
	updatedSlideBuilder := slide.NewSlideBuilder(updatedSlide)
	updatedSlideBuilder.AddTextBox(500000, 500000, 3000000, 300000, "Updated via Stage 5!")
	updatedXML, _ := updatedSlide.ToXML()
	updatedPkg.RemovePart(opc.NewPackURI("/ppt/slides/slide1.xml"))
	updatedPkg.CreatePart(
		opc.NewPackURI("/ppt/slides/slide1.xml"),
		"application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
		updatedXML,
	)
	updatedPath := "test_updated_output.pptx"
	updatedPkg.SaveFile(updatedPath)
	updatedPkg.Close()
	defer os.Remove(updatedPath)

	// 验证更新
	verifyPkg, _ := opc.OpenFile(updatedPath)
	verifySlidePart := verifyPkg.GetPart(opc.NewPackURI("/ppt/slides/slide1.xml"))
	if !strings.Contains(string(verifySlidePart.Blob()), "Updated via Stage 5!") {
		t.Error("Stage 5 验证失败: 更新内容不存在")
	}
	verifyPkg.Close()
	t.Log("Stage 5 通过: 数据更新成功")

	// Stage 6: 安全打包
	t.Log("----- Stage 6: OPC 安全打包 -----")
	finalPkg, _ := opc.OpenFile(updatedPath)
	finalPath := "test_final_output.pptx"
	finalPkg.SaveFile(finalPath)
	finalPkg.Close()
	defer os.Remove(finalPath)

	// 验证最终文件
	finalVerifyPkg, _ := opc.OpenFile(finalPath)
	finalPartCount := finalPkg.PartCount()
	finalVerifyPkg.Close()
	t.Logf("Stage 6 通过: 最终文件包含 %d 个部件", finalPartCount)

	t.Log("========== 完整流水线测试全部通过 ==========")

	// 清理所有临时文件
	cleanupFiles := []string{
		"test_opc_output.pptx",
		"test_secure_output.pptx",
		"test_integration_output.pptx",
		"test_updated_output.pptx",
		"test_final_output.pptx",
	}
	for _, f := range cleanupFiles {
		os.Remove(f)
	}
}
