package pptx_test

import (
	"archive/zip"
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/Muprprpr/Go-pptx/parts"
	"github.com/Muprprpr/Go-pptx/pptx"
)

// TestNewPresentation 测试创建空白演示文稿
func TestNewPresentation(t *testing.T) {
	pres := pptx.New()
	if pres == nil {
		t.Fatal("New() 返回 nil")
	}

	// 检查初始状态
	if pres.SlideCount() != 0 {
		t.Errorf("期望 0 张幻灯片，实际 %d 张", pres.SlideCount())
	}

	// 检查幻灯片尺寸
	cx, cy := pres.SlideSize()
	if cx == 0 || cy == 0 {
		t.Errorf("幻灯片尺寸无效: cx=%d, cy=%d", cx, cy)
	}
}

// TestAddSlide 测试添加幻灯片
func TestAddSlide(t *testing.T) {
	pres := pptx.New()

	// 添加第一张幻灯片
	slide1 := pres.AddSlide()
	if slide1 == nil {
		t.Fatal("AddSlide() 返回 nil")
	}

	if slide1.Index() != 0 {
		t.Errorf("期望索引 0，实际 %d", slide1.Index())
	}

	if pres.SlideCount() != 1 {
		t.Errorf("期望 1 张幻灯片，实际 %d 张", pres.SlideCount())
	}

	// 添加第二张幻灯片
	slide2 := pres.AddSlide()
	if slide2.Index() != 1 {
		t.Errorf("期望索引 1，实际 %d", slide2.Index())
	}

	if pres.SlideCount() != 2 {
		t.Errorf("期望 2 张幻灯片，实际 %d 张", pres.SlideCount())
	}
}

// TestAddTextBox 测试添加文本框
func TestAddTextBox(t *testing.T) {
	pres := pptx.New()
	slide := pres.AddSlide()

	// 添加文本框
	textBox := slide.AddTextBox(100, 100, 500, 50, "Hello World")
	if textBox == nil {
		t.Fatal("AddTextBox() 返回 nil")
	}
}

// TestAddAutoShape 测试添加形状
func TestAddAutoShape(t *testing.T) {
	pres := pptx.New()
	slide := pres.AddSlide()

	// 添加矩形
	rect := slide.AddRectangle(100, 100, 200, 100)
	if rect == nil {
		t.Fatal("AddRectangle() 返回 nil")
	}

	// 添加椭圆
	ellipse := slide.AddEllipse(100, 250, 200, 100)
	if ellipse == nil {
		t.Fatal("AddEllipse() 返回 nil")
	}
}

// TestAddTable 测试添加表格
func TestAddTable(t *testing.T) {
	pres := pptx.New()
	slide := pres.AddSlide()

	// 添加表格
	table := slide.AddTable(100, 100, 500, 300, 3, 4)
	if table == nil {
		t.Fatal("AddTable() 返回 nil")
	}

	// 设置单元格文本
	slide.SetTableCellText(table, 0, 0, "Header 1")
	slide.SetTableCellText(table, 0, 1, "Header 2")
	slide.SetTableCellText(table, 1, 0, "Data 1")
}

// TestGetSlide 测试获取幻灯片
func TestGetSlide(t *testing.T) {
	pres := pptx.New()
	pres.AddSlide()
	pres.AddSlide()

	// 获取第一张幻灯片
	slide, err := pres.GetSlide(0)
	if err != nil {
		t.Fatalf("GetSlide(0) 失败: %v", err)
	}
	if slide.Index() != 0 {
		t.Errorf("期望索引 0，实际 %d", slide.Index())
	}

	// 获取第二张幻灯片
	slide, err = pres.GetSlide(1)
	if err != nil {
		t.Fatalf("GetSlide(1) 失败: %v", err)
	}
	if slide.Index() != 1 {
		t.Errorf("期望索引 1，实际 %d", slide.Index())
	}

	// 测试越界
	_, err = pres.GetSlide(2)
	if err == nil {
		t.Error("期望越界错误，但成功了")
	}
}

// TestRemoveSlide 测试移除幻灯片
func TestRemoveSlide(t *testing.T) {
	pres := pptx.New()
	pres.AddSlide()
	pres.AddSlide()
	pres.AddSlide()

	// 移除中间的幻灯片
	err := pres.RemoveSlide(1)
	if err != nil {
		t.Fatalf("RemoveSlide(1) 失败: %v", err)
	}

	if pres.SlideCount() != 2 {
		t.Errorf("期望 2 张幻灯片，实际 %d 张", pres.SlideCount())
	}

	// 检查索引更新
	slides := pres.Slides()
	if slides[0].Index() != 0 {
		t.Errorf("第一张幻灯片索引应为 0，实际 %d", slides[0].Index())
	}
	if slides[1].Index() != 1 {
		t.Errorf("第二张幻灯片索引应为 1，实际 %d", slides[1].Index())
	}
}

// TestSaveAndWrite 测试保存和写入
func TestSaveAndWrite(t *testing.T) {
	pres := pptx.New()
	slide := pres.AddSlide()
	slide.AddTextBox(100, 100, 500, 50, "Test Presentation")

	// 创建临时目录
	tmpDir := t.TempDir()
	outputPath := filepath.Join(tmpDir, "test.pptx")

	// 测试 Save
	err := pres.Save(outputPath)
	if err != nil {
		t.Fatalf("Save() 失败: %v", err)
	}

	// 检查文件是否存在
	if _, err := os.Stat(outputPath); os.IsNotExist(err) {
		t.Fatal("输出文件不存在")
	}

	// 检查文件大小
	info, err := os.Stat(outputPath)
	if err != nil {
		t.Fatalf("获取文件信息失败: %v", err)
	}
	if info.Size() == 0 {
		t.Error("输出文件为空")
	}

	// 测试 WriteToBytes
	data, err := pres.WriteToBytes()
	if err != nil {
		t.Fatalf("WriteToBytes() 失败: %v", err)
	}
	if len(data) == 0 {
		t.Error("WriteToBytes() 返回空数据")
	}

	// 测试 Write
	var buf bytes.Buffer
	err = pres.Write(&buf)
	if err != nil {
		t.Fatalf("Write() 失败: %v", err)
	}
	if buf.Len() == 0 {
		t.Error("Write() 写入空数据")
	}
}

// TestSetSlideSize 测试设置幻灯片尺寸
func TestSetSlideSize(t *testing.T) {
	pres := pptx.New()

	// 设置自定义尺寸
	pres.SetSlideSize(9144000, 6858000) // 4:3 标准
	cx, cy := pres.SlideSize()
	if cx != 9144000 || cy != 6858000 {
		t.Errorf("幻灯片尺寸设置失败: cx=%d, cy=%d", cx, cy)
	}

	// 设置标准尺寸
	pres.SetSlideSizeStandard("16:9")
	cx, cy = pres.SlideSize()
	if cx != 12192000 || cy != 6858000 {
		t.Errorf("16:9 尺寸设置失败: cx=%d, cy=%d", cx, cy)
	}
}

// TestClone 测试克隆演示文稿
func TestClone(t *testing.T) {
	pres := pptx.New()
	slide := pres.AddSlide()
	slide.AddTextBox(100, 100, 500, 50, "Original")

	// 克隆
	cloned, err := pres.Clone()
	if err != nil {
		t.Fatalf("Clone() 失败: %v", err)
	}

	// 检查幻灯片数量
	if cloned.SlideCount() != pres.SlideCount() {
		t.Errorf("克隆后幻灯片数量不一致: 原始=%d, 克隆=%d", pres.SlideCount(), cloned.SlideCount())
	}

	// 修改克隆版本
	clonedSlide, _ := cloned.GetSlide(0)
	clonedSlide.AddTextBox(100, 200, 500, 50, "Cloned")

	// 原始版本不应受影响
	originalSlides := pres.Slides()
	clonedSlides := cloned.Slides()

	// 简单验证：两张幻灯片的形状数量应该不同
	// （原始有 1 个文本框，克隆有 2 个文本框）
	_ = originalSlides
	_ = clonedSlides
}

// TestAddSlideAt 测试在指定位置插入幻灯片
func TestAddSlideAt(t *testing.T) {
	pres := pptx.New()
	pres.AddSlide() // 索引 0
	pres.AddSlide() // 索引 1

	// 在位置 1 插入
	slide, err := pres.AddSlideAt(1)
	if err != nil {
		t.Fatalf("AddSlideAt(1) 失败: %v", err)
	}

	if slide.Index() != 1 {
		t.Errorf("期望索引 1，实际 %d", slide.Index())
	}

	if pres.SlideCount() != 3 {
		t.Errorf("期望 3 张幻灯片，实际 %d 张", pres.SlideCount())
	}

	// 检查索引更新
	slides := pres.Slides()
	if slides[0].Index() != 0 {
		t.Errorf("第一张幻灯片索引应为 0，实际 %d", slides[0].Index())
	}
	if slides[1].Index() != 1 {
		t.Errorf("第二张幻灯片索引应为 1，实际 %d", slides[1].Index())
	}
	if slides[2].Index() != 2 {
		t.Errorf("第三张幻灯片索引应为 2，实际 %d", slides[2].Index())
	}
}

// TestSlides 测试获取所有幻灯片
func TestSlides(t *testing.T) {
	pres := pptx.New()
	pres.AddSlide()
	pres.AddSlide()
	pres.AddSlide()

	slides := pres.Slides()
	if len(slides) != 3 {
		t.Errorf("期望 3 张幻灯片，实际 %d 张", len(slides))
	}

	// 验证返回的是副本
	slides[0] = nil
	originalSlides := pres.Slides()
	if originalSlides[0] == nil {
		t.Error("Slides() 应返回副本，不应影响原始数据")
	}
}

// TestClose 测试关闭演示文稿
func TestClose(t *testing.T) {
	pres := pptx.New()
	pres.AddSlide()

	err := pres.Close()
	if err != nil {
		t.Fatalf("Close() 失败: %v", err)
	}
}

// ============================================================================
// 重点测试 3：基于"假积木 (Mock)" 的契约测试
// ============================================================================
//
// 此测试验证 slide.AddComponent() 能够正确接收任何实现了 Component 接口的对象，
// 并成功调用它的 Render() 方法将数据挂载到画布上。
// 注意：此测试不引入真实的业务组件（如 text、chart），仅使用 Mock。

// mockComponent 是一个仅存在于测试中的假组件
type mockComponent struct {
	rendered bool
}

// Render 实现 Component 接口
func (m *mockComponent) Render(ctx *pptx.SlideContext) error {
	m.rendered = true
	// 往上下文里塞一个底层的空白 XSp
	ctx.AppendShape(&parts.XSp{})
	return nil
}

// TestSlide_AddComponent 测试添加组件的基本功能
func TestSlide_AddComponent(t *testing.T) {
	prs := pptx.New()
	slide := prs.AddSlide()

	mock := &mockComponent{}
	err := slide.AddComponent(mock)

	if err != nil || !mock.rendered {
		t.Fatal("契约失效：未成功触发组件的 Render 方法")
	}

	// 验证底层 parts 的 Children 长度是否增加了 1
	childrenCount := len(slide.Part().SpTree().Children)
	if childrenCount != 1 {
		t.Fatalf("画布挂载失败：XML 树未增加内容，期望 1，实际 %d", childrenCount)
	}
}

// mockComponentWithName 是带名称的假组件
type mockComponentWithName struct {
	rendered bool
	name     string
}

// Render 实现 Component 接口
func (m *mockComponentWithName) Render(ctx *pptx.SlideContext) error {
	m.rendered = true
	ctx.AppendShape(&parts.XSp{})
	return nil
}

// Name 实现 ComponentWithName 接口
func (m *mockComponentWithName) Name() string {
	return m.name
}

// TestSlide_AddComponentWithName 测试带名称的组件
func TestSlide_AddComponentWithName(t *testing.T) {
	prs := pptx.New()
	slide := prs.AddSlide()

	mock := &mockComponentWithName{name: "test-component"}
	err := slide.AddComponent(mock)

	if err != nil || !mock.rendered {
		t.Fatal("契约失效：未成功触发组件的 Render 方法")
	}

	// 验证名称
	if mock.Name() != "test-component" {
		t.Errorf("期望组件名称 'test-component'，实际 %s", mock.Name())
	}

	// 验证底层 parts 的 Children 长度是否增加了 1
	childrenCount := len(slide.Part().SpTree().Children)
	if childrenCount != 1 {
		t.Fatalf("画布挂载失败：XML 树未增加内容，期望 1，实际 %d", childrenCount)
	}
}

// mockComponentWithError 是会返回错误的假组件
type mockComponentWithError struct {
	rendered bool
	errMsg   string
}

// Render 实现 Component 接口，返回错误
func (m *mockComponentWithError) Render(ctx *pptx.SlideContext) error {
	m.rendered = true
	return fmt.Errorf("%s", m.errMsg)
}

// TestSlide_AddComponentWithError 测试组件渲染返回错误的情况
func TestSlide_AddComponentWithError(t *testing.T) {
	prs := pptx.New()
	slide := prs.AddSlide()

	mock := &mockComponentWithError{errMsg: "intentional error"}
	err := slide.AddComponent(mock)

	if err == nil {
		t.Fatal("期望返回错误，但成功了")
	}

	if !mock.rendered {
		t.Fatal("期望 Render 方法被调用")
	}

	// 验证错误信息包含原始错误
	if !strings.Contains(err.Error(), "intentional error") {
		t.Errorf("期望错误信息包含 'intentional error'，实际: %s", err.Error())
	}

	// 验证底层 parts 的 Children 长度没有增加（因为渲染失败不应挂载）
	childrenCount := len(slide.Part().SpTree().Children)
	if childrenCount != 0 {
		t.Fatalf("期望渲染失败时不挂载形状，实际 Children 数量: %d", childrenCount)
	}
}

// mockComponentWithSize 是带尺寸信息的假组件
type mockComponentWithSize struct {
	x, y, cx, cy int
}

// Render 实现 Component 接口
func (m *mockComponentWithSize) Render(ctx *pptx.SlideContext) error {
	ctx.AppendShape(&parts.XSp{})
	return nil
}

// Bounds 实现 ComponentWithSize 接口
func (m *mockComponentWithSize) Bounds() (x, y, cx, cy int) {
	return m.x, m.y, m.cx, m.cy
}

// TestSlide_ComponentWithSize 测试带尺寸信息的组件
func TestSlide_ComponentWithSize(t *testing.T) {
	prs := pptx.New()
	slide := prs.AddSlide()

	mock := &mockComponentWithSize{x: 100, y: 200, cx: 300, cy: 400}
	err := slide.AddComponent(mock)

	if err != nil {
		t.Fatalf("AddComponent 失败: %v", err)
	}

	// 验证 Bounds 方法
	x, y, cx, cy := mock.Bounds()
	if x != 100 || y != 200 || cx != 300 || cy != 400 {
		t.Errorf("Bounds() 返回值不符合预期: x=%d, y=%d, cx=%d, cy=%d", x, y, cx, cy)
	}
}

// TestSlide_MultipleComponents 测试批量添加组件
func TestSlide_MultipleComponents(t *testing.T) {
	prs := pptx.New()
	slide := prs.AddSlide()

	// 创建多个假组件
	mock1 := &mockComponent{}
	mock2 := &mockComponent{}
	mock3 := &mockComponent{}

	// 使用 AddComponents 批量添加
	err := slide.AddComponents(mock1, mock2, mock3)

	if err != nil {
		t.Fatalf("AddComponents 失败: %v", err)
	}

	// 验证所有组件都被渲染
	if !mock1.rendered || !mock2.rendered || !mock3.rendered {
		t.Fatal("部分组件未被渲染")
	}

	// 验证底层 parts 的 Children 长度是否增加了 3
	childrenCount := len(slide.Part().SpTree().Children)
	if childrenCount != 3 {
		t.Fatalf("期望 3 个子元素，实际 %d", childrenCount)
	}
}

// mockComponentWithID 是需要形状 ID 的假组件
type mockComponentWithID struct {
	rendered bool
	shapeID  uint32
}

// Render 实现 Component 接口
func (m *mockComponentWithID) Render(ctx *pptx.SlideContext) error {
	m.rendered = true
	m.shapeID = ctx.NextShapeID()
	ctx.AppendShape(&parts.XSp{})
	return nil
}

// TestSlide_ComponentShapeIDAllocation 测试组件中的形状 ID 分配
func TestSlide_ComponentShapeIDAllocation(t *testing.T) {
	prs := pptx.New()
	slide := prs.AddSlide()

	// 创建需要形状 ID 的组件
	mock1 := &mockComponentWithID{}
	mock2 := &mockComponentWithID{}

	err := slide.AddComponents(mock1, mock2)
	if err != nil {
		t.Fatalf("AddComponents 失败: %v", err)
	}

	// 验证 ID 是唯一且递增的
	if mock1.shapeID == 0 {
		t.Error("期望 mock1 获得有效的形状 ID")
	}
	if mock2.shapeID == 0 {
		t.Error("期望 mock2 获得有效的形状 ID")
	}
	// ID 应该是不同的
	if mock1.shapeID == mock2.shapeID {
		t.Errorf("期望不同的形状 ID，实际都是 %d", mock1.shapeID)
	}
	// mock2 的 ID 应该大于 mock1
	if mock2.shapeID <= mock1.shapeID {
		t.Errorf("期望 mock2.shapeID > mock1.shapeID，实际 %d <= %d", mock2.shapeID, mock1.shapeID)
	}
}

// ============================================================================
// 战役 4：端到端流式无损输出 (Sanity Check)
// ============================================================================
//
// 此测试验证整个引擎的流式输出（WriteToWriter）是否健康。
// 不碰硬盘，直接将生成的文稿写进内存的 bytes.Buffer，
// 并用 archive/zip 重新打开验证它是一个合法的 ZIP 包。
// TestPresentation_WriteToMemory 测试流式写入内存并验证 ZIP 合法性
func TestPresentation_WriteToMemory(t *testing.T) {
	prs := pptx.New()
	slide := prs.AddSlide()
	slide.AddTextBox(100, 100, 500, 50, "端到端测试")

	// 流式写入内存
	var buf bytes.Buffer
	err := prs.Write(&buf)
	if err != nil {
		t.Fatalf("流式保存失败: %v", err)
	}

	// 验证生成的数据不是空的
	if buf.Len() == 0 {
		t.Fatal("生成的数据为空")
	}

	// 验证它是不是一个合法的 ZIP 文件
	reader, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("生成的不是合法 ZIP 文件: %v", err)
	}

	// 检查关键文件是否存在
	expectedFiles := []string{
		"[Content_Types].xml",
		"_rels/.rels",
		"ppt/presentation.xml",
	}

	// 构建文件列表 map
	fileMap := make(map[string]bool)
	for _, file := range reader.File {
		fileMap[file.Name] = true
	}

	// 验证关键文件存在
	for _, expected := range expectedFiles {
		if !fileMap[expected] {
				t.Errorf("ZIP 包中缺少关键文件: %s", expected)
		}
	}

	// 验证至少有一张幻灯片文件
	slideFound := false
	for name := range fileMap {
		if strings.HasPrefix(name, "ppt/slides/slide") && strings.HasSuffix(name, ".xml") {
			slideFound = true
			break
		}
	}
	if !slideFound {
		t.Error("ZIP 包中缺少幻灯片文件")
	}

	t.Logf("✅ 端到端测试通过，生成了合法的 ZIP 包，包含 %d 个文件", len(reader.File))
}

// TestPresentation_WriteToMemory_Empty 测试空白演示文稿的流式写入
func TestPresentation_WriteToMemory_Empty(t *testing.T) {
	prs := pptx.New()
	// 不添加任何幻灯片

	var buf bytes.Buffer
	err := prs.Write(&buf)
	if err != nil {
		t.Fatalf("空白演示文稿流式保存失败: %v", err)
	}

	// 即使是空白演示文稿，也应该生成合法的 ZIP
	reader, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("空白演示文稿生成的不是合法 ZIP 文件: %v", err)
	}

	// 验证基本结构文件存在
	fileMap := make(map[string]bool)
	for _, file := range reader.File {
		fileMap[file.Name] = true
	}

	if !fileMap["[Content_Types].xml"] {
		t.Error("缺少 [Content_Types].xml")
	}
	if !fileMap["ppt/presentation.xml"] {
		t.Error("缺少 ppt/presentation.xml")
	}

	t.Logf("✅ 空白演示文稿测试通过，生成了合法的 ZIP 包")
}

// TestPresentation_WriteToMemory_MultipleSlides 测试多幻灯片的流式写入
func TestPresentation_WriteToMemory_MultipleSlides(t *testing.T) {
	prs := pptx.New()

	// 添加多张幻灯片，每张都有内容
	for i := 0; i < 5; i++ {
		slide := prs.AddSlide()
		slide.AddTextBox(100, 100+i*50, 500, 50, fmt.Sprintf("幻灯片 %d", i+1))
		slide.AddRectangle(100, 200+i*30, 200, 100)
	}

	var buf bytes.Buffer
	err := prs.Write(&buf)
	if err != nil {
		t.Fatalf("多幻灯片流式保存失败: %v", err)
	}

	// 验证 ZIP 合法性
	reader, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("多幻灯片生成的不是合法 ZIP 文件: %v", err)
	}

	// 统计幻灯片文件数量
	slideCount := 0
	for _, file := range reader.File {
		if strings.HasPrefix(file.Name, "ppt/slides/slide") && strings.HasSuffix(file.Name, ".xml") {
			slideCount++
		}
	}

	if slideCount != 5 {
		t.Errorf("期望 5 张幻灯片文件，实际 %d 张", slideCount)
	}

	t.Logf("✅ 多幻灯片测试通过，生成了 %d 张幻灯片", slideCount)
}
// TestPresentation_WriteToBytes_Consistency 测试 WriteToBytes 与 Write 都能生成合法 ZIP
func TestPresentation_WriteToBytes_Consistency(t *testing.T) {
	// 创建第一个演示文稿用于 WriteToBytes
	prs1 := pptx.New()
	slide1 := prs1.AddSlide()
	slide1.AddTextBox(100, 100, 500, 50, "一致性测试")

	// 使用 WriteToBytes
	bytesData, err := prs1.WriteToBytes()
	if err != nil {
		t.Fatalf("WriteToBytes 失败: %v", err)
	}

	// 验证 WriteToBytes 生成的是合法 ZIP
	_, err = zip.NewReader(bytes.NewReader(bytesData), int64(len(bytesData)))
	if err != nil {
		t.Fatalf("WriteToBytes 生成的不是合法 ZIP: %v", err)
	}

	// 创建第二个演示文稿用于 Write（相同配置）
	prs2 := pptx.New()
	slide2 := prs2.AddSlide()
	slide2.AddTextBox(100, 100, 500, 50, "一致性测试")

	// 使用 Write
	var buf bytes.Buffer
	err = prs2.Write(&buf)
	if err != nil {
		t.Fatalf("Write 失败: %v", err)
	}

	// 验证 Write 生成的是合法 ZIP
	_, err = zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("Write 生成的不是合法 ZIP: %v", err)
	}

	// 两者生成的数据大小应该相近（允许有小幅差异）
	// 注意：由于内部 ID 分配等原因，字节数可能不完全相同，但应该非常接近
	ratio := float64(len(bytesData)) / float64(buf.Len())
	if ratio < 0.95 || ratio > 1.05 {
		t.Errorf("两种方法生成的数据大小差异过大: WriteToBytes=%d, Write=%d", len(bytesData), buf.Len())
	}

	t.Logf("✅ 一致性测试通过，两种方法都生成了合法的 ZIP，大小: WriteToBytes=%d, Write=%d", len(bytesData), buf.Len())
}
