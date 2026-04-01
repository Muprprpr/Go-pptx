package parts_test

import (
	"encoding/xml"
	"strings"
	"testing"

	"github.com/Muprprpr/Go-pptx/parts"
)

// 一段典型的、由 PowerPoint 原生生成的最小化 app.xml
const validAppXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <TotalTime>0</TotalTime>
  <Words>150</Words>
  <Application>Microsoft Office PowerPoint</Application>
  <PresentationFormat>宽屏</PresentationFormat>
  <Paragraphs>20</Paragraphs>
  <Slides>5</Slides>
  <Notes>0</Notes>
  <HiddenSlides>0</HiddenSlides>
  <MMClips>0</MMClips>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size="4" baseType="variant">
      <vt:variant><vt:lpstr>主题</vt:lpstr></vt:variant>
      <vt:variant><vt:i4>1</vt:i4></vt:variant>
      <vt:variant><vt:lpstr>幻灯片标题</vt:lpstr></vt:variant>
      <vt:variant><vt:i4>5</vt:i4></vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="6" baseType="lpstr">
      <vt:lpstr>PowerPoint 演示文稿</vt:lpstr>
      <vt:lpstr>第一页</vt:lpstr>
      <vt:lpstr>第二页</vt:lpstr>
      <vt:lpstr>第三页</vt:lpstr>
      <vt:lpstr>第四页</vt:lpstr>
      <vt:lpstr>第五页</vt:lpstr>
    </vt:vector>
  </TitlesOfParts>
  <Company>Microsoft Corporation</Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>`

// 1. 往返无损与修改测试 (Round-Trip & Mutate)
func TestAppProperties_RoundTripAndMutate(t *testing.T) {
	appProps := &parts.XMLAppProps{}
	err := xml.Unmarshal([]byte(validAppXML), appProps)
	if err != nil {
		t.Fatalf("解析合法 app.xml 失败: %v", err)
	}

	// 验证解析是否成功
	if appProps.Application != "Microsoft Office PowerPoint" {
		t.Errorf("期望 Application 为 'Microsoft Office PowerPoint', 得到 '%s'", appProps.Application)
	}
	if *appProps.Slides != 5 {
		t.Errorf("期望 Slides 为 5, 得到 %d", *appProps.Slides)
	}

	// 模拟底层修改行为
	appProps.Application = "Go-pptx Engine"
	appProps.Company = "My AI Company"
	*appProps.Slides = 99

	// 重新序列化
	outputBytes, err := xml.Marshal(appProps)
	if err != nil {
		t.Fatalf("序列化 app.xml 失败: %v", err)
	}
	outputXML := string(outputBytes)

	// 验证修改后的值是否正确写入
	if !strings.Contains(outputXML, "<Application>Go-pptx Engine</Application>") {
		t.Error("Application 字段未能正确修改并序列化")
	}
	if !strings.Contains(outputXML, "<Company>My AI Company</Company>") {
		t.Error("Company 字段未能正确修改并序列化")
	}
	if !strings.Contains(outputXML, "<Slides>99</Slides>") {
		t.Error("Slides 字段未能正确修改并序列化")
	}
	// 验证复杂嵌套结构没有丢失 (HeadingPairs 和 vt:vector)
	if !strings.Contains(outputXML, "<HeadingPairs>") || !strings.Contains(outputXML, "vt:vector") {
		t.Error("复杂嵌套元素 (HeadingPairs/vt:vector) 在往返序列化时丢失")
	}

	t.Log("✅ App 属性往返与修改测试通过")
}

// 2. 命名空间与根节点测试 (Namespace)
func TestAppProperties_Namespaces(t *testing.T) {
	appProps := &parts.XMLAppProps{
		Application: "Go-pptx Engine",
		XmlnsProp:   parts.NamespaceExtendedProperties,
		XmlnsVt:     parts.NamespaceDocPropsVTypes,
	}

	outputBytes, err := xml.Marshal(appProps)
	if err != nil {
		t.Fatalf("序列化失败: %v", err)
	}
	outputXML := string(outputBytes)

	// OOXML 极其看重这两个命名空间
	expectedNS1 := `xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"`
	expectedNS2 := `xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"`

	if !strings.Contains(outputXML, expectedNS1) {
		t.Errorf("缺失主命名空间扩展属性声明: %s", expectedNS1)
	}
	if !strings.Contains(outputXML, expectedNS2) {
		t.Errorf("缺失 VT 命名空间声明: %s", expectedNS2)
	}

	t.Log("✅ App 命名空间测试通过")
}

// 3. 标签省略测试 (Omitempty)
func TestAppProperties_Omitempty(t *testing.T) {
	// 创建一个除了 Application 外全空的结构体
	appProps := &parts.XMLAppProps{
		Application: "Go-pptx Engine",
	}

	outputBytes, err := xml.Marshal(appProps)
	if err != nil {
		t.Fatalf("序列化失败: %v", err)
	}
	outputXML := string(outputBytes)

	// 验证未赋值的字段不会生成空的标签，例如 <Company></Company>
	if strings.Contains(outputXML, "<Company>") {
		t.Error("未赋值的 Company 字段不应出现在 XML 中, 检查 struct tag 是否缺少 omitempty")
	}
	if strings.Contains(outputXML, "<Manager>") {
		t.Error("未赋值的 Manager 字段不应出现在 XML 中")
	}

	t.Log("✅ App 标签省略测试通过")
}
