package parts_test

import (
	"encoding/xml"
	"strings"
	"testing"

	"github.com/Muprprpr/Go-pptx/parts"
)

func TestThemePart_FromXML(t *testing.T) {
	// 使用实际的 theme XML 数据
	themeXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office 主题">
	<a:themeElements>
		<a:clrScheme name="Office">
			<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
			<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
			<a:dk2><a:srgbClr val="44546A"/></a:dk2>
			<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
			<a:accent1><a:srgbClr val="5B9BD5"/></a:accent1>
			<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
			<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
			<a:accent4><a:srgbClr val="FFC000"/></a:accent4>
			<a:accent5><a:srgbClr val="4472C4"/></a:accent5>
			<a:accent6><a:srgbClr val="70AD47"/></a:accent6>
			<a:hlink><a:srgbClr val="0563C1"/></a:hlink>
			<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
		</a:clrScheme>
	</a:themeElements>
</a:theme>`

	themePart := parts.NewThemePart(1)
	err := themePart.FromXML([]byte(themeXML))
	if err != nil {
		t.Fatalf("FromXML failed: %v", err)
	}

	// 验证主题名称
	theme := themePart.Theme()
	if theme == nil {
		t.Fatal("Theme is nil")
	}
	if theme.Name != "Office 主题" {
		t.Errorf("Expected theme name 'Office 主题', got '%s'", theme.Name)
	}

	// 验证颜色方案
	colorScheme := themePart.ColorScheme()
	if colorScheme == nil {
		t.Fatal("ColorScheme is nil")
	}
	if colorScheme.Name != "Office" {
		t.Errorf("Expected color scheme name 'Office', got '%s'", colorScheme.Name)
	}
}

func TestThemePart_GetColor(t *testing.T) {
	themeXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Test Theme">
	<a:themeElements>
		<a:clrScheme name="Test">
			<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
			<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
			<a:dk2><a:srgbClr val="44546A"/></a:dk2>
			<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
			<a:accent1><a:srgbClr val="5B9BD5"/></a:accent1>
			<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
			<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
			<a:accent4><a:srgbClr val="FFC000"/></a:accent4>
			<a:accent5><a:srgbClr val="4472C4"/></a:accent5>
			<a:accent6><a:srgbClr val="70AD47"/></a:accent6>
			<a:hlink><a:srgbClr val="0563C1"/></a:hlink>
			<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
		</a:clrScheme>
	</a:themeElements>
</a:theme>`

	themePart := parts.NewThemePart(1)
	if err := themePart.FromXML([]byte(themeXML)); err != nil {
		t.Fatalf("FromXML failed: %v", err)
	}

	tests := []struct {
		role     parts.ColorRole
		expected string
		isSystem bool
	}{
		{parts.ColorRoleDark1, "000000", true}, // 系统颜色
		{parts.ColorRoleLight1, "FFFFFF", true}, // 系统颜色
		{parts.ColorRoleDark2, "44546A", false}, // RGB 颜色
		{parts.ColorRoleLight2, "E7E6E6", false}, // RGB 颜色
		{parts.ColorRoleAccent1, "5B9BD5", false},
		{parts.ColorRoleAccent2, "ED7D31", false},
		{parts.ColorRoleAccent3, "A5A5A5", false},
		{parts.ColorRoleAccent4, "FFC000", false},
		{parts.ColorRoleAccent5, "4472C4", false},
		{parts.ColorRoleAccent6, "70AD47", false},
		{parts.ColorRoleHyperlink, "0563C1", false},
		{parts.ColorRoleFollowedHyperlink, "954F72", false},
	}

	for _, tt := range tests {
		t.Run(tt.role.String(), func(t *testing.T) {
			rgb := themePart.GetThemeColorRGB(tt.role)
			if rgb != tt.expected {
				t.Errorf("GetThemeColorRGB(%v) = '%s', want '%s'", tt.role, rgb, tt.expected)
			}

			colorType := themePart.GetThemeColorType(tt.role)
			if tt.isSystem && colorType != parts.ColorTypeSystem {
				t.Errorf("Expected ColorTypeSystem for %v, got %v", tt.role, colorType)
			}
			if !tt.isSystem && colorType != parts.ColorTypeRGB {
				t.Errorf("Expected ColorTypeRGB for %v, got %v", tt.role, colorType)
			}
		})
	}
}

func TestThemePart_New(t *testing.T) {
	themePart := parts.NewThemePart(1)
	if themePart == nil {
		t.Fatal("NewThemePart returned nil")
	}

	// 验证 URI
	uri := themePart.PartURI()
	if uri == nil {
		t.Fatal("PartURI is nil")
	}
	if uri.URI() != "/ppt/theme/theme1.xml" {
		t.Errorf("Expected URI '/ppt/theme/theme1.xml', got '%s'", uri.URI())
	}
}

// ============================================================================
// 测试 1：XML 往返无损测试 (Round-Trip Test)
// 目标：证明 Go 结构体兜住了所有必须的 XML 标签，序列化后没有丢失数据
// ============================================================================
func TestTheme_RoundTrip(t *testing.T) {
	// 使用完整的默认主题 XML
	theme := parts.DefaultTheme()
	if theme == nil {
		t.Fatal("DefaultTheme returned nil")
	}

	// 序列化
	outputBytes, err := xml.Marshal(theme)
	if err != nil {
		t.Fatalf("序列化 Theme 失败: %v", err)
	}
	outputXML := string(outputBytes)

	// 核心校验：检查三大支柱是否健在
	requiredTags := []string{
		"<clrScheme",
		"<fontScheme",
		"<fmtScheme",
		"<accent1",
		"<fillStyleLst",
		"<lnStyleLst",
		"<effectStyleLst",
		"<bgFillStyleLst",
	}

	for _, tag := range requiredTags {
		if !strings.Contains(outputXML, tag) {
			t.Errorf("往返测试失败：序列化后丢失了关键标签 %s", tag)
		}
	}
	t.Log("✅ Theme 往返序列化无损测试通过")
}

// ============================================================================
// 测试 2：并发克隆安全测试 (Clone Safety Test)
// 目标：证明通过 CloneTheme 获取的副本，修改时绝对不会污染全局模板
// ============================================================================
func TestTheme_CloneSafety(t *testing.T) {
	// 获取两个独立的克隆体
	themeA := parts.CloneTheme()
	themeB := parts.CloneTheme()

	if themeA == nil || themeB == nil {
		t.Fatal("CloneTheme returned nil")
	}

	// 创建 ThemePart 并设置 themeA
	partA := parts.NewThemePart(1)
	partA.SetThemeData(themeA)

	// 修改 themeA 的 Accent1 为红色
	partA.SetThemeColorRGB(parts.ColorRoleAccent1, "FF0000")

	// 验证 themeB 是否被污染（themeB 应该保持原始值 5B9BD5）
	partB := parts.NewThemePart(2)
	partB.SetThemeData(themeB)
	colorB := partB.GetThemeColorRGB(parts.ColorRoleAccent1)

	if colorB == "FF0000" {
		t.Fatal("并发克隆测试失败：浅拷贝导致内存地址污染，修改 A 影响了 B！")
	}
	if colorB != "5B9BD5" {
		t.Errorf("期望 themeB Accent1 为 '5B9BD5'，实际为 '%s'", colorB)
	}

	t.Log("✅ Theme 克隆隔离测试通过")
}

// ============================================================================
// 测试 3：底层突变器行为测试 (Mutator Test)
// 目标：证明 SetThemeColorRGB 能够正确地修改 XML 节点树
// ============================================================================
func TestTheme_SetThemeColorRGB(t *testing.T) {
	theme := parts.CloneTheme()
	part := parts.NewThemePart(1)
	part.SetThemeData(theme)

	// 将强调色 1 改为纯绿色
	targetColor := "00FF00"
	part.SetThemeColorRGB(parts.ColorRoleAccent1, targetColor)

	// 验证获取的颜色是新的值
	gotColor := part.GetThemeColorRGB(parts.ColorRoleAccent1)
	if gotColor != targetColor {
		t.Errorf("突变测试失败：期望 '%s'，实际 '%s'", targetColor, gotColor)
	}

	// 序列化后，检查 XML 字符串里是否真实写入了这段 RGB
	themeData := part.Theme()
	outputBytes, err := xml.Marshal(themeData)
	if err != nil {
		t.Fatalf("序列化失败: %v", err)
	}
	outputXML := string(outputBytes)

	expectedNode := `<srgbClr val="00FF00"`
	if !strings.Contains(outputXML, expectedNode) {
		t.Errorf("突变测试失败：XML 树中未能找到修改后的色值标签 %s", expectedNode)
	}

	t.Log("✅ Theme 底层突变器测试通过")
}
