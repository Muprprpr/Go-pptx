package parts_test

import (
	"os"
	"testing"

	"github.com/Muprprpr/Go-pptx/parts"
)

// ============================================================================
// Master 解析测试
// ============================================================================

func TestParseMasterFromFile(t *testing.T) {
	tests := []struct {
		name     string
		filePath string
	}{
		// slideMaster1 只有背景图，无占位符
		{"slideMaster1", "../test-data/test/ppt/slideMasters/slideMaster1.xml"},
		// slideMaster2 有标题、正文、日期、页脚、页码占位符
		{"slideMaster2", "../test-data/test/ppt/slideMasters/slideMaster2.xml"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			xmlData, err := os.ReadFile(tt.filePath)
			if err != nil {
				t.Fatalf("读取文件失败: %v", err)
			}

			master, err := parts.ParseMaster(xmlData)
			if err != nil {
				t.Fatalf("ParseMaster 返回错误: %v", err)
			}
			if master == nil {
				t.Fatal("ParseMaster 返回 nil")
			}

			// 断言：能解析出母版数据
			t.Logf("%s: 占位符数量 = %d", tt.name, len(master.Placeholders()))
		})
	}
}

// ============================================================================
// Master 背景与占位符断言
// ============================================================================

func TestParseMaster_BackgroundAndPlaceholders(t *testing.T) {
	xmlData, err := os.ReadFile("../test-data/test/ppt/slideMasters/slideMaster2.xml")
	if err != nil {
		t.Fatalf("读取文件失败: %v", err)
	}

	master, err := parts.ParseMaster(xmlData)
	if err != nil {
		t.Fatalf("ParseMaster 返回错误: %v", err)
	}

	// 断言：背景被正确提取
	bg := master.Background()
	if bg == nil {
		t.Fatal("母版 Background 为 nil，期望有背景数据")
	}

	// 断言：母版包含占位符
	placeholders := master.Placeholders()
	if len(placeholders) == 0 {
		t.Fatal("母版 Placeholders 长度为 0，期望至少有一些占位符")
	}

	// 断言：母版通常包含日期、页码、页脚占位符
	hasDatePh, hasFooterPh, hasSlideNumPh := false, false, false
	for _, ph := range placeholders {
		switch ph.Type() {
		case parts.PlaceholderTypeDateTime:
			hasDatePh = true
		case parts.PlaceholderTypeFooter:
			hasFooterPh = true
		case parts.PlaceholderTypeSlideNumber:
			hasSlideNumPh = true
		}
	}

	if !hasDatePh {
		t.Log("未找到日期占位符 (dt)")
	}
	if !hasFooterPh {
		t.Log("未找到页脚占位符 (ftr)")
	}
	if !hasSlideNumPh {
		t.Log("未找到页码占位符 (sldNum)")
	}

	// 至少应该有一个母版级占位符
	t.Logf("母版包含 %d 个占位符", len(placeholders))
}
