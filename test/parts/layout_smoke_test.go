package parts_test

import (
	"os"
	"testing"

	"github.com/Muprprpr/Go-pptx/parts"
)

// ============================================================================
// Layout 冒烟测试
// ============================================================================

func TestParseLayoutFromFile(t *testing.T) {
	tests := []struct {
		name     string
		filePath string
	}{
		{"slideLayout5", "../test-data/test/ppt/slideLayouts/slideLayout5.xml"},
		{"slideLayout7", "../test-data/test/ppt/slideLayouts/slideLayout7.xml"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			xmlData, err := os.ReadFile(tt.filePath)
			if err != nil {
				t.Fatalf("读取文件失败: %v", err)
			}

			layout, err := parts.ParseLayout(xmlData)
			if err != nil {
				t.Fatalf("ParseLayout 返回错误: %v", err)
			}
			if layout == nil {
				t.Fatal("ParseLayout 返回 nil")
			}

			placeholders := layout.Placeholders()
			if len(placeholders) == 0 {
				t.Error("Placeholders 长度为 0，期望至少有 1 个占位符")
			}
		})
	}
}

// ============================================================================
// Layout 精准坐标与类型断言
// ============================================================================

func TestParseLayout_PlaceholderDetails(t *testing.T) {
	xmlData, err := os.ReadFile("../test-data/test/ppt/slideLayouts/slideLayout5.xml")
	if err != nil {
		t.Fatalf("读取文件失败: %v", err)
	}

	layout, err := parts.ParseLayout(xmlData)
	if err != nil {
		t.Fatalf("ParseLayout 返回错误: %v", err)
	}

	// 断言：获取标题占位符
	titlePh := layout.PlaceholderByType(parts.PlaceholderTypeTitle)
	if titlePh == nil {
		t.Fatal("未找到 title 类型的占位符")
	}

	// 断言：标题占位符坐标和尺寸必须大于 0
	if titlePh.X() <= 0 {
		t.Errorf("标题占位符 X = %d, 期望 > 0", titlePh.X())
	}
	if titlePh.Y() <= 0 {
		t.Errorf("标题占位符 Y = %d, 期望 > 0", titlePh.Y())
	}
	if titlePh.Cx() <= 0 {
		t.Errorf("标题占位符 Cx (宽度) = %d, 期望 > 0", titlePh.Cx())
	}
	if titlePh.Cy() <= 0 {
		t.Errorf("标题占位符 Cy (高度) = %d, 期望 > 0", titlePh.Cy())
	}

	// 尝试获取 body 占位符（可能不存在）
	bodyPh := layout.PlaceholderByType(parts.PlaceholderTypeBody)
	if bodyPh != nil {
		// 断言：正文占位符坐标和尺寸必须大于 0
		if bodyPh.X() <= 0 {
			t.Errorf("正文占位符 X = %d, 期望 > 0", bodyPh.X())
		}
		if bodyPh.Y() <= 0 {
			t.Errorf("正文占位符 Y = %d, 期望 > 0", bodyPh.Y())
		}
		if bodyPh.Cx() <= 0 {
			t.Errorf("正文占位符 Cx (宽度) = %d, 期望 > 0", bodyPh.Cx())
		}
		if bodyPh.Cy() <= 0 {
			t.Errorf("正文占位符 Cy (高度) = %d, 期望 > 0", bodyPh.Cy())
		}
	}
}
