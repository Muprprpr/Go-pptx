package parts_test

import (
	"testing"

	"github.com/Muprprpr/Go-pptx/parts"
)

// ============================================================================
// Core Properties 解析测试
// ============================================================================

func TestParseCoreProps(t *testing.T) {
	tests := []struct {
		name           string
		xmlData        string
		wantTitle      string
		wantCreator    string
		wantSubject    string
		wantCreated    string
		wantModified   string
		wantKeywords   string
		wantLastModBy  string
		wantRevision   string
		wantCategory   string
		wantError      bool
	}{
		{
			name: "正常解析-完整核心属性",
			xmlData: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>演示文稿标题</dc:title>
  <dc:creator>张三</dc:creator>
  <dc:subject>测试主题</dc:subject>
  <dc:description>这是一个测试文档</dc:description>
  <cp:keywords>测试,示例,PPT</cp:keywords>
  <cp:lastModifiedBy>李四</cp:lastModifiedBy>
  <cp:revision>3</cp:revision>
  <cp:category>演示文稿</cp:category>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-01-15T10:30:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-16T14:45:00Z</dcterms:modified>
</cp:coreProperties>`,
			wantTitle:     "演示文稿标题",
			wantCreator:   "张三",
			wantSubject:   "测试主题",
			wantKeywords:  "测试,示例,PPT",
			wantLastModBy: "李四",
			wantRevision:  "3",
			wantCategory:  "演示文稿",
			wantCreated:   "2024-01-15T10:30:00Z",
			wantModified:  "2024-01-16T14:45:00Z",
		},
		{
			name: "正常解析-仅基础字段",
			xmlData: `<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>简单标题</dc:title>
  <dc:creator>作者名</dc:creator>
</cp:coreProperties>`,
			wantTitle:   "简单标题",
			wantCreator: "作者名",
		},
		{
			name: "正常解析-带版本信息",
			xmlData: `<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>版本测试</dc:title>
  <dc:creator>开发团队</dc:creator>
  <cp:revision>15</cp:revision>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-03-20T08:00:00Z</dcterms:created>
</cp:coreProperties>`,
			wantTitle:    "版本测试",
			wantCreator:  "开发团队",
			wantRevision: "15",
			wantCreated:  "2024-03-20T08:00:00Z",
		},
		{
			name: "边界-空核心属性元素",
			xmlData: `<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>`,
		},
		{
			name: "边界-仅命名空间声明",
			xmlData: `<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/">
</cp:coreProperties>`,
		},
		{
			name: "边界-仅时间戳",
			xmlData: `<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-12-31T23:59:59Z</dcterms:modified>
</cp:coreProperties>`,
			wantCreated:  "2024-01-01T00:00:00Z",
			wantModified:  "2024-12-31T23:59:59Z",
		},
		{
			name:      "错误-无效XML",
			xmlData:   `<invalid><unclosed>`,
			wantError: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cp, err := parts.ParseCoreProps([]byte(tt.xmlData))

			if tt.wantError {
				if err == nil {
					t.Error("期望返回错误，但解析成功")
				}
				return
			}

			if err != nil {
				t.Fatalf("解析失败: %v", err)
			}

			if cp == nil {
				t.Fatal("ParseCoreProps 返回 nil")
			}

			// 验证基础字段
			if cp.Title != tt.wantTitle {
				t.Errorf("Title = %q, want %q", cp.Title, tt.wantTitle)
			}
			if cp.Creator != tt.wantCreator {
				t.Errorf("Creator = %q, want %q", cp.Creator, tt.wantCreator)
			}
			if cp.Subject != tt.wantSubject {
				t.Errorf("Subject = %q, want %q", cp.Subject, tt.wantSubject)
			}
			if cp.Keywords != tt.wantKeywords {
				t.Errorf("Keywords = %q, want %q", cp.Keywords, tt.wantKeywords)
			}
			if cp.LastModifiedBy != tt.wantLastModBy {
				t.Errorf("LastModifiedBy = %q, want %q", cp.LastModifiedBy, tt.wantLastModBy)
			}
			if cp.Revision != tt.wantRevision {
				t.Errorf("Revision = %q, want %q", cp.Revision, tt.wantRevision)
			}
			if cp.Category != tt.wantCategory {
				t.Errorf("Category = %q, want %q", cp.Category, tt.wantCategory)
			}

			// 验证时间字段
			if gotCreated := cp.GetCreated(); gotCreated != tt.wantCreated {
				t.Errorf("Created = %q, want %q", gotCreated, tt.wantCreated)
			}
			if gotModified := cp.GetModified(); gotModified != tt.wantModified {
				t.Errorf("Modified = %q, want %q", gotModified, tt.wantModified)
			}
		})
	}
}

// ============================================================================
// Core Properties 序列化测试
// ============================================================================

func TestCorePropsToXML(t *testing.T) {
	t.Run("序列化-完整属性", func(t *testing.T) {
		cp := parts.NewXMLCoreProperties()
		cp.Title = "测试标题"
		cp.Creator = "测试作者"
		cp.Subject = "测试主题"
		cp.Keywords = "关键词1,关键词2"
		cp.LastModifiedBy = "修改者"
		cp.Revision = "2"
		cp.SetCreated("2024-01-15T10:30:00Z")
		cp.SetModified("2024-01-16T14:45:00Z")

		data, err := cp.ToXML()
		if err != nil {
			t.Fatalf("序列化失败: %v", err)
		}

		// 验证可以重新解析
		parsed, err := parts.ParseCoreProps(data)
		if err != nil {
			t.Fatalf("重新解析失败: %v", err)
		}

		if parsed.Title != cp.Title {
			t.Errorf("Title 不匹配: got %q, want %q", parsed.Title, cp.Title)
		}
		if parsed.Creator != cp.Creator {
			t.Errorf("Creator 不匹配: got %q, want %q", parsed.Creator, cp.Creator)
		}
		if parsed.GetCreated() != cp.GetCreated() {
			t.Errorf("Created 不匹配: got %q, want %q", parsed.GetCreated(), cp.GetCreated())
		}
	})

	t.Run("序列化-空属性", func(t *testing.T) {
		cp := parts.NewXMLCoreProperties()

		data, err := cp.ToXML()
		if err != nil {
			t.Fatalf("序列化失败: %v", err)
		}

		// 验证可以重新解析
		parsed, err := parts.ParseCoreProps(data)
		if err != nil {
			t.Fatalf("重新解析失败: %v", err)
		}

		if parsed.Title != "" {
			t.Errorf("空 Title 应为空字符串, got %q", parsed.Title)
		}
	})
}
