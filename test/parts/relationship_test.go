package parts_test

import (
	"testing"

	"github.com/Muprprpr/Go-pptx/parts"
)

// ============================================================================
// Relationships 解析测试
// ============================================================================

func TestParseRelationships(t *testing.T) {
	tests := []struct {
		name           string
		xmlData        string
		wantCount      int
		wantRID1Type   string
		wantRID1Target string
		wantError      bool
	}{
		{
			name: "正常解析-图片和视频关系",
			xmlData: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="../media/media1.mp4"/>
</Relationships>`,
			wantCount:      2,
			wantRID1Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
			wantRID1Target: "../media/image1.png",
		},
		{
			name: "正常解析-外部链接",
			xmlData: `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>`,
			wantCount:      1,
			wantRID1Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
			wantRID1Target: "https://example.com",
		},
		{
			name: "边界-空关系集合",
			xmlData: `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`,
			wantCount: 0,
		},
		{
			name:      "错误-无效XML",
			xmlData:   `<invalid><unclosed>`,
			wantError: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			rs, err := parts.ParseRelationships([]byte(tt.xmlData))

			if tt.wantError {
				if err == nil {
					t.Error("期望返回错误，但解析成功")
				}
				return
			}

			if err != nil {
				t.Fatalf("解析失败: %v", err)
			}

			if rs == nil {
				t.Fatal("ParseRelationships 返回 nil")
			}

			if rs.Count() != tt.wantCount {
				t.Errorf("Count = %d, want %d", rs.Count(), tt.wantCount)
			}

			if tt.wantCount > 0 && tt.wantRID1Type != "" {
				rel := rs.GetByID("rId1")
				if rel == nil {
					t.Fatal("GetByID(\"rId1\") 返回 nil")
				}
				if rel.Type != tt.wantRID1Type {
					t.Errorf("Type = %q, want %q", rel.Type, tt.wantRID1Type)
				}
				if rel.Target != tt.wantRID1Target {
					t.Errorf("Target = %q, want %q", rel.Target, tt.wantRID1Target)
				}
			}
		})
	}
}

// ============================================================================
// Relationship 辅助方法测试
// ============================================================================

func TestXMLRelationshipsMethods(t *testing.T) {
	t.Run("GetByType", func(t *testing.T) {
		xmlData := `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="../media/media1.mp4"/>
</Relationships>`

		rs, err := parts.ParseRelationships([]byte(xmlData))
		if err != nil {
			t.Fatalf("解析失败: %v", err)
		}

		images := rs.GetByType(parts.RelTypeImage)
		if len(images) != 2 {
			t.Errorf("GetByType(image) 返回 %d 个, want 2", len(images))
		}

		videos := rs.GetByType(parts.RelTypeMedia)
		if len(videos) != 1 {
			t.Errorf("GetByType(video) 返回 %d 个, want 1", len(videos))
		}
	})

	t.Run("GetByTarget", func(t *testing.T) {
		xmlData := `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>`

		rs, err := parts.ParseRelationships([]byte(xmlData))
		if err != nil {
			t.Fatalf("解析失败: %v", err)
		}

		rel := rs.GetByTarget("../media/image1.png")
		if rel == nil {
			t.Fatal("GetByTarget 返回 nil")
		}
		if rel.ID != "rId1" {
			t.Errorf("ID = %q, want %q", rel.ID, "rId1")
		}
	})

	t.Run("IsExternal", func(t *testing.T) {
		xmlData := `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>`

		rs, err := parts.ParseRelationships([]byte(xmlData))
		if err != nil {
			t.Fatalf("解析失败: %v", err)
		}

		external := rs.GetByID("rId1")
		if !external.IsExternal() {
			t.Error("rId1 应为外部链接")
		}

		internal := rs.GetByID("rId2")
		if internal.IsExternal() {
			t.Error("rId2 应为内部链接")
		}
	})
}
