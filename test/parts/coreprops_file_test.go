package parts_test

import (
	"os"
	"testing"

	"github.com/Muprprpr/Go-pptx/parts"
)

// ============================================================================
// CoreProperties 真实文件解析测试
// ============================================================================

func TestParseCorePropsFromFile(t *testing.T) {
	// 读取真实 core.xml 文件
	xmlData, err := os.ReadFile("../test-data/test/docProps/core.xml")
	if err != nil {
		t.Fatalf("读取文件失败: %v", err)
	}

	// 解析
	cp, err := parts.ParseCoreProps(xmlData)
	if err != nil {
		t.Fatalf("ParseCoreProps 返回错误: %v", err)
	}
	if cp == nil {
		t.Fatal("ParseCoreProps 返回 nil")
	}

	// 断言核心字段非空
	if cp.Title == "" {
		t.Error("Title 为空")
	}
	if cp.Creator == "" {
		t.Error("Creator 为空")
	}
	if cp.LastModifiedBy == "" {
		t.Error("LastModifiedBy 为空")
	}
	if cp.Revision == "" {
		t.Error("Revision 为空")
	}
	if cp.GetCreated() == "" {
		t.Error("Created 为空")
	}
	if cp.GetModified() == "" {
		t.Error("Modified 为空")
	}

	// 验证具体值
	if cp.Title != "PowerPoint 演示文稿" {
		t.Errorf("Title = %q, want %q", cp.Title, "PowerPoint 演示文稿")
	}
	if cp.Creator != "优品PPT" {
		t.Errorf("Creator = %q, want %q", cp.Creator, "优品PPT")
	}
	if cp.LastModifiedBy != "kan" {
		t.Errorf("LastModifiedBy = %q, want %q", cp.LastModifiedBy, "kan")
	}
	if cp.Revision != "91" {
		t.Errorf("Revision = %q, want %q", cp.Revision, "91")
	}
	if cp.GetCreated() != "2019-05-16T00:04:14Z" {
		t.Errorf("Created = %q, want %q", cp.GetCreated(), "2019-05-16T00:04:14Z")
	}
	if cp.GetModified() != "2022-05-30T10:23:18Z" {
		t.Errorf("Modified = %q, want %q", cp.GetModified(), "2022-05-30T10:23:18Z")
	}

	t.Logf("解析成功: Title=%q, Creator=%q, Revision=%q", cp.Title, cp.Creator, cp.Revision)
}
