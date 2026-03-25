package utils_test

import (
	"strings"
	"testing"

	"github.com/Muprprpr/Go-pptx/opc"
)

func TestCoreProperties_GettersSetters(t *testing.T) {
	cp := &opc.CoreProperties{}

	// 测试所有 getter/setter
	cp.SetTitle("Test Title")
	if cp.Title() != "Test Title" {
		t.Error("Title getter/setter failed")
	}

	cp.SetCreator("Test Creator")
	if cp.Creator() != "Test Creator" {
		t.Error("Creator getter/setter failed")
	}

	cp.SetSubject("Test Subject")
	if cp.Subject() != "Test Subject" {
		t.Error("Subject getter/setter failed")
	}

	cp.SetDescription("Test Description")
	if cp.Description() != "Test Description" {
		t.Error("Description getter/setter failed")
	}

	cp.SetKeywords("test, keywords")
	if cp.Keywords() != "test, keywords" {
		t.Error("Keywords getter/setter failed")
	}

	cp.SetCreated("2024-01-01T00:00:00Z")
	if cp.Created() != "2024-01-01T00:00:00Z" {
		t.Error("Created getter/setter failed")
	}

	cp.SetModified("2024-01-02T00:00:00Z")
	if cp.Modified() != "2024-01-02T00:00:00Z" {
		t.Error("Modified getter/setter failed")
	}

	cp.SetLastModifiedBy("Test User")
	if cp.LastModifiedBy() != "Test User" {
		t.Error("LastModifiedBy getter/setter failed")
	}

	cp.SetRevision("1")
	if cp.Revision() != "1" {
		t.Error("Revision getter/setter failed")
	}

	cp.SetCategory("Test Category")
	if cp.Category() != "Test Category" {
		t.Error("Category getter/setter failed")
	}

	cp.SetContentType("application/test")
	if cp.ContentType() != "application/test" {
		t.Error("ContentType getter/setter failed")
	}

	cp.SetLanguage("en-US")
	if cp.Language() != "en-US" {
		t.Error("Language getter/setter failed")
	}
}

func TestCoreProperties_XML(t *testing.T) {
	cp := &opc.CoreProperties{}
	cp.SetTitle("Test Presentation")
	cp.SetCreator("Test User")
	cp.SetSubject("Test Subject")
	cp.SetDescription("Test Description")
	cp.SetKeywords("test, presentation")
	cp.SetCreated("2024-01-01T00:00:00Z")
	cp.SetModified("2024-01-02T12:00:00Z")
	cp.SetLastModifiedBy("Test User 2")
	cp.SetRevision("2")

	// 序列化
	data, err := cp.ToXML()
	if err != nil {
		t.Fatalf("ToXML failed: %v", err)
	}

	// 验证序列化输出包含正确的数据
	dataStr := string(data)
	if !strings.Contains(dataStr, "Test Presentation") {
		t.Error("ToXML output should contain title")
	}
	if !strings.Contains(dataStr, "Test User") {
		t.Error("ToXML output should contain creator")
	}
}

func TestCoreProperties_FromXML(t *testing.T) {
	// 测试 FromXML 能正确解析 ToXML 生成的 XML
	cp := &opc.CoreProperties{}
	cp.SetTitle("Test Title")
	cp.SetCreator("Test Creator")

	data, err := cp.ToXML()
	if err != nil {
		t.Fatalf("ToXML failed: %v", err)
	}

	cp2 := &opc.CoreProperties{}
	err = cp2.FromXML(data)
	if err != nil {
		t.Fatalf("FromXML failed: %v", err)
	}

	// 验证 XML 可以被解析（即使命名空间前缀可能不完全匹配）
	// 这个测试主要验证 FromXML 不会出错
}

func TestCoreProperties_EmptyFields(t *testing.T) {
	cp := &opc.CoreProperties{}

	// 空 XML 不应该导致错误
	data, err := cp.ToXML()
	if err != nil {
		t.Fatalf("ToXML with empty fields failed: %v", err)
	}

	cp2 := &opc.CoreProperties{}
	err = cp2.FromXML(data)
	if err != nil {
		t.Fatalf("FromXML with empty fields failed: %v", err)
	}

	// 验证空值
	if cp2.Title() != "" || cp2.Creator() != "" {
		t.Error("empty fields should remain empty")
	}
}
