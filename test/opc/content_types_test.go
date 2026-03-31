package opc_test

import (
	"testing"

	"github.com/Muprprpr/Go-pptx/opc"
)

func TestContentTypes_New(t *testing.T) {
	ct := opc.NewContentTypes()
	if ct == nil {
		t.Fatal("NewContentTypes returned nil")
	}

	// 检查默认内容类型已初始化
	if ct.GetDefault(".xml") != opc.ContentTypeXML {
		t.Error("default .xml content type not set")
	}
	if ct.GetDefault(".rels") != opc.ContentTypeRelationships {
		t.Error("default .rels content type not set")
	}
}

func TestContentTypes_AddDefault(t *testing.T) {
	ct := opc.NewContentTypes()
	ct.AddDefault(".custom", "application/custom")

	if ct.GetDefault(".custom") != "application/custom" {
		t.Error("failed to add default content type")
	}
}

func TestContentTypes_AddOverride(t *testing.T) {
	ct := opc.NewContentTypes()
	uri := opc.NewPackURI("/ppt/presentation.xml")
	ct.AddOverride(uri, opc.ContentTypePresentation)

	if ct.GetOverride(uri) != opc.ContentTypePresentation {
		t.Error("failed to add override content type")
	}
}

func TestContentTypes_GetContentType(t *testing.T) {
	ct := opc.NewContentTypes()

	// 测试默认类型
	xmlURI := opc.NewPackURI("/ppt/slides/slide1.xml")
	if ct.GetContentType(xmlURI) != opc.ContentTypeXML {
		t.Error("failed to get default content type for .xml")
	}

	// 测试覆盖类型
	pptURI := opc.NewPackURI("/ppt/presentation.xml")
	ct.AddOverride(pptURI, opc.ContentTypePresentation)
	if ct.GetContentType(pptURI) != opc.ContentTypePresentation {
		t.Error("failed to get override content type")
	}

	// 测试未知扩展名
	unknownURI := opc.NewPackURI("/unknown/file.xyz")
	if ct.GetContentType(unknownURI) != opc.ContentTypeDefault {
		t.Error("unknown extension should return default content type")
	}
}

func TestContentTypes_RemoveOverride(t *testing.T) {
	ct := opc.NewContentTypes()
	uri := opc.NewPackURI("/ppt/presentation.xml")
	ct.AddOverride(uri, opc.ContentTypePresentation)

	ct.RemoveOverride(uri)
	if ct.GetOverride(uri) != "" {
		t.Error("failed to remove override")
	}
}

func TestContentTypes_Defaults(t *testing.T) {
	ct := opc.NewContentTypes()
	defaults := ct.Defaults()

	if len(defaults) == 0 {
		t.Error("defaults should not be empty")
	}

	// 确保返回的是副本
	defaults[".test"] = "test"
	if ct.GetDefault(".test") != "" {
		t.Error("modifying returned map should not affect original")
	}
}

func TestContentTypes_Overrides(t *testing.T) {
	ct := opc.NewContentTypes()
	uri := opc.NewPackURI("/ppt/presentation.xml")
	ct.AddOverride(uri, opc.ContentTypePresentation)

	overrides := ct.Overrides()
	if len(overrides) != 1 {
		t.Fatalf("expected 1 override, got %d", len(overrides))
	}

	// 确保返回的是副本
	overrides["/test"] = "test"
	if ct.GetOverride(opc.NewPackURI("/test")) != "" {
		t.Error("modifying returned map should not affect original")
	}
}

func TestContentTypes_XML(t *testing.T) {
	ct := opc.NewContentTypes()
	ct.AddDefault("custom", "application/custom") // 扩展名不带点
	uri := opc.NewPackURI("/ppt/presentation.xml")
	ct.AddOverride(uri, opc.ContentTypePresentation)

	// 序列化
	data, err := ct.ToXML()
	if err != nil {
		t.Fatalf("ToXML failed: %v", err)
	}

	// 反序列化
	ct2 := opc.NewContentTypes()
	err = ct2.FromXML(data)
	if err != nil {
		t.Fatalf("FromXML failed: %v", err)
	}

	// 验证 - FromXML 存储的是不带点的扩展名
	if ct2.GetDefault("custom") != "application/custom" {
		t.Error("custom default not preserved after XML round-trip")
	}
	if ct2.GetOverride(uri) != opc.ContentTypePresentation {
		t.Error("override not preserved after XML round-trip")
	}
}

func TestContentTypes_FromXML(t *testing.T) {
	xmlData := `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
</Types>`

	ct := &opc.ContentTypes{}
	err := ct.FromXML([]byte(xmlData))
	if err != nil {
		t.Fatalf("FromXML failed: %v", err)
	}

	// FromXML 存储的是不带点的扩展名
	if ct.GetDefault("xml") != opc.ContentTypeXML {
		t.Error("failed to parse Default element")
	}
	if ct.GetOverride(opc.NewPackURI("/ppt/presentation.xml")) != opc.ContentTypePresentation {
		t.Error("failed to parse Override element")
	}
}

func TestGetContentTypeByExtension(t *testing.T) {
	tests := []struct {
		ext      string
		expected string
	}{
		{".xml", opc.ContentTypeXML},
		{".png", opc.ContentTypePNG},
		{".jpg", opc.ContentTypeJPEG},
		{".rels", opc.ContentTypeRelationships},
		{".unknown", opc.ContentTypeDefault},
	}

	for _, tt := range tests {
		t.Run(tt.ext, func(t *testing.T) {
			result := opc.GetContentTypeByExtension(tt.ext)
			if result != tt.expected {
				t.Errorf("GetContentTypeByExtension(%q) = %q, want %q", tt.ext, result, tt.expected)
			}
		})
	}
}

func TestGetExtensionByContentType(t *testing.T) {
	tests := []struct {
		ct       string
		expected string
	}{
		{opc.ContentTypePNG, ".png"},
		{opc.ContentTypeJPEG, ".jpg"},
		{opc.ContentTypeXML, ".xml"},
		{"application/unknown", ".bin"},
	}

	for _, tt := range tests {
		t.Run(tt.ct, func(t *testing.T) {
			result := opc.GetExtensionByContentType(tt.ct)
			if result != tt.expected {
				t.Errorf("GetExtensionByContentType(%q) = %q, want %q", tt.ct, result, tt.expected)
			}
		})
	}
}
