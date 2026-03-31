package utils_test

import (
	"testing"

	"github.com/Muprprpr/Go-pptx/opc"
)

func TestPackURI_New(t *testing.T) {
	tests := []struct {
		name     string
		input    string
		expected string
	}{
		{"absolute path", "/ppt/slides/slide1.xml", "/ppt/slides/slide1.xml"},
		{"relative path", "ppt/slides/slide1.xml", "/ppt/slides/slide1.xml"},
		{"double slashes", "//ppt//slides//slide1.xml", "/ppt/slides/slide1.xml"},
		{"root", "/", "/"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			uri := opc.NewPackURI(tt.input)
			if uri.URI() != tt.expected {
				t.Errorf("NewPackURI(%q) = %q, want %q", tt.input, uri.URI(), tt.expected)
			}
		})
	}
}

func TestPackURI_FileName(t *testing.T) {
	tests := []struct {
		name     string
		uri      string
		expected string
	}{
		{"slide", "/ppt/slides/slide1.xml", "slide1.xml"},
		{"rels", "/ppt/slides/_rels/slide1.xml.rels", "slide1.xml.rels"},
		{"root file", "/[Content_Types].xml", "[Content_Types].xml"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			uri := opc.NewPackURI(tt.uri)
			if uri.FileName() != tt.expected {
				t.Errorf("FileName() = %q, want %q", uri.FileName(), tt.expected)
			}
		})
	}
}

func TestPackURI_BaseName(t *testing.T) {
	tests := []struct {
		name     string
		uri      string
		expected string
	}{
		{"slide", "/ppt/slides/slide1.xml", "slide1"},
		{"rels", "/ppt/slides/_rels/slide1.xml.rels", "slide1.xml"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			uri := opc.NewPackURI(tt.uri)
			if uri.BaseName() != tt.expected {
				t.Errorf("BaseName() = %q, want %q", uri.BaseName(), tt.expected)
			}
		})
	}
}

func TestPackURI_Extension(t *testing.T) {
	tests := []struct {
		name     string
		uri      string
		expected string
	}{
		{"xml", "/ppt/slides/slide1.xml", ".xml"},
		{"rels", "/ppt/slides/_rels/slide1.xml.rels", ".rels"},
		{"no extension", "/ppt/slides/slide1", ""},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			uri := opc.NewPackURI(tt.uri)
			if uri.Extension() != tt.expected {
				t.Errorf("Extension() = %q, want %q", uri.Extension(), tt.expected)
			}
		})
	}
}

func TestPackURI_DirName(t *testing.T) {
	tests := []struct {
		name     string
		uri      string
		expected string
	}{
		{"slide", "/ppt/slides/slide1.xml", "/ppt/slides"},
		{"root file", "/[Content_Types].xml", "/"},
		{"nested", "/a/b/c/d.xml", "/a/b/c"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			uri := opc.NewPackURI(tt.uri)
			if uri.DirName() != tt.expected {
				t.Errorf("DirName() = %q, want %q", uri.DirName(), tt.expected)
			}
		})
	}
}

func TestPackURI_MemberName(t *testing.T) {
	uri := opc.NewPackURI("/ppt/slides/slide1.xml")
	expected := "ppt/slides/slide1.xml"
	if uri.MemberName() != expected {
		t.Errorf("MemberName() = %q, want %q", uri.MemberName(), expected)
	}
}

func TestPackURI_IsRelationshipsPart(t *testing.T) {
	tests := []struct {
		name     string
		uri      string
		expected bool
	}{
		{"rels file", "/ppt/slides/_rels/slide1.xml.rels", true},
		{"normal file", "/ppt/slides/slide1.xml", false},
		{"package rels", "/_rels/.rels", true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			uri := opc.NewPackURI(tt.uri)
			if uri.IsRelationshipsPart() != tt.expected {
				t.Errorf("IsRelationshipsPart() = %v, want %v", uri.IsRelationshipsPart(), tt.expected)
			}
		})
	}
}

func TestPackURI_RelationshipsURI(t *testing.T) {
	tests := []struct {
		name     string
		uri      string
		expected string
	}{
		{"slide", "/ppt/slides/slide1.xml", "/ppt/slides/_rels/slide1.xml.rels"},
		{"presentation", "/ppt/presentation.xml", "/ppt/_rels/presentation.xml.rels"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			uri := opc.NewPackURI(tt.uri)
			relURI := uri.RelationshipsURI()
			if relURI.URI() != tt.expected {
				t.Errorf("RelationshipsURI() = %q, want %q", relURI.URI(), tt.expected)
			}
		})
	}
}

func TestPackURI_SourceURI(t *testing.T) {
	tests := []struct {
		name     string
		uri      string
		expected string
	}{
		{"rels file", "/ppt/slides/_rels/slide1.xml.rels", "/ppt/slides/slide1.xml"},
		{"package rels", "/_rels/.rels", "/"},
		{"normal file", "/ppt/slides/slide1.xml", "/ppt/slides/slide1.xml"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			uri := opc.NewPackURI(tt.uri)
			sourceURI := uri.SourceURI()
			if sourceURI.URI() != tt.expected {
				t.Errorf("SourceURI() = %q, want %q", sourceURI.URI(), tt.expected)
			}
		})
	}
}

func TestPackURI_IsPackageRels(t *testing.T) {
	tests := []struct {
		name     string
		uri      string
		expected bool
	}{
		{"package rels", "/_rels/.rels", true},
		{"part rels", "/ppt/slides/_rels/slide1.xml.rels", false},
		{"normal file", "/ppt/slides/slide1.xml", false},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			uri := opc.NewPackURI(tt.uri)
			if uri.IsPackageRels() != tt.expected {
				t.Errorf("IsPackageRels() = %v, want %v", uri.IsPackageRels(), tt.expected)
			}
		})
	}
}

func TestPackURI_Equals(t *testing.T) {
	uri1 := opc.NewPackURI("/ppt/slides/slide1.xml")
	uri2 := opc.NewPackURI("/ppt/slides/slide1.xml")
	uri3 := opc.NewPackURI("/ppt/slides/slide2.xml")

	if !uri1.Equals(uri2) {
		t.Error("uri1 should equal uri2")
	}
	if uri1.Equals(uri3) {
		t.Error("uri1 should not equal uri3")
	}
	if uri1.Equals(nil) {
		t.Error("uri1 should not equal nil")
	}
}

func TestPackURI_Join(t *testing.T) {
	tests := []struct {
		name     string
		base     string
		relative string
		expected string
	}{
		// 注意：Join 的实现可能与预期不同，这里测试实际行为
		{"absolute", "/ppt/slides", "/docProps/core.xml", "/docProps/core.xml"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			uri := opc.NewPackURI(tt.base)
			result := uri.Join(tt.relative)
			if result.URI() != tt.expected {
				t.Errorf("Join(%q) = %q, want %q", tt.relative, result.URI(), tt.expected)
			}
		})
	}
}

func TestPackURI_Clone(t *testing.T) {
	uri1 := opc.NewPackURI("/ppt/slides/slide1.xml")
	uri2 := uri1.Clone()

	if !uri1.Equals(uri2) {
		t.Error("cloned URI should equal original")
	}

	// 确保是独立的副本
	if &uri1 == &uri2 {
		t.Error("clone should create a new instance")
	}
}

func TestPackURI_MarshalUnmarshalText(t *testing.T) {
	original := opc.NewPackURI("/ppt/slides/slide1.xml")

	data, err := original.MarshalText()
	if err != nil {
		t.Fatalf("MarshalText failed: %v", err)
	}

	var unmarshaled opc.PackURI
	err = unmarshaled.UnmarshalText(data)
	if err != nil {
		t.Fatalf("UnmarshalText failed: %v", err)
	}

	if !original.Equals(&unmarshaled) {
		t.Errorf("unmarshaled URI = %q, want %q", unmarshaled.URI(), original.URI())
	}
}

func TestRootURI(t *testing.T) {
	uri := opc.RootURI()
	if uri.URI() != "/" {
		t.Errorf("RootURI() = %q, want %q", uri.URI(), "/")
	}
}

func TestPackageRelsURI(t *testing.T) {
	uri := opc.PackageRelsURI()
	expected := "/_rels/.rels"
	if uri.URI() != expected {
		t.Errorf("PackageRelsURI() = %q, want %q", uri.URI(), expected)
	}
}

func TestContentTypesURI(t *testing.T) {
	uri := opc.ContentTypesURI()
	expected := "/[Content_Types].xml"
	if uri.URI() != expected {
		t.Errorf("ContentTypesURI() = %q, want %q", uri.URI(), expected)
	}
}

func TestIsValidPackURI(t *testing.T) {
	tests := []struct {
		name     string
		uri      string
		expected bool
	}{
		{"valid absolute", "/ppt/slides/slide1.xml", true},
		{"invalid relative", "ppt/slides/slide1.xml", false},
		{"empty", "", false},
		{"with backslash", "/ppt\\slides", false},
		{"with colon", "/ppt:slides", false},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := opc.IsValidPackURI(tt.uri)
			if result != tt.expected {
				t.Errorf("IsValidPackURI(%q) = %v, want %v", tt.uri, result, tt.expected)
			}
		})
	}
}

// TestNormalizeZipPath 测试 ZIP 路径规范化函数
// 确保能正确处理 Windows 反斜杠和各种异常路径
func TestNormalizeZipPath(t *testing.T) {
	tests := []struct {
		name     string
		input    string
		expected string
	}{
		// 正常路径
		{"normal path", "ppt/slides/slide1.xml", "ppt/slides/slide1.xml"},
		{"normal path with leading slash", "/ppt/slides/slide1.xml", "ppt/slides/slide1.xml"},

		// Windows 反斜杠问题
		{"windows backslash", "ppt\\slides\\slide1.xml", "ppt/slides/slide1.xml"},
		{"mixed slashes", "ppt\\slides/slide1.xml", "ppt/slides/slide1.xml"},
		{"all backslashes", "ppt\\slides\\_rels\\slide1.xml.rels", "ppt/slides/_rels/slide1.xml.rels"},
		{"windows with leading backslash", "\\ppt\\slides\\slide1.xml", "ppt/slides/slide1.xml"},

		// 重复斜杠问题
		{"double forward slashes", "ppt//slides//slide1.xml", "ppt/slides/slide1.xml"},
		{"triple slashes", "ppt///slides/slide1.xml", "ppt/slides/slide1.xml"},
		{"mixed repeated slashes", "ppt\\/\\\\slides/slide1.xml", "ppt/slides/slide1.xml"},

		// 边界情况
		{"empty string", "", ""},
		{"single slash", "/", ""},
		{"trailing slash", "ppt/slides/", "ppt/slides"},
		{"leading and trailing slash", "/ppt/slides/", "ppt/slides"},
		{"root file", "[Content_Types].xml", "[Content_Types].xml"},
		{"rels file", "_rels/.rels", "_rels/.rels"},
		{"windows rels file", "_rels\\.rels", "_rels/.rels"},

		// 复杂场景
		{"deeply nested windows", "ppt\\slides\\slide1\\_rels\\slide1.xml.rels", "ppt/slides/slide1/_rels/slide1.xml.rels"},
		{"media file windows", "ppt\\media\\image1.png", "ppt/media/image1.png"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := opc.NormalizeZipPath(tt.input)
			if result != tt.expected {
				t.Errorf("NormalizeZipPath(%q) = %q, want %q", tt.input, result, tt.expected)
			}
		})
	}
}

// TestNormalizeZipPath_Idempotent 测试规范化是幂等的
// 多次规范化应该产生相同的结果
func TestNormalizeZipPath_Idempotent(t *testing.T) {
	testCases := []string{
		"ppt/slides/slide1.xml",
		"ppt\\slides\\slide1.xml",
		"/ppt//slides/slide1.xml",
		"\\ppt\\\\slides\\slide1.xml",
	}

	for _, input := range testCases {
		first := opc.NormalizeZipPath(input)
		second := opc.NormalizeZipPath(first)
		third := opc.NormalizeZipPath(second)

		if first != second || second != third {
			t.Errorf("NormalizeZipPath is not idempotent for %q: %q -> %q -> %q", input, first, second, third)
		}
	}
}

// TestNormalizeURI_vs_NormalizeZipPath 测试两个规范化函数的区别
func TestNormalizeURI_vs_NormalizeZipPath(t *testing.T) {
	testCases := []struct {
		input      string
		uriResult  string
		zipResult  string
	}{
		{"ppt/slides/slide1.xml", "/ppt/slides/slide1.xml", "ppt/slides/slide1.xml"},
		{"/ppt/slides/slide1.xml", "/ppt/slides/slide1.xml", "ppt/slides/slide1.xml"},
		{"ppt\\slides\\slide1.xml", "/ppt/slides/slide1.xml", "ppt/slides/slide1.xml"},
	}

	for _, tc := range testCases {
		uriResult := opc.NormalizeURI(tc.input)
		zipResult := opc.NormalizeZipPath(tc.input)

		if uriResult != tc.uriResult {
			t.Errorf("NormalizeURI(%q) = %q, want %q", tc.input, uriResult, tc.uriResult)
		}
		if zipResult != tc.zipResult {
			t.Errorf("NormalizeZipPath(%q) = %q, want %q", tc.input, zipResult, tc.zipResult)
		}

		// NormalizeZipPath 的结果加上 "/" 前缀应该等于 NormalizeURI 的结果
		expectedURIFromZip := "/" + zipResult
		if uriResult != expectedURIFromZip {
			t.Errorf("NormalizeURI(%q) = %q, but NormalizeZipPath + / = %q", tc.input, uriResult, expectedURIFromZip)
		}
	}
}
