package utils_test

import (
	"strings"
	"testing"

	"github.com/Muprprpr/Go-pptx/opc"
)

func TestPart_New(t *testing.T) {
	uri := opc.NewPackURI("/ppt/slides/slide1.xml")
	blob := []byte("<slide/>")
	part := opc.NewPart(uri, opc.ContentTypeSlide, blob)

	if part == nil {
		t.Fatal("NewPart returned nil")
	}
	if part.PartURI().URI() != uri.URI() {
		t.Errorf("PartURI() = %q, want %q", part.PartURI().URI(), uri.URI())
	}
	if part.ContentType() != opc.ContentTypeSlide {
		t.Errorf("ContentType() = %q, want %q", part.ContentType(), opc.ContentTypeSlide)
	}
	if string(part.Blob()) != string(blob) {
		t.Error("Blob() does not match original")
	}
	if !part.IsDirty() {
		t.Error("new part should be dirty")
	}
}

func TestPart_NewPartFromReader(t *testing.T) {
	uri := opc.NewPackURI("/ppt/slides/slide1.xml")
	reader := strings.NewReader("<slide/>")
	part, err := opc.NewPartFromReader(uri, opc.ContentTypeSlide, reader)

	if err != nil {
		t.Fatalf("NewPartFromReader failed: %v", err)
	}
	if part == nil {
		t.Fatal("NewPartFromReader returned nil")
	}
	if string(part.Blob()) != "<slide/>" {
		t.Errorf("Blob() = %q, want %q", string(part.Blob()), "<slide/>")
	}
}

func TestPart_SetContentType(t *testing.T) {
	uri := opc.NewPackURI("/ppt/slides/slide1.xml")
	part := opc.NewPart(uri, opc.ContentTypeSlide, []byte{})
	part.SetDirty(false)

	part.SetContentType(opc.ContentTypeSlideLayout)
	if part.ContentType() != opc.ContentTypeSlideLayout {
		t.Error("SetContentType failed")
	}
	if !part.IsDirty() {
		t.Error("SetContentType should mark part as dirty")
	}
}

func TestPart_SetBlob(t *testing.T) {
	uri := opc.NewPackURI("/ppt/slides/slide1.xml")
	part := opc.NewPart(uri, opc.ContentTypeSlide, []byte{})
	part.SetDirty(false)

	newBlob := []byte("<newSlide/>")
	part.SetBlob(newBlob)
	if string(part.Blob()) != string(newBlob) {
		t.Error("SetBlob failed")
	}
	if !part.IsDirty() {
		t.Error("SetBlob should mark part as dirty")
	}
}

func TestPart_SetBlobFromReader(t *testing.T) {
	uri := opc.NewPackURI("/ppt/slides/slide1.xml")
	part := opc.NewPart(uri, opc.ContentTypeSlide, []byte{})
	part.SetDirty(false)

	reader := strings.NewReader("<newSlide/>")
	err := part.SetBlobFromReader(reader)
	if err != nil {
		t.Fatalf("SetBlobFromReader failed: %v", err)
	}
	if string(part.Blob()) != "<newSlide/>" {
		t.Error("SetBlobFromReader failed")
	}
	if !part.IsDirty() {
		t.Error("SetBlobFromReader should mark part as dirty")
	}
}

func TestPart_Reader(t *testing.T) {
	uri := opc.NewPackURI("/ppt/slides/slide1.xml")
	blob := []byte("<slide/>")
	part := opc.NewPart(uri, opc.ContentTypeSlide, blob)

	reader := part.Reader()
	buf := make([]byte, len(blob))
	n, err := reader.Read(buf)
	if err != nil {
		t.Fatalf("Read failed: %v", err)
	}
	if n != len(blob) {
		t.Errorf("Read returned %d bytes, want %d", n, len(blob))
	}
}

func TestPart_Size(t *testing.T) {
	uri := opc.NewPackURI("/ppt/slides/slide1.xml")
	blob := []byte("<slide/>")
	part := opc.NewPart(uri, opc.ContentTypeSlide, blob)

	if part.Size() != len(blob) {
		t.Errorf("Size() = %d, want %d", part.Size(), len(blob))
	}
}

func TestPart_SetDirty(t *testing.T) {
	uri := opc.NewPackURI("/ppt/slides/slide1.xml")
	part := opc.NewPart(uri, opc.ContentTypeSlide, []byte{})

	part.SetDirty(false)
	if part.IsDirty() {
		t.Error("SetDirty(false) failed")
	}

	part.SetDirty(true)
	if !part.IsDirty() {
		t.Error("SetDirty(true) failed")
	}
}

func TestPart_Relationships(t *testing.T) {
	uri := opc.NewPackURI("/ppt/slides/slide1.xml")
	part := opc.NewPart(uri, opc.ContentTypeSlide, []byte{})

	rels := part.Relationships()
	if rels == nil {
		t.Fatal("Relationships() returned nil")
	}

	// 添加关系
	rel, err := part.AddRelationship(opc.RelTypeImage, "../media/image1.png", false)
	if err != nil {
		t.Fatalf("AddRelationship failed: %v", err)
	}
	if rel == nil {
		t.Fatal("AddRelationship returned nil")
	}
	if !part.IsDirty() {
		t.Error("AddRelationship should mark part as dirty")
	}
}

func TestPart_RemoveRelationship(t *testing.T) {
	uri := opc.NewPackURI("/ppt/slides/slide1.xml")
	part := opc.NewPart(uri, opc.ContentTypeSlide, []byte{})
	part.AddRelationship(opc.RelTypeImage, "../media/image1.png", false)
	part.SetDirty(false)

	err := part.RemoveRelationship("rId1")
	if err != nil {
		t.Fatalf("RemoveRelationship failed: %v", err)
	}
	if !part.IsDirty() {
		t.Error("RemoveRelationship should mark part as dirty")
	}
	if part.HasRelationships() {
		t.Error("part should have no relationships after removal")
	}
}

func TestPart_GetRelatedPart(t *testing.T) {
	uri := opc.NewPackURI("/ppt/slides/slide1.xml")
	part := opc.NewPart(uri, opc.ContentTypeSlide, []byte{})
	part.AddRelationship(opc.RelTypeImage, "../media/image1.png", false)

	targetURI := part.GetRelatedPart("rId1")
	if targetURI == nil {
		t.Fatal("GetRelatedPart returned nil")
	}

	// 获取不存在的关系
	if part.GetRelatedPart("rId999") != nil {
		t.Error("GetRelatedPart for non-existent rID should return nil")
	}
}

func TestPart_HasRelationships(t *testing.T) {
	uri := opc.NewPackURI("/ppt/slides/slide1.xml")
	part := opc.NewPart(uri, opc.ContentTypeSlide, []byte{})

	if part.HasRelationships() {
		t.Error("new part should have no relationships")
	}

	part.AddRelationship(opc.RelTypeImage, "../media/image1.png", false)
	if !part.HasRelationships() {
		t.Error("part should have relationships after adding")
	}
}

func TestPart_RelationshipsURI(t *testing.T) {
	uri := opc.NewPackURI("/ppt/slides/slide1.xml")
	part := opc.NewPart(uri, opc.ContentTypeSlide, []byte{})

	relURI := part.RelationshipsURI()
	expected := "/ppt/slides/_rels/slide1.xml.rels"
	if relURI.URI() != expected {
		t.Errorf("RelationshipsURI() = %q, want %q", relURI.URI(), expected)
	}
}

func TestPart_XML(t *testing.T) {
	uri := opc.NewPackURI("/ppt/slides/slide1.xml")
	xmlData := []byte(`<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`)
	part := opc.NewPart(uri, opc.ContentTypeSlide, xmlData)

	// 测试 UnmarshalBlob
	var slide struct {
		XMLName struct{} `xml:"sld"`
	}
	err := part.UnmarshalBlob(&slide)
	if err != nil {
		t.Fatalf("UnmarshalBlob failed: %v", err)
	}

	// 测试 MarshalToBlob
	type newSlide struct {
		XMLName struct{} `xml:"sld"`
	}
	part.SetDirty(false)
	err = part.MarshalToBlob(&newSlide{})
	if err != nil {
		t.Fatalf("MarshalToBlob failed: %v", err)
	}
	if !part.IsDirty() {
		t.Error("MarshalToBlob should mark part as dirty")
	}
}

func TestPart_Clone(t *testing.T) {
	uri := opc.NewPackURI("/ppt/slides/slide1.xml")
	blob := []byte("<slide/>")
	part := opc.NewPart(uri, opc.ContentTypeSlide, blob)
	part.AddRelationship(opc.RelTypeImage, "../media/image1.png", false)

	clone := part.Clone()
	if clone == nil {
		t.Fatal("Clone returned nil")
	}

	// 验证克隆是独立的
	if clone == part {
		t.Error("clone should be a different instance")
	}

	// 修改原始不应该影响克隆
	part.SetBlob([]byte("<modified/>"))
	if string(clone.Blob()) != "<slide/>" {
		t.Error("modifying original should not affect clone")
	}
}

// ===== PartCollection 测试 =====

func TestPartCollection_New(t *testing.T) {
	pc := opc.NewPartCollection()
	if pc == nil {
		t.Fatal("NewPartCollection returned nil")
	}
	if pc.Count() != 0 {
		t.Error("new collection should be empty")
	}
}

func TestPartCollection_Add(t *testing.T) {
	pc := opc.NewPartCollection()
	uri := opc.NewPackURI("/ppt/slides/slide1.xml")
	part := opc.NewPart(uri, opc.ContentTypeSlide, []byte{})

	err := pc.Add(part)
	if err != nil {
		t.Fatalf("Add failed: %v", err)
	}
	if pc.Count() != 1 {
		t.Errorf("Count() = %d, want 1", pc.Count())
	}

	// 添加重复 URI 应该失败
	duplicatePart := opc.NewPart(uri, opc.ContentTypeSlide, []byte{})
	err = pc.Add(duplicatePart)
	if err == nil {
		t.Error("adding duplicate URI should fail")
	}
}

func TestPartCollection_Get(t *testing.T) {
	pc := opc.NewPartCollection()
	uri := opc.NewPackURI("/ppt/slides/slide1.xml")
	part := opc.NewPart(uri, opc.ContentTypeSlide, []byte{})
	pc.Add(part)

	got := pc.Get(uri)
	if got == nil {
		t.Fatal("Get returned nil")
	}
	if got.PartURI().URI() != uri.URI() {
		t.Error("Get returned wrong part")
	}

	// 获取不存在的部件
	nonExistent := opc.NewPackURI("/ppt/slides/slide999.xml")
	if pc.Get(nonExistent) != nil {
		t.Error("Get for non-existent URI should return nil")
	}
}

func TestPartCollection_GetByStr(t *testing.T) {
	pc := opc.NewPartCollection()
	uri := opc.NewPackURI("/ppt/slides/slide1.xml")
	part := opc.NewPart(uri, opc.ContentTypeSlide, []byte{})
	pc.Add(part)

	got := pc.GetByStr("/ppt/slides/slide1.xml")
	if got == nil {
		t.Fatal("GetByStr returned nil")
	}
}

func TestPartCollection_Remove(t *testing.T) {
	pc := opc.NewPartCollection()
	uri := opc.NewPackURI("/ppt/slides/slide1.xml")
	part := opc.NewPart(uri, opc.ContentTypeSlide, []byte{})
	pc.Add(part)

	err := pc.Remove(uri)
	if err != nil {
		t.Fatalf("Remove failed: %v", err)
	}
	if pc.Count() != 0 {
		t.Error("part should be removed")
	}

	// 删除不存在的部件应该失败
	err = pc.Remove(uri)
	if err == nil {
		t.Error("removing non-existent part should fail")
	}
}

func TestPartCollection_Contains(t *testing.T) {
	pc := opc.NewPartCollection()
	uri := opc.NewPackURI("/ppt/slides/slide1.xml")
	part := opc.NewPart(uri, opc.ContentTypeSlide, []byte{})
	pc.Add(part)

	if !pc.Contains(uri) {
		t.Error("should contain added part")
	}

	nonExistent := opc.NewPackURI("/ppt/slides/slide999.xml")
	if pc.Contains(nonExistent) {
		t.Error("should not contain non-existent part")
	}
}

func TestPartCollection_All(t *testing.T) {
	pc := opc.NewPartCollection()
	uri1 := opc.NewPackURI("/ppt/slides/slide1.xml")
	uri2 := opc.NewPackURI("/ppt/slides/slide2.xml")
	pc.Add(opc.NewPart(uri1, opc.ContentTypeSlide, []byte{}))
	pc.Add(opc.NewPart(uri2, opc.ContentTypeSlide, []byte{}))

	all := pc.All()
	if len(all) != 2 {
		t.Errorf("All() returned %d parts, want 2", len(all))
	}
}

func TestPartCollection_URIs(t *testing.T) {
	pc := opc.NewPartCollection()
	uri1 := opc.NewPackURI("/ppt/slides/slide1.xml")
	uri2 := opc.NewPackURI("/ppt/slides/slide2.xml")
	pc.Add(opc.NewPart(uri1, opc.ContentTypeSlide, []byte{}))
	pc.Add(opc.NewPart(uri2, opc.ContentTypeSlide, []byte{}))

	uris := pc.URIs()
	if len(uris) != 2 {
		t.Errorf("URIs() returned %d URIs, want 2", len(uris))
	}
}

func TestPartCollection_GetByType(t *testing.T) {
	pc := opc.NewPartCollection()
	uri1 := opc.NewPackURI("/ppt/slides/slide1.xml")
	uri2 := opc.NewPackURI("/ppt/slides/slide2.xml")
	uri3 := opc.NewPackURI("/ppt/theme/theme1.xml")
	pc.Add(opc.NewPart(uri1, opc.ContentTypeSlide, []byte{}))
	pc.Add(opc.NewPart(uri2, opc.ContentTypeSlide, []byte{}))
	pc.Add(opc.NewPart(uri3, opc.ContentTypeTheme, []byte{}))

	slides := pc.GetByType(opc.ContentTypeSlide)
	if len(slides) != 2 {
		t.Errorf("GetByType(slide) returned %d, want 2", len(slides))
	}

	themes := pc.GetByType(opc.ContentTypeTheme)
	if len(themes) != 1 {
		t.Errorf("GetByType(theme) returned %d, want 1", len(themes))
	}
}

func TestPartCollection_Clear(t *testing.T) {
	pc := opc.NewPartCollection()
	uri := opc.NewPackURI("/ppt/slides/slide1.xml")
	pc.Add(opc.NewPart(uri, opc.ContentTypeSlide, []byte{}))

	pc.Clear()
	if pc.Count() != 0 {
		t.Error("Clear should remove all parts")
	}
}

func TestPartCollection_DirtyParts(t *testing.T) {
	pc := opc.NewPartCollection()
	uri1 := opc.NewPackURI("/ppt/slides/slide1.xml")
	uri2 := opc.NewPackURI("/ppt/slides/slide2.xml")

	part1 := opc.NewPart(uri1, opc.ContentTypeSlide, []byte{})
	part2 := opc.NewPart(uri2, opc.ContentTypeSlide, []byte{})
	pc.Add(part1)
	pc.Add(part2)

	// 所有新部件都是 dirty
	dirty := pc.DirtyParts()
	if len(dirty) != 2 {
		t.Errorf("DirtyParts() returned %d, want 2", len(dirty))
	}

	// 清除 dirty 标记
	part1.SetDirty(false)
	part2.SetDirty(false)

	dirty = pc.DirtyParts()
	if len(dirty) != 0 {
		t.Errorf("DirtyParts() after clearing should be 0, got %d", len(dirty))
	}
}

func TestDefaultPartFactory(t *testing.T) {
	factory := &opc.DefaultPartFactory{}
	uri := opc.NewPackURI("/ppt/slides/slide1.xml")
	blob := []byte("<slide/>")

	part, err := factory.CreatePart(uri, opc.ContentTypeSlide, blob)
	if err != nil {
		t.Fatalf("CreatePart failed: %v", err)
	}
	if part == nil {
		t.Fatal("CreatePart returned nil")
	}
}
