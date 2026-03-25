package utils_test

import (
	"testing"

	"github.com/Muprprpr/Go-pptx/opc"
)

func TestRelationship_New(t *testing.T) {
	source := opc.NewPackURI("/ppt/presentation.xml")
	rel := opc.NewRelationship("rId1", opc.RelTypeSlide, "/ppt/slides/slide1.xml", false, source)

	if rel.RID() != "rId1" {
		t.Errorf("RID() = %q, want %q", rel.RID(), "rId1")
	}
	if rel.Type() != opc.RelTypeSlide {
		t.Errorf("Type() = %q, want %q", rel.Type(), opc.RelTypeSlide)
	}
	if rel.IsExternal() {
		t.Error("internal relationship should not be external")
	}
	if rel.TargetMode() != "Internal" {
		t.Errorf("TargetMode() = %q, want %q", rel.TargetMode(), "Internal")
	}
}

func TestRelationship_External(t *testing.T) {
	source := opc.NewPackURI("/ppt/slides/slide1.xml")
	rel := opc.NewRelationship("rId2", opc.RelTypeHyperlink, "http://example.com", true, source)

	if !rel.IsExternal() {
		t.Error("external relationship should be external")
	}
	if rel.TargetMode() != "External" {
		t.Errorf("TargetMode() = %q, want %q", rel.TargetMode(), "External")
	}
}

func TestRelationship_TargetURI(t *testing.T) {
	source := opc.NewPackURI("/ppt/presentation.xml")
	rel := opc.NewRelationship("rId1", opc.RelTypeSlide, "/ppt/slides/slide1.xml", false, source)

	target := rel.TargetURI()
	if target.URI() != "/ppt/slides/slide1.xml" {
		t.Errorf("TargetURI() = %q, want %q", target.URI(), "/ppt/slides/slide1.xml")
	}
}

func TestRelationship_TargetRef(t *testing.T) {
	source := opc.NewPackURI("/ppt/presentation.xml")
	rel := opc.NewRelationship("rId1", opc.RelTypeSlide, "/ppt/slides/slide1.xml", false, source)

	ref := rel.TargetRef()
	if ref == "" {
		t.Error("TargetRef should not be empty")
	}

	// 外部关系
	externalRel := opc.NewRelationship("rId2", opc.RelTypeHyperlink, "http://example.com", true, source)
	extRef := externalRel.TargetRef()
	if extRef != "http://example.com" {
		t.Errorf("external TargetRef = %q, want %q", extRef, "http://example.com")
	}
}

func TestRelationship_Equals(t *testing.T) {
	source := opc.NewPackURI("/ppt/presentation.xml")
	rel1 := opc.NewRelationship("rId1", opc.RelTypeSlide, "/ppt/slides/slide1.xml", false, source)
	rel2 := opc.NewRelationship("rId1", opc.RelTypeSlide, "/ppt/slides/slide1.xml", false, source)
	rel3 := opc.NewRelationship("rId2", opc.RelTypeSlide, "/ppt/slides/slide1.xml", false, source)

	if !rel1.Equals(rel2) {
		t.Error("identical relationships should be equal")
	}
	if rel1.Equals(rel3) {
		t.Error("different relationships should not be equal")
	}
	if rel1.Equals(nil) {
		t.Error("relationship should not equal nil")
	}
}

func TestRelationships_New(t *testing.T) {
	source := opc.NewPackURI("/ppt/presentation.xml")
	rels := opc.NewRelationships(source)

	if rels == nil {
		t.Fatal("NewRelationships returned nil")
	}
	if rels.Count() != 0 {
		t.Error("new relationships should be empty")
	}
}

func TestRelationships_Add(t *testing.T) {
	source := opc.NewPackURI("/ppt/presentation.xml")
	rels := opc.NewRelationships(source)
	rel := opc.NewRelationship("rId1", opc.RelTypeSlide, "/ppt/slides/slide1.xml", false, source)

	err := rels.Add(rel)
	if err != nil {
		t.Fatalf("Add failed: %v", err)
	}
	if rels.Count() != 1 {
		t.Errorf("Count() = %d, want 1", rels.Count())
	}

	// 添加重复的 rID 应该失败
	rel2 := opc.NewRelationship("rId1", opc.RelTypeSlide, "/ppt/slides/slide2.xml", false, source)
	err = rels.Add(rel2)
	if err == nil {
		t.Error("adding duplicate rID should fail")
	}
}

func TestRelationships_AddNew(t *testing.T) {
	source := opc.NewPackURI("/ppt/presentation.xml")
	rels := opc.NewRelationships(source)

	rel, err := rels.AddNew(opc.RelTypeSlide, "/ppt/slides/slide1.xml", false)
	if err != nil {
		t.Fatalf("AddNew failed: %v", err)
	}
	if rel.RID() != "rId1" {
		t.Errorf("first rID = %q, want %q", rel.RID(), "rId1")
	}

	rel2, err := rels.AddNew(opc.RelTypeSlide, "/ppt/slides/slide2.xml", false)
	if err != nil {
		t.Fatalf("AddNew failed: %v", err)
	}
	if rel2.RID() != "rId2" {
		t.Errorf("second rID = %q, want %q", rel2.RID(), "rId2")
	}
}

func TestRelationships_Get(t *testing.T) {
	source := opc.NewPackURI("/ppt/presentation.xml")
	rels := opc.NewRelationships(source)
	rel, _ := rels.AddNew(opc.RelTypeSlide, "/ppt/slides/slide1.xml", false)

	got := rels.Get("rId1")
	if got == nil {
		t.Fatal("Get returned nil")
	}
	if got.RID() != rel.RID() {
		t.Error("Get returned wrong relationship")
	}

	// 获取不存在的 rID
	if rels.Get("rId999") != nil {
		t.Error("Get for non-existent rID should return nil")
	}
}

func TestRelationships_GetByType(t *testing.T) {
	source := opc.NewPackURI("/ppt/presentation.xml")
	rels := opc.NewRelationships(source)
	rels.AddNew(opc.RelTypeSlide, "/ppt/slides/slide1.xml", false)
	rels.AddNew(opc.RelTypeSlide, "/ppt/slides/slide2.xml", false)
	rels.AddNew(opc.RelTypeTheme, "/ppt/theme/theme1.xml", false)

	slides := rels.GetByType(opc.RelTypeSlide)
	if len(slides) != 2 {
		t.Errorf("GetByType(slide) returned %d, want 2", len(slides))
	}

	themes := rels.GetByType(opc.RelTypeTheme)
	if len(themes) != 1 {
		t.Errorf("GetByType(theme) returned %d, want 1", len(themes))
	}

	images := rels.GetByType(opc.RelTypeImage)
	if len(images) != 0 {
		t.Error("GetByType for non-existent type should return empty")
	}
}

func TestRelationships_GetByTarget(t *testing.T) {
	source := opc.NewPackURI("/ppt/presentation.xml")
	rels := opc.NewRelationships(source)
	rels.AddNew(opc.RelTypeSlide, "/ppt/slides/slide1.xml", false)

	rel := rels.GetByTarget(opc.NewPackURI("/ppt/slides/slide1.xml"))
	if rel == nil {
		t.Fatal("GetByTarget returned nil")
	}

	// 获取不存在的目标
	if rels.GetByTarget(opc.NewPackURI("/ppt/slides/slide999.xml")) != nil {
		t.Error("GetByTarget for non-existent target should return nil")
	}
}

func TestRelationships_Remove(t *testing.T) {
	source := opc.NewPackURI("/ppt/presentation.xml")
	rels := opc.NewRelationships(source)
	rels.AddNew(opc.RelTypeSlide, "/ppt/slides/slide1.xml", false)

	err := rels.Remove("rId1")
	if err != nil {
		t.Fatalf("Remove failed: %v", err)
	}
	if rels.Count() != 0 {
		t.Error("relationship should be removed")
	}

	// 删除不存在的 rID 应该失败
	err = rels.Remove("rId999")
	if err == nil {
		t.Error("removing non-existent rID should fail")
	}
}

func TestRelationships_Contains(t *testing.T) {
	source := opc.NewPackURI("/ppt/presentation.xml")
	rels := opc.NewRelationships(source)
	rels.AddNew(opc.RelTypeSlide, "/ppt/slides/slide1.xml", false)

	if !rels.Contains("rId1") {
		t.Error("should contain rId1")
	}
	if rels.Contains("rId999") {
		t.Error("should not contain rId999")
	}
}

func TestRelationships_All(t *testing.T) {
	source := opc.NewPackURI("/ppt/presentation.xml")
	rels := opc.NewRelationships(source)
	rels.AddNew(opc.RelTypeSlide, "/ppt/slides/slide1.xml", false)
	rels.AddNew(opc.RelTypeSlide, "/ppt/slides/slide2.xml", false)

	all := rels.All()
	if len(all) != 2 {
		t.Errorf("All() returned %d, want 2", len(all))
	}
}

func TestRelationships_NextRID(t *testing.T) {
	source := opc.NewPackURI("/ppt/presentation.xml")
	rels := opc.NewRelationships(source)

	// 空集合应该返回 rId1
	if rels.NextRID() != "rId1" {
		t.Error("first NextRID should be rId1")
	}

	rels.AddNew(opc.RelTypeSlide, "/ppt/slides/slide1.xml", false)
	if rels.NextRID() != "rId2" {
		t.Error("second NextRID should be rId2")
	}
}

func TestRelationships_Clone(t *testing.T) {
	source := opc.NewPackURI("/ppt/presentation.xml")
	rels := opc.NewRelationships(source)
	rels.AddNew(opc.RelTypeSlide, "/ppt/slides/slide1.xml", false)

	clone := rels.Clone()
	if clone.Count() != rels.Count() {
		t.Error("clone should have same count")
	}

	// 修改克隆不应该影响原始
	clone.AddNew(opc.RelTypeSlide, "/ppt/slides/slide2.xml", false)
	if rels.Count() == clone.Count() {
		t.Error("modifying clone should not affect original")
	}
}

func TestRelationships_XML(t *testing.T) {
	source := opc.NewPackURI("/ppt/presentation.xml")
	rels := opc.NewRelationships(source)
	rels.AddNew(opc.RelTypeSlide, "/ppt/slides/slide1.xml", false)

	// 序列化
	data, err := rels.ToXML()
	if err != nil {
		t.Fatalf("ToXML failed: %v", err)
	}

	// 反序列化
	rels2 := opc.NewRelationships(source)
	err = rels2.FromXML(data)
	if err != nil {
		t.Fatalf("FromXML failed: %v", err)
	}

	if rels2.Count() != 1 {
		t.Errorf("Count after round-trip = %d, want 1", rels2.Count())
	}

	rel := rels2.Get("rId1")
	if rel == nil {
		t.Fatal("rId1 not found after round-trip")
	}
	if rel.Type() != opc.RelTypeSlide {
		t.Errorf("Type after round-trip = %q, want %q", rel.Type(), opc.RelTypeSlide)
	}
}

func TestRelationships_FromXML(t *testing.T) {
	xmlData := `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml" TargetMode="External"/>
</Relationships>`

	source := opc.NewPackURI("/ppt/presentation.xml")
	rels := opc.NewRelationships(source)
	err := rels.FromXML([]byte(xmlData))
	if err != nil {
		t.Fatalf("FromXML failed: %v", err)
	}

	if rels.Count() != 2 {
		t.Errorf("Count = %d, want 2", rels.Count())
	}

	rel1 := rels.Get("rId1")
	if rel1 == nil || rel1.IsExternal() {
		t.Error("rId1 should be internal")
	}

	rel2 := rels.Get("rId2")
	if rel2 == nil || !rel2.IsExternal() {
		t.Error("rId2 should be external")
	}
}

func TestParseRelationshipsFromXML(t *testing.T) {
	xmlData := `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`

	source := opc.NewPackURI("/ppt/presentation.xml")
	rels, err := opc.ParseRelationshipsFromXML([]byte(xmlData), source)
	if err != nil {
		t.Fatalf("ParseRelationshipsFromXML failed: %v", err)
	}

	if rels.Count() != 1 {
		t.Errorf("Count = %d, want 1", rels.Count())
	}
}
