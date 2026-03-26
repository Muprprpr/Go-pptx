package utils_test

import (
	"os"
	"strings"
	"sync"
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

// ===== 并发 ID 分配测试 =====

func TestRelationships_ConcurrentIDAllocation(t *testing.T) {
	source := opc.NewPackURI("/ppt/presentation.xml")
	rels := opc.NewRelationships(source)

	// 并发添加关系
	const goroutines = 10
	const relationsPerGoroutine = 100

	var wg sync.WaitGroup
	rIDs := make([][]string, goroutines)

	for i := 0; i < goroutines; i++ {
		rIDs[i] = make([]string, 0, relationsPerGoroutine)
		wg.Add(1)
		go func(idx int) {
			defer wg.Done()
			for j := 0; j < relationsPerGoroutine; j++ {
				rel, err := rels.AddNew(opc.RelTypeSlide, "/ppt/slides/slide.xml", false)
				if err != nil {
					t.Errorf("goroutine %d: AddNew failed: %v", idx, err)
					return
				}
				rIDs[idx] = append(rIDs[idx], rel.RID())
			}
		}(i)
	}

	wg.Wait()

	// 验证总关系数
	if rels.Count() != goroutines*relationsPerGoroutine {
		t.Errorf("Count = %d, want %d", rels.Count(), goroutines*relationsPerGoroutine)
	}

	// 验证所有 ID 都是唯一的
	idSet := make(map[string]bool)
	for _, ids := range rIDs {
		for _, id := range ids {
			if idSet[id] {
				t.Errorf("duplicate rID found: %s", id)
			}
			idSet[id] = true
		}
	}

	// 验证 ID 格式正确
	for id := range idSet {
		if !strings.HasPrefix(id, "rId") {
			t.Errorf("invalid rID format: %s", id)
		}
	}
}

func TestRelationships_InitRIDCounter(t *testing.T) {
	source := opc.NewPackURI("/ppt/presentation.xml")

	// 从 XML 加载包含现有关系的集合
	xmlData := `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide3.xml"/>
</Relationships>`

	rels := opc.NewRelationships(source)
	err := rels.FromXML([]byte(xmlData))
	if err != nil {
		t.Fatalf("FromXML failed: %v", err)
	}

	// FromXML 会自动初始化计数器为最大值 5
	// NextRID 预览下一个（Load + 1），不消耗计数器
	nextRID := rels.NextRID()
	if nextRID != "rId6" {
		t.Errorf("NextRID after InitRIDCounter = %s, want rId6", nextRID)
	}

	// AddNew 分配下一个 ID（使用 Add，返回新值）
	// 第一个 AddNew 应该得到 rId6
	rel, err := rels.AddNew(opc.RelTypeSlide, "/ppt/slides/slide4.xml", false)
	if err != nil {
		t.Fatalf("AddNew failed: %v", err)
	}
	if rel.RID() != "rId6" {
		t.Errorf("new relationship RID = %s, want rId6", rel.RID())
	}

	// 第二个 AddNew 应该得到 rId7
	rel2, err := rels.AddNew(opc.RelTypeSlide, "/ppt/slides/slide5.xml", false)
	if err != nil {
		t.Fatalf("AddNew failed: %v", err)
	}
	if rel2.RID() != "rId7" {
		t.Errorf("second new relationship RID = %s, want rId7", rel2.RID())
	}
}

func TestRelationships_ClonePreservesCounter(t *testing.T) {
	source := opc.NewPackURI("/ppt/presentation.xml")
	rels := opc.NewRelationships(source)

	// 添加一些关系
	rels.AddNew(opc.RelTypeSlide, "/ppt/slides/slide1.xml", false)
	rels.AddNew(opc.RelTypeSlide, "/ppt/slides/slide2.xml", false)

	// 克隆
	cloned := rels.Clone()

	// 克隆后的 NextRID 应该与原始相同
	if rels.NextRID() != cloned.NextRID() {
		t.Errorf("cloned NextRID = %s, want %s", cloned.NextRID(), rels.NextRID())
	}

	// 在克隆中添加关系不应该影响原始
	cloned.AddNew(opc.RelTypeSlide, "/ppt/slides/slide3.xml", false)
	if rels.Count() != 2 {
		t.Errorf("original count changed after clone modification")
	}
}

// TestParseRelationshipsFromFile 测试从真实 .rels 文件解析媒体关系
func TestParseRelationshipsFromFile(t *testing.T) {
	// 读取真实的 .rels 文件
	data, err := os.ReadFile("../test-data/test/ppt/slides/_rels/slide4.xml.rels")
	if err != nil {
		t.Fatalf("读取 .rels 文件失败: %v", err)
	}

	// 反序列化为 Relationships 结构体
	source := opc.NewPackURI("/ppt/slides/slide4.xml")
	rels, err := opc.ParseRelationshipsFromXML(data, source)

	// 断言解析无 error 且结果不为 nil
	if err != nil {
		t.Fatalf("ParseRelationshipsFromXML 失败: %v", err)
	}
	if rels == nil {
		t.Fatal("ParseRelationshipsFromXML 返回 nil")
	}

	// 断言解析出的 Relationship 切片长度大于 0
	allRels := rels.All()
	if len(allRels) == 0 {
		t.Fatal("解析出的 Relationship 切片长度为 0")
	}
	t.Logf("共解析到 %d 个关系", len(allRels))

	// 遍历切片，找到 Type 包含 image 或 media 的节点
	var foundMediaRel bool
	for _, rel := range allRels {
		relType := rel.Type()
		// 检查类型是否包含 image 或 media
		if strings.Contains(relType, "image") || strings.Contains(relType, "media") {
			foundMediaRel = true

			// 断言 Id (rId) 非空
			rID := rel.RID()
			if rID == "" {
				t.Error("媒体关系的 Id (rId) 为空")
			} else {
				t.Logf("找到媒体关系: Id=%s", rID)
			}

			// 断言 Target 非空
			target := rel.TargetURI()
			if target == nil || target.URI() == "" {
				t.Error("媒体关系的 Target 为空")
			} else {
				t.Logf("  Target=%s", target.URI())
			}

			// 断言 Type 非空
			if relType == "" {
				t.Error("媒体关系的 Type 为空")
			} else {
				t.Logf("  Type=%s", relType)
			}
		}
	}

	// 确保至少找到一个媒体关系
	if !foundMediaRel {
		t.Error("未找到 Type 包含 image 或 media 的关系")
	}
}
