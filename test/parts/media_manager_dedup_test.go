package parts_test

import (
	"testing"

	"github.com/Muprprpr/Go-pptx/slide"
)

// ============================================================================
// MediaManager 去重测试
// ============================================================================

func TestMediaManager_Deduplication(t *testing.T) {
	mgr := slide.NewMediaManager()

	// 第一次添加
	rID1, res1 := mgr.AddMediaAuto("company_logo.png", []byte("logo_data"))
	if rID1 != "rId1" {
		t.Errorf("第一次添加 rID = %q, want rId1", rID1)
	}

	// 第二次添加相同文件名和相同数据
	rID2, res2 := mgr.AddMediaAuto("company_logo.png", []byte("logo_data"))

	// 验证：第二次应该返回相同的 rID
	if rID2 != rID1 {
		t.Errorf("重复添加 rID = %q, want %q (应复用已有资源)", rID2, rID1)
	}

	// 验证：返回的是同一个资源对象
	if res2 != res1 {
		t.Error("重复添加应返回相同的资源对象")
	}

	// 验证：资源池总数应该为 1
	if mgr.Count() != 1 {
		t.Errorf("Count = %d, want 1 (去重失败)", mgr.Count())
	}
}

func TestMediaManager_Deduplication_Multiple(t *testing.T) {
	mgr := slide.NewMediaManager()

	// 模拟 PPT 中每页都有相同的 logo
	const pages = 100
	for i := 0; i < pages; i++ {
		mgr.AddMediaAuto("company_logo.png", []byte("logo_data"))
	}

	// 验证：只存储了一份
	if mgr.Count() != 1 {
		t.Errorf("添加 %d 次相同文件后 Count = %d, want 1", pages, mgr.Count())
	}
}

func TestMediaManager_Deduplication_DifferentFiles(t *testing.T) {
	mgr := slide.NewMediaManager()

	// 添加不同文件
	rID1, _ := mgr.AddMediaAuto("logo.png", []byte("data1"))
	rID2, _ := mgr.AddMediaAuto("banner.png", []byte("data2"))
	rID3, _ := mgr.AddMediaAuto("icon.png", []byte("data3"))

	// 验证：不同文件应该有不同的 rID
	if rID1 == rID2 || rID2 == rID3 || rID1 == rID3 {
		t.Error("不同文件不应共享 rID")
	}

	// 验证：计数为 3
	if mgr.Count() != 3 {
		t.Errorf("Count = %d, want 3", mgr.Count())
	}

	// 再次添加已存在的文件
	rID1Again, _ := mgr.AddMediaAuto("logo.png", []byte("data1"))

	// 验证：应该复用已有的 rID
	if rID1Again != rID1 {
		t.Errorf("重复添加 logo.png 返回 rID = %q, want %q", rID1Again, rID1)
	}

	// 验证：计数仍为 3
	if mgr.Count() != 3 {
		t.Errorf("去重后 Count = %d, want 3", mgr.Count())
	}
}

func TestMediaManager_Deduplication_SameNameDifferentData(t *testing.T) {
	mgr := slide.NewMediaManager()

	// 第一次添加
	rID1, res1 := mgr.AddMediaAuto("image.png", []byte("original_data"))

	// 第二次添加相同文件名但不同数据（基于 Hash 去重，内容不同则创建新资源）
	rID2, res2 := mgr.AddMediaAuto("image.png", []byte("different_data"))

	// 验证：不同内容应该有不同的 rID
	if rID2 == rID1 {
		t.Errorf("不同内容不应共享 rID, got %q == %q", rID2, rID1)
	}

	// 验证：返回的是各自的数据
	if string(res1.Data()) != "original_data" {
		t.Errorf("res1.Data = %q, want original_data", res1.Data())
	}
	if string(res2.Data()) != "different_data" {
		t.Errorf("res2.Data = %q, want different_data", res2.Data())
	}

	// 验证：计数为 2
	if mgr.Count() != 2 {
		t.Errorf("Count = %d, want 2", mgr.Count())
	}
}

func TestMediaManager_Deduplication_DifferentNameSameContent(t *testing.T) {
	mgr := slide.NewMediaManager()

	sameData := []byte("identical_logo_bytes")

	// 不同文件名，相同内容
	rID1, res1 := mgr.AddMediaAuto("logo_v1.png", sameData)
	rID2, res2 := mgr.AddMediaAuto("logo_v2.png", sameData)
	rID3, res3 := mgr.AddMediaAuto("company_logo.png", sameData)

	// 验证：相同内容应该共享相同的 rID（基于 Hash 去重）
	if rID2 != rID1 {
		t.Errorf("相同内容不同文件名应复用 rID, got rID1=%q, rID2=%q", rID1, rID2)
	}
	if rID3 != rID1 {
		t.Errorf("相同内容不同文件名应复用 rID, got rID1=%q, rID3=%q", rID1, rID3)
	}

	// 验证：返回的是同一个资源对象
	if res2 != res1 || res3 != res1 {
		t.Error("相同内容应返回相同的资源对象")
	}

	// 验证：计数为 1
	if mgr.Count() != 1 {
		t.Errorf("Count = %d, want 1 (Hash 去重)", mgr.Count())
	}
}

func TestMediaManager_Deduplication_HashConsistency(t *testing.T) {
	mgr := slide.NewMediaManager()

	data := []byte("test_content")

	// 多次添加相同内容
	rID1, res1 := mgr.AddMediaAuto("a.png", data)
	rID2, res2 := mgr.AddMediaAuto("b.jpg", data)
	rID3, res3 := mgr.AddMediaAuto("c.mp4", data)

	// 验证：所有都应该返回相同的 rID
	if rID1 != rID2 || rID2 != rID3 {
		t.Errorf("Hash 去重失败: rID1=%q, rID2=%q, rID3=%q", rID1, rID2, rID3)
	}

	// 验证：所有资源对象相同
	if res1 != res2 || res2 != res3 {
		t.Error("相同内容应返回相同的资源对象")
	}

	// 验证：资源的 Hash 值相同
	if res1.Hash() == "" {
		t.Error("Hash 不应为空")
	}
	if res1.Hash() != res2.Hash() || res2.Hash() != res3.Hash() {
		t.Error("相同内容的 Hash 值应该相同")
	}
}
