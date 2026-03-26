package parts_test

import (
	"testing"

	"github.com/Muprprpr/Go-pptx/parts"
)

// ============================================================================
// MediaManager 基础 CRUD 测试
// ============================================================================

func TestMediaManager_AddAndGet(t *testing.T) {
	// 初始化空的 MediaManager
	mgr := parts.NewMediaManager()

	// 添加图片
	rID1, img := mgr.AddMediaAuto("logo.png", []byte("fake_image_data"))
	if rID1 != "rId1" {
		t.Errorf("第一次添加 rID = %q, want %q", rID1, "rId1")
	}
	if img.ContentType() != "image/png" {
		t.Errorf("图片 ContentType = %q, want %q", img.ContentType(), "image/png")
	}
	if img.MediaType() != parts.MediaTypeImage {
		t.Errorf("图片 MediaType = %v, want %v", img.MediaType(), parts.MediaTypeImage)
	}

	// 添加视频
	rID2, video := mgr.AddMediaAuto("demo.mp4", []byte("fake_video_data"))
	if rID2 != "rId2" {
		t.Errorf("第二次添加 rID = %q, want %q", rID2, "rId2")
	}
	if video.ContentType() != "video/mp4" {
		t.Errorf("视频 ContentType = %q, want %q", video.ContentType(), "video/mp4")
	}
	if video.MediaType() != parts.MediaTypeVideo {
		t.Errorf("视频 MediaType = %v, want %v", video.MediaType(), parts.MediaTypeVideo)
	}

	// 验证 GetMedia 能正确取回
	retrieved := mgr.GetMedia("rId1")
	if retrieved == nil {
		t.Fatal("GetMedia(\"rId1\") 返回 nil")
	}
	if string(retrieved.Data()) != "fake_image_data" {
		t.Errorf("Data = %q, want %q", string(retrieved.Data()), "fake_image_data")
	}
	if retrieved.FileName() != "logo.png" {
		t.Errorf("FileName = %q, want %q", retrieved.FileName(), "logo.png")
	}

	// 验证计数
	if mgr.Count() != 2 {
		t.Errorf("Count = %d, want 2", mgr.Count())
	}
}

// ============================================================================
// MIME 类型推断测试
// ============================================================================

func TestMediaManager_ContentTypeInference(t *testing.T) {
	mgr := parts.NewMediaManager()

	tests := []struct {
		fileName        string
		wantContentType string
		wantMediaType   parts.MediaType
	}{
		{"image.png", "image/png", parts.MediaTypeImage},
		{"photo.jpg", "image/jpeg", parts.MediaTypeImage},
		{"photo.jpeg", "image/jpeg", parts.MediaTypeImage},
		{"animation.gif", "image/gif", parts.MediaTypeImage},
		{"icon.bmp", "image/bmp", parts.MediaTypeImage},
		{"modern.webp", "image/webp", parts.MediaTypeImage},
		{"video.mp4", "video/mp4", parts.MediaTypeVideo},
		{"clip.webm", "video/webm", parts.MediaTypeVideo},
		{"movie.avi", "video/x-msvideo", parts.MediaTypeVideo},
		{"audio.mp3", "audio/mpeg", parts.MediaTypeAudio},
		{"sound.wav", "audio/wav", parts.MediaTypeAudio},
		{"music.aac", "audio/aac", parts.MediaTypeAudio},
		{"unknown.xyz", "application/octet-stream", parts.MediaTypeUnknown},
	}

	for _, tt := range tests {
		t.Run(tt.fileName, func(t *testing.T) {
			// 使用唯一数据，避免去重干扰类型推断测试
			data := []byte("data_for_" + tt.fileName)
			_, res := mgr.AddMediaAuto(tt.fileName, data)
			if res.ContentType() != tt.wantContentType {
				t.Errorf("ContentType = %q, want %q", res.ContentType(), tt.wantContentType)
			}
			if res.MediaType() != tt.wantMediaType {
				t.Errorf("MediaType = %v, want %v", res.MediaType(), tt.wantMediaType)
			}
		})
	}
}

// ============================================================================
// 索引查找测试
// ============================================================================

func TestMediaManager_GetByIndex(t *testing.T) {
	mgr := parts.NewMediaManager()

	// 添加媒体
	mgr.AddMediaAuto("logo.png", []byte("img_data"))
	mgr.AddMediaAuto("sound.mp3", []byte("audio_data"))

	t.Run("GetMediaByFileName", func(t *testing.T) {
		res := mgr.GetMediaByFileName("logo.png")
		if res == nil {
			t.Fatal("GetMediaByFileName 返回 nil")
		}
		if res.ContentType() != "image/png" {
			t.Errorf("ContentType = %q, want image/png", res.ContentType())
		}
	})

	t.Run("GetMediaByTarget", func(t *testing.T) {
		res := mgr.GetMediaByTarget("ppt/media/logo.png")
		if res == nil {
			t.Fatal("GetMediaByTarget 返回 nil")
		}
		if res.RID() != "rId1" {
			t.Errorf("RID = %q, want rId1", res.RID())
		}
	})

	t.Run("HasMedia", func(t *testing.T) {
		if !mgr.HasMedia("rId1") {
			t.Error("HasMedia(\"rId1\") = false, want true")
		}
		if mgr.HasMedia("rId999") {
			t.Error("HasMedia(\"rId999\") = true, want false")
		}
	})

	t.Run("HasMediaByFileName", func(t *testing.T) {
		if !mgr.HasMediaByFileName("logo.png") {
			t.Error("HasMediaByFileName(\"logo.png\") = false, want true")
		}
	})
}

// ============================================================================
// Remove 和 Clear 测试
// ============================================================================

func TestMediaManager_RemoveAndClear(t *testing.T) {
	t.Run("RemoveMedia", func(t *testing.T) {
		mgr := parts.NewMediaManager()
		mgr.AddMediaAuto("test.png", []byte("data"))

		if !mgr.RemoveMedia("rId1") {
			t.Error("RemoveMedia 返回 false")
		}
		if mgr.HasMedia("rId1") {
			t.Error("删除后 HasMedia 仍返回 true")
		}
		if mgr.Count() != 0 {
			t.Errorf("删除后 Count = %d, want 0", mgr.Count())
		}
	})

	t.Run("Clear", func(t *testing.T) {
		mgr := parts.NewMediaManager()
		mgr.AddMediaAuto("a.png", []byte("a"))
		mgr.AddMediaAuto("b.mp3", []byte("b"))
		mgr.AddMediaAuto("c.mp4", []byte("c"))

		mgr.Clear()
		if mgr.Count() != 0 {
			t.Errorf("Clear 后 Count = %d, want 0", mgr.Count())
		}
		if mgr.HasMedia("rId1") {
			t.Error("Clear 后 HasMedia 仍返回 true")
		}
	})
}
