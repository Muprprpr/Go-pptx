package opc_test

import (
	"testing"

	"github.com/Muprprpr/Go-pptx/opc"
)

// TestResourcePool_Basic 测试资源池基本功能
func TestResourcePool_Basic(t *testing.T) {
	pool := opc.GetGlobalPool()
	if pool == nil {
		t.Fatal("GetGlobalPool returned nil")
	}

	// 清理之前的测试数据
	pool.ReleaseAll()

	// 测试 GetOrLoad
	callCount := 0
	data, err := pool.GetOrLoad("/ppt/media/image1.png", opc.ContentTypePNG, func() ([]byte, error) {
		callCount++
		return []byte{0x01, 0x02, 0x03}, nil
	})
	if err != nil {
		t.Fatalf("GetOrLoad failed: %v", err)
	}
	if callCount != 1 {
		t.Error("loader should be called once")
	}
	if len(data) != 3 {
		t.Errorf("data length should be 3, got %d", len(data))
	}

	// 再次获取（应该使用缓存）
	callCount = 0
	data2, err := pool.GetOrLoad("/ppt/media/image1.png", opc.ContentTypePNG, func() ([]byte, error) {
		callCount++
		return nil, nil
	})
	if err != nil {
		t.Fatalf("GetOrLoad second call failed: %v", err)
	}
	if callCount != 0 {
		t.Error("loader should not be called again (should use cache)")
	}

	// 验证数据相同（zero-copy）
	if len(data2) != len(data) {
		t.Error("data2 length should match data")
	}
	for i := range data {
		if data[i] != data2[i] {
		t.Errorf("data2[%d] should equal data[%d]", i, i)
		}
	}

	// 测试 Release（需要释放两次，因为调用了两次 GetOrLoad）
	pool.Release("/ppt/media/image1.png")
	pool.Release("/ppt/media/image1.png")
	stats := pool.Stats()
	if stats["media"] != 0 {
		t.Errorf("media count should be 0 after release, got %d", stats["media"])
	}

	// 清理
	pool.ReleaseAll()

		t.Log("Resource pool basic test passed")
}

// TestResourcePool_ContentTypeCategories 测试不同内容类型的分类
func TestResourcePool_ContentTypeCategories(t *testing.T) {
	// 测试不可变内容类型判断
	testCases := []struct {
		contentType string
		expected   bool
	}{
		{opc.ContentTypePNG, true},
		{opc.ContentTypeJPEG, true},
		{opc.ContentTypeTheme, true},
		{opc.ContentTypeSlideMaster, true},
		{opc.ContentTypeSlide, false}, // slide 是可变的
		{opc.ContentTypePresentation, false},
	}

	for _, tc := range testCases {
		result := opc.IsImmutableContentType(tc.contentType)
		if result != tc.expected {
			t.Errorf("IsImmutableContentType(%s) = %v, want %v", tc.contentType, result, tc.expected)
		}
	}

	t.Log("Content type categorization test passed")
}

// TestPart_CloneShared 测试 Part 的 zero-copy 克隆
func TestPart_CloneShared(t *testing.T) {
	// 创建原始数据
	originalData := []byte{0x01, 0x02, 0x03, 0x04, 0x05}
	uri := opc.NewPackURI("/ppt/media/image1.png")

	// 创建原始 Part
	original := opc.NewPart(uri, opc.ContentTypePNG, originalData)

	// 使用 CloneShared 进行 zero-copy 克隆
	cloned := original.CloneShared()

	// 验证克隆不为 nil
	if cloned == nil {
		t.Fatal("CloneShared returned nil")
	}

	// 验证是不可变的
	if !cloned.IsImmutable() {
		t.Error("cloned part should be immutable")
	}

	// 验证 URI 相同（共享指针）
	if cloned.PartURI() != original.PartURI() {
		t.Error("URI should be shared")
	}

	// 验证 Blob 返回相同的数据
	blob := cloned.Blob()
	if len(blob) != len(originalData) {
		t.Errorf("blob length mismatch: got %d, want %d", len(blob), len(originalData))
	}

	// 验证内容相同
	for i := range originalData {
		if blob[i] != originalData[i] {
			t.Errorf("blob content mismatch at index %d", i)
		}
	}

	t.Log("Part CloneShared test passed")
}

// TestPart_Clone_DeepCopy 测试 Part 的深拷贝
func TestPart_Clone_DeepCopy(t *testing.T) {
	// 创建原始数据
	originalData := []byte{0x01, 0x02, 0x03, 0x04, 0x05}
	uri := opc.NewPackURI("/ppt/slides/slide1.xml")

	// 创建原始 Part
	original := opc.NewPart(uri, opc.ContentTypeSlide, originalData)

	// 使用 Clone 进行深拷贝
	cloned := original.Clone()

	// 验证克隆不为 nil
	if cloned == nil {
		t.Fatal("Clone returned nil")
	}

	// 验证不是不可变的
	if cloned.IsImmutable() {
		t.Error("cloned part should not be immutable")
	}

	// 验证 Blob 是深拷贝的
	originalBlob := original.Blob()
	clonedBlob := cloned.Blob()

	// 修改克隆的数据不应影响原始数据
	if len(clonedBlob) > 0 {
		clonedBlob[0] = 0xFF
		if originalBlob[0] == 0xFF {
			t.Error("modifying cloned blob should not affect original")
		}
	}

	t.Log("Part Clone deep copy test passed")
}

// TestPackage_Clone_SmartCloning 测试 Package 的智能克隆
func TestPackage_Clone_SmartCloning(t *testing.T) {
	// 创建一个包含不同类型部件的 Package
	pkg := opc.NewPackage()

	// 添加一个图片部件（应该使用 zero-copy）
	imageData := []byte{0x01, 0x02, 0x03, 0x04, 0x05}
	imageURI := opc.NewPackURI("/ppt/media/image1.png")
	pkg.CreatePart(imageURI, opc.ContentTypePNG, imageData)

	// 添加一个幻灯片部件（应该使用深拷贝）
	slideData := []byte("<slide>test</slide>")
	slideURI := opc.NewPackURI("/ppt/slides/slide1.xml")
	pkg.CreatePart(slideURI, opc.ContentTypeSlide, slideData)

	// 克隆 Package
	clonedPkg := pkg.Clone()

	// 验证克隆不为 nil
	if clonedPkg == nil {
		t.Fatal("Clone returned nil")
	}

	// 验证图片部件使用了 zero-copy
	originalImagePart := pkg.GetPart(imageURI)
	clonedImagePart := clonedPkg.GetPart(imageURI)

	if originalImagePart == nil || clonedImagePart == nil {
		t.Fatal("image parts should exist")
	}

	// 图片应该是不可变的（使用了 zero-copy）
	if !clonedImagePart.IsImmutable() {
		t.Error("cloned image part should be immutable (zero-copy)")
	}

	// 幻灯片应该是可变的（使用了深拷贝）
	clonedSlidePart := clonedPkg.GetPart(slideURI)
	if clonedSlidePart == nil {
		t.Fatal("cloned slide part should exist")
	}
	if clonedSlidePart.IsImmutable() {
		t.Error("cloned slide part should not be immutable (deep copy)")
	}

	t.Log("Package Clone smart cloning test passed")
}
