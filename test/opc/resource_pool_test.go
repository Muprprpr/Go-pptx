package opc_test

import (
	"archive/zip"
	"bytes"
	"testing"
	"time"

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

// TestZipEntry_Timestamp 测试 ZIP 条目的时间戳是否正确设置
// 验证：
// 1. 时间戳不为零值（解决 Windows 资源管理器 MS-DOS 时间 bug）
// 2. 时间戳是当前时间（允许几秒误差）
// 3. 时间戳可以正确转换为北京时间（UTC+8）
func TestZipEntry_Timestamp(t *testing.T) {
	// 记录创建前的时间（北京时间）
	beijingLoc, _ := time.LoadLocation("Asia/Shanghai")
	beforeCreate := time.Now().In(beijingLoc)

	// 创建一个简单的 Package
	pkg := opc.NewPackage()

	// 添加一个部件
	slideURI := opc.NewPackURI("/ppt/slides/slide1.xml")
	pkg.CreatePart(slideURI, opc.ContentTypeSlide, []byte("<slide/>"))

	// 添加关系
	pkg.AddRelationship(opc.RelTypeOfficeDocument, "/ppt/presentation.xml", false)

	// 保存到字节数组
	data, err := pkg.SaveToBytes()
	if err != nil {
		t.Fatalf("SaveToBytes failed: %v", err)
	}

	// 记录创建后的时间
	afterCreate := time.Now().In(beijingLoc)

	// 读取 ZIP 文件并检查时间戳
	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("Failed to read ZIP: %v", err)
	}

	// 检查每个文件的时间戳
	for _, file := range reader.File {
		modTime := file.Modified

		// 1. 时间戳不应为零值
		if modTime.IsZero() {
			t.Errorf("File %s has zero modification time", file.Name)
			continue
		}

		// 2. 时间戳应在 beforeCreate 和 afterCreate 之间（允许 1 秒误差）
		modTimeUTC := modTime.UTC()
		beforeUTC := beforeCreate.UTC()
		afterUTC := afterCreate.UTC()

		if modTimeUTC.Before(beforeUTC.Add(-1 * time.Second)) {
			t.Errorf("File %s modification time %v is before creation time %v",
				file.Name, modTimeUTC, beforeUTC)
		}
		if modTimeUTC.After(afterUTC.Add(1 * time.Second)) {
			t.Errorf("File %s modification time %v is after creation time %v",
				file.Name, modTimeUTC, afterUTC)
		}

		// 3. 转换为北京时间并验证
		modTimeBeijing := modTime.In(beijingLoc)

		// 验证北京时间与 UTC 时间差为 8 小时（或根据夏令时调整）
		_, beijingOffset := modTimeBeijing.Zone()
		expectedOffset := 8 * 60 * 60 // 8 小时（秒）
		if beijingOffset != expectedOffset {
			t.Logf("Warning: Beijing offset is %d seconds, expected %d (may vary by DST)",
				beijingOffset, expectedOffset)
		}

		t.Logf("File: %s", file.Name)
		t.Logf("  UTC Time:   %v", modTimeUTC.Format("2006-01-02 15:04:05 MST"))
		t.Logf("  Beijing:    %v", modTimeBeijing.Format("2006-01-02 15:04:05 MST"))
		t.Logf("  Unix:       %d", modTime.Unix())
	}

	t.Log("ZIP entry timestamp test passed")
}

// TestZipEntry_TimestampNotZero 专门测试时间戳不为零
// 这是解决 Windows 资源管理器 MS-DOS 时间解析 bug 的关键测试
func TestZipEntry_TimestampNotZero(t *testing.T) {
	pkg := opc.NewPackage()

	// 添加多个不同类型的部件
	testParts := []struct {
		uri         string
		contentType string
		data        []byte
	}{
		{"/ppt/slides/slide1.xml", opc.ContentTypeSlide, []byte("<slide/>")},
		{"/ppt/media/image1.png", opc.ContentTypePNG, []byte{0x89, 0x50, 0x4E, 0x47}},
		{"/docProps/core.xml", opc.ContentTypeCoreProperties, []byte("<coreProperties/>")},
	}

	for _, tp := range testParts {
		uri := opc.NewPackURI(tp.uri)
		pkg.CreatePart(uri, tp.contentType, tp.data)
	}

	// 保存
	data, err := pkg.SaveToBytes()
	if err != nil {
		t.Fatalf("SaveToBytes failed: %v", err)
	}

	// 读取并验证
	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("Failed to read ZIP: %v", err)
	}

	zeroTimeCount := 0
	validTimeCount := 0

	for _, file := range reader.File {
		if file.Modified.IsZero() {
			zeroTimeCount++
			t.Errorf("File %s has zero modification time (will cause Explorer bug)", file.Name)
		} else {
			validTimeCount++
		}
	}

	if zeroTimeCount > 0 {
		t.Errorf("Found %d files with zero timestamp (Windows Explorer bug risk)", zeroTimeCount)
	}

	t.Logf("Valid timestamps: %d, Zero timestamps: %d", validTimeCount, zeroTimeCount)

	if zeroTimeCount == 0 {
		t.Log("All ZIP entries have valid timestamps - Windows Explorer compatible")
	}
}
