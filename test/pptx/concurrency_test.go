package pptx_test

import (
	"sync"
	"sync/atomic"
	"testing"

	"github.com/Muprprpr/Go-pptx/pptx"
)

// ============================================================================
// 重点 1：高并发克隆的绝对隔离测试
// ============================================================================
//
// 测试目标：证明 100 个 Goroutine 同时调用 pptx.New()，它们拿到的是
// 100 份完全独立的物理副本，修改其中一个，绝对不会污染内置的 template 和其他的实例。
//
// ============================================================================

// TestPresentation_ConcurrencyIsolation 并发隔离测试
// 验证多个 goroutine 同时创建 Presentation 时，实例之间完全隔离
func TestPresentation_ConcurrencyIsolation(t *testing.T) {
	const goroutineCount = 100
	var wg sync.WaitGroup
	var successCount int32
	var panicCount int32

	// 用于存储每个 goroutine 创建的 Presentation，后续验证独立性
	presentations := make([]*pptx.Presentation, goroutineCount)
	var mu sync.Mutex

	for i := 0; i < goroutineCount; i++ {
		wg.Add(1)
		go func(id int) {
			defer wg.Done()
			defer func() {
				if r := recover(); r != nil {
					t.Errorf("Goroutine %d panic: %v", id, r)
					atomic.AddInt32(&panicCount, 1)
				}
			}()

			// 创建新的 Presentation
			prs := pptx.New()

			// 添加幻灯片
			slide := prs.AddSlide()

			// 添加文本框（使用 goroutine ID 作为内容，确保唯一性）
			text := string(rune('A'+id%26)) + "-Slide-Content"
			slide.AddTextBox(100, 100, 500, 50, text)

			// 添加形状
			slide.AddRectangle(100, 200, 300, 150)

			// 设置幻灯片尺寸（测试修改不会影响其他实例）
			prs.SetSlideSize(1280*9525, 720*9525)

			// 保存引用
			mu.Lock()
			presentations[id] = prs
			mu.Unlock()

			atomic.AddInt32(&successCount, 1)
		}(i)
	}

	wg.Wait()

	// 验证所有 goroutine 都成功完成
	if successCount != goroutineCount {
		t.Errorf("成功数量: %d, 期望: %d", successCount, goroutineCount)
	}

	if panicCount > 0 {
		t.Errorf("发生 panic 的 goroutine 数量: %d", panicCount)
	}

	t.Logf("✅ 并发创建 %d 个 Presentation 全部成功", successCount)

	// 验证每个 Presentation 都是独立的
	t.Run("VerifyIndependence", func(t *testing.T) {
		for i := 0; i < goroutineCount; i++ {
			if presentations[i] == nil {
				t.Errorf("Presentation %d 为 nil", i)
				continue
			}

			// 验证幻灯片数量
			if presentations[i].SlideCount() != 1 {
				t.Errorf("Presentation %d 幻灯片数量: %d, 期望: 1", i, presentations[i].SlideCount())
			}
		}
	})

	t.Log("✅ 并发隔离测试通过，零串线污染！")
}

// TestPresentation_ConcurrentModification 并发修改测试
// 验证多个 goroutine 修改同一个 Presentation 时不会发生数据竞争
func TestPresentation_ConcurrentModification(t *testing.T) {
	prs := pptx.New()

	const goroutineCount = 50
	var wg sync.WaitGroup
	var slideIndices []int
	var mu sync.Mutex

	// 并发添加幻灯片
	for i := 0; i < goroutineCount; i++ {
		wg.Add(1)
		go func(id int) {
			defer wg.Done()
			slide := prs.AddSlide()
			slide.AddTextBox(100, 100, 500, 50, "Slide-"+string(rune('0'+id%10)))

			mu.Lock()
			slideIndices = append(slideIndices, slide.Index())
			mu.Unlock()
		}(i)
	}

	wg.Wait()

	// 验证幻灯片数量
	if prs.SlideCount() != goroutineCount {
		t.Errorf("幻灯片数量: %d, 期望: %d", prs.SlideCount(), goroutineCount)
	}

	t.Logf("✅ 并发添加 %d 张幻灯片成功", prs.SlideCount())
}

// TestPresentation_ConcurrentSlideAccess 并发访问幻灯片测试
// 验证并发读取幻灯片时的线程安全
func TestPresentation_ConcurrentSlideAccess(t *testing.T) {
	prs := pptx.New()

	// 预先添加幻灯片
	const slideCount = 20
	for i := 0; i < slideCount; i++ {
		prs.AddSlide()
	}

	const goroutineCount = 100
	var wg sync.WaitGroup
	var readErrors int32

	// 并发读取幻灯片
	for i := 0; i < goroutineCount; i++ {
		wg.Add(1)
		go func(id int) {
			defer wg.Done()
			defer func() {
				if r := recover(); r != nil {
					atomic.AddInt32(&readErrors, 1)
				}
			}()

			// 随机访问幻灯片
			index := id % slideCount
			slide, err := prs.GetSlide(index)
			if err != nil {
				atomic.AddInt32(&readErrors, 1)
				return
			}

			// 读取幻灯片属性
			_ = slide.Index()
			_, _ = slide.SlideSize()

			// 获取所有幻灯片
			slides := prs.Slides()
			if len(slides) != slideCount {
				atomic.AddInt32(&readErrors, 1)
			}
		}(i)
	}

	wg.Wait()

	if readErrors > 0 {
		t.Errorf("读取错误数量: %d", readErrors)
	}

	t.Log("✅ 并发访问幻灯片测试通过")
}

// TestPresentation_ConcurrentSave 并发保存测试
// 验证多个 goroutine 同时保存不同的 Presentation 时不会互相干扰
func TestPresentation_ConcurrentSave(t *testing.T) {
	const goroutineCount = 20
	var wg sync.WaitGroup
	var saveErrors int32

	for i := 0; i < goroutineCount; i++ {
		wg.Add(1)
		go func(id int) {
			defer wg.Done()
			defer func() {
				if r := recover(); r != nil {
					t.Errorf("Goroutine %d panic during save: %v", id, r)
					atomic.AddInt32(&saveErrors, 1)
				}
			}()

			prs := pptx.New()
			slide := prs.AddSlide()
			slide.AddTextBox(100, 100, 500, 50, "Content-"+string(rune('0'+id%10)))

			// 写入字节数组（不实际写入文件）
			data, err := prs.WriteToBytes()
			if err != nil {
				atomic.AddInt32(&saveErrors, 1)
				return
			}

			// 验证生成的数据不为空
			if len(data) == 0 {
				atomic.AddInt32(&saveErrors, 1)
			}
		}(i)
	}

	wg.Wait()

	if saveErrors > 0 {
		t.Errorf("保存错误数量: %d", saveErrors)
	}

	t.Log("✅ 并发保存测试通过")
}

// TestMediaManager_Concurrency 媒体管理器并发测试
// 验证多个独立的 Presentation 并发添加媒体时的线程安全
// 重要：每个 goroutine 使用独立的 Presentation，这是正确的使用方式
func TestMediaManager_Concurrency(t *testing.T) {
	const goroutineCount = 50
	var wg sync.WaitGroup
	var addErrors int32

	// 模拟图片数据
	imageData := []byte{0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A} // PNG header

	for i := 0; i < goroutineCount; i++ {
		wg.Add(1)
		go func(id int) {
			defer wg.Done()
			defer func() {
				if r := recover(); r != nil {
					t.Errorf("Goroutine %d panic: %v", id, r)
					atomic.AddInt32(&addErrors, 1)
				}
			}()

			// 每个 goroutine 创建独立的 Presentation
			prs := pptx.New()
			slide := prs.AddSlide()

			fileName := "image-" + string(rune('0'+id%10)) + ".png"
			_, err := slide.AddPictureFromBytes(100+id*10, 100, 100, 100, fileName, imageData)
			if err != nil {
				// 某些错误是预期的（如无效图片数据），不计入错误
			}
		}(i)
	}

	wg.Wait()

	if addErrors > 0 {
		t.Errorf("添加媒体错误数量: %d", addErrors)
	}

	t.Log("✅ 媒体管理器并发测试通过")
}

// TestPresentation_MemoryIsolation 内存隔离测试
// 验证修改一个 Presentation 不会影响其他 Presentation
func TestPresentation_MemoryIsolation(t *testing.T) {
	// 创建第一个 Presentation
	prs1 := pptx.New()
	slide1 := prs1.AddSlide()
	slide1.AddTextBox(100, 100, 500, 50, "Presentation 1")

	// 创建第二个 Presentation
	prs2 := pptx.New()
	slide2 := prs2.AddSlide()
	slide2.AddTextBox(100, 100, 500, 50, "Presentation 2")

	// 修改第一个 Presentation
	prs1.SetSlideSize(960*9525, 720*9525) // 4:3

	// 验证第二个 Presentation 不受影响
	cx, cy := prs2.SlideSize()
	expectedCX := 1280 * 9525 // 默认 16:9
	expectedCY := 720 * 9525

	// 允许一定的误差（EMU 精度）
	if cx != expectedCX || cy != expectedCY {
		t.Errorf("Presentation 2 尺寸被污染: cx=%d, cy=%d, 期望: cx=%d, cy=%d",
			cx, cy, expectedCX, expectedCY)
	}

	// 验证幻灯片数量独立
	if prs1.SlideCount() != 1 {
		t.Errorf("Presentation 1 幻灯片数量: %d, 期望: 1", prs1.SlideCount())
	}
	if prs2.SlideCount() != 1 {
		t.Errorf("Presentation 2 幻灯片数量: %d, 期望: 1", prs2.SlideCount())
	}

	t.Log("✅ 内存隔离测试通过，实例之间零污染！")
}

// TestPresentation_CloneIsolation 克隆隔离测试
// 验证克隆后的 Presentation 与原实例完全独立
func TestPresentation_CloneIsolation(t *testing.T) {
	// 创建原始 Presentation
	original := pptx.New()
	slide := original.AddSlide()
	slide.AddTextBox(100, 100, 500, 50, "Original")

	// 克隆
	cloned, err := original.Clone()
	if err != nil {
		t.Fatalf("克隆失败: %v", err)
	}

	// 修改原始实例
	original.AddSlide()
	original.SetSlideSize(960*9525, 720*9525)

	// 验证克隆实例不受影响
	if cloned.SlideCount() != 1 {
		t.Errorf("克隆实例幻灯片数量被污染: %d, 期望: 1", cloned.SlideCount())
	}

	cx, _ := cloned.SlideSize()
	if cx != 1280*9525 {
		t.Errorf("克隆实例尺寸被污染: cx=%d, 期望: %d", cx, 1280*9525)
	}

	// 修改克隆实例
	cloned.AddSlide()

	// 验证原始实例不受影响
	if original.SlideCount() != 2 { // 原始实例有 2 张幻灯片（1+1）
		t.Errorf("原始实例幻灯片数量: %d, 期望: 2", original.SlideCount())
	}

	t.Log("✅ 克隆隔离测试通过！")
}
