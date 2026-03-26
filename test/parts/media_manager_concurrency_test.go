package parts_test

import (
	"math/rand"
	"strconv"
	"sync"
	"testing"

	"github.com/Muprprpr/Go-pptx/parts"
)

// ============================================================================
// MediaManager 并发安全性测试
// ============================================================================

func TestMediaManager_Concurrency(t *testing.T) {
	mgr := parts.NewMediaManager()
	const goroutines = 100
	const opsPerGoroutine = 50

	var wg sync.WaitGroup
	wg.Add(goroutines)

	// 预先添加一些媒体供读取
	for i := 0; i < 20; i++ {
		rID := "rId" + strconv.Itoa(i+1)
		fileName := "preset_" + strconv.Itoa(i) + ".png"
		mgr.AddMediaWithBytes(rID, fileName, "image/png", "ppt/media/"+fileName, []byte("preset_data"))
	}

	// 启动并发 goroutine
	for g := 0; g < goroutines; g++ {
		go func(goroutineID int) {
			defer wg.Done()

			for i := 0; i < opsPerGoroutine; i++ {
				op := rand.Intn(3) // 0: Add, 1: Get, 2: Remove

				switch op {
				case 0:
					// 随机添加
					n := rand.Intn(1000)
					fileName := "img_" + strconv.Itoa(n) + ".png"
					mgr.AddMediaAuto(fileName, []byte("data_"+strconv.Itoa(n)))

				case 1:
					// 随机读取
					rID := "rId" + strconv.Itoa(rand.Intn(50)+1)
					mgr.GetMedia(rID)
					mgr.GetMediaByFileName("preset_" + strconv.Itoa(rand.Intn(20)) + ".png")
					mgr.HasMedia(rID)

				case 2:
					// 随机删除
					rID := "rId" + strconv.Itoa(rand.Intn(30)+1)
					mgr.RemoveMedia(rID)
				}
			}
		}(g)
	}

	wg.Wait()

	// 验证管理器仍然可用
	_ = mgr.Count()
	_ = mgr.AllMedia()
	_ = mgr.ListRIDs()
}

// ============================================================================
// 并发读写分离测试
// ============================================================================

func TestMediaManager_ConcurrentReadWrite(t *testing.T) {
	mgr := parts.NewMediaManager()
	const writers = 20
	const readers = 80
	const iterations = 100

	var wg sync.WaitGroup
	wg.Add(writers + readers)

	// 写者 goroutine
	for w := 0; w < writers; w++ {
		go func(id int) {
			defer wg.Done()
			for i := 0; i < iterations; i++ {
				fileName := "writer_" + strconv.Itoa(id) + "_" + strconv.Itoa(i) + ".png"
				mgr.AddMediaAuto(fileName, []byte("data"))
			}
		}(w)
	}

	// 读者 goroutine
	for r := 0; r < readers; r++ {
		go func(id int) {
			defer wg.Done()
			for i := 0; i < iterations; i++ {
				// 随机读取各种方法
				rID := "rId" + strconv.Itoa(rand.Intn(100)+1)
				mgr.GetMedia(rID)
				mgr.HasMedia(rID)
				mgr.Count()
				mgr.AllImages()
				mgr.ListFileNames()
			}
		}(r)
	}

	wg.Wait()
}

// ============================================================================
// 并发计数器测试
// ============================================================================

func TestMediaManager_ConcurrentCount(t *testing.T) {
	mgr := parts.NewMediaManager()
	const ops = 1000

	var wg sync.WaitGroup
	wg.Add(ops)

	for i := 0; i < ops; i++ {
		go func(n int) {
			defer wg.Done()
			fileName := "count_test_" + strconv.Itoa(n) + ".png"
			// 使用唯一数据避免去重，确保每个操作都添加新资源
			data := []byte("data_for_count_test_" + strconv.Itoa(n))
			mgr.AddMediaAuto(fileName, data)
		}(i)
	}

	wg.Wait()

	// 最终计数应等于操作数
	if mgr.Count() != ops {
		t.Errorf("Count = %d, want %d", mgr.Count(), ops)
	}
}
