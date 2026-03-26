package parts_test

import (
	"os"
	"sync"
	"testing"

	"github.com/Muprprpr/Go-pptx/parts"
)

// ============================================================================
// MasterCache 并发读取测试
// ============================================================================

func TestMasterCache_ConcurrentRead(t *testing.T) {
	// 从真实文件初始化缓存
	cache := parts.NewMasterCache()
	masters, layouts := loadTestMasterData()
	cache.Init(masters, layouts)

	// 启动并发读取
	const readers = 100
	const readsPerGoroutine = 50
	var wg sync.WaitGroup
	wg.Add(readers)

	for i := 0; i < readers; i++ {
		go func(readerID int) {
			defer wg.Done()

			for j := 0; j < readsPerGoroutine; j++ {
				// 读取版式
				if layout, ok := cache.GetLayout("layout_1"); ok {
					_ = layout.Placeholders()
					_ = layout.PlaceholderByType(parts.PlaceholderTypeTitle)
					_ = layout.PlaceholderByType(parts.PlaceholderTypeBody)
				}

				// 读取母版
				if master, ok := cache.GetMaster("master_1"); ok {
					_ = master.Placeholders()
					_ = master.Background()
				}

				// 读取占位符
				if ph, ok := cache.GetPlaceholder("layout_1", "title"); ok {
					_ = ph.X()
					_ = ph.Y()
					_ = ph.Cx()
					_ = ph.Cy()
				}

				// 读取统计
				_ = cache.LayoutCount()
				_ = cache.MasterCount()
				_ = cache.LayoutExists("layout_1")
				_ = cache.MasterExists("master_1")

				// 读取列表
				_ = cache.ListLayoutIDs()
				_ = cache.ListMasterIDs()
			}
		}(i)
	}

	wg.Wait()
}

// ============================================================================
// MasterCache 高并发压力测试
// ============================================================================

func TestMasterCache_HighConcurrency(t *testing.T) {
	cache := parts.NewMasterCache()
	masters, layouts := loadTestMasterData()
	cache.Init(masters, layouts)

	// 100 个 goroutine，每个执行 100 次读取
	const totalGoroutines = 100
	const iterations = 100
	var wg sync.WaitGroup
	wg.Add(totalGoroutines)

	for i := 0; i < totalGoroutines; i++ {
		go func(id int) {
			defer wg.Done()

			for j := 0; j < iterations; j++ {
				// 并发读取所有数据
				layout, _ := cache.GetLayout("layout_1")
				if layout != nil {
					_ = layout.PlaceholderCount()
					_ = layout.Placeholders()
				}

				master, _ := cache.GetMaster("master_1")
				if master != nil {
					_ = master.PlaceholderCount()
				}

				// 索引查询
				if _, ok := cache.GetLayoutByName("13_Title Slide"); ok {}
				if _, ok := cache.GetPlaceholderByID("layout_1", "ph_8"); ok {}
				if _, ok := cache.GetPlaceholder("layout_1", "title"); ok {}

				// 批量读取
				_ = cache.AllLayouts()
				_ = cache.AllMasters()
			}
		}(i)
	}

	wg.Wait()
}

// ============================================================================
// MasterCache 辅助函数
// ============================================================================

func loadTestMasterData() ([]*parts.SlideMasterData, []*parts.SlideLayoutData) {
	var masters []*parts.SlideMasterData
	var layouts []*parts.SlideLayoutData

	// 加载母版
	masterXML, _ := os.ReadFile("../test-data/test/ppt/slideMasters/slideMaster2.xml")
	if master, err := parts.ParseMaster(masterXML); err == nil {
		masters = append(masters, master)
	}

	// 加载版式
	layoutXML, _ := os.ReadFile("../test-data/test/ppt/slideLayouts/slideLayout5.xml")
	if layout, err := parts.ParseLayout(layoutXML); err == nil {
		layouts = append(layouts, layout)
	}

	return masters, layouts
}
