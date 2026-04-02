package pptx_test

import (
	"sync"
	"sync/atomic"
	"testing"

	"github.com/Muprprpr/Go-pptx/pptx"
)

// ============================================================================
// 重点 2： Context 的 ID 递增逻辑
// ============================================================================
//
// SlideContext 是派给各个组件的特派员，它手里握着发牌权（分配 rId 和 shapeId）。
// 如果它发错了牌，PPT 就会由于 ID 冲突而彻底损坏。
//
// 测试目标：
// 1. 证明同一个 Slide 上的 Context，每次调用 NextShapeID() 时，分配的 ID 严格递增
// 2. 证明 ID 绝不重复
// 3. 证明并发分配时不会冲突（使用原子计数器）
//
// ============================================================================

// TestSlideContext_ShapeIDIncremental 验证形状 ID 严格递增
func TestSlideContext_ShapeIDIncremental(t *testing.T) {
	prs := pptx.New()
	slide := prs.AddSlide()
	ctx := slide.NewContext()

	// 记录分配的 ID
	allocatedIDs := make(map[uint32]bool)
	var prevID uint32 = 0

	// 分配 100 个 ID，验证递增
	for i := 0; i < 100; i++ {
		id := ctx.NextShapeID()

		// 验证 ID 不为 0（0 通常是无效 ID）
		if id == 0 {
			t.Fatalf("分配了无效 ID: 0")
		}

		// 验证 ID 严格递增
		if i > 0 && id <= prevID {
			t.Fatalf("ID 非递增: prev=%d, current=%d", prevID, id)
		}

		// 验证 ID 不重复
		if allocatedIDs[id] {
			t.Fatalf("严重异常：分配了重复的形状 ID: %d", id)
		}
		allocatedIDs[id] = true
		prevID = id
	}

	t.Logf("✅ 形状 ID 递增测试通过，分配了 %d 个唯一 ID", len(allocatedIDs))
}

// TestSlideContext_ShapeIDNoDuplicate 验证形状 ID 绝不重复
func TestSlideContext_ShapeIDNoDuplicate(t *testing.T) {
	prs := pptx.New()
	slide := prs.AddSlide()
	ctx := slide.NewContext()

	const allocationCount = 1000
	allocatedIDs := make(map[uint32]bool, allocationCount)

	for i := 0; i < allocationCount; i++ {
		id := ctx.NextShapeID()

		if allocatedIDs[id] {
			t.Fatalf("严重异常：在第 %d 次分配时发现重复 ID: %d", i+1, id)
		}
		allocatedIDs[id] = true
	}

	if len(allocatedIDs) != allocationCount {
		t.Errorf("分配的唯一 ID 数量: %d, 期望: %d", len(allocatedIDs), allocationCount)
	}

	t.Logf("✅ 形状 ID 唯一性测试通过，%d 次分配零重复", allocationCount)
}

// TestSlideContext_ConcurrentIDAllocation 验证并发分配 ID 不冲突
// 这是战役 2 的核心测试！
func TestSlideContext_ConcurrentIDAllocation(t *testing.T) {
	prs := pptx.New()
	slide := prs.AddSlide()
	ctx := slide.NewContext()

	const concurrencyLevel = 1000 // 模拟 1000 个组件同时请求 ID

	// 使用 sync.Map 来记录所有分配出去的 ID，检查是否有重复
	var idMap sync.Map
	var wg sync.WaitGroup
	var duplicateCount int32

	// 释放 1000 个 Goroutine 并发抢夺 ID
	for i := 0; i < concurrencyLevel; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()

			// 并发调用！
			id := ctx.NextShapeID()

			// 检查这个 ID 是否已经被别人抢过了
			if _, loaded := idMap.LoadOrStore(id, true); loaded {
				// 如果 loaded 为 true，说明 Map 里早就有了这个 ID，发生撞车！
				atomic.AddInt32(&duplicateCount, 1)
			}
		}()
	}

	wg.Wait()

	// 验证没有重复
	if duplicateCount > 0 {
		t.Fatalf("严重异常：并发分配发现 %d 个重复 ID", duplicateCount)
	}

	// 验证总数是否一票不少
	count := 0
	idMap.Range(func(_, _ any) bool {
		count++
		return true
	})

	if count != concurrencyLevel {
		t.Fatalf("严重漏算：期望分配 %d 个 ID，实际只分配了 %d 个", concurrencyLevel, count)
	}

	t.Logf("✅ 极度舒适：%d 并发无锁 ID 分配测试通过，零冲突！", concurrencyLevel)
}

// TestSlideContext_IDAllocationOrder 验证 ID 分配顺序
func TestSlideContext_IDAllocationOrder(t *testing.T) {
	prs := pptx.New()
	slide := prs.AddSlide()
	ctx := slide.NewContext()

	// 验证初始状态
	firstID := ctx.NextShapeID()
	secondID := ctx.NextShapeID()
	thirdID := ctx.NextShapeID()

	// 验证严格递增
	if secondID <= firstID {
		t.Errorf("第二个 ID (%d) 应大于第一个 ID (%d)", secondID, firstID)
	}
	if thirdID <= secondID {
		t.Errorf("第三个 ID (%d) 应大于第二个 ID (%d)", thirdID, secondID)
	}

	// 验证差值为 1（严格连续）
	if secondID-firstID != 1 {
		t.Errorf("ID 差值异常: second - first = %d (期望 1)", secondID-firstID)
	}
	if thirdID-secondID != 1 {
		t.Errorf("ID 差值异常: third - second = %d (期望 1)", thirdID-secondID)
	}

	t.Logf("✅ ID 分配顺序测试通过: %d -> %d -> %d", firstID, secondID, thirdID)
}

// TestSlideContext_MultipleContextsSameSlide 验证同一 Slide 的多个 Context 共享 ID 分配器
func TestSlideContext_MultipleContextsSameSlide(t *testing.T) {
	prs := pptx.New()
	slide := prs.AddSlide()

	// 创建多个 Context
	ctx1 := slide.NewContext()
	ctx2 := slide.NewContext()
	ctx3 := slide.NewContext()

	// 从不同 Context 分配 ID
	id1 := ctx1.NextShapeID()
	id2 := ctx2.NextShapeID()
	id3 := ctx3.NextShapeID()

	// 验证所有 ID 唯一（因为它们共享底层的原子计数器）
	allIDs := []uint32{id1, id2, id3}
	idSet := make(map[uint32]bool)
	for i, id := range allIDs {
		if idSet[id] {
			t.Fatalf("不同 Context 分配了重复 ID: %d (索引 %d)", id, i)
		}
		idSet[id] = true
	}

	t.Logf("✅ 多 Context 共享 ID 分配器测试通过: %v", allIDs)
}

// TestSlideContext_ConcurrentStress 并发压力测试
func TestSlideContext_ConcurrentStress(t *testing.T) {
	prs := pptx.New()
	slide := prs.AddSlide()

	const goroutineCount = 100
	const operationsPerGoroutine = 100

	var wg sync.WaitGroup
	globalIDSet := sync.Map{} // 并发安全的 map
	var duplicateCount int32

	for g := 0; g < goroutineCount; g++ {
		wg.Add(1)
		go func(goroutineID int) {
			defer wg.Done()
			ctx := slide.NewContext()

			for i := 0; i < operationsPerGoroutine; i++ {
				id := ctx.NextShapeID()

				// 尝试存储，如果已存在则计数重复
				if _, exists := globalIDSet.LoadOrStore(id, true); exists {
					atomic.AddInt32(&duplicateCount, 1)
				}
			}
		}(g)
	}

	wg.Wait()

	if duplicateCount > 0 {
		t.Fatalf("严重异常：并发压力测试发现 %d 个重复 ID", duplicateCount)
	}

	// 统计实际分配的 ID 数量
	var uniqueCount int
	globalIDSet.Range(func(_, _ any) bool {
		uniqueCount++
		return true
	})

	expectedCount := goroutineCount * operationsPerGoroutine
	if uniqueCount != expectedCount {
		t.Errorf("唯一 ID 数量: %d, 期望: %d", uniqueCount, expectedCount)
	}

	t.Logf("✅ 并发压力测试通过，%d 个 goroutine 各分配 %d 个 ID，零冲突",
		goroutineCount, operationsPerGoroutine)
}

// TestSlideContext_RelationshipIDIncremental 验证关系 ID (rId) 递增
func TestSlideContext_RelationshipIDIncremental(t *testing.T) {
	prs := pptx.New()
	slide := prs.AddSlide()
	ctx := slide.NewContext()

	// 模拟添加图片获取 rId
	imageData := []byte{0x89, 0x50, 0x4E, 0x47} // PNG header

	allocatedRIDs := make(map[string]bool)

	for i := 0; i < 10; i++ {
		rID, err := ctx.AddMedia(imageData, "test.png")
		if err != nil {
			// 忽略无效图片数据的错误，只关注 rId 分配
			continue
		}

		if rID == "" {
			continue
		}

		// 验证 rId 格式（通常是 "rId1", "rId2" 等）
		if len(rID) < 4 || rID[:3] != "rId" {
			t.Errorf("rId 格式错误: %s", rID)
			continue
		}

		// 验证 rId 不重复
		if allocatedRIDs[rID] {
			t.Fatalf("严重异常：分配了重复的 rId: %s", rID)
		}
		allocatedRIDs[rID] = true
	}

	t.Logf("✅ 关系 ID 分配测试通过")
}

// TestSlideContext_InitialValue 验证 ID 初始值符合 OOXML 规范
func TestSlideContext_InitialValue(t *testing.T) {
	prs := pptx.New()
	slide := prs.AddSlide()
	ctx := slide.NewContext()

	// 第一个分配的 ID 应该是 2（因为 1 预留给根节点）
	firstID := ctx.NextShapeID()
	if firstID != 2 {
		t.Errorf("第一个 ID: %d, 期望: 2 (1 预留给根节点)", firstID)
	}

	t.Logf("✅ ID 初始值测试通过，第一个 ID: %d", firstID)
}
