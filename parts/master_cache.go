package parts

import (
	"sync"
)

// ============================================================================
// 母版/版式只读缓存 - 高并发无锁读取
// ============================================================================
//
// 设计原则：
// 1. 一次写入，到处读取 - 使用 sync.Once 确保初始化只执行一次
// 2. 初始化后冻结写入，后续读取无需加锁，性能最优
// 3. 面向并发安全设计，适用于高并发/流式生成场景
// ============================================================================

// MasterCache 母版/版式只读缓存
// 初始化后所有字段只读，支持无锁并发访问
type MasterCache struct {
	// 初始化控制
	once sync.Once

	// 只读数据（初始化后冻结）
	masters  map[string]*SlideMasterData // key: masterID
	layouts  map[string]*SlideLayoutData // key: layoutID

	// 辅助索引（初始化时构建）
	layoutByName map[string]string // layoutName -> layoutID
	masterByName map[string]string // masterName -> masterID

	// 占位符索引（初始化时构建）
	// key 格式: "layoutID:phType" 或 "masterID:phType"
	placeholderIndex map[string]*Placeholder
}

// ============================================================================
// 全局缓存实例
// ============================================================================

// 全局默认缓存实例
var defaultCache = &MasterCache{
	masters:          make(map[string]*SlideMasterData),
	layouts:          make(map[string]*SlideLayoutData),
	layoutByName:     make(map[string]string),
	masterByName:     make(map[string]string),
	placeholderIndex: make(map[string]*Placeholder),
}

// DefaultCache 返回全局默认缓存实例
func DefaultCache() *MasterCache {
	return defaultCache
}

// NewMasterCache 创建新的母版缓存实例
func NewMasterCache() *MasterCache {
	return &MasterCache{
		masters:          make(map[string]*SlideMasterData),
		layouts:          make(map[string]*SlideLayoutData),
		layoutByName:     make(map[string]string),
		masterByName:     make(map[string]string),
		placeholderIndex: make(map[string]*Placeholder),
	}
}

// ============================================================================
// 初始化方法（仅调用一次）
// ============================================================================

// Init 使用提供的数据初始化缓存（仅执行一次）
// 后续调用将被忽略
func (c *MasterCache) Init(masters []*SlideMasterData, layouts []*SlideLayoutData) {
	c.once.Do(func() {
		c.buildIndex(masters, layouts)
	})
}

// InitFunc 延迟初始化，接受初始化函数
// 函数仅在第一次访问时执行
func (c *MasterCache) InitFunc(initFn func() ([]*SlideMasterData, []*SlideLayoutData)) {
	c.once.Do(func() {
		masters, layouts := initFn()
		c.buildIndex(masters, layouts)
	})
}

// buildIndex 构建索引（内部方法，仅初始化时调用）
func (c *MasterCache) buildIndex(masters []*SlideMasterData, layouts []*SlideLayoutData) {
	// 索引母版
	for _, master := range masters {
		if master == nil {
			continue
		}
		c.masters[master.id] = master
		if master.name != "" {
			c.masterByName[master.name] = master.id
		}

		// 索引母版级占位符
		for phID, ph := range master.placeholders {
			key := master.id + ":" + ph.placeholderType.String()
			c.placeholderIndex[key] = ph
			// 同时按 ID 索引
			c.placeholderIndex[master.id+":"+phID] = ph
		}
	}

	// 索引版式
	for _, layout := range layouts {
		if layout == nil {
			continue
		}
		c.layouts[layout.id] = layout
		if layout.name != "" {
			c.layoutByName[layout.name] = layout.id
		}

		// 索引版式级占位符
		for phID, ph := range layout.placeholders {
			key := layout.id + ":" + ph.placeholderType.String()
			c.placeholderIndex[key] = ph
			// 同时按 ID 索引
			c.placeholderIndex[layout.id+":"+phID] = ph
		}
	}
}

// ============================================================================
// 读取接口 - 无锁并发安全
// ============================================================================

// GetMaster 根据 ID 获取母版
func (c *MasterCache) GetMaster(masterID string) (*SlideMasterData, bool) {
	m, ok := c.masters[masterID]
	return m, ok
}

// GetMasterByName 根据名称获取母版
func (c *MasterCache) GetMasterByName(name string) (*SlideMasterData, bool) {
	if id, ok := c.masterByName[name]; ok {
		return c.GetMaster(id)
	}
	return nil, false
}

// GetLayout 根据 ID 获取版式
func (c *MasterCache) GetLayout(layoutID string) (*SlideLayoutData, bool) {
	l, ok := c.layouts[layoutID]
	return l, ok
}

// GetLayoutByName 根据名称获取版式
func (c *MasterCache) GetLayoutByName(name string) (*SlideLayoutData, bool) {
	if id, ok := c.layoutByName[name]; ok {
		return c.GetLayout(id)
	}
	return nil, false
}

// GetPlaceholder 根据版式 ID 和占位符类型获取占位符
// phType 可以是 PlaceholderType.String() 的值，如 "title", "body" 等
func (c *MasterCache) GetPlaceholder(layoutID, phType string) (*Placeholder, bool) {
	key := layoutID + ":" + phType
	ph, ok := c.placeholderIndex[key]
	return ph, ok
}

// GetPlaceholderByID 根据版式 ID 和占位符 ID 获取占位符
func (c *MasterCache) GetPlaceholderByID(layoutID, placeholderID string) (*Placeholder, bool) {
	key := layoutID + ":" + placeholderID
	ph, ok := c.placeholderIndex[key]
	return ph, ok
}

// GetMasterPlaceholder 根据母版 ID 和占位符类型获取占位符
func (c *MasterCache) GetMasterPlaceholder(masterID, phType string) (*Placeholder, bool) {
	key := masterID + ":" + phType
	ph, ok := c.placeholderIndex[key]
	return ph, ok
}

// ============================================================================
// 批量读取接口
// ============================================================================

// AllMasters 返回所有母版（只读）
func (c *MasterCache) AllMasters() map[string]*SlideMasterData {
	return c.masters
}

// AllLayouts 返回所有版式（只读）
func (c *MasterCache) AllLayouts() map[string]*SlideLayoutData {
	return c.layouts
}

// MasterCount 返回母版数量
func (c *MasterCache) MasterCount() int {
	return len(c.masters)
}

// LayoutCount 返回版式数量
func (c *MasterCache) LayoutCount() int {
	return len(c.layouts)
}

// ============================================================================
// 辅助方法
// ============================================================================

// LayoutExists 检查版式是否存在
func (c *MasterCache) LayoutExists(layoutID string) bool {
	_, ok := c.layouts[layoutID]
	return ok
}

// MasterExists 检查母版是否存在
func (c *MasterCache) MasterExists(masterID string) bool {
	_, ok := c.masters[masterID]
	return ok
}

// ListLayoutIDs 列出所有版式 ID
func (c *MasterCache) ListLayoutIDs() []string {
	ids := make([]string, 0, len(c.layouts))
	for id := range c.layouts {
		ids = append(ids, id)
	}
	return ids
}

// ListMasterIDs 列出所有母版 ID
func (c *MasterCache) ListMasterIDs() []string {
	ids := make([]string, 0, len(c.masters))
	for id := range c.masters {
		ids = append(ids, id)
	}
	return ids
}

// ListLayoutNames 列出所有版式名称
func (c *MasterCache) ListLayoutNames() []string {
	names := make([]string, 0, len(c.layoutByName))
	for name := range c.layoutByName {
		names = append(names, name)
	}
	return names
}

// ============================================================================
// 全局便捷函数（操作默认缓存）
// ============================================================================

// InitDefaultCache 初始化全局默认缓存
func InitDefaultCache(masters []*SlideMasterData, layouts []*SlideLayoutData) {
	defaultCache.Init(masters, layouts)
}

// GetLayout 从默认缓存获取版式
func GetLayout(layoutID string) (*SlideLayoutData, bool) {
	return defaultCache.GetLayout(layoutID)
}

// GetLayoutByName 从默认缓存根据名称获取版式
func GetLayoutByName(name string) (*SlideLayoutData, bool) {
	return defaultCache.GetLayoutByName(name)
}

// GetMaster 从默认缓存获取母版
func GetMaster(masterID string) (*SlideMasterData, bool) {
	return defaultCache.GetMaster(masterID)
}

// GetPlaceholder 从默认缓存获取占位符
func GetPlaceholder(layoutID, phType string) (*Placeholder, bool) {
	return defaultCache.GetPlaceholder(layoutID, phType)
}
