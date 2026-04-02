// Package pptx 提供 PPTX 文件的高级操作接口
package pptx

import (
	"archive/zip"
	"fmt"
	"io"
	"path"
	"sort"
	"strings"
	"sync"

	"github.com/Muprprpr/Go-pptx/parts"
)

// ============================================================================
// MasterCache - 母版/版式只读缓存
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
	masters map[string]*parts.SlideMasterData // key: masterID
	layouts map[string]*parts.SlideLayoutData // key: layoutID

	// 辅助索引（初始化时构建）
	layoutByName map[string]string // layoutName -> layoutID
	masterByName map[string]string // masterName -> masterID

	// 占位符索引（初始化时构建）
	// key 格式: "layoutID:phType" 或 "masterID:phType"
	placeholderIndex map[string]*parts.Placeholder
}

// NewMasterCache 创建新的母版缓存实例
func NewMasterCache() *MasterCache {
	return &MasterCache{
		masters:          make(map[string]*parts.SlideMasterData),
		layouts:          make(map[string]*parts.SlideLayoutData),
		layoutByName:     make(map[string]string),
		masterByName:     make(map[string]string),
		placeholderIndex: make(map[string]*parts.Placeholder),
	}
}

// ============================================================================
// 初始化方法（仅调用一次）
// ============================================================================

// Init 使用提供的数据初始化缓存（仅执行一次）
// 后续调用将被忽略
func (c *MasterCache) Init(masters []*parts.SlideMasterData, layouts []*parts.SlideLayoutData) {
	c.once.Do(func() {
		c.buildIndex(masters, layouts)
	})
}

// InitFunc 延迟初始化，接受初始化函数
// 函数仅在第一次访问时执行
func (c *MasterCache) InitFunc(initFn func() ([]*parts.SlideMasterData, []*parts.SlideLayoutData)) {
	c.once.Do(func() {
		masters, layouts := initFn()
		c.buildIndex(masters, layouts)
	})
}

// buildIndex 构建索引（内部方法，仅初始化时调用）
func (c *MasterCache) buildIndex(masters []*parts.SlideMasterData, layouts []*parts.SlideLayoutData) {
	// 索引母版
	for _, master := range masters {
		if master == nil {
			continue
		}
		c.masters[master.ID()] = master
		if master.Name() != "" {
			c.masterByName[master.Name()] = master.ID()
		}

		// 索引母版级占位符
		for phID, ph := range master.Placeholders() {
			key := master.ID() + ":" + ph.Type().String()
			c.placeholderIndex[key] = ph
			// 同时按 ID 索引
			c.placeholderIndex[master.ID()+":"+phID] = ph
		}
	}

	// 索引版式
	for _, layout := range layouts {
		if layout == nil {
			continue
		}
		c.layouts[layout.ID()] = layout
		if layout.Name() != "" {
			c.layoutByName[layout.Name()] = layout.ID()
		}

		// 索引版式级占位符
		for phID, ph := range layout.Placeholders() {
			key := layout.ID() + ":" + ph.Type().String()
			c.placeholderIndex[key] = ph
			// 同时按 ID 索引
			c.placeholderIndex[layout.ID()+":"+phID] = ph
		}
	}
}

// ============================================================================
// 读取接口 - 无锁并发安全
// ============================================================================

// GetMaster 根据 ID 获取母版
func (c *MasterCache) GetMaster(masterID string) (*parts.SlideMasterData, bool) {
	m, ok := c.masters[masterID]
	return m, ok
}

// GetMasterByName 根据名称获取母版
func (c *MasterCache) GetMasterByName(name string) (*parts.SlideMasterData, bool) {
	if id, ok := c.masterByName[name]; ok {
		return c.GetMaster(id)
	}
	return nil, false
}

// GetLayout 根据 ID 获取版式
func (c *MasterCache) GetLayout(layoutID string) (*parts.SlideLayoutData, bool) {
	l, ok := c.layouts[layoutID]
	return l, ok
}

// GetLayoutByName 根据名称获取版式
func (c *MasterCache) GetLayoutByName(name string) (*parts.SlideLayoutData, bool) {
	if id, ok := c.layoutByName[name]; ok {
		return c.GetLayout(id)
	}
	return nil, false
}

// GetPlaceholder 根据版式 ID 和占位符类型获取占位符
// phType 可以是 PlaceholderType.String() 的值，如 "title", "body" 等
func (c *MasterCache) GetPlaceholder(layoutID, phType string) (*parts.Placeholder, bool) {
	key := layoutID + ":" + phType
	ph, ok := c.placeholderIndex[key]
	return ph, ok
}

// GetPlaceholderByID 根据版式 ID 和占位符 ID 获取占位符
func (c *MasterCache) GetPlaceholderByID(layoutID, placeholderID string) (*parts.Placeholder, bool) {
	key := layoutID + ":" + placeholderID
	ph, ok := c.placeholderIndex[key]
	return ph, ok
}

// GetMasterPlaceholder 根据母版 ID 和占位符类型获取占位符
func (c *MasterCache) GetMasterPlaceholder(masterID, phType string) (*parts.Placeholder, bool) {
	key := masterID + ":" + phType
	ph, ok := c.placeholderIndex[key]
	return ph, ok
}

// ============================================================================
// 批量读取接口
// ============================================================================

// AllMasters 返回所有母版（只读）
func (c *MasterCache) AllMasters() map[string]*parts.SlideMasterData {
	return c.masters
}

// AllLayouts 返回所有版式（只读）
func (c *MasterCache) AllLayouts() map[string]*parts.SlideLayoutData {
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
// MasterManager - 母版/版式管理器（门面模式）
// ============================================================================
//
// 作为外部 API 调用的入口，负责：
// 1. 从 ZIP 文件加载母版和版式
// 2. 解析 XML 并转换为只读数据结构
// 3. 填充到 MasterCache 供高并发读取
// ============================================================================

// MasterManager 母版/版式管理器
type MasterManager struct {
	cache *MasterCache
}

// NewMasterManager 创建新的母版管理器
func NewMasterManager() *MasterManager {
	return &MasterManager{
		cache: NewMasterCache(),
	}
}

// NewMasterManagerWithCache 使用指定缓存创建母版管理器
func NewMasterManagerWithCache(cache *MasterCache) *MasterManager {
	return &MasterManager{
		cache: cache,
	}
}

// Cache 返回内部缓存（只读）
func (m *MasterManager) Cache() *MasterCache {
	return m.cache
}

// ============================================================================
// 从 ZIP 加载
// ============================================================================

// LoadFromZip 从 ZIP Reader 加载母版和版式
// 遍历 ZIP 内的 /ppt/slideMasters/ 和 /ppt/slideLayouts/ 目录
func (m *MasterManager) LoadFromZip(zipReader *zip.Reader) error {
	var masters []*parts.SlideMasterData
	var layouts []*parts.SlideLayoutData

	// 收集母版文件
	masterFiles := m.collectFiles(zipReader, "ppt/slideMasters/", "slideMaster")
	layoutFiles := m.collectFiles(zipReader, "ppt/slideLayouts/", "slideLayout")

	// 按文件名排序，确保顺序一致
	sort.Slice(masterFiles, func(i, j int) bool {
		return masterFiles[i].name < masterFiles[j].name
	})
	sort.Slice(layoutFiles, func(i, j int) bool {
		return layoutFiles[i].name < layoutFiles[j].name
	})

	// 解析母版
	for _, f := range masterFiles {
		data, err := m.readFile(f.file)
		if err != nil {
			return fmt.Errorf("读取母版文件 %s 失败: %w", f.name, err)
		}

		master, err := parts.ParseMaster(data)
		if err != nil {
			return fmt.Errorf("解析母版 %s 失败: %w", f.name, err)
		}

		masters = append(masters, master)
	}

	// 解析版式
	for _, f := range layoutFiles {
		data, err := m.readFile(f.file)
		if err != nil {
			return fmt.Errorf("读取版式文件 %s 失败: %w", f.name, err)
		}

		layout, err := parts.ParseLayout(data)
		if err != nil {
			return fmt.Errorf("解析版式 %s 失败: %w", f.name, err)
		}

		layouts = append(layouts, layout)
	}

	// 初始化缓存（仅执行一次）
	m.cache.Init(masters, layouts)

	return nil
}

// LoadFromZipFile 从 ZIP 文件路径加载
func (m *MasterManager) LoadFromZipFile(filePath string) error {
	reader, err := zip.OpenReader(filePath)
	if err != nil {
		return fmt.Errorf("打开 ZIP 文件失败: %w", err)
	}
	defer reader.Close()

	return m.LoadFromZip(&reader.Reader)
}

// ============================================================================
// 文件收集辅助
// ============================================================================

type zipFileEntry struct {
	name string
	file *zip.File
}

// collectFiles 收集指定目录下匹配前缀的 XML 文件
func (m *MasterManager) collectFiles(zipReader *zip.Reader, dir, prefix string) []zipFileEntry {
	var files []zipFileEntry

	for _, f := range zipReader.File {
		// 检查是否在目标目录
		fileDir := path.Dir(f.Name)
		if fileDir != dir && !strings.HasPrefix(fileDir, dir) {
			continue
		}

		// 检查文件名前缀和扩展名
		fileName := path.Base(f.Name)
		if !strings.HasPrefix(fileName, prefix) {
			continue
		}
		if path.Ext(fileName) != ".xml" {
			continue
		}

		files = append(files, zipFileEntry{
			name: f.Name,
			file: f,
		})
	}

	return files
}

// readFile 读取 ZIP 文件内容
func (m *MasterManager) readFile(f *zip.File) ([]byte, error) {
	rc, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	return io.ReadAll(rc)
}

// ============================================================================
// 便捷访问方法（委托给 Cache）
// ============================================================================

// GetLayout 获取版式
func (m *MasterManager) GetLayout(layoutID string) (*parts.SlideLayoutData, bool) {
	return m.cache.GetLayout(layoutID)
}

// GetLayoutByName 根据名称获取版式
func (m *MasterManager) GetLayoutByName(name string) (*parts.SlideLayoutData, bool) {
	return m.cache.GetLayoutByName(name)
}

// GetMaster 获取母版
func (m *MasterManager) GetMaster(masterID string) (*parts.SlideMasterData, bool) {
	return m.cache.GetMaster(masterID)
}

// GetMasterByName 根据名称获取母版
func (m *MasterManager) GetMasterByName(name string) (*parts.SlideMasterData, bool) {
	return m.cache.GetMasterByName(name)
}

// GetPlaceholder 获取占位符
func (m *MasterManager) GetPlaceholder(layoutID, phType string) (*parts.Placeholder, bool) {
	return m.cache.GetPlaceholder(layoutID, phType)
}

// AllLayouts 返回所有版式
func (m *MasterManager) AllLayouts() map[string]*parts.SlideLayoutData {
	return m.cache.AllLayouts()
}

// AllMasters 返回所有母版
func (m *MasterManager) AllMasters() map[string]*parts.SlideMasterData {
	return m.cache.AllMasters()
}

// LayoutCount 返回版式数量
func (m *MasterManager) LayoutCount() int {
	return m.cache.LayoutCount()
}

// MasterCount 返回母版数量
func (m *MasterManager) MasterCount() int {
	return m.cache.MasterCount()
}

// ListLayoutIDs 列出所有版式 ID
func (m *MasterManager) ListLayoutIDs() []string {
	return m.cache.ListLayoutIDs()
}

// ListLayoutNames 列出所有版式名称
func (m *MasterManager) ListLayoutNames() []string {
	return m.cache.ListLayoutNames()
}
