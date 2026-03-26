package parts

import (
	"crypto/md5"
	"encoding/hex"
	"io"
	"path/filepath"
	"strconv"
	"sync"
	"sync/atomic"
)

// ============================================================================
// MediaManager - 媒体资源管理器（并发安全缓存）
// ============================================================================
//
// 设计原则：
// 1. 一次写入，到处读取 - 初始化后主要操作是读取
// 2. 读取优化 - 使用 sync.RWMutex，读操作无需阻塞
// 3. 双重索引 - 按 rID 和 fileName 都能快速查找
// ============================================================================

// MediaManager 媒体资源管理器
// 维护 PPTX 中所有媒体资源的并发安全缓存
type MediaManager struct {
	// 主存储：rID -> *MediaResource
	byRID sync.Map

	// 辅助索引：fileName -> rID（用于按文件名查找）
	byName sync.Map

	// 辅助索引：target -> rID（用于按路径查找）
	byTarget sync.Map

	// 辅助索引：contentHash -> rID（用于按内容去重）
	byHash sync.Map

	// 计数器
	count int64

	// 自增 ID 计数器（用于生成 rId1, rId2...）
	nextID int64

	// 初始化标记（用于确保一次性加载）
	once sync.Once

	// 写入锁（仅用于批量操作时的互斥）
	writeMu sync.Mutex
}

// ============================================================================
// 全局默认实例
// ============================================================================

var defaultMediaManager = NewMediaManager()

// DefaultMediaManager 返回全局默认媒体管理器
func DefaultMediaManager() *MediaManager {
	return defaultMediaManager
}

// ============================================================================
// 构造函数
// ============================================================================

// NewMediaManager 创建新的媒体资源管理器
func NewMediaManager() *MediaManager {
	return &MediaManager{}
}

// ============================================================================
// 写入方法
// ============================================================================

// AddMedia 添加媒体资源到缓存
// 返回资源的 rID，如果已存在则返回现有 rID
func (m *MediaManager) AddMedia(resource *MediaResource) string {
	if resource == nil {
		return ""
	}

	rID := resource.RID()
	if rID == "" {
		return ""
	}

	// 检查是否已存在
	if _, loaded := m.byRID.LoadOrStore(rID, resource); loaded {
		return rID // 已存在，直接返回
	}

	// 建立辅助索引
	if resource.FileName() != "" {
		m.byName.Store(resource.FileName(), rID)
	}
	if resource.Target() != "" {
		m.byTarget.Store(resource.Target(), rID)
	}

	// 增加计数（原子操作）
	atomic.AddInt64(&m.count, 1)

	return rID
}

// AddMediaWithBytes 从字节数据添加媒体资源
func (m *MediaManager) AddMediaWithBytes(rID, fileName, contentType, target string, data []byte) *MediaResource {
	resource := NewMediaResourceFromBytes(fileName, contentType, target, data)
	resource.SetRID(rID)
	m.AddMedia(resource)
	return resource
}

// AddMediaWithReader 从 Reader 添加媒体资源
func (m *MediaManager) AddMediaWithReader(rID, fileName, contentType, target string, reader io.Reader, size int64) *MediaResource {
	resource := NewMediaResourceFromReader(fileName, contentType, target, reader, size)
	resource.SetRID(rID)
	m.AddMedia(resource)
	return resource
}

// AddMediaAuto 自动推断 MIME 类型并生成自增 rID
// 如果相同内容已存在（基于 Hash），则返回已有资源（去重）
// 返回生成的 rID 和创建的 MediaResource
func (m *MediaManager) AddMediaAuto(fileName string, data []byte) (string, *MediaResource) {
	// 计算内容 Hash
	contentHash := computeHash(data)

	// 去重检查：如果相同内容已存在，直接返回已有资源
	if existing := m.GetMediaByHash(contentHash); existing != nil {
		return existing.RID(), existing
	}

	// 生成自增 rID
	id := atomic.AddInt64(&m.nextID, 1)
	rID := formatRID(id)

	// 推断 MIME 类型
	contentType := inferContentType(fileName)

	// 生成 target 路径
	target := "ppt/media/" + fileName

	// 创建资源
	resource := NewMediaResourceFromBytes(fileName, contentType, target, data)
	resource.SetRID(rID)
	resource.SetHash(contentHash)

	// 添加到缓存
	m.AddMedia(resource)

	// 建立 Hash 索引
	m.byHash.Store(contentHash, rID)

	return rID, resource
}

// computeHash 计算数据的 MD5 Hash
func computeHash(data []byte) string {
	hash := md5.Sum(data)
	return hex.EncodeToString(hash[:])
}

// formatRID 格式化 rID（如 rId1, rId2）
func formatRID(id int64) string {
	return "rId" + strconv.FormatInt(id, 10)
}

// inferContentType 根据文件扩展名推断 MIME 类型
func inferContentType(fileName string) string {
	ext := filepath.Ext(fileName)
	switch ext {
	case ".png":
		return "image/png"
	case ".jpg", ".jpeg":
		return "image/jpeg"
	case ".gif":
		return "image/gif"
	case ".bmp":
		return "image/bmp"
	case ".webp":
		return "image/webp"
	case ".svg":
		return "image/svg+xml"
	case ".mp4":
		return "video/mp4"
	case ".webm":
		return "video/webm"
	case ".avi":
		return "video/x-msvideo"
	case ".wmv":
		return "video/x-ms-wmv"
	case ".mov":
		return "video/quicktime"
	case ".mp3":
		return "audio/mpeg"
	case ".wav":
		return "audio/wav"
	case ".aac":
		return "audio/aac"
	case ".ogg":
		return "audio/ogg"
	default:
		return "application/octet-stream"
	}
}

// RemoveMedia 移除媒体资源
func (m *MediaManager) RemoveMedia(rID string) bool {
	if rID == "" {
		return false
	}

	// 先获取资源以便清理索引
	val, ok := m.byRID.Load(rID)
	if !ok {
		return false
	}

	resource := val.(*MediaResource)

	// 清理辅助索引
	if resource.FileName() != "" {
		m.byName.Delete(resource.FileName())
	}
	if resource.Target() != "" {
		m.byTarget.Delete(resource.Target())
	}

	// 删除主存储
	m.byRID.Delete(rID)

	// 减少计数（原子操作）
	atomic.AddInt64(&m.count, -1)

	return true
}

// Clear 清空所有媒体资源
func (m *MediaManager) Clear() {
	m.writeMu.Lock()
	defer m.writeMu.Unlock()

	m.byRID = sync.Map{}
	m.byName = sync.Map{}
	m.byTarget = sync.Map{}
	atomic.StoreInt64(&m.count, 0)
}

// ============================================================================
// 读取方法（并发安全，无锁读取）
// ============================================================================

// GetMedia 根据 rID 获取媒体资源
func (m *MediaManager) GetMedia(rID string) *MediaResource {
	if rID == "" {
		return nil
	}

	val, ok := m.byRID.Load(rID)
	if !ok {
		return nil
	}
	return val.(*MediaResource)
}

// GetMediaByFileName 根据文件名获取媒体资源
func (m *MediaManager) GetMediaByFileName(fileName string) *MediaResource {
	if fileName == "" {
		return nil
	}

	ridVal, ok := m.byName.Load(fileName)
	if !ok {
		return nil
	}

	return m.GetMedia(ridVal.(string))
}

// GetMediaByTarget 根据目标路径获取媒体资源
func (m *MediaManager) GetMediaByTarget(target string) *MediaResource {
	if target == "" {
		return nil
	}

	ridVal, ok := m.byTarget.Load(target)
	if !ok {
		return nil
	}

	return m.GetMedia(ridVal.(string))
}

// GetMediaByHash 根据内容 Hash 获取媒体资源（用于去重）
func (m *MediaManager) GetMediaByHash(hash string) *MediaResource {
	if hash == "" {
		return nil
	}

	ridVal, ok := m.byHash.Load(hash)
	if !ok {
		return nil
	}

	return m.GetMedia(ridVal.(string))
}

// HasMedia 检查媒体资源是否存在
func (m *MediaManager) HasMedia(rID string) bool {
	_, ok := m.byRID.Load(rID)
	return ok
}

// HasMediaByFileName 检查文件名是否存在
func (m *MediaManager) HasMediaByFileName(fileName string) bool {
	_, ok := m.byName.Load(fileName)
	return ok
}

// ============================================================================
// 批量读取方法
// ============================================================================

// AllMedia 返回所有媒体资源（返回新切片，线程安全）
func (m *MediaManager) AllMedia() []*MediaResource {
	result := make([]*MediaResource, 0, m.Count())
	m.byRID.Range(func(key, value interface{}) bool {
		result = append(result, value.(*MediaResource))
		return true
	})
	return result
}

// AllMediaByType 返回指定类型的所有媒体资源
func (m *MediaManager) AllMediaByType(mediaType MediaType) []*MediaResource {
	result := make([]*MediaResource, 0)
	m.byRID.Range(func(key, value interface{}) bool {
		res := value.(*MediaResource)
		if res.MediaType() == mediaType {
			result = append(result, res)
		}
		return true
	})
	return result
}

// AllImages 返回所有图片资源
func (m *MediaManager) AllImages() []*MediaResource {
	return m.AllMediaByType(MediaTypeImage)
}

// AllAudio 返回所有音频资源
func (m *MediaManager) AllAudio() []*MediaResource {
	return m.AllMediaByType(MediaTypeAudio)
}

// AllVideo 返回所有视频资源
func (m *MediaManager) AllVideo() []*MediaResource {
	return m.AllMediaByType(MediaTypeVideo)
}

// ============================================================================
// 统计方法
// ============================================================================

// Count 返回媒体资源总数
func (m *MediaManager) Count() int64 {
	return atomic.LoadInt64(&m.count)
}

// CountByType 返回指定类型的媒体资源数量
func (m *MediaManager) CountByType(mediaType MediaType) int64 {
	var count int64
	m.byRID.Range(func(key, value interface{}) bool {
		if value.(*MediaResource).MediaType() == mediaType {
			count++
		}
		return true
	})
	return count
}

// CountImages 返回图片数量
func (m *MediaManager) CountImages() int64 {
	return m.CountByType(MediaTypeImage)
}

// CountAudio 返回音频数量
func (m *MediaManager) CountAudio() int64 {
	return m.CountByType(MediaTypeAudio)
}

// CountVideo 返回视频数量
func (m *MediaManager) CountVideo() int64 {
	return m.CountByType(MediaTypeVideo)
}

// ============================================================================
// RID 列表
// ============================================================================

// ListRIDs 返回所有 rID
func (m *MediaManager) ListRIDs() []string {
	result := make([]string, 0, m.Count())
	m.byRID.Range(func(key, value interface{}) bool {
		result = append(result, key.(string))
		return true
	})
	return result
}

// ListFileNames 返回所有文件名
func (m *MediaManager) ListFileNames() []string {
	result := make([]string, 0, m.Count())
	m.byName.Range(func(key, value interface{}) bool {
		result = append(result, key.(string))
		return true
	})
	return result
}

// ListTargets 返回所有目标路径
func (m *MediaManager) ListTargets() []string {
	result := make([]string, 0, m.Count())
	m.byTarget.Range(func(key, value interface{}) bool {
		result = append(result, key.(string))
		return true
	})
	return result
}

// ============================================================================
// 全局便捷函数
// ============================================================================

// AddMedia 向默认管理器添加媒体资源
func AddMedia(resource *MediaResource) string {
	return defaultMediaManager.AddMedia(resource)
}

// GetMedia 从默认管理器获取媒体资源
func GetMedia(rID string) *MediaResource {
	return defaultMediaManager.GetMedia(rID)
}

// GetMediaByFileName 从默认管理器根据文件名获取媒体资源
func GetMediaByFileName(fileName string) *MediaResource {
	return defaultMediaManager.GetMediaByFileName(fileName)
}

// GetMediaByTarget 从默认管理器根据目标路径获取媒体资源
func GetMediaByTarget(target string) *MediaResource {
	return defaultMediaManager.GetMediaByTarget(target)
}

// ClearMedia 清空默认管理器
func ClearMedia() {
	defaultMediaManager.Clear()
}
