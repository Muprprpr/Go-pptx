// Package pptx 提供 PPTX 文件的高级操作接口
package pptx

import (
	"crypto/md5"
	"encoding/hex"
	"io"
	"path/filepath"
	"strconv"
	"sync"
	"sync/atomic"

	"github.com/Muprprpr/Go-pptx/parts"
)

// ============================================================================
// MediaManager - 媒体资源管理器（并发安全缓存 + 跨幻灯片去重）
// ============================================================================
//
// 设计原则：
// 1. 一次写入，到处读取 - 初始化后主要操作是读取
// 2. 读取优化 - 使用 sync.Map，读操作无需阻塞
// 3. 多重索引 - 按 rID、fileName、target、hash 都能快速查找
// 4. 内容去重 - 基于内容 Hash 自动去重，相同媒体只存储一份
// 5. 跨幻灯片引用 - 同一媒体可被多个幻灯片引用，每页有独立 rId
//
// 使用场景：
// - AI 在 10 页幻灯片里都插入了同一个公司 Logo
// - Hash 判定后，ZIP 包里只存一份 image1.png
// - 但每页幻灯片获得独立的 rId（如 rId5, rId12, rId23...）
// - 极大减小生成文件的体积！
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

	// 媒体文件计数器（用于生成 image1.png, image2.png...）
	mediaFileID int64

	// 初始化标记（用于确保一次性加载）
	once sync.Once

	// 写入锁（仅用于批量操作时的互斥）
	writeMu sync.Mutex

	// ============================================================================
	// 跨幻灯片引用支持
	// ============================================================================

	// 幻灯片级关系映射：slideIndex -> {localRID -> globalRID}
	// 每个幻灯片有自己的 rId 命名空间
	slideRelations sync.Map // map[int]*SlideMediaIndex

	// 全局媒体存储：hash -> *MediaResource
	// 存储去重后的媒体资源
	globalMedia sync.Map
}

// SlideMediaIndex 幻灯片媒体索引
// 管理单个幻灯片的媒体引用
type SlideMediaIndex struct {
	// 幻灯片索引
	slideIndex int

	// 本地 rID -> 全局资源 Hash
	localToHash sync.Map

	// 全局 Hash -> 本地 rID（反向索引）
	hashToLocal sync.Map

	// 本地 rID 计数器
	nextLocalID int64
}

// NewSlideMediaIndex 创建幻灯片媒体索引
func NewSlideMediaIndex(slideIndex int) *SlideMediaIndex {
	return &SlideMediaIndex{
		slideIndex: slideIndex,
	}
}

// NewMediaManager 创建新的媒体资源管理器
func NewMediaManager() *MediaManager {
	return &MediaManager{}
}

// ============================================================================
// 写入方法
// ============================================================================

// AddMedia 添加媒体资源到缓存
// 返回资源的 rID，如果已存在则返回现有 rID
func (m *MediaManager) AddMedia(resource *parts.MediaResource) string {
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
func (m *MediaManager) AddMediaWithBytes(rID, fileName, contentType, target string, data []byte) *parts.MediaResource {
	resource := parts.NewMediaResourceFromBytes(fileName, contentType, target, data)
	resource.SetRID(rID)
	m.AddMedia(resource)
	return resource
}

// AddMediaWithReader 从 Reader 添加媒体资源
func (m *MediaManager) AddMediaWithReader(rID, fileName, contentType, target string, reader io.Reader, size int64) *parts.MediaResource {
	resource := parts.NewMediaResourceFromReader(fileName, contentType, target, reader, size)
	resource.SetRID(rID)
	m.AddMedia(resource)
	return resource
}

// AddMediaAuto 自动推断 MIME 类型并生成自增 rID
// 如果相同内容已存在（基于 Hash），则返回已有资源（去重）
// 返回生成的 rID 和创建的 MediaResource
func (m *MediaManager) AddMediaAuto(fileName string, data []byte) (string, *parts.MediaResource) {
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
	resource := parts.NewMediaResourceFromBytes(fileName, contentType, target, data)
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

	resource := val.(*parts.MediaResource)

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
	m.byHash = sync.Map{}
	atomic.StoreInt64(&m.count, 0)
}

// ============================================================================
// 读取方法（并发安全，无锁读取）
// ============================================================================

// GetMedia 根据 rID 获取媒体资源
func (m *MediaManager) GetMedia(rID string) *parts.MediaResource {
	if rID == "" {
		return nil
	}

	val, ok := m.byRID.Load(rID)
	if !ok {
		return nil
	}
	return val.(*parts.MediaResource)
}

// GetMediaByFileName 根据文件名获取媒体资源
func (m *MediaManager) GetMediaByFileName(fileName string) *parts.MediaResource {
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
func (m *MediaManager) GetMediaByTarget(target string) *parts.MediaResource {
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
func (m *MediaManager) GetMediaByHash(hash string) *parts.MediaResource {
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
func (m *MediaManager) AllMedia() []*parts.MediaResource {
	result := make([]*parts.MediaResource, 0, m.Count())
	m.byRID.Range(func(key, value interface{}) bool {
		result = append(result, value.(*parts.MediaResource))
		return true
	})
	return result
}

// AllMediaByType 返回指定类型的所有媒体资源
func (m *MediaManager) AllMediaByType(mediaType parts.MediaType) []*parts.MediaResource {
	result := make([]*parts.MediaResource, 0)
	m.byRID.Range(func(key, value interface{}) bool {
		res := value.(*parts.MediaResource)
		if res.MediaType() == mediaType {
			result = append(result, res)
		}
		return true
	})
	return result
}

// AllImages 返回所有图片资源
func (m *MediaManager) AllImages() []*parts.MediaResource {
	return m.AllMediaByType(parts.MediaTypeImage)
}

// AllAudio 返回所有音频资源
func (m *MediaManager) AllAudio() []*parts.MediaResource {
	return m.AllMediaByType(parts.MediaTypeAudio)
}

// AllVideo 返回所有视频资源
func (m *MediaManager) AllVideo() []*parts.MediaResource {
	return m.AllMediaByType(parts.MediaTypeVideo)
}

// ============================================================================
// 统计方法
// ============================================================================

// Count 返回媒体资源总数
func (m *MediaManager) Count() int64 {
	return atomic.LoadInt64(&m.count)
}

// CountByType 返回指定类型的媒体资源数量
func (m *MediaManager) CountByType(mediaType parts.MediaType) int64 {
	var count int64
	m.byRID.Range(func(key, value interface{}) bool {
		if value.(*parts.MediaResource).MediaType() == mediaType {
			count++
		}
		return true
	})
	return count
}

// CountImages 返回图片数量
func (m *MediaManager) CountImages() int64 {
	return m.CountByType(parts.MediaTypeImage)
}

// CountAudio 返回音频数量
func (m *MediaManager) CountAudio() int64 {
	return m.CountByType(parts.MediaTypeAudio)
}

// CountVideo 返回视频数量
func (m *MediaManager) CountVideo() int64 {
	return m.CountByType(parts.MediaTypeVideo)
}

// ============================================================================
// 列表方法
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
	m.byTarget.Range(func(key, value any) bool {
		result = append(result, key.(string))
		return true
	})
	return result
}

// ============================================================================
// 跨幻灯片媒体引用方法
// ============================================================================
//
// 核心功能：支持同一媒体被多个幻灯片引用，每页有独立的 rId
//
// 使用场景：
// - AI 在 10 页幻灯片里都插入了同一个公司 Logo
// - Hash 判定后，ZIP 包里只存一份 image1.png
// - 但每页幻灯片获得独立的 rId（如 rId5, rId12, rId23...）
// ============================================================================

// AddMediaForSlide 为指定幻灯片添加媒体（支持跨幻灯片去重）
// 返回该幻灯片的本地 rId 和全局媒体资源
//
// 使用示例：
//
//	// 第 1 页插入 Logo
//	rId1, _ := mediaManager.AddMediaForSlide(0, logoData, "logo.png")
//	// 返回: rId1="rId1", 全局存储 image1.png
//
//	// 第 2 页插入同一个 Logo
//	rId2, _ := mediaManager.AddMediaForSlide(1, logoData, "logo.png")
//	// 返回: rId2="rId1"（该幻灯片的本地 rId）, 复用 image1.png
//
//	// 最终 ZIP 包中只有一份 image1.png，但两张幻灯片都有各自的 rId 引用
func (m *MediaManager) AddMediaForSlide(slideIndex int, data []byte, fileName string) (string, *parts.MediaResource) {
	// 1. 计算内容 Hash
	contentHash := computeHash(data)

	// 2. 检查全局媒体是否已存在
	globalResource, _ := m.getOrCreateGlobalMedia(contentHash, data, fileName)

	// 3. 获取或创建幻灯片媒体索引
	slideIndexer := m.getOrCreateSlideIndex(slideIndex)

	// 4. 为该幻灯片生成本地 rId（如果尚未引用此媒体）
	localRID, _ := slideIndexer.getOrCreateLocalRID(contentHash, globalResource)

	return localRID, globalResource
}

// getOrCreateGlobalMedia 获取或创建全局媒体资源
func (m *MediaManager) getOrCreateGlobalMedia(contentHash string, data []byte, fileName string) (*parts.MediaResource, bool) {
	// 检查是否已存在
	if val, ok := m.globalMedia.Load(contentHash); ok {
		return val.(*parts.MediaResource), true
	}

	// 创建新的全局媒体资源
	m.writeMu.Lock()
	defer m.writeMu.Unlock()

	// 双重检查
	if val, ok := m.globalMedia.Load(contentHash); ok {
		return val.(*parts.MediaResource), true
	}

	// 生成媒体文件名（image1.png, image2.png...）
	mediaFileID := atomic.AddInt64(&m.mediaFileID, 1)
	ext := filepath.Ext(fileName)
	mediaFileName := "image" + strconv.FormatInt(mediaFileID, 10) + ext

	// 推断 MIME 类型
	contentType := inferContentType(mediaFileName)

	// 生成目标路径
	target := "ppt/media/" + mediaFileName

	// 创建资源
	resource := parts.NewMediaResourceFromBytes(mediaFileName, contentType, target, data)
	resource.SetHash(contentHash)

	// 存储到全局媒体池
	m.globalMedia.Store(contentHash, resource)

	// 同时更新传统索引（兼容旧代码）
	m.byHash.Store(contentHash, contentHash)

	return resource, false
}

// getOrCreateSlideIndex 获取或创建幻灯片媒体索引
func (m *MediaManager) getOrCreateSlideIndex(slideIndex int) *SlideMediaIndex {
	if val, ok := m.slideRelations.Load(slideIndex); ok {
		return val.(*SlideMediaIndex)
	}

	m.writeMu.Lock()
	defer m.writeMu.Unlock()

	// 双重检查
	if val, ok := m.slideRelations.Load(slideIndex); ok {
		return val.(*SlideMediaIndex)
	}

	indexer := NewSlideMediaIndex(slideIndex)
	m.slideRelations.Store(slideIndex, indexer)
	return indexer
}

// GetSlideMediaIndex 获取幻灯片媒体索引
func (m *MediaManager) GetSlideMediaIndex(slideIndex int) *SlideMediaIndex {
	if val, ok := m.slideRelations.Load(slideIndex); ok {
		return val.(*SlideMediaIndex)
	}
	return nil
}

// GetGlobalMediaByHash 根据 Hash 获取全局媒体资源
func (m *MediaManager) GetGlobalMediaByHash(hash string) *parts.MediaResource {
	if val, ok := m.globalMedia.Load(hash); ok {
		return val.(*parts.MediaResource)
	}
	return nil
}

// AllGlobalMedia 返回所有全局媒体资源（去重后的）
func (m *MediaManager) AllGlobalMedia() []*parts.MediaResource {
	result := make([]*parts.MediaResource, 0)
	m.globalMedia.Range(func(key, value any) bool {
		result = append(result, value.(*parts.MediaResource))
		return true
	})
	return result
}

// GlobalMediaCount 返回全局媒体资源数量（去重后）
func (m *MediaManager) GlobalMediaCount() int64 {
	var count int64
	m.globalMedia.Range(func(key, value any) bool {
		count++
		return true
	})
	return count
}

// SlideCount 返回引用媒体的幻灯片数量
func (m *MediaManager) SlideCount() int64 {
	var count int64
	m.slideRelations.Range(func(key, value any) bool {
		count++
		return true
	})
	return count
}

// ============================================================================
// SlideMediaIndex 方法
// ============================================================================

// getOrCreateLocalRID 获取或创建本地 rId
func (smi *SlideMediaIndex) getOrCreateLocalRID(contentHash string, resource *parts.MediaResource) (string, bool) {
	// 检查该幻灯片是否已引用此媒体
	if val, ok := smi.hashToLocal.Load(contentHash); ok {
		return val.(string), true
	}

	// 生成本地 rId
	id := atomic.AddInt64(&smi.nextLocalID, 1)
	localRID := "rId" + strconv.FormatInt(id, 10)

	// 建立映射
	smi.hashToLocal.Store(contentHash, localRID)
	smi.localToHash.Store(localRID, contentHash)

	// 设置资源的 rId（这是该幻灯片的本地 rId）
	resource.SetRID(localRID)

	return localRID, false
}

// GetLocalRIDByHash 根据 Hash 获取本地 rId
func (smi *SlideMediaIndex) GetLocalRIDByHash(hash string) string {
	if val, ok := smi.hashToLocal.Load(hash); ok {
		return val.(string)
	}
	return ""
}

// GetHashByLocalRID 根据本地 rId 获取 Hash
func (smi *SlideMediaIndex) GetHashByLocalRID(localRID string) string {
	if val, ok := smi.localToHash.Load(localRID); ok {
		return val.(string)
	}
	return ""
}

// AllLocalRIDs 返回所有本地 rId
func (smi *SlideMediaIndex) AllLocalRIDs() []string {
	result := make([]string, 0)
	smi.hashToLocal.Range(func(key, value any) bool {
		result = append(result, value.(string))
		return true
	})
	return result
}

// LocalRefCount 返回本地引用数量
func (smi *SlideMediaIndex) LocalRefCount() int64 {
	var count int64
	smi.hashToLocal.Range(func(key, value any) bool {
		count++
		return true
	})
	return count
}

// ============================================================================
// 统计与诊断方法
// ============================================================================

// DeduplicationStats 去重统计信息
type DeduplicationStats struct {
	// 全局媒体数量（实际存储）
	GlobalMediaCount int64

	// 总引用次数（所有幻灯片的引用总和）
	TotalReferences int64

	// 幻灯片数量
	SlideCount int64

	// 节省的存储空间（字节）
	SavedBytes int64

	// 去重率（0.0 - 1.0）
	DeduplicationRate float64
}

// GetDeduplicationStats 获取去重统计信息
func (m *MediaManager) GetDeduplicationStats() DeduplicationStats {
	stats := DeduplicationStats{}

	// 统计全局媒体数量
	stats.GlobalMediaCount = m.GlobalMediaCount()

	// 统计幻灯片数量
	stats.SlideCount = m.SlideCount()

	// 统计总引用次数
	m.slideRelations.Range(func(key, value any) bool {
		smi := value.(*SlideMediaIndex)
		stats.TotalReferences += smi.LocalRefCount()
		return true
	})

	// 计算去重率
	if stats.TotalReferences > 0 {
		stats.DeduplicationRate = 1.0 - float64(stats.GlobalMediaCount)/float64(stats.TotalReferences)
	}

	// 计算节省的字节数
	m.globalMedia.Range(func(key, value any) bool {
		res := value.(*parts.MediaResource)
		// 统计该媒体被引用的次数
		refCount := int64(0)
		m.slideRelations.Range(func(k, v any) bool {
			smi := v.(*SlideMediaIndex)
			if smi.GetLocalRIDByHash(key.(string)) != "" {
				refCount++
			}
			return true
		})
		if refCount > 1 {
			stats.SavedBytes += int64(len(res.Data())) * (refCount - 1)
		}
		return true
	})

	return stats
}
