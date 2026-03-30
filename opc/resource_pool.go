package opc

import (
	"sync"
)

// ResourcePool 全局资源池，管理可共享的静态资源
// 使用 zero-copy 策略，多个 Package 可以共享相同的二进制数据
type ResourcePool struct {
	mu      sync.RWMutex
	themes  map[string][]byte // theme URI -> blob
	masters map[string][]byte // master URI -> blob
	layouts map[string][]byte // layout URI -> blob
	media   map[string][]byte // media URI -> blob (图片、音视频等)
	fonts   map[string][]byte // font URI -> blob

	// 引用计数，用于追踪资源使用情况
	refCount map[string]int
}

// globalPool 全局资源池单例
var globalPool = &ResourcePool{
	themes:   make(map[string][]byte),
	masters:  make(map[string][]byte),
	layouts:  make(map[string][]byte),
	media:    make(map[string][]byte),
	fonts:    make(map[string][]byte),
	refCount: make(map[string]int),
}

// GetGlobalPool 获取全局资源池
func GetGlobalPool() *ResourcePool {
	return globalPool
}

// GetOrLoad 获取或加载资源（全局唯一实例，zero-copy）
// loader 函数仅在资源不存在时调用一次
func (p *ResourcePool) GetOrLoad(uri string, contentType string, loader func() ([]byte, error)) ([]byte, error) {
	// 先尝试读锁快速路径
	p.mu.RLock()
	if data, ok := p.getResource(uri, contentType); ok {
		p.mu.RUnlock()
		p.incrementRef(uri)
		return data, nil
	}
	p.mu.RUnlock()

	// 需要加载，使用写锁
	p.mu.Lock()
	defer p.mu.Unlock()

	// double check
	if data, ok := p.getResource(uri, contentType); ok {
		p.incrementRefLocked(uri)
		return data, nil
	}

	// 加载资源
	data, err := loader()
	if err != nil {
		return nil, err
	}

	// 存储到对应的池中
	p.storeResource(uri, contentType, data)
	p.incrementRefLocked(uri)

	return data, nil
}

// getResource 从池中获取资源（需持有锁）
func (p *ResourcePool) getResource(uri string, contentType string) ([]byte, bool) {
	var data []byte
	var ok bool

	switch {
	case contentType == ContentTypeTheme || contentType == ContentTypeThemeOverride:
		data, ok = p.themes[uri]
	case contentType == ContentTypeSlideMaster:
		data, ok = p.masters[uri]
	case contentType == ContentTypeSlideLayout:
		data, ok = p.layouts[uri]
	case contentType == ContentTypeFont:
		data, ok = p.fonts[uri]
	case IsLargeBinaryContentType(contentType):
		data, ok = p.media[uri]
	default:
		return nil, false
	}

	return data, ok
}

// storeResource 存储资源到池中（需持有写锁）
func (p *ResourcePool) storeResource(uri string, contentType string, data []byte) {
	switch {
	case contentType == ContentTypeTheme || contentType == ContentTypeThemeOverride:
		p.themes[uri] = data
	case contentType == ContentTypeSlideMaster:
		p.masters[uri] = data
	case contentType == ContentTypeSlideLayout:
		p.layouts[uri] = data
	case contentType == ContentTypeFont:
		p.fonts[uri] = data
	case IsLargeBinaryContentType(contentType):
		p.media[uri] = data
	}
}

// incrementRef 增加引用计数（需持有读锁）
func (p *ResourcePool) incrementRef(uri string) {
	p.mu.Lock()
	p.refCount[uri]++
	p.mu.Unlock()
}

// incrementRefLocked 增加引用计数（已持有写锁）
func (p *ResourcePool) incrementRefLocked(uri string) {
	p.refCount[uri]++
}

// Release 释放资源引用（引用计数减一）
// 当引用计数归零时，资源会被移除
func (p *ResourcePool) Release(uri string) {
	p.mu.Lock()
	defer p.mu.Unlock()

	if count, ok := p.refCount[uri]; ok {
		if count <= 1 {
			// 引用计数归零，移除资源
			delete(p.refCount, uri)
			delete(p.themes, uri)
			delete(p.masters, uri)
			delete(p.layouts, uri)
			delete(p.media, uri)
			delete(p.fonts, uri)
		} else {
			p.refCount[uri] = count - 1
		}
	}
}

// ReleaseAll 释放所有资源（慎用！）
func (p *ResourcePool) ReleaseAll() {
	p.mu.Lock()
	defer p.mu.Unlock()

	p.themes = make(map[string][]byte)
	p.masters = make(map[string][]byte)
	p.layouts = make(map[string][]byte)
	p.media = make(map[string][]byte)
	p.fonts = make(map[string][]byte)
	p.refCount = make(map[string]int)
}

// Stats 返回资源池统计信息
func (p *ResourcePool) Stats() map[string]int {
	p.mu.RLock()
	defer p.mu.RUnlock()

	return map[string]int{
		"themes":  len(p.themes),
		"masters": len(p.masters),
		"layouts": len(p.layouts),
		"media":   len(p.media),
		"fonts":   len(p.fonts),
		"total":   len(p.themes) + len(p.masters) + len(p.layouts) + len(p.media) + len(p.fonts),
	}
}

// Prefetch 预加载资源到池中
// 用于提前加载已知会使用的资源，避免运行时加载延迟
func (p *ResourcePool) Prefetch(resources map[string]func() ([]byte, error)) error {
	for uri, loader := range resources {
		if _, err := p.GetOrLoad(uri, "", loader); err != nil {
			return err
		}
	}
	return nil
}

// CreateSharedPart 从资源池创建共享部件
// 如果资源不在池中，会使用 loader 加载
func (p *ResourcePool) CreateSharedPart(uri *PackURI, contentType string, loader func() ([]byte, error)) (*Part, error) {
	data, err := p.GetOrLoad(uri.URI(), contentType, loader)
	if err != nil {
		return nil, err
	}
	return NewSharedPart(uri, contentType, data), nil
}
