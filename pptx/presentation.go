// Package pptx 提供 PPTX 文件的高级操作接口
// 作为人类开发者和 AI 调用的绝对入口
package pptx

import (
	"bytes"
	"fmt"
	"io"
	"sync"
	"sync/atomic"

	"github.com/Muprprpr/Go-pptx/opc"
	"github.com/Muprprpr/Go-pptx/parts"
)

// ============================================================================
// Presentation - 总控门面
// ============================================================================
//
// 这是人类开发者和 AI 调用的绝对入口。
// 负责包装底层的 opc.Package，提供高层业务方法。
//
// 核心职责：
// 1. 从模板懒加载并 Clone 出线程安全的副本
// 2. 为每个实例创建专用的 MediaManager（防止并发图片串线）
// 3. 自动管理幻灯片 ID 注册和 .rels 路由
// 4. 提供流式输出能力（支持 HTTP 响应）
//
// 使用示例：
//
//	pres := pptx.New()
//	slide1 := pres.AddSlide()
//	slide1.AddTextBox(100, 100, 500, 50, "Hello World")
//	pres.Save("output.pptx")
//
// ============================================================================

// Presentation PPTX 演示文稿总控门面
type Presentation struct {
	// 底层 OPC 包
	pkg *opc.Package

	// 演示文稿部件（presentation.xml）
	presentationPart *parts.PresentationPart

	// 幻灯片列表（按顺序）
	slides []*Slide

	// 媒体管理器（实例专用，防止并发串线）
	mediaManager *MediaManager

	// 母版管理器
	masterManager *MasterManager

	// 母版缓存（从模板加载的母版/布局信息）
	masterCache *MasterCache

	// 幻灯片计数器（用于生成 slide1.xml, slide2.xml 等）
	slideCounter int32

	// 图表计数器（用于生成 chart1.xml, chart2.xml 等）
	chartCounter int32

	// 关系 ID 计数器
	relCounter int32

	// 并发安全锁
	mu sync.RWMutex
}

// ============================================================================
// 构造函数
// ============================================================================

// New 创建空白演示文稿
// 使用默认模板（16:9 宽屏）
func New() *Presentation {
	pres := &Presentation{
		pkg:              opc.NewPackage(),
		presentationPart: parts.NewPresentationPart(),
		slides:           make([]*Slide, 0),
		mediaManager:     NewMediaManager(),
		masterManager:    NewMasterManager(),
		slideCounter:     0,
		relCounter:       0,
	}

	// 初始化包结构
	pres.initPackage()

	return pres
}

// NewWithTemplate 从模板创建演示文稿
// name: 模板名称（如 TemplateDefault, TemplateBlank 等）
func NewWithTemplate(name TemplateType) (*Presentation, error) {
	// 从模板管理器加载模板
	pkg, err := globalTemplateManager.LoadTemplate(name)
	if err != nil {
		return nil, fmt.Errorf("加载模板失败: %w", err)
	}

	pres := &Presentation{
		pkg:           pkg,
		slides:        make([]*Slide, 0),
		mediaManager:  NewMediaManager(),
		masterManager: NewMasterManager(),
		slideCounter:  0,
		relCounter:    0,
	}

	// 从包中解析 presentation.xml
	if err := pres.loadPresentationPart(); err != nil {
		return nil, fmt.Errorf("解析演示文稿部件失败: %w", err)
	}

	// 获取母版缓存
	pres.masterCache = globalTemplateManager.GetMasterCache()

	return pres, nil
}

// NewFromBytes 从字节数据创建演示文稿
func NewFromBytes(data []byte) (*Presentation, error) {
	reader := bytes.NewReader(data)
	pkg, err := opc.Open(reader, int64(len(data)))
	if err != nil {
		return nil, fmt.Errorf("解析 PPTX 数据失败: %w", err)
	}

	pres := &Presentation{
		pkg:           pkg,
		slides:        make([]*Slide, 0),
		mediaManager:  NewMediaManager(),
		masterManager: NewMasterManager(),
		slideCounter:  0,
		relCounter:    0,
	}

	if err := pres.loadPresentationPart(); err != nil {
		return nil, fmt.Errorf("解析演示文稿部件失败: %w", err)
	}

	return pres, nil
}

// NewFromFile 从文件创建演示文稿
func NewFromFile(path string) (*Presentation, error) {
	pkg, err := opc.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf("打开 PPTX 文件失败: %w", err)
	}

	pres := &Presentation{
		pkg:           pkg,
		slides:        make([]*Slide, 0),
		mediaManager:  NewMediaManager(),
		masterManager: NewMasterManager(),
		slideCounter:  0,
		relCounter:    0,
	}

	if err := pres.loadPresentationPart(); err != nil {
		return nil, fmt.Errorf("解析演示文稿部件失败: %w", err)
	}

	return pres, nil
}

// ============================================================================
// 初始化方法
// ============================================================================

// initPackage 初始化 OPC 包结构
func (p *Presentation) initPackage() {
	// 添加 presentation.xml
	uri := opc.NewPackURI("/ppt/presentation.xml")
	blob, _ := p.presentationPart.ToXML()
	part := opc.NewPart(uri, opc.ContentTypePresentation, blob)
	_ = p.pkg.AddPart(part)

	// 添加包级别关系（指向 presentation.xml）
	_, _ = p.pkg.AddRelationship(opc.RelTypeOfficeDocument, "ppt/presentation.xml", false)
}

// loadPresentationPart 从包中加载 presentation.xml
func (p *Presentation) loadPresentationPart() error {
	// 通过关系类型获取 presentation 部件
	part := p.pkg.GetPartByRelType(opc.RelTypeOfficeDocument)
	if part == nil {
		// 如果不存在，创建新的
		p.presentationPart = parts.NewPresentationPart()
		return nil
	}

	// 解析 XML
	p.presentationPart = parts.NewPresentationPart()
	if err := p.presentationPart.FromXML(part.Blob()); err != nil {
		return err
	}

	// 更新幻灯片计数器
	p.slideCounter = int32(p.presentationPart.SlideCount())

	return nil
}

// ============================================================================
// 幻灯片管理 - 核心方法
// ============================================================================

// AddSlide 添加新幻灯片
// layout: 可选的布局名称（如 "title", "blank", "titleAndContent" 等）
// 如果不指定布局，使用空白布局
//
// 内部自动：
// - 向 presentation.xml 的 <p:sldIdLst> 注册新 ID
// - 向 .rels 申请路由
// - 返回高层的 *Slide 对象
func (p *Presentation) AddSlide(layout ...string) *Slide {
	p.mu.Lock()
	defer p.mu.Unlock()

	// 分配幻灯片编号
	slideNum := int(atomic.AddInt32(&p.slideCounter, 1))

	// 创建幻灯片部件
	slidePart := parts.NewSlidePart(slideNum)

	// 确定布局
	layoutRId := ""
	if len(layout) > 0 && layout[0] != "" {
		// 查找布局
		if p.masterCache != nil {
			if layoutData, ok := p.masterCache.GetLayoutByName(layout[0]); ok {
				// 创建布局关系
				layoutRId = p.allocateRelID()
				// TODO: 添加布局关系到幻灯片
				_ = layoutData
			}
		}
	}
	slidePart.SetLayoutRId(layoutRId)

	// 创建幻灯片 URI
	slideURI := opc.NewPackURI(fmt.Sprintf("/ppt/slides/slide%d.xml", slideNum))
	slidePart.SetURI(slideURI)

	// 添加到包
	slideBlob, _ := slidePart.ToXML()
	slidePartOPC := opc.NewPart(slideURI, opc.ContentTypeSlide, slideBlob)
	_ = p.pkg.AddPart(slidePartOPC)

	// 添加关系到 presentation.xml
	slideRelID := p.allocateRelID()
	_ = slideRelID // 关系 ID 由 PresentationPart 内部管理

	// 注册到 PresentationPart（内部会自动分配 slide ID）
	_ = p.presentationPart.AddSlide(layoutRId, slidePart)

	// 创建高层 Slide 对象
	s := &Slide{
		presentation: p,
		part:         slidePart,
		builder:      NewSlideBuilder(slidePart),
		mediaManager: p.mediaManager,
		index:        len(p.slides),
	}
	// 初始化原子计数器（OOXML 规范中 shapeId 从 2 开始，1 预留给根节点）
	s.shapeIDCounter.Store(1)

	p.slides = append(p.slides, s)

	return s
}

// AddSlideAt 在指定位置插入幻灯片
func (p *Presentation) AddSlideAt(index int, layout ...string) (*Slide, error) {
	p.mu.Lock()
	defer p.mu.Unlock()

	if index < 0 || index > len(p.slides) {
		return nil, fmt.Errorf("索引 %d 超出范围 [0, %d]", index, len(p.slides))
	}

	// 创建幻灯片
	slideNum := int(atomic.AddInt32(&p.slideCounter, 1))
	slidePart := parts.NewSlidePart(slideNum)

	// 确定布局
	layoutRId := ""
	if len(layout) > 0 && layout[0] != "" && p.masterCache != nil {
		if layoutData, ok := p.masterCache.GetLayoutByName(layout[0]); ok {
			_ = layoutData // TODO: 设置布局关系
		}
	}
	slidePart.SetLayoutRId(layoutRId)

	// 设置 URI
	slideURI := opc.NewPackURI(fmt.Sprintf("/ppt/slides/slide%d.xml", slideNum))
	slidePart.SetURI(slideURI)

	// 添加到包
	slideBlob, _ := slidePart.ToXML()
	_ = p.pkg.AddPart(opc.NewPart(slideURI, opc.ContentTypeSlide, slideBlob))

	// 注册到 PresentationPart
	_ = p.presentationPart.AddSlide(layoutRId, slidePart)

	// 创建高层对象
	s := &Slide{
		presentation: p,
		part:         slidePart,
		builder:      NewSlideBuilder(slidePart),
		mediaManager: p.mediaManager,
		index:        index,
	}
	// 初始化原子计数器（OOXML 规范中 shapeId 从 2 开始，1 预留给根节点）
	s.shapeIDCounter.Store(1)

	// 插入到指定位置
	p.slides = append(p.slides[:index], append([]*Slide{s}, p.slides[index:]...)...)

	// 更新后续幻灯片的索引
	for i := index + 1; i < len(p.slides); i++ {
		p.slides[i].index = i
	}

	return s, nil
}

// GetSlide 获取指定索引的幻灯片
func (p *Presentation) GetSlide(index int) (*Slide, error) {
	p.mu.RLock()
	defer p.mu.RUnlock()

	if index < 0 || index >= len(p.slides) {
		return nil, fmt.Errorf("索引 %d 超出范围 [0, %d)", index, len(p.slides))
	}

	return p.slides[index], nil
}

// RemoveSlide 移除指定索引的幻灯片
func (p *Presentation) RemoveSlide(index int) error {
	p.mu.Lock()
	defer p.mu.Unlock()

	if index < 0 || index >= len(p.slides) {
		return fmt.Errorf("索引 %d 超出范围 [0, %d)", index, len(p.slides))
	}

	// 获取要移除的幻灯片
	s := p.slides[index]

	// 从包中移除部件
	_ = p.pkg.RemovePart(s.part.PartURI())

	// 从 presentation.xml 中移除
	_ = p.presentationPart.RemoveSlide(index)

	// 从切片中移除
	p.slides = append(p.slides[:index], p.slides[index+1:]...)

	// 更新后续幻灯片的索引
	for i := index; i < len(p.slides); i++ {
		p.slides[i].index = i
	}

	return nil
}

// SlideCount 返回幻灯片数量
func (p *Presentation) SlideCount() int {
	p.mu.RLock()
	defer p.mu.RUnlock()
	return len(p.slides)
}

// Slides 返回所有幻灯片
func (p *Presentation) Slides() []*Slide {
	p.mu.RLock()
	defer p.mu.RUnlock()

	result := make([]*Slide, len(p.slides))
	copy(result, p.slides)
	return result
}

// ============================================================================
// 保存方法
// ============================================================================

// Save 将演示文稿保存到文件
// 触发底层的序列化和 ZIP 打包
func (p *Presentation) Save(path string) error {
	p.mu.RLock()
	defer p.mu.RUnlock()

	// 同步所有幻灯片到包
	if err := p.syncSlides(); err != nil {
		return err
	}

	// 同步 presentation.xml
	if err := p.syncPresentationPart(); err != nil {
		return err
	}

	// 保存到文件
	return p.pkg.SaveFile(path)
}

// Write 将演示文稿写入 io.Writer
// 这是为高并发流式输出（如 HTTP 响应）准备的杀手锏
func (p *Presentation) Write(w io.Writer) error {
	p.mu.RLock()
	defer p.mu.RUnlock()

	// 同步所有幻灯片到包
	if err := p.syncSlides(); err != nil {
		return err
	}

	// 同步 presentation.xml
	if err := p.syncPresentationPart(); err != nil {
		return err
	}

	// 写入到 Writer
	return p.pkg.Save(w)
}

// WriteToBytes 将演示文稿写入字节数组
func (p *Presentation) WriteToBytes() ([]byte, error) {
	p.mu.RLock()
	defer p.mu.RUnlock()

	// 同步所有幻灯片到包
	if err := p.syncSlides(); err != nil {
		return nil, err
	}

	// 同步 presentation.xml
	if err := p.syncPresentationPart(); err != nil {
		return nil, err
	}

	// 写入到字节数组
	return p.pkg.SaveToBytes()
}

// ============================================================================
// 同步方法
// ============================================================================

// syncSlides 同步所有幻灯片到 OPC 包
func (p *Presentation) syncSlides() error {
	for _, s := range p.slides {
		blob, err := s.part.ToXML()
		if err != nil {
			return fmt.Errorf("序列化幻灯片 %d 失败: %w", s.index+1, err)
		}

		// 更新或创建部件
		uri := s.part.PartURI()
		existingPart := p.pkg.GetPart(uri)
		if existingPart != nil {
			existingPart.SetBlob(blob)
		} else {
			part := opc.NewPart(uri, opc.ContentTypeSlide, blob)
			_ = p.pkg.AddPart(part)
		}
	}

	return nil
}

// syncPresentationPart 同步 presentation.xml 到 OPC 包
func (p *Presentation) syncPresentationPart() error {
	blob, err := p.presentationPart.ToXML()
	if err != nil {
		return fmt.Errorf("序列化 presentation.xml 失败: %w", err)
	}

	uri := opc.NewPackURI("/ppt/presentation.xml")
	existingPart := p.pkg.GetPart(uri)
	if existingPart != nil {
		existingPart.SetBlob(blob)
	} else {
		part := opc.NewPart(uri, opc.ContentTypePresentation, blob)
		_ = p.pkg.AddPart(part)
	}

	return nil
}

// ============================================================================
// 辅助方法
// ============================================================================

// allocateRelID 分配关系 ID
func (p *Presentation) allocateRelID() string {
	id := atomic.AddInt32(&p.relCounter, 1)
	return fmt.Sprintf("rId%d", id)
}

// Package 返回底层 OPC 包（高级用法）
func (p *Presentation) Package() *opc.Package {
	return p.pkg
}

// PresentationPart 返回演示文稿部件
func (p *Presentation) PresentationPart() *parts.PresentationPart {
	return p.presentationPart
}

// MediaManager 返回媒体管理器
func (p *Presentation) MediaManager() *MediaManager {
	return p.mediaManager
}

// MasterCache 返回母版缓存
func (p *Presentation) MasterCache() *MasterCache {
	return p.masterCache
}

// SetSlideSize 设置幻灯片尺寸
func (p *Presentation) SetSlideSize(cx, cy int) {
	p.mu.Lock()
	defer p.mu.Unlock()
	p.presentationPart.SetSlideSize(parts.SlideSize{Cx: cx, Cy: cy})
}

// SetSlideSizeStandard 设置标准幻灯片尺寸
func (p *Presentation) SetSlideSizeStandard(name string) {
	size := parts.NewSlideSizeFromStandard(name)
	p.SetSlideSize(size.Cx, size.Cy)
}

// SlideSize 返回当前幻灯片尺寸
func (p *Presentation) SlideSize() (int, int) {
	size := p.presentationPart.SlideSize()
	return size.Cx, size.Cy
}

// Close 关闭演示文稿，释放资源
func (p *Presentation) Close() error {
	p.mu.Lock()
	defer p.mu.Unlock()

	p.slides = nil
	p.mediaManager = nil
	p.masterCache = nil

	return p.pkg.Close()
}

// ============================================================================
// 克隆方法
// ============================================================================

// Clone 克隆演示文稿
// 返回一个完全独立的副本
func (p *Presentation) Clone() (*Presentation, error) {
	p.mu.RLock()
	defer p.mu.RUnlock()

	// 克隆底层包
	newPkg := p.pkg.Clone()

	// 创建新演示文稿
	newPres := &Presentation{
		pkg:           newPkg,
		slides:        make([]*Slide, len(p.slides)),
		mediaManager:  NewMediaManager(),
		masterManager: p.masterManager,
		masterCache:   p.masterCache,
		slideCounter:  p.slideCounter,
		relCounter:    p.relCounter,
	}

	// 克隆 presentation part
	newPres.presentationPart = parts.NewPresentationPart()
	presPartData, err := p.presentationPart.ToXML()
	if err != nil {
		return nil, err
	}
	if err := newPres.presentationPart.FromXML(presPartData); err != nil {
		return nil, err
	}

	// 克隆幻灯片
	for i, s := range p.slides {
		newSlidePart := parts.NewSlidePartWithURI(s.part.PartURI())
		slideData, err := s.part.ToXML()
		if err != nil {
			return nil, err
		}
		if err := newSlidePart.FromXML(slideData); err != nil {
			return nil, err
		}

		newPres.slides[i] = &Slide{
			presentation: newPres,
			part:         newSlidePart,
			builder:      NewSlideBuilder(newSlidePart),
			mediaManager: newPres.mediaManager,
			index:        i,
		}
	}

	return newPres, nil
}
