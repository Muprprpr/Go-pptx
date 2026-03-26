package parts

import (
	"encoding/xml"
	"fmt"
	"sync"
	"sync/atomic"

	"github.com/Muprprpr/Go-pptx/opc"
)

// SlideIDStart 是 Slide ID 的起始值
const SlideIDStart = 256

// SlideSize 幻灯片尺寸
type SlideSize struct {
	Cx int // 宽度，单位 EMU (English Metric Units)
	Cy int // 高度，单位 EMU
}

// StandardSlideSizes 标准幻灯片尺寸
var StandardSlideSizes = struct {
	// 16:9 宽屏 (12192000 x 6858000 EMU)
	Wide16x9 SlideSize
	// 4:3 标准 (9144000 x 6858000 EMU)
	Standard4x3 SlideSize
}{
	Wide16x9:    SlideSize{Cx: 12192000, Cy: 6858000},
	Standard4x3: SlideSize{Cx: 9144000, Cy: 6858000},
}

// PresentationPart 对应 /ppt/presentation.xml
// 是整个 PPTX 逻辑上的根节点
type PresentationPart struct {
	uri *opc.PackURI

	// 幻灯片管理
	slideIDs      []uint32 // 分配过的 slide ID 列表
	slideIDNext   uint32   // 下一个可分配的 slide ID（原子操作）
	slideCount    int32    // 当前幻灯片数量（原子操作）

	// 母版和布局管理
	slideMasterIDs []string // 母版 rId 列表
	slideLayoutIDs []string // 布局 rId 列表（与 slide 一一对应）

	// 全局属性
	slideSize     SlideSize // 幻灯片尺寸
	notesMasterID string    // 备注母版 rId
	themeID       string     // 主题 rId

	mu sync.RWMutex
}

// NewPresentationPart 创建新的演示文稿部件
func NewPresentationPart() *PresentationPart {
	return &PresentationPart{
		uri:          opc.NewPackURI("/ppt/presentation.xml"),
		slideIDs:     make([]uint32, 0),
		slideMasterIDs: make([]string, 0),
		slideLayoutIDs: make([]string, 0),
		slideSize:    StandardSlideSizes.Wide16x9,
		slideIDNext:  SlideIDStart,
	}
}

// NewPresentationPartWithSize 创建演示文稿并设置尺寸
func NewPresentationPartWithSize(size SlideSize) *PresentationPart {
	p := NewPresentationPart()
	p.slideSize = size
	return p
}

// PartURI 返回部件 URI
func (p *PresentationPart) PartURI() *opc.PackURI {
	return p.uri
}

// SlideSize 返回幻灯片尺寸
func (p *PresentationPart) SlideSize() SlideSize {
	p.mu.RLock()
	defer p.mu.RUnlock()
	return p.slideSize
}

// SetSlideSize 设置幻灯片尺寸
func (p *PresentationPart) SetSlideSize(size SlideSize) {
	p.mu.Lock()
	defer p.mu.Unlock()
	p.slideSize = size
}

// SlideCount 返回幻灯片数量
func (p *PresentationPart) SlideCount() int32 {
	return atomic.LoadInt32(&p.slideCount)
}

// SlideIDAt 返回指定索引的 slide ID
func (p *PresentationPart) SlideIDAt(index int) (uint32, error) {
	p.mu.RLock()
	defer p.mu.RUnlock()

	if index < 0 || index >= len(p.slideIDs) {
		return 0, fmt.Errorf("slide index out of range: %d", index)
	}
	return p.slideIDs[index], nil
}

// SlideIDs 返回所有 slide ID
func (p *PresentationPart) SlideIDs() []uint32 {
	p.mu.RLock()
	defer p.mu.RUnlock()

	ids := make([]uint32, len(p.slideIDs))
	copy(ids, p.slideIDs)
	return ids
}

// SlideMasterIDs 返回所有母版 rId
func (p *PresentationPart) SlideMasterIDs() []string {
	p.mu.RLock()
	defer p.mu.RUnlock()

	ids := make([]string, len(p.slideMasterIDs))
	copy(ids, p.slideMasterIDs)
	return ids
}

// AddSlideMaster 添加母版
// 返回分配的 rId
func (p *PresentationPart) AddSlideMaster(rId string) {
	p.mu.Lock()
	defer p.mu.Unlock()
	p.slideMasterIDs = append(p.slideMasterIDs, rId)
}

// allocateSlideID 原子分配一个新的 slide ID
func (p *PresentationPart) allocateSlideID() uint32 {
	return atomic.AddUint32(&p.slideIDNext, 1)
}

// AddSlide 添加幻灯片
// layout 是关联的布局 rId，slidePart 是实际的幻灯片部件
func (p *PresentationPart) AddSlide(layoutRId string, slidePart *SlidePart) error {
	p.mu.Lock()
	defer p.mu.Unlock()

	// 分配 slide ID
	slideID := p.allocateSlideID()

	// 更新内部状态
	p.slideIDs = append(p.slideIDs, slideID)
	p.slideLayoutIDs = append(p.slideLayoutIDs, layoutRId)
	atomic.AddInt32(&p.slideCount, 1)

	return nil
}

// RemoveSlide 移除幻灯片（按索引）
func (p *PresentationPart) RemoveSlide(index int) error {
	p.mu.Lock()
	defer p.mu.Unlock()

	if index < 0 || index >= len(p.slideIDs) {
		return fmt.Errorf("slide index out of range: %d", index)
	}

	// 从切片中移除
	p.slideIDs = append(p.slideIDs[:index], p.slideIDs[index+1:]...)
	p.slideLayoutIDs = append(p.slideLayoutIDs[:index], p.slideLayoutIDs[index+1:]...)
	atomic.AddInt32(&p.slideCount, -1)

	return nil
}

// XPresentation 对应 presentation.xml 的完整 XML 结构
// 注意：XML 标签使用不带前缀的名称，因为解析前会去除命名空间前缀
type XPresentation struct {
	XMLName xml.Name `xml:"presentation"`

	// 兼容设置
	Compatibility *XCompatibility `xml:"compatSpt,omitempty"`

	// 幻灯片尺寸（必须）
	SldSz *XSldSz `xml:"sldSz"`

	// 备注尺寸
	NotesSz *XSldSz `xml:"notesSz,omitempty"`

	// 幻灯片 ID 列表（必须）
	SldIdLst *XSldIdLst `xml:"sldIdLst"`

	// 母版 ID 列表
	SldMasterIdLst *XSldMasterIdLst `xml:"sldMasterIdLst,omitempty"`

	// 备注母版 ID 列表
	NotesMasterIdLst *XSldMasterIdLst `xml:"notesMasterIdLst,omitempty"`

	// 打印设置
	PrintSettings *XPrintSettings `xml:"printSettings,omitempty"`
}

// XCompatibility 兼容设置
type XCompatibility struct {
	CompatMode string `xml:"compatMode,attr,omitempty"`
}

// XSldSz 幻灯片尺寸
type XSldSz struct {
	Cx int `xml:"cx,attr"`
	Cy int `xml:"cy,attr"`
}

// XSldIdLst 幻灯片 ID 列表
type XSldIdLst struct {
	SldIds []XSldId `xml:"sldId"`
}

// XSldId 单个幻灯片 ID
type XSldId struct {
	Id  uint32 `xml:"id,attr"`
	RId string `xml:"rid,attr"`
}

// XSldMasterIdLst 母版 ID 列表
type XSldMasterIdLst struct {
	SldMasterIds []XSldMasterId `xml:"sldMasterId"`
}

// XSldMasterId 单个母版 ID
type XSldMasterId struct {
	Id  uint32 `xml:"id,attr"`
	RId string `xml:"rid,attr"`
}

// XPrintSettings 打印设置
type XPrintSettings struct {
	OutputOptions *XOutputOptions `xml:"outputOptions,omitempty"`
}

// XOutputOptions 输出选项
type XOutputOptions struct {
	UsePrintFml     *bool `xml:"usePrintFml,attr,omitempty"`
	CloneLinkedObjs *bool `xml:"cloneLinkedObjs,attr,omitempty"`
}

// ToXML 将 PresentationPart 序列化为 XML
func (p *PresentationPart) ToXML() ([]byte, error) {
	p.mu.RLock()
	defer p.mu.RUnlock()

	xp := XPresentation{
		SldSz: &XSldSz{
			Cx: p.slideSize.Cx,
			Cy: p.slideSize.Cy,
		},
	}

	// 构建幻灯片列表
	if len(p.slideIDs) > 0 {
		xp.SldIdLst = &XSldIdLst{
			SldIds: make([]XSldId, len(p.slideIDs)),
		}
		for i := range p.slideIDs {
			xp.SldIdLst.SldIds[i] = XSldId{
				Id:  p.slideIDs[i],
				RId: p.slideLayoutIDs[i],
			}
		}
	}

	// 构建母版列表
	if len(p.slideMasterIDs) > 0 {
		xp.SldMasterIdLst = &XSldMasterIdLst{
			SldMasterIds: make([]XSldMasterId, len(p.slideMasterIDs)),
		}
		// 母版 ID 从 1 开始
		for i, rId := range p.slideMasterIDs {
			xp.SldMasterIdLst.SldMasterIds[i] = XSldMasterId{
				Id:  uint32(i + 1),
				RId: rId,
			}
		}
	}

	output, err := xml.Marshal(&xp)
	if err != nil {
		return nil, err
	}
	return append([]byte(XMLDeclaration), output...), nil
}

// FromXML 从 XML 反序列化为 PresentationPart
func (p *PresentationPart) FromXML(data []byte) error {
	// 去除命名空间前缀以兼容 Go 的 xml.Unmarshal
	cleanData, err := StripNamespacePrefixes(data)
	if err != nil {
		return fmt.Errorf("failed to clean XML: %w", err)
	}

	var xp XPresentation
	if err := xml.Unmarshal(cleanData, &xp); err != nil {
		return fmt.Errorf("failed to unmarshal presentation XML: %w", err)
	}

	p.mu.Lock()
	defer p.mu.Unlock()

	// 解析幻灯片尺寸
	if xp.SldSz != nil {
		p.slideSize = SlideSize{
			Cx: xp.SldSz.Cx,
			Cy: xp.SldSz.Cy,
		}
	}

	// 解析幻灯片列表
	p.slideIDs = make([]uint32, 0)
	p.slideLayoutIDs = make([]string, 0)
	if xp.SldIdLst != nil {
		for _, sldId := range xp.SldIdLst.SldIds {
			p.slideIDs = append(p.slideIDs, sldId.Id)
			p.slideLayoutIDs = append(p.slideLayoutIDs, sldId.RId)
		}
	}

	// 更新 slide ID 计数器
	if len(p.slideIDs) > 0 {
		maxID := p.slideIDs[0]
		for _, id := range p.slideIDs {
			if id > maxID {
				maxID = id
			}
		}
		p.slideIDNext = maxID + 1
	}

	// 解析母版列表
	p.slideMasterIDs = make([]string, 0)
	if xp.SldMasterIdLst != nil {
		for _, masterId := range xp.SldMasterIdLst.SldMasterIds {
			p.slideMasterIDs = append(p.slideMasterIDs, masterId.RId)
		}
	}

	// 更新幻灯片计数
	p.slideCount = int32(len(p.slideIDs))

	return nil
}

// Presentation 辅助函数

// NewSlideSizeFromStandard 根据标准尺寸名称创建 SlideSize
func NewSlideSizeFromStandard(name string) SlideSize {
	switch name {
	case "16:9", "wide", "widescreen":
		return StandardSlideSizes.Wide16x9
	case "4:3", "standard":
		return StandardSlideSizes.Standard4x3
	default:
		return StandardSlideSizes.Wide16x9
	}
}

// EMUFromPoints 将磅值转换为 EMU
func EMUFromPoints(points float64) int {
	return int(points * 12700)
}

// PointsFromEMU 将 EMU 转换为磅值
func PointsFromEMU(emu int) float64 {
	return float64(emu) / 12700.0
}

// EMUFromInches 将英寸转换为 EMU
func EMUFromInches(inches float64) int {
	return int(inches * 914400)
}

// InchesFromEMU 将 EMU 转换为英寸
func InchesFromEMU(emu int) float64 {
	return float64(emu) / 914400.0
}

// EMUFromMM 将毫米转换为 EMU
func EMUFromMM(mm float64) int {
	return int(mm * 36000)
}

// MMFromEMU 将 EMU 转换为毫米
func MMFromEMU(emu int) float64 {
	return float64(emu) / 36000.0
}
