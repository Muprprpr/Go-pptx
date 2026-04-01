package parts

// ============================================================================
// AppPropsPart - 应用程序属性部件
// ============================================================================
//
// 对应 /docProps/app.xml
// 包含应用程序相关的元数据（公司、管理者、幻灯片数等）
//
// ============================================================================

import (
	"encoding/xml"
	"fmt"
	"sync"

	"github.com/Muprprpr/Go-pptx/opc"
)

// AppPropsPart 应用程序属性部件
type AppPropsPart struct {
	uri       *opc.PackURI
	appProps  *XMLAppProps
	mu        sync.RWMutex
}

// NewAppPropsPart 创建新的应用程序属性部件
func NewAppPropsPart() *AppPropsPart {
	return &AppPropsPart{
		uri: opc.NewPackURI("/docProps/app.xml"),
		appProps: &XMLAppProps{
			Application: "Microsoft Office PowerPoint",
			AppVersion:  "15.0000",
		},
	}
}

// PartURI 返回部件 URI
func (a *AppPropsPart) PartURI() *opc.PackURI {
	return a.uri
}

// AppProps 返回应用属性数据
func (a *AppPropsPart) AppProps() *XMLAppProps {
	a.mu.RLock()
	defer a.mu.RUnlock()
	return a.appProps
}

// ============================================================================
// 属性访问方法
// ============================================================================

// GetAppApplication 获取应用程序名称
func (a *AppPropsPart) GetAppApplication() string {
	a.mu.RLock()
	defer a.mu.RUnlock()
	if a.appProps == nil {
		return ""
	}
	return a.appProps.Application
}

// GetAppVersion 获取应用程序版本
func (a *AppPropsPart) GetAppVersion() string {
	a.mu.RLock()
	defer a.mu.RUnlock()
	if a.appProps == nil {
		return ""
	}
	return a.appProps.AppVersion
}

// GetAppCompany 获取公司名称
func (a *AppPropsPart) GetAppCompany() string {
	a.mu.RLock()
	defer a.mu.RUnlock()
	if a.appProps == nil {
		return ""
	}
	return a.appProps.Company
}

// GetAppManager 获取管理者
func (a *AppPropsPart) GetAppManager() string {
	a.mu.RLock()
	defer a.mu.RUnlock()
	if a.appProps == nil {
		return ""
	}
	return a.appProps.Manager
}

// GetAppSlideCount 获取幻灯片数量
func (a *AppPropsPart) GetAppSlideCount() int {
	a.mu.RLock()
	defer a.mu.RUnlock()
	if a.appProps == nil || a.appProps.Slides == nil {
		return 0
	}
	return *a.appProps.Slides
}

// GetAppWordCount 获取字数
func (a *AppPropsPart) GetAppWordCount() int {
	a.mu.RLock()
	defer a.mu.RUnlock()
	if a.appProps == nil || a.appProps.Words == nil {
		return 0
	}
	return *a.appProps.Words
}

// GetAppTotalTime 获取总编辑时间（分钟）
func (a *AppPropsPart) GetAppTotalTime() int {
	a.mu.RLock()
	defer a.mu.RUnlock()
	if a.appProps == nil || a.appProps.TotalTime == nil {
		return 0
	}
	return *a.appProps.TotalTime
}

// ============================================================================
// 属性设置方法
// ============================================================================

// SetAppCompany 设置公司名称
func (a *AppPropsPart) SetAppCompany(company string) {
	a.mu.Lock()
	defer a.mu.Unlock()
	a.ensureAppProps()
	a.appProps.Company = company
}

// SetAppManager 设置管理者
func (a *AppPropsPart) SetAppManager(manager string) {
	a.mu.Lock()
	defer a.mu.Unlock()
	a.ensureAppProps()
	a.appProps.Manager = manager
}

// SetAppApplication 设置应用程序名称
func (a *AppPropsPart) SetAppApplication(app string) {
	a.mu.Lock()
	defer a.mu.Unlock()
	a.ensureAppProps()
	a.appProps.Application = app
}

// SetAppVersion 设置应用程序版本
func (a *AppPropsPart) SetAppVersion(version string) {
	a.mu.Lock()
	defer a.mu.Unlock()
	a.ensureAppProps()
	a.appProps.AppVersion = version
}

// SetAppSlideCount 设置幻灯片数量
func (a *AppPropsPart) SetAppSlideCount(count int) {
	a.mu.Lock()
	defer a.mu.Unlock()
	a.ensureAppProps()
	a.appProps.Slides = &count
}

// SetAppWordCount 设置字数
func (a *AppPropsPart) SetAppWordCount(count int) {
	a.mu.Lock()
	defer a.mu.Unlock()
	a.ensureAppProps()
	a.appProps.Words = &count
}

// SetAppTotalTime 设置总编辑时间（分钟）
func (a *AppPropsPart) SetAppTotalTime(minutes int) {
	a.mu.Lock()
	defer a.mu.Unlock()
	a.ensureAppProps()
	a.appProps.TotalTime = &minutes
}

// SetAppNoteCount 设置备注数量
func (a *AppPropsPart) SetAppNoteCount(count int) {
	a.mu.Lock()
	defer a.mu.Unlock()
	a.ensureAppProps()
	a.appProps.Notes = &count
}

// SetAppHiddenSlideCount 设置隐藏幻灯片数量
func (a *AppPropsPart) SetAppHiddenSlideCount(count int) {
	a.mu.Lock()
	defer a.mu.Unlock()
	a.ensureAppProps()
	a.appProps.HiddenSlides = &count
}

// SetAppMMClipCount 设置多媒体剪辑数量
func (a *AppPropsPart) SetAppMMClipCount(count int) {
	a.mu.Lock()
	defer a.mu.Unlock()
	a.ensureAppProps()
	a.appProps.MMClips = &count
}

// SetAppParagraphCount 设置段落数量
func (a *AppPropsPart) SetAppParagraphCount(count int) {
	a.mu.Lock()
	defer a.mu.Unlock()
	a.ensureAppProps()
	a.appProps.Paragraphs = &count
}

// SetAppCharacterCount 设置字符数量
func (a *AppPropsPart) SetAppCharacterCount(count int) {
	a.mu.Lock()
	defer a.mu.Unlock()
	a.ensureAppProps()
	a.appProps.Characters = &count
}

// SetAppHyperlinkBase 设置超链接基础
func (a *AppPropsPart) SetAppHyperlinkBase(base string) {
	a.mu.Lock()
	defer a.mu.Unlock()
	a.ensureAppProps()
	a.appProps.HyperlinkBase = base
}

// SetAppLinksUpToDate 设置链接是否最新
func (a *AppPropsPart) SetAppLinksUpToDate(upToDate bool) {
	a.mu.Lock()
	defer a.mu.Unlock()
	a.ensureAppProps()
	a.appProps.LinksUpToDate = &upToDate
}

// SetAppSharedDoc 设置是否共享文档
func (a *AppPropsPart) SetAppSharedDoc(shared bool) {
	a.mu.Lock()
	defer a.mu.Unlock()
	a.ensureAppProps()
	a.appProps.SharedDoc = &shared
}

// SetAppHeadingPairs 设置标题对
func (a *AppPropsPart) SetAppHeadingPairs(headingPairs *XMLHeadingPairs) {
	a.mu.Lock()
	defer a.mu.Unlock()
	a.ensureAppProps()
	a.appProps.HeadingPairs = headingPairs
}

// SetAppTitlesOfParts 设置部件标题
func (a *AppPropsPart) SetAppTitlesOfParts(titlesOfParts *XMLTitlesOfParts) {
	a.mu.Lock()
	defer a.mu.Unlock()
	a.ensureAppProps()
	a.appProps.TitlesOfParts = titlesOfParts
}

// ensureAppProps 确保应用属性结构存在（调用时需持有锁）
func (a *AppPropsPart) ensureAppProps() {
	if a.appProps == nil {
		a.appProps = &XMLAppProps{
			Application: "Microsoft Office PowerPoint",
			AppVersion:  "15.0000",
		}
	}
}

// ============================================================================
// XML 序列化/反序列化
// ============================================================================

// ToXML 将应用属性序列化为 XML
func (a *AppPropsPart) ToXML() ([]byte, error) {
	a.mu.RLock()
	defer a.mu.RUnlock()

	if a.appProps == nil {
		return nil, fmt.Errorf("app data is nil")
	}

	// 设置命名空间
	a.appProps.XmlnsProp = NamespaceExtendedProperties
	a.appProps.XmlnsVt = NamespaceDocPropsVTypes

	output, err := xml.MarshalIndent(a.appProps, "", "  ")
	if err != nil {
		return nil, err
	}

	return append([]byte(XMLDeclaration), output...), nil
}

// FromXML 从 XML 反序列化应用属性
func (a *AppPropsPart) FromXML(data []byte) error {
	var app XMLAppProps
	if err := xml.Unmarshal(data, &app); err != nil {
		return fmt.Errorf("failed to unmarshal app XML: %w", err)
	}

	a.mu.Lock()
	defer a.mu.Unlock()
	a.appProps = &app

	return nil
}

// ParseAppProps 从 XML 字节解析应用属性
func ParseAppProps(data []byte) (*XMLAppProps, error) {
	var app XMLAppProps
	if err := xml.Unmarshal(data, &app); err != nil {
		return nil, fmt.Errorf("failed to parse app: %w", err)
	}
	return &app, nil
}
