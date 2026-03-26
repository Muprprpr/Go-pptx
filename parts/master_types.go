package parts

// ============================================================================
// PPT 母版和版式核心数据结构
// ============================================================================
//
// 设计原则：
// 1. 所有结构体字段均为只读（小写字段通过构造函数初始化，大写字段为不可变值）
// 2. 针对高并发读取优化，无需加锁即可安全读取
// 3. 数据在解析时一次性构建，之后不再修改
// ============================================================================

// PlaceholderType 占位符类型枚举
// 对应 XML: <p:ph type="...">
type PlaceholderType int8

const (
	PlaceholderTypeNone      PlaceholderType = iota // 未指定
	PlaceholderTypeTitle                            // 标题
	PlaceholderTypeBody                             // 正文/内容
	PlaceholderTypeCenterTitle                      // 居中标题
	PlaceholderTypeSubTitle                         // 副标题
	PlaceholderTypeDateTime                         // 日期时间
	PlaceholderTypeSlideNumber                      // 幻灯片编号
	PlaceholderTypeFooter                           // 页脚
	PlaceholderTypeHeader                           // 页眉
	PlaceholderTypeObject                           // 对象
	PlaceholderTypeChart                            // 图表
	PlaceholderTypeTable                            // 表格
	PlaceholderTypeClipArt                          // 剪贴画
	PlaceholderTypeOrgChart                         // 组织结构图
	PlaceholderTypeMedia                            // 媒体
	PlaceholderTypeSlideImage                       // 幻灯片图像
	PlaceholderTypePicture                          // 图片
)

// String 返回占位符类型的字符串表示
func (t PlaceholderType) String() string {
	switch t {
	case PlaceholderTypeTitle:
		return "title"
	case PlaceholderTypeBody:
		return "body"
	case PlaceholderTypeCenterTitle:
		return "ctrTitle"
	case PlaceholderTypeSubTitle:
		return "subTitle"
	case PlaceholderTypeDateTime:
		return "dt"
	case PlaceholderTypeSlideNumber:
		return "sldNum"
	case PlaceholderTypeFooter:
		return "ftr"
	case PlaceholderTypeHeader:
		return "hdr"
	case PlaceholderTypeObject:
		return "obj"
	case PlaceholderTypeChart:
		return "chart"
	case PlaceholderTypeTable:
		return "tbl"
	case PlaceholderTypeClipArt:
		return "clipArt"
	case PlaceholderTypeOrgChart:
		return "dgm"
	case PlaceholderTypeMedia:
		return "media"
	case PlaceholderTypeSlideImage:
		return "sldImg"
	case PlaceholderTypePicture:
		return "pic"
	default:
		return ""
	}
}

// TextStyle 文本样式（只读）
// 用于定义占位符中文本的默认字体、大小、颜色等
type TextStyle struct {
	fontName  string // 字体名称
	fontSize  int32  // 字体大小（单位：百分之一磅，100 = 1pt）
	bold      bool   // 是否粗体
	italic    bool   // 是否斜体
	underline bool   // 是否下划线
	colorRGB  string // 文本颜色（RGB 十六进制，如 "FF0000"）
}

// Placeholder 占位符 - 母版/版式中定义的可填充区域（只读）
// 对应 XML: <p:sp> with <p:nvSpPr><p:nvPr><p:ph ...>
type Placeholder struct {
	id              string          // 占位符唯一标识符（XML 中的 idx 或内部生成）
	placeholderType PlaceholderType // 占位符类型
	x               int64           // X 坐标（EMU 单位）
	y               int64           // Y 坐标（EMU 单位）
	cx              int64           // 宽度（EMU 单位）
	cy              int64           // 高度（EMU 单位）
	rotation        int32           // 旋转角度（1/60000 度）
	defaultStyle    *TextStyle      // 默认文本样式（可为 nil）
}

// ============================================================================
// Background 背景相关结构
// ============================================================================

// BackgroundType 背景类型枚举
// 对应 XML: <p:bg> 下的不同子元素
type BackgroundType int8

const (
	BackgroundTypeNone       BackgroundType = iota // 无背景
	BackgroundTypeSolidColor                      // 纯色背景
	BackgroundTypeGradient                        // 渐变背景
	BackgroundTypePattern                         // 图案填充
	BackgroundTypePicture                         // 图片背景
	BackgroundTypeThemeColor                      // 主题色背景（如 bg1, tx1）
)

// Background 背景定义（只读）
// 对应 XML: <p:bg> 或 <p:cSld><p:bg>
// 设计说明：使用独立字段存储不同类型的值，避免接口带来的动态分配
type Background struct {
	backgroundType BackgroundType // 背景类型

	// 纯色背景属性（当 backgroundType == BackgroundTypeSolidColor 时有效）
	solidColorRGB string // RGB 十六进制颜色值，如 "FFFFFF"

	// 渐变背景属性（当 backgroundType == BackgroundTypeGradient 时有效）
	gradientAngle   int32   // 渐变角度（度）
	gradientColors  []GradientStop // 渐变色标列表

	// 图片背景属性（当 backgroundType == BackgroundTypePicture 时有效）
	pictureRId     string // 图片关系 ID（指向媒体资源）
	pictureURI     string // 图片内部 URI 路径

	// 通用属性
	opacity float32 // 不透明度 (0.0 - 1.0)，默认 1.0
}

// GradientStop 渐变色标（只读）
// 对应 XML: <a:gs>
type GradientStop struct {
	position float32 // 位置 (0.0 - 1.0)，表示渐变中的百分比位置
	colorRGB string  // RGB 十六进制颜色值
}

// ============================================================================
// SlideLayout 版式相关结构（只读数据结构）
// ============================================================================
//
// 注意：SlideLayoutType 已在 slide_types.go 中定义，此处直接使用
// ============================================================================

// SlideLayoutData 版式只读数据（用于模板系统）
// 对应 XML: /ppt/slideLayouts/slideLayoutN.xml
// 与 SlideLayoutPart 不同，这是纯数据结构，无 XML 读写能力
type SlideLayoutData struct {
	id           string                  // 版式唯一标识符（内部生成）
	name         string                  // 版式名称（显示在 PowerPoint 版式选择器中）
	layoutType   SlideLayoutType         // 版式类型（复用 slide_types.go 中的定义）
	background   *Background             // 背景（可为 nil，表示使用母版背景）
	masterId     string                  // 所属母版的 ID
	placeholders map[string]*Placeholder // 占位符集合，key 为占位符 ID
}

// ============================================================================
// SlideMaster 母版相关结构
// ============================================================================

// SlideMasterData 母版只读数据（用于模板系统）
// 对应 XML: /ppt/slideMasters/slideMasterN.xml
// 母版是幻灯片模板的顶层容器，包含一个或多个版式
type SlideMasterData struct {
	id           string                  // 母版唯一标识符（内部生成）
	name         string                  // 母版名称
	background   *Background             // 母版级背景（可为 nil）
	placeholders map[string]*Placeholder // 母版级占位符（可为 nil），定义全局占位符样式
	layouts      []*SlideLayoutData      // 包含的版式列表
}

// ============================================================================
// Placeholder 访问器方法
// ============================================================================

// ID 返回占位符唯一标识符
func (p *Placeholder) ID() string { return p.id }

// Type 返回占位符类型
func (p *Placeholder) Type() PlaceholderType { return p.placeholderType }

// X 返回 X 坐标（EMU 单位）
func (p *Placeholder) X() int64 { return p.x }

// Y 返回 Y 坐标（EMU 单位）
func (p *Placeholder) Y() int64 { return p.y }

// Cx 返回宽度（EMU 单位）
func (p *Placeholder) Cx() int64 { return p.cx }

// Cy 返回高度（EMU 单位）
func (p *Placeholder) Cy() int64 { return p.cy }

// Rotation 返回旋转角度（1/60000 度）
func (p *Placeholder) Rotation() int32 { return p.rotation }

// DefaultStyle 返回默认文本样式（可能为 nil）
func (p *Placeholder) DefaultStyle() *TextStyle { return p.defaultStyle }

// Bounds 返回边界矩形 (x, y, cx, cy)
func (p *Placeholder) Bounds() (x, y, cx, cy int64) {
	return p.x, p.y, p.cx, p.cy
}

// ============================================================================
// TextStyle 访问器方法
// ============================================================================

// FontName 返回字体名称
func (s *TextStyle) FontName() string { return s.fontName }

// FontSize 返回字体大小（百分之一磅，100 = 1pt）
func (s *TextStyle) FontSize() int32 { return s.fontSize }

// Bold 返回是否粗体
func (s *TextStyle) Bold() bool { return s.bold }

// Italic 返回是否斜体
func (s *TextStyle) Italic() bool { return s.italic }

// Underline 返回是否下划线
func (s *TextStyle) Underline() bool { return s.underline }

// ColorRGB 返回文本颜色（RGB 十六进制）
func (s *TextStyle) ColorRGB() string { return s.colorRGB }

// ============================================================================
// Background 访问器方法
// ============================================================================

// Type 返回背景类型
func (b *Background) Type() BackgroundType { return b.backgroundType }

// SolidColorRGB 返回纯色背景的 RGB 值（仅当 Type == BackgroundTypeSolidColor 时有效）
func (b *Background) SolidColorRGB() string { return b.solidColorRGB }

// GradientAngle 返回渐变角度（仅当 Type == BackgroundTypeGradient 时有效）
func (b *Background) GradientAngle() int32 { return b.gradientAngle }

// GradientColors 返回渐变色标列表（仅当 Type == BackgroundTypeGradient 时有效）
func (b *Background) GradientColors() []GradientStop { return b.gradientColors }

// PictureRId 返回图片关系 ID（仅当 Type == BackgroundTypePicture 时有效）
func (b *Background) PictureRId() string { return b.pictureRId }

// PictureURI 返回图片内部 URI（仅当 Type == BackgroundTypePicture 时有效）
func (b *Background) PictureURI() string { return b.pictureURI }

// Opacity 返回不透明度 (0.0 - 1.0)
func (b *Background) Opacity() float32 { return b.opacity }

// ============================================================================
// GradientStop 访问器方法
// ============================================================================

// Position 返回色标位置 (0.0 - 1.0)
func (g *GradientStop) Position() float32 { return g.position }

// ColorRGB 返回色标颜色（RGB 十六进制）
func (g *GradientStop) ColorRGB() string { return g.colorRGB }

// ============================================================================
// SlideLayoutData 访问器方法
// ============================================================================

// ID 返回版式唯一标识符
func (l *SlideLayoutData) ID() string { return l.id }

// Name 返回版式名称
func (l *SlideLayoutData) Name() string { return l.name }

// LayoutType 返回版式类型
func (l *SlideLayoutData) LayoutType() SlideLayoutType { return l.layoutType }

// Background 返回背景（可能为 nil）
func (l *SlideLayoutData) Background() *Background { return l.background }

// MasterID 返回所属母版的 ID
func (l *SlideLayoutData) MasterID() string { return l.masterId }

// Placeholders 返回占位符集合
func (l *SlideLayoutData) Placeholders() map[string]*Placeholder { return l.placeholders }

// PlaceholderByID 根据 ID 获取占位符（可能为 nil）
func (l *SlideLayoutData) PlaceholderByID(id string) *Placeholder {
	return l.placeholders[id]
}

// PlaceholderCount 返回占位符数量
func (l *SlideLayoutData) PlaceholderCount() int { return len(l.placeholders) }

// PlaceholderByType 根据类型获取第一个匹配的占位符
func (l *SlideLayoutData) PlaceholderByType(phType PlaceholderType) *Placeholder {
	for _, ph := range l.placeholders {
		if ph.placeholderType == phType {
			return ph
		}
	}
	return nil
}

// TitlePlaceholder 获取标题占位符（便捷方法）
func (l *SlideLayoutData) TitlePlaceholder() *Placeholder {
	return l.PlaceholderByType(PlaceholderTypeTitle)
}

// BodyPlaceholder 获取正文占位符（便捷方法）
func (l *SlideLayoutData) BodyPlaceholder() *Placeholder {
	return l.PlaceholderByType(PlaceholderTypeBody)
}

// ============================================================================
// SlideMasterData 访问器方法
// ============================================================================

// ID 返回母版唯一标识符
func (m *SlideMasterData) ID() string { return m.id }

// Name 返回母版名称
func (m *SlideMasterData) Name() string { return m.name }

// Background 返回背景（可能为 nil）
func (m *SlideMasterData) Background() *Background { return m.background }

// Placeholders 返回母版级占位符集合
func (m *SlideMasterData) Placeholders() map[string]*Placeholder { return m.placeholders }

// PlaceholderByID 根据 ID 获取占位符（可能为 nil）
func (m *SlideMasterData) PlaceholderByID(id string) *Placeholder {
	return m.placeholders[id]
}

// PlaceholderCount 返回占位符数量
func (m *SlideMasterData) PlaceholderCount() int { return len(m.placeholders) }

// Layouts 返回版式列表
func (m *SlideMasterData) Layouts() []*SlideLayoutData { return m.layouts }

// LayoutCount 返回版式数量
func (m *SlideMasterData) LayoutCount() int { return len(m.layouts) }

// LayoutByID 根据 ID 获取版式（可能为 nil）
func (m *SlideMasterData) LayoutByID(id string) *SlideLayoutData {
	for _, layout := range m.layouts {
		if layout.id == id {
			return layout
		}
	}
	return nil
}
