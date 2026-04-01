package parts

// ============================================================================
// ChartPart - 图表部件
// ============================================================================
//
// 对应 /ppt/charts/chartN.xml
//
// 设计策略：模板 + 占位符
// - 图表 XML 结构极其复杂（几百种元素组合），不适合用 Go Struct 完整映射
// - 采用字符串模板 + 占位符策略，高层组件通过替换占位符注入数据
// - 提供常见图表类型的预定义模板
//
// ============================================================================

import (
	"fmt"
	"strings"
	"sync"

	"github.com/Muprprpr/Go-pptx/opc"
)

// ChartPart 图表部件
type ChartPart struct {
	uri         *opc.PackURI
	template    string        // XML 模板（含占位符）
	externalRID string        // 外部数据关系 ID（如嵌入的 Excel）
	mu          sync.RWMutex
}

// NewChartPart 创建新的图表部件
func NewChartPart(id int) *ChartPart {
	return &ChartPart{
		uri:      opc.NewPackURI(fmt.Sprintf("/ppt/charts/chart%d.xml", id)),
		template: ChartTemplateBar, // 默认柱状图
	}
}

// NewChartPartWithURI 使用指定 URI 创建图表部件
func NewChartPartWithURI(uri *opc.PackURI) *ChartPart {
	return &ChartPart{
		uri:      uri,
		template: ChartTemplateBar,
	}
}

// NewChartPartWithType 创建指定类型的图表部件
func NewChartPartWithType(id int, chartType ChartType) *ChartPart {
	return &ChartPart{
		uri:      opc.NewPackURI(fmt.Sprintf("/ppt/charts/chart%d.xml", id)),
		template: GetChartTemplate(chartType),
	}
}

// PartURI 返回部件 URI
func (c *ChartPart) PartURI() *opc.PackURI {
	return c.uri
}

// ============================================================================
// 模板操作方法
// ============================================================================

// SetTemplate 设置图表模板
func (c *ChartPart) SetTemplate(tmpl string) {
	c.mu.Lock()
	defer c.mu.Unlock()
	c.template = tmpl
}

// SetRawXML 设置原始 XML（等同于 SetTemplate，语义更清晰）
func (c *ChartPart) SetRawXML(xml []byte) {
	c.mu.Lock()
	defer c.mu.Unlock()
	c.template = string(xml)
}

// Template 返回当前模板
func (c *ChartPart) Template() string {
	c.mu.RLock()
	defer c.mu.RUnlock()
	return c.template
}

// ReplacePlaceholder 替换单个占位符
// placeholder: 占位符名称（不含 {{}}）
// value: 替换值
func (c *ChartPart) ReplacePlaceholder(placeholder, value string) {
	c.mu.Lock()
	defer c.mu.Unlock()
	c.template = strings.ReplaceAll(c.template, "{{"+placeholder+"}}", value)
}

// ReplacePlaceholders 批量替换占位符
func (c *ChartPart) ReplacePlaceholders(replacements map[string]string) {
	c.mu.Lock()
	defer c.mu.Unlock()
	for k, v := range replacements {
		c.template = strings.ReplaceAll(c.template, "{{"+k+"}}", v)
	}
}

// ============================================================================
// 外部数据引用
// ============================================================================

// SetExternalDataRID 设置外部数据关系 ID
func (c *ChartPart) SetExternalDataRID(rid string) {
	c.mu.Lock()
	defer c.mu.Unlock()
	c.externalRID = rid
}

// GetExternalDataRID 获取外部数据关系 ID
func (c *ChartPart) GetExternalDataRID() string {
	c.mu.RLock()
	defer c.mu.RUnlock()
	return c.externalRID
}

// HasExternalData 检查是否有外部数据
func (c *ChartPart) HasExternalData() bool {
	c.mu.RLock()
	defer c.mu.RUnlock()
	return c.externalRID != ""
}

// ============================================================================
// XML 序列化
// ============================================================================

// ToXML 将图表序列化为 XML
func (c *ChartPart) ToXML() ([]byte, error) {
	c.mu.RLock()
	defer c.mu.RUnlock()

	// 如果有外部数据引用，注入到模板中
	xml := c.template
	if c.externalRID != "" {
		// 在 </c:chartSpace> 前插入 externalData
		externalDataTag := fmt.Sprintf(`<c:externalData xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="%s"/>`, c.externalRID)
		xml = strings.Replace(xml, "</c:chartSpace>", externalDataTag+"</c:chartSpace>", 1)
	}

	return []byte(xml), nil
}

// FromXML 从 XML 加载图表（直接作为模板存储）
func (c *ChartPart) FromXML(data []byte) error {
	c.mu.Lock()
	defer c.mu.Unlock()
	c.template = string(data)
	return nil
}

// ============================================================================
// 图表类型枚举
// ============================================================================

// ChartType 图表类型
type ChartType int

const (
	ChartTypeBar ChartType = iota    // 柱状图
	ChartTypePie                     // 饼图
	ChartTypeLine                    // 折线图
	ChartTypeArea                    // 面积图
	ChartTypeScatter                 // 散点图
	ChartTypeDoughnut                // 环形图
)

// GetChartTemplate 获取图表模板
func GetChartTemplate(chartType ChartType) string {
	switch chartType {
	case ChartTypeBar:
		return ChartTemplateBar
	case ChartTypePie:
		return ChartTemplatePie
	case ChartTypeLine:
		return ChartTemplateLine
	case ChartTypeArea:
		return ChartTemplateArea
	case ChartTypeScatter:
		return ChartTemplateScatter
	case ChartTypeDoughnut:
		return ChartTemplateDoughnut
	default:
		return ChartTemplateBar
	}
}

// ============================================================================
// 数据结构（用于模板填充的辅助结构）
// ============================================================================

// ChartSeriesData 图表系列数据
type ChartSeriesData struct {
	Name   string   // 系列名称
	Values []string // 数值列表
}

// ChartCategoryData 图表分类数据
type ChartCategoryData struct {
	Categories []string       // 分类标签
	Series     []ChartSeriesData // 系列数据
}
