package parts_test

import (
	"strings"
	"testing"

	"github.com/Muprprpr/Go-pptx/opc"
	"github.com/Muprprpr/Go-pptx/parts"
)

// 这是一个极其经典的【无 Excel 图表 XML】(路线 C 的核心图纸)
// 注意看：里面只有 strCache 和 numCache，绝对没有 <c:externalData> 标签！
const routeCChartXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:cat>
            <c:strRef>
              <c:strCache>
                <c:ptCount val="2"/>
                <c:pt idx="0"><c:v>Q1</c:v></c:pt>
                <c:pt idx="1"><c:v>Q2</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:numCache>
                <c:ptCount val="2"/>
                <c:pt idx="0"><c:v>150</c:v></c:pt>
                <c:pt idx="1"><c:v>200</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`

// ============================================================================
// 测试 1：ChartPart 原始 XML 承载测试
// 目标：证明 ChartPart 能原封不动地保存我们利用模板引擎生成的 XML
// ============================================================================
func TestChartPart_RawXMLCarrier(t *testing.T) {
	// 使用 SetRawXML 方法加载原始 XML
	chartPart := parts.NewChartPart(1)
	chartPart.SetRawXML([]byte(routeCChartXML))

	// 模拟序列化过程
	outputBytes, err := chartPart.ToXML()
	if err != nil {
		t.Fatalf("图表序列化失败: %v", err)
	}
	outputXML := string(outputBytes)

	// 核心验证 1：必须包含缓存数据
	if !strings.Contains(outputXML, "<c:strCache>") || !strings.Contains(outputXML, "Q1") {
		t.Error("丢失了字符串缓存数据 (strCache)")
	}
	if !strings.Contains(outputXML, "<c:numCache>") || !strings.Contains(outputXML, "200") {
		t.Error("丢失了数字缓存数据 (numCache)")
	}

	// 核心验证 2：【路线 C 护城河】绝对不能包含 externalData
	if strings.Contains(outputXML, "<c:externalData") {
		t.Fatal("严重错误：路线 C 的图表中出现了外部 Excel 引用标签！")
	}

	t.Log("✅ ChartPart 原始数据承载测试通过")
}

// ============================================================================
// 测试 2：OPC 关系纯净度测试 (验证没有挂载 Excel)
// 目标：证明在打包图表时，没有意外生成指向 embedding 文件夹的 Excel 关系
// ============================================================================
func TestChartPart_NoExcelRelationship(t *testing.T) {
	pkg := opc.NewPackage()

	// 1. 创建并写入图表部件
	chartURI := opc.NewPackURI("/ppt/charts/chart1.xml")
	chartPartOp, err := pkg.CreatePart(chartURI, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", []byte(routeCChartXML))
	if err != nil {
		t.Fatalf("创建 Chart Part 失败: %v", err)
	}

	// 2. 检查这个 chartPartOp 底下挂载的关系
	// 在路线 C 中，图表自身不应该拥有任何关系（因为它不依赖 Excel）
	rels := chartPartOp.Relationships()
	if rels != nil && rels.Count() > 0 {
		// 遍历检查是否有 Excel 类型的关联
		for _, rel := range rels.All() {
			if strings.Contains(rel.Type(), "officeDocument/2006/relationships/package") {
				t.Fatalf("架构违规：发现图表挂载了 Excel 关系 -> %s", rel.TargetURI())
			}
		}
	}

	t.Log("✅ Chart OPC 关系纯净度测试通过 (无 Excel 依赖)")
}

// ============================================================================
// 测试 3：占位符替换测试
// 目标：证明模板占位符策略能正确注入数据
// ============================================================================
func TestChartPart_PlaceholderReplacement(t *testing.T) {
	chartPart := parts.NewChartPartWithType(1, parts.ChartTypeBar)

	// 替换占位符
	chartPart.ReplacePlaceholder("CHART_TITLE", "销售报表")
	chartPart.ReplacePlaceholder("SERIES_NAME", "2024年")
	chartPart.ReplacePlaceholder("CAT_COUNT", "3")
	chartPart.ReplacePlaceholder("CAT_COUNT_PLUS_1", "4")

	// 序列化
	outputBytes, err := chartPart.ToXML()
	if err != nil {
		t.Fatalf("图表序列化失败: %v", err)
	}
	outputXML := string(outputBytes)

	// 验证替换成功
	if !strings.Contains(outputXML, "销售报表") {
		t.Error("CHART_TITLE 占位符替换失败")
	}
	if !strings.Contains(outputXML, "2024年") {
		t.Error("SERIES_NAME 占位符替换失败")
	}
	if !strings.Contains(outputXML, `<c:ptCount val="3"/>`) {
		t.Error("CAT_COUNT 占位符替换失败")
	}

	t.Log("✅ ChartPart 占位符替换测试通过")
}

// ============================================================================
// 测试 4：外部数据引用测试
// 目标：证明可以正确设置外部 Excel 数据引用
// ============================================================================
func TestChartPart_ExternalDataReference(t *testing.T) {
	chartPart := parts.NewChartPart(1)

	// 初始状态不应该有外部数据
	if chartPart.HasExternalData() {
		t.Error("新创建的图表不应该有外部数据")
	}

	// 设置外部数据引用
	chartPart.SetExternalDataRID("rId1")

	// 验证
	if !chartPart.HasExternalData() {
		t.Error("设置后应该有外部数据")
	}
	if chartPart.GetExternalDataRID() != "rId1" {
		t.Errorf("期望 rId1，得到 %s", chartPart.GetExternalDataRID())
	}

	// 序列化后应该包含 externalData 标签
	outputBytes, err := chartPart.ToXML()
	if err != nil {
		t.Fatalf("图表序列化失败: %v", err)
	}
	outputXML := string(outputBytes)

	if !strings.Contains(outputXML, `<c:externalData`) {
		t.Error("序列化后应包含 externalData 标签")
	}
	if !strings.Contains(outputXML, `r:id="rId1"`) {
		t.Error("externalData 标签应包含正确的 r:id")
	}

	t.Log("✅ ChartPart 外部数据引用测试通过")
}
