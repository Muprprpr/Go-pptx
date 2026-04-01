package parts

// ============================================================================
// Chart XML 模板 - 对应 /ppt/charts/chartN.xml
// ============================================================================
//
// 设计策略：字符串模板 + 占位符
// - 图表 XML 结构极其复杂，不适合用 Go Struct 完整映射
// - 采用预定义模板 + 占位符策略，高层组件通过替换占位符注入数据
//
// 常用占位符：
//   {{CHART_TITLE}}     - 图表标题
//   {{CATEGORIES}}      - 分类数据（XML 片段）
//   {{SERIES_NAME}}     - 系列名称
//   {{SERIES_VALUES}}   - 系列数值（XML 片段）
//   {{SER_COUNT}}       - 系列数量
//   {{CAT_COUNT}}       - 分类数量
//
// ============================================================================

// ============================================================================
// 柱状图模板
// ============================================================================

// ChartTemplateBar 柱状图模板
const ChartTemplateBar = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{{CHART_TITLE}}</a:t></a:r></a:p></c:rich></c:tx>
    </c:title>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{{SERIES_NAME}}</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat>
            <c:strRef>
              <c:f>Sheet1!$A$2:$A${{CAT_COUNT_PLUS_1}}</c:f>
              <c:strCache>
                <c:ptCount val="{{CAT_COUNT}}"/>
{{CATEGORIES}}
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$2:$B${{CAT_COUNT_PLUS_1}}</c:f>
              <c:numCache>
                <c:formatCode>General</c:formatCode>
                <c:ptCount val="{{CAT_COUNT}}"/>
{{SERIES_VALUES}}
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:scaling/><c:delete val="0"/>
        <c:axPos val="b"/>
        <c:crossAx val="2"/>
        <c:crosses val="autoZero"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling/><c:delete val="0"/>
        <c:axPos val="l"/>
        <c:crossAx val="1"/>
        <c:crosses val="autoZero"/>
      </c:valAx>
    </c:plotArea>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>`

// ============================================================================
// 饼图模板
// ============================================================================

// ChartTemplatePie 饼图模板
const ChartTemplatePie = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{{CHART_TITLE}}</a:t></a:r></a:p></c:rich></c:tx>
    </c:title>
    <c:plotArea>
      <c:layout/>
      <c:pieChart>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{{SERIES_NAME}}</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat>
            <c:strRef>
              <c:f>Sheet1!$A$2:$A${{CAT_COUNT_PLUS_1}}</c:f>
              <c:strCache>
                <c:ptCount val="{{CAT_COUNT}}"/>
{{CATEGORIES}}
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$2:$B${{CAT_COUNT_PLUS_1}}</c:f>
              <c:numCache>
                <c:formatCode>General</c:formatCode>
                <c:ptCount val="{{CAT_COUNT}}"/>
{{SERIES_VALUES}}
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:firstSliceAng val="0"/>
      </c:pieChart>
    </c:plotArea>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>`

// ============================================================================
// 折线图模板
// ============================================================================

// ChartTemplateLine 折线图模板
const ChartTemplateLine = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{{CHART_TITLE}}</a:t></a:r></a:p></c:rich></c:tx>
    </c:title>
    <c:plotArea>
      <c:layout/>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{{SERIES_NAME}}</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:marker><c:symbol val="none"/></c:marker>
          <c:cat>
            <c:strRef>
              <c:f>Sheet1!$A$2:$A${{CAT_COUNT_PLUS_1}}</c:f>
              <c:strCache>
                <c:ptCount val="{{CAT_COUNT}}"/>
{{CATEGORIES}}
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$2:$B${{CAT_COUNT_PLUS_1}}</c:f>
              <c:numCache>
                <c:formatCode>General</c:formatCode>
                <c:ptCount val="{{CAT_COUNT}}"/>
{{SERIES_VALUES}}
              </c:numCache>
            </c:numRef>
          </c:val>
          <c:smooth val="0"/>
        </c:ser>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:lineChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:scaling/><c:delete val="0"/>
        <c:axPos val="b"/>
        <c:crossAx val="2"/>
        <c:crosses val="autoZero"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling/><c:delete val="0"/>
        <c:axPos val="l"/>
        <c:crossAx val="1"/>
        <c:crosses val="autoZero"/>
      </c:valAx>
    </c:plotArea>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>`

// ============================================================================
// 面积图模板
// ============================================================================

// ChartTemplateArea 面积图模板
const ChartTemplateArea = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{{CHART_TITLE}}</a:t></a:r></a:p></c:rich></c:tx>
    </c:title>
    <c:plotArea>
      <c:layout/>
      <c:areaChart>
        <c:grouping val="standard"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{{SERIES_NAME}}</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat>
            <c:strRef>
              <c:f>Sheet1!$A$2:$A${{CAT_COUNT_PLUS_1}}</c:f>
              <c:strCache>
                <c:ptCount val="{{CAT_COUNT}}"/>
{{CATEGORIES}}
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$2:$B${{CAT_COUNT_PLUS_1}}</c:f>
              <c:numCache>
                <c:formatCode>General</c:formatCode>
                <c:ptCount val="{{CAT_COUNT}}"/>
{{SERIES_VALUES}}
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:areaChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:scaling/><c:delete val="0"/>
        <c:axPos val="b"/>
        <c:crossAx val="2"/>
        <c:crosses val="autoZero"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling/><c:delete val="0"/>
        <c:axPos val="l"/>
        <c:crossAx val="1"/>
        <c:crosses val="autoZero"/>
      </c:valAx>
    </c:plotArea>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>`

// ============================================================================
// 散点图模板
// ============================================================================

// ChartTemplateScatter 散点图模板
const ChartTemplateScatter = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{{CHART_TITLE}}</a:t></a:r></a:p></c:rich></c:tx>
    </c:title>
    <c:plotArea>
      <c:layout/>
      <c:scatterChart>
        <c:scatterStyle val="marker"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{{SERIES_NAME}}</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:marker><c:symbol val="circle"/><c:size val="7"/></c:marker>
          <c:xVal>
            <c:numRef>
              <c:f>Sheet1!$A$2:$A${{CAT_COUNT_PLUS_1}}</c:f>
              <c:numCache>
                <c:formatCode>General</c:formatCode>
                <c:ptCount val="{{CAT_COUNT}}"/>
{{X_VALUES}}
              </c:numCache>
            </c:numRef>
          </c:xVal>
          <c:yVal>
            <c:numRef>
              <c:f>Sheet1!$B$2:$B${{CAT_COUNT_PLUS_1}}</c:f>
              <c:numCache>
                <c:formatCode>General</c:formatCode>
                <c:ptCount val="{{CAT_COUNT}}"/>
{{Y_VALUES}}
              </c:numCache>
            </c:numRef>
          </c:yVal>
        </c:ser>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:scatterChart>
      <c:valAx>
        <c:axId val="1"/>
        <c:scaling/><c:delete val="0"/>
        <c:axPos val="b"/>
        <c:crossAx val="2"/>
        <c:crosses val="autoZero"/>
      </c:valAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling/><c:delete val="0"/>
        <c:axPos val="l"/>
        <c:crossAx val="1"/>
        <c:crosses val="autoZero"/>
      </c:valAx>
    </c:plotArea>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>`

// ============================================================================
// 环形图模板
// ============================================================================

// ChartTemplateDoughnut 环形图模板
const ChartTemplateDoughnut = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{{CHART_TITLE}}</a:t></a:r></a:p></c:rich></c:tx>
    </c:title>
    <c:plotArea>
      <c:layout/>
      <c:doughnutChart>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{{SERIES_NAME}}</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat>
            <c:strRef>
              <c:f>Sheet1!$A$2:$A${{CAT_COUNT_PLUS_1}}</c:f>
              <c:strCache>
                <c:ptCount val="{{CAT_COUNT}}"/>
{{CATEGORIES}}
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$2:$B${{CAT_COUNT_PLUS_1}}</c:f>
              <c:numCache>
                <c:formatCode>General</c:formatCode>
                <c:ptCount val="{{CAT_COUNT}}"/>
{{SERIES_VALUES}}
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:firstSliceAng val="0"/>
        <c:holeSize val="50"/>
      </c:doughnutChart>
    </c:plotArea>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>`

// ============================================================================
// 占位符常量
// ============================================================================

const (
	PlaceholderChartTitle      = "{{CHART_TITLE}}"
	PlaceholderCategories     = "{{CATEGORIES}}"
	PlaceholderSeriesName     = "{{SERIES_NAME}}"
	PlaceholderSeriesValues   = "{{SERIES_VALUES}}"
	PlaceholderCatCount       = "{{CAT_COUNT}}"
	PlaceholderCatCountPlus1  = "{{CAT_COUNT_PLUS_1}}"
	PlaceholderXValues        = "{{X_VALUES}}"
	PlaceholderYValues        = "{{Y_VALUES}}"
)
