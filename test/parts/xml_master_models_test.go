package parts_test

import (
	"encoding/xml"
	"testing"

	"github.com/Muprprpr/Go-pptx/parts"
)

// ============================================================================
// 基础坐标与尺寸测试 - XMLOffset 和 XMLExtents
// ============================================================================

func TestParseXMLOffset(t *testing.T) {
	tests := []struct {
		name      string
		xmlData   string
		wantX     int64
		wantY     int64
		wantError bool
	}{
		{
			name:    "正常解析-正数坐标",
			xmlData: `<a:off x="1524000" y="1143000"/>`,
			wantX:   1524000,
			wantY:   1143000,
		},
		{
			name:    "正常解析-零值坐标",
			xmlData: `<a:off x="0" y="0"/>`,
			wantX:   0,
			wantY:   0,
		},
		{
			name:    "正常解析-大数值",
			xmlData: `<a:off x="9144000" y="6858000"/>`,
			wantX:   9144000,
			wantY:   6858000,
		},
		{
			name:    "边界-缺失x属性Go默认零值",
			xmlData: `<a:off y="1143000"/>`,
			wantX:   0,
			wantY:   1143000,
		},
		{
			name:    "边界-缺失y属性Go默认零值",
			xmlData: `<a:off x="1524000"/>`,
			wantX:   1524000,
			wantY:   0,
		},
		{
			name:    "边界-空元素Go默认零值",
			xmlData: `<a:off/>`,
			wantX:   0,
			wantY:   0,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			var off parts.XMLOffset
			err := xml.Unmarshal([]byte(tt.xmlData), &off)

			if tt.wantError {
				if err == nil {
					t.Error("期望返回错误，但解析成功")
				}
				return
			}

			if err != nil {
				t.Fatalf("解析失败: %v", err)
			}

			if off.X != tt.wantX {
				t.Errorf("X = %d, want %d", off.X, tt.wantX)
			}
			if off.Y != tt.wantY {
				t.Errorf("Y = %d, want %d", off.Y, tt.wantY)
			}
		})
	}
}

func TestParseXMLExtents(t *testing.T) {
	tests := []struct {
		name      string
		xmlData   string
		wantCx    int64
		wantCy    int64
		wantError bool
	}{
		{
			name:    "正常解析-标准尺寸",
			xmlData: `<a:ext cx="6858000" cy="5143500"/>`,
			wantCx:  6858000,
			wantCy:  5143500,
		},
		{
			name:    "正常解析-零值尺寸",
			xmlData: `<a:ext cx="0" cy="0"/>`,
			wantCx:  0,
			wantCy:  0,
		},
		{
			name:    "正常解析-宽屏尺寸",
			xmlData: `<a:ext cx="12192000" cy="6858000"/>`,
			wantCx:  12192000,
			wantCy:  6858000,
		},
		{
			name:    "边界-缺失cx属性Go默认零值",
			xmlData: `<a:ext cy="5143500"/>`,
			wantCx:  0,
			wantCy:  5143500,
		},
		{
			name:    "边界-缺失cy属性Go默认零值",
			xmlData: `<a:ext cx="6858000"/>`,
			wantCx:  6858000,
			wantCy:  0,
		},
		{
			name:    "边界-空元素Go默认零值",
			xmlData: `<a:ext/>`,
			wantCx:  0,
			wantCy:  0,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			var ext parts.XMLExtents
			err := xml.Unmarshal([]byte(tt.xmlData), &ext)

			if tt.wantError {
				if err == nil {
					t.Error("期望返回错误，但解析成功")
				}
				return
			}

			if err != nil {
				t.Fatalf("解析失败: %v", err)
			}

			if ext.Cx != tt.wantCx {
				t.Errorf("Cx = %d, want %d", ext.Cx, tt.wantCx)
			}
			if ext.Cy != tt.wantCy {
				t.Errorf("Cy = %d, want %d", ext.Cy, tt.wantCy)
			}
		})
	}
}

// ============================================================================
// 组合变换结构测试 - XMLTransform
// ============================================================================

// xmlSpPrWrapper 用于测试 XMLTransform 的包装结构（通过 p:spPr）
type xmlSpPrWrapper struct {
	Xfrm *parts.XMLTransform `xml:"xfrm"`
}

func TestParseXMLTransform(t *testing.T) {
	tests := []struct {
		name      string
		xmlData   string
		wantX     int64
		wantY     int64
		wantCx    int64
		wantCy    int64
		hasOff    bool
		hasExt    bool
	}{
		{
			name:    "正常解析-完整变换",
			xmlData: `<spPr><xfrm><off x="1524000" y="1143000"/><ext cx="6858000" cy="5143500"/></xfrm></spPr>`,
			wantX:   1524000,
			wantY:   1143000,
			wantCx:  6858000,
			wantCy:  5143500,
			hasOff:  true,
			hasExt:  true,
		},
		{
			name:    "正常解析-宽屏尺寸",
			xmlData: `<spPr><xfrm><off x="0" y="0"/><ext cx="12192000" cy="6858000"/></xfrm></spPr>`,
			wantX:   0,
			wantY:   0,
			wantCx:  12192000,
			wantCy:  6858000,
			hasOff:  true,
			hasExt:  true,
		},
		{
			name:    "边界-仅有坐标无尺寸",
			xmlData: `<spPr><xfrm><off x="9144000" y="6858000"/></xfrm></spPr>`,
			wantX:   9144000,
			wantY:   6858000,
			wantCx:  0,
			wantCy:  0,
			hasOff:  true,
			hasExt:  false,
		},
		{
			name:    "边界-仅有尺寸无坐标",
			xmlData: `<spPr><xfrm><ext cx="4572000" cy="3429000"/></xfrm></spPr>`,
			wantX:   0,
			wantY:   0,
			wantCx:  4572000,
			wantCy:  3429000,
			hasOff:  false,
			hasExt:  true,
		},
		{
			name:    "边界-空变换元素",
			xmlData: `<spPr><xfrm/></spPr>`,
			wantX:   0,
			wantY:   0,
			wantCx:  0,
			wantCy:  0,
			hasOff:  false,
			hasExt:  false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			var wrapper xmlSpPrWrapper
			err := xml.Unmarshal([]byte(tt.xmlData), &wrapper)
			if err != nil {
				t.Fatalf("解析失败: %v", err)
			}

			if wrapper.Xfrm == nil {
				t.Fatal("Xfrm 为 nil")
			}

			xfrm := wrapper.Xfrm

			// 检查 Off 是否存在
			if tt.hasOff {
				if xfrm.Off == nil {
					t.Fatal("期望 Off 不为 nil")
				}
				if xfrm.Off.X != tt.wantX {
					t.Errorf("Off.X = %d, want %d", xfrm.Off.X, tt.wantX)
				}
				if xfrm.Off.Y != tt.wantY {
					t.Errorf("Off.Y = %d, want %d", xfrm.Off.Y, tt.wantY)
				}
			} else {
				if xfrm.Off != nil {
					t.Errorf("期望 Off 为 nil，实际为 %+v", xfrm.Off)
				}
			}

			// 检查 Ext 是否存在
			if tt.hasExt {
				if xfrm.Ext == nil {
					t.Fatal("期望 Ext 不为 nil")
				}
				if xfrm.Ext.Cx != tt.wantCx {
					t.Errorf("Ext.Cx = %d, want %d", xfrm.Ext.Cx, tt.wantCx)
				}
				if xfrm.Ext.Cy != tt.wantCy {
					t.Errorf("Ext.Cy = %d, want %d", xfrm.Ext.Cy, tt.wantCy)
				}
			} else {
				if xfrm.Ext != nil {
					t.Errorf("期望 Ext 为 nil，实际为 %+v", xfrm.Ext)
				}
			}
		})
	}
}

// ============================================================================
// 占位符结构测试 - XMLPlaceholder
// ============================================================================

func TestParseXMLPlaceholder(t *testing.T) {
	tests := []struct {
		name     string
		xmlData  string
		wantType string
		wantIdx  string
	}{
		{
			name:     "标准标题占位符",
			xmlData:  `<ph type="title"/>`,
			wantType: "title",
			wantIdx:  "",
		},
		{
			name:     "带索引的正文占位符",
			xmlData:  `<ph type="body" idx="1"/>`,
			wantType: "body",
			wantIdx:  "1",
		},
		{
			name:     "日期占位符-带额外sz属性",
			xmlData:  `<ph type="dt" sz="half"/>`,
			wantType: "dt",
			wantIdx:  "",
		},
		{
			name:     "幻灯片编号占位符",
			xmlData:  `<ph type="sldNum"/>`,
			wantType: "sldNum",
			wantIdx:  "",
		},
		{
			name:     "页脚占位符",
			xmlData:  `<ph type="ftr"/>`,
			wantType: "ftr",
			wantIdx:  "",
		},
		{
			name:     "居中标题占位符",
			xmlData:  `<ph type="ctrTitle"/>`,
			wantType: "ctrTitle",
			wantIdx:  "",
		},
		{
			name:     "副标题占位符",
			xmlData:  `<ph type="subTitle"/>`,
			wantType: "subTitle",
			wantIdx:  "",
		},
		{
			name:     "图表占位符",
			xmlData:  `<ph type="chart" idx="2"/>`,
			wantType: "chart",
			wantIdx:  "2",
		},
		{
			name:     "表格占位符",
			xmlData:  `<ph type="tbl"/>`,
			wantType: "tbl",
			wantIdx:  "",
		},
		{
			name:     "图片占位符",
			xmlData:  `<ph type="pic"/>`,
			wantType: "pic",
			wantIdx:  "",
		},
		{
			name:     "仅有idx无type",
			xmlData:  `<ph idx="0"/>`,
			wantType: "",
			wantIdx:  "0",
		},
		{
			name:     "边界-空占位符元素",
			xmlData:  `<ph/>`,
			wantType: "",
			wantIdx:  "",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			var ph parts.XMLPlaceholder
			err := xml.Unmarshal([]byte(tt.xmlData), &ph)
			if err != nil {
				t.Fatalf("解析失败: %v", err)
			}

			if ph.Type != tt.wantType {
				t.Errorf("Type = %q, want %q", ph.Type, tt.wantType)
			}
			if ph.Idx != tt.wantIdx {
				t.Errorf("Idx = %q, want %q", ph.Idx, tt.wantIdx)
			}
		})
	}
}
