package parts

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
)

// XMLDeclaration OPC 包中所有 XML 文件的标准声明头
const XMLDeclaration = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`

// StripNamespacePrefixes 处理 XML 数据，去除命名空间前缀使其兼容 Go 的 xml.Unmarshal
// Go 的 xml.Unmarshal 无法处理带前缀的 XML 命名空间（如 <p:presentation>）
// 此函数将 <p:xxx> 转换为 <xxx>，同时将带前缀的属性转换为无冒号形式（如 r:id -> rid）
func StripNamespacePrefixes(data []byte) ([]byte, error) {
	var buf bytes.Buffer
	decoder := xml.NewDecoder(bytes.NewReader(data))

	// 命名空间 URI -> 前缀 的映射（全局，跨所有元素）
	nsToPrefix := make(map[string]string)

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("XML token error: %w", err)
		}

		switch v := token.(type) {
		case xml.StartElement:
			// 首先收集 xmlns 声明，建立 URI -> 前缀映射
			for _, attr := range v.Attr {
				if attr.Name.Space == "xmlns" {
					// xmlns:p="..." -> prefix="p", URI=attr.Value
					nsToPrefix[attr.Value] = attr.Name.Local
				} else if attr.Name.Local == "xmlns" {
					// xmlns="..." -> 默认命名空间，前缀为空
					nsToPrefix[attr.Value] = ""
				}
			}

			// 去除元素名前缀（如 p:presentation -> presentation）
			buf.WriteString("<")
			buf.WriteString(v.Name.Local)

			// 处理属性 - 将 r:id 转换为 rid（去掉冒号），去除 xmlns 声明
			for _, attr := range v.Attr {
				// 跳过 xmlns 声明
				if attr.Name.Space == "xmlns" || attr.Name.Local == "xmlns" {
					continue
				}
				buf.WriteString(" ")

				// 如果属性有命名空间，查找对应的前缀并拼接（去掉冒号）
				if attr.Name.Space != "" {
					if prefix, ok := nsToPrefix[attr.Name.Space]; ok && prefix != "" {
						buf.WriteString(prefix)
						// 注意：不写冒号，直接拼接，如 r:id -> rid
					}
				}
				buf.WriteString(attr.Name.Local)
				buf.WriteString("=\"")
				buf.WriteString(attr.Value)
				buf.WriteString("\"")
			}
			buf.WriteString(">")
		case xml.EndElement:
			buf.WriteString("</")
			buf.WriteString(v.Name.Local)
			buf.WriteString(">")
		case xml.CharData:
			buf.Write(v)
		case xml.Comment:
			buf.WriteString("<!--")
			buf.Write(v)
			buf.WriteString("-->")
		case xml.ProcInst:
			// 跳过 XML 声明
			if v.Target == "xml" {
				continue
			}
		}
	}

	return buf.Bytes(), nil
}
