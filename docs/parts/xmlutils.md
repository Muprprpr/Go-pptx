# XML 工具函数

本模块提供 XML 处理工具函数，主要用于解决 Go 标准库 `encoding/xml` 在处理带命名空间前缀的 XML 时的兼容性问题。

## 概述

Office Open XML (OOXML) 格式的文件使用带命名空间前缀的 XML 元素和属性，例如：

```xml
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
  </p:sldIdLst>
</p:presentation>
```

Go 的 `xml.Unmarshal` 无法正确处理这种格式，因为它：
1. 无法匹配带前缀的元素名（如 `<p:presentation>`）
2. 无法匹配带前缀的属性名（如 `r:id`）

## 核心函数

### StripNamespacePrefixes

```go
func StripNamespacePrefixes(data []byte) ([]byte, error)
```

处理 XML 数据，去除命名空间前缀使其兼容 Go 的 `xml.Unmarshal`。

**转换规则：**

| 原始 XML | 转换后 |
|----------|--------|
| `<p:presentation>` | `<presentation>` |
| `<a:solidFill>` | `<solidFill>` |
| `r:id="rId1"` | `rid="rId1"` |
| `xmlns:p="..."` | (移除) |

**使用示例：**

```go
// 从 PPTX 文件读取的原始 XML 数据
rawXML := slidePart.Blob()

// 去除命名空间前缀
cleanXML, err := parts.StripNamespacePrefixes(rawXML)
if err != nil {
    return err
}

// 现在可以正确解析
var slide XSlide
if err := xml.Unmarshal(cleanXML, &slide); err != nil {
    return err
}
```

## XML 结构体标签约定

使用 `StripNamespacePrefixes` 后，结构体标签应使用以下约定：

### 元素名
- 不带前缀：`xml:"presentation"` 而非 `xml:"p:presentation"`
- 不带命名空间 URI：`xml:"spTree"` 而非 `xml:"http://... spTree"`

### 属性名
- 合并前缀到属性名：`xml:"rid,attr"` 而非 `xml:"r:id,attr"`
- 示例：`r:id` → `rid`，`r:embed` → `rembed`

```go
// 正确的标签格式
type XSldId struct {
    Id  uint32 `xml:"id,attr"`    // 匹配 id="256"
    RId string `xml:"rid,attr"`   // 匹配转换后的 rid="rId2"
}

// 错误的标签格式（会导致解析失败）
type XSldId struct {
    Id  uint32 `xml:"id,attr"`
    RId string `xml:"r:id,attr"`  // 无法匹配 rid
}
```

## 常量

### XMLDeclaration

```go
const XMLDeclaration = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
```

OPC 包中所有 XML 文件的标准 XML 声明头。

## 内部实现

### 处理流程

1. **收集命名空间映射**：遍历所有 `xmlns` 属性，建立 URI → 前缀的映射
2. **转换元素名**：去除元素名的前缀（如 `p:presentation` → `presentation`）
3. **转换属性名**：将属性的前缀合并到属性名（如 `r:id` → `rid`）
4. **移除 xmlns 声明**：删除所有命名空间声明属性

### 代码示例

```go
func stripNamespacePrefixes(data []byte) ([]byte, error) {
    var buf bytes.Buffer
    decoder := xml.NewDecoder(bytes.NewReader(data))
    nsToPrefix := make(map[string]string)

    for {
        token, err := decoder.Token()
        if err == io.EOF {
            break
        }
        // ... 处理各种 token 类型
    }
    return buf.Bytes(), nil
}
```

## 注意事项

1. **性能考虑**：此函数会复制整个 XML 数据，对于大文件可能有内存开销
2. **保留信息**：命名空间 URI 信息会丢失，但前缀信息保留在属性名中
3. **双向转换**：此函数仅用于读取（反序列化），写入时使用完整命名空间
