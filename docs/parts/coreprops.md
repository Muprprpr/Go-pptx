# Core Properties 接口文档

> 对应 `/docProps/core.xml`，基于 Dublin Core 元数据标准

## 命名空间

| 常量 | 值 |
|------|-----|
| `NamespaceCoreProperties` | `http://schemas.openxmlformats.org/package/2006/metadata/core-properties` |
| `NamespaceDublinCore` | `http://purl.org/dc/elements/1.1/` |
| `NamespaceDublinCoreTerms` | `http://purl.org/dc/terms/` |
| `NamespaceXMLSchema` | `http://www.w3.org/2001/XMLSchema-instance` |

## 结构体

### XMLCoreProperties

核心属性 XML 结构体，对应 `core.xml` 文件。

| 字段 | XML 路径 | 类型 | 说明 |
|------|----------|------|------|
| `Title` | `dc:title` | `string` | 文档标题 |
| `Creator` | `dc:creator` | `string` | 创建者 |
| `Subject` | `dc:subject` | `string` | 主题 |
| `Description` | `dc:description` | `string` | 描述 |
| `Created` | `dcterms:created` | `*XMLW3CDTFDate` | 创建时间 |
| `Modified` | `dcterms:modified` | `*XMLW3CDTFDate` | 修改时间 |
| `Keywords` | `cp:keywords` | `string` | 关键词 |
| `LastModifiedBy` | `cp:lastModifiedBy` | `string` | 最后修改者 |
| `Revision` | `cp:revision` | `string` | 修订号 |
| `Category` | `cp:category` | `string` | 类别 |
| `ContentType` | `cp:contentType` | `string` | 内容类型 |
| `Version` | `cp:version` | `string` | 版本 |
| `Identifier` | `cp:identifier` | `string` | 标识符 |
| `Language` | `dc:language` | `string` | 语言 |

### XMLW3CDTFDate

W3CDTF 格式日期元素，对应 `<dcterms:created xsi:type="dcterms:W3CDTF">`。

| 字段 | XML 属性 | 类型 | 说明 |
|------|----------|------|------|
| `Type` | `xsi:type` | `string` | 类型标识，固定为 `dcterms:W3CDTF` |
| `Value` | chardata | `string` | 日期值，格式：`YYYY-MM-DDThh:mm:ssZ` |

## 构造函数

### NewXMLCoreProperties

```go
func NewXMLCoreProperties() *XMLCoreProperties
```

创建带默认命名空间的核心属性结构体。

## 辅助方法

### SetCreated

```go
func (cp *XMLCoreProperties) SetCreated(value string)
```

设置创建时间，格式为 W3CDTF。

### SetModified

```go
func (cp *XMLCoreProperties) SetModified(value string)
```

设置修改时间，格式为 W3CDTF。

### GetCreated

```go
func (cp *XMLCoreProperties) GetCreated() string
```

获取创建时间值。

### GetModified

```go
func (cp *XMLCoreProperties) GetModified() string
```

获取修改时间值。

### ToXML

```go
func (cp *XMLCoreProperties) ToXML() ([]byte, error)
```

将核心属性序列化为 XML 字节。

### ParseCoreProperties

```go
func ParseCoreProperties(data []byte) (*XMLCoreProperties, error)
```

从 XML 字节解析核心属性。

### ParseCoreProps

```go
func ParseCoreProps(data []byte) (*XMLCoreProperties, error)
```

`ParseCoreProperties` 的简写别名。
