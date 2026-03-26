# Relationship 模块接口文档

> 对应 `*.rels` 文件，OpenXML 关系定义

---

## 概述

本模块提供两种关系管理方式：

1. **简单 XML 结构** (`XMLRelationships`/`XMLRelationship`)：轻量级的 XML 序列化/反序列化
2. **OPC 关系系统** (`opc.Relationships`)：完整的包关系管理，支持路径解析、线程安全 ID 分配

> **相关文档**：如需关系路径解析和完整的 OPC 关系管理，请参阅 [OPC 关系解析](../opc/relationship_resolution.md)。

---

## 概述

关系文件位置示例：
- 包级别: `/_rels/.rels`
- 幻灯片: `/ppt/slides/_rels/slide1.xml.rels`
- 母版: `/ppt/slideMasters/_rels/slideMaster1.xml.rels`

命名空间: `http://schemas.openxmlformats.org/package/2006/relationships`

---

## 常量定义

### 命名空间

```go
NamespaceRelationships = "http://schemas.openxmlformats.org/package/2006/relationships"
```

### 目标模式

```go
TargetModeInternal = "Internal"  // 内部目标
TargetModeExternal = "External"   // 外部目标
```

### 常用关系类型

```go
RelTypeImage       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
RelTypeHyperlink   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
RelTypeSlide       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
RelTypeSlideLayout = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
RelTypeSlideMaster = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
RelTypeTheme       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
RelTypeNotesSlide  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
RelTypeComments    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
RelTypeChart       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
RelTypeTable       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
RelTypeMedia       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"
RelTypeAudio       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
```

---

## 数据结构

### XMLRelationships

关系集合根节点，对应 XML: `<Relationships xmlns="...">...</Relationships>`

| 字段 | 类型 | 说明 |
|------|------|------|
| `Xmlns` | `string` | 命名空间 |
| `Relationships` | `[]XMLRelationship` | 关系列表 |

### XMLRelationship

单个关系，对应 XML: `<Relationship Id="rId1" Type="..." Target="..."/>`

| 字段 | XML 属性 | 类型 | 说明 |
|------|----------|------|------|
| `ID` | `Id` | `string` | 关系 ID（如 `rId1`, `rId2`） |
| `Type` | `Type` | `string` | 关系类型 URI |
| `Target` | `Target` | `string` | 目标路径（相对或绝对） |
| `TargetMode` | `TargetMode` | `string` | `Internal`（默认）或 `External` |

#### 方法

```go
func (r *XMLRelationship) IsExternal() bool
```

检查是否为外部关系。

---

## 构造函数

### NewXMLRelationships

```go
func NewXMLRelationships() *XMLRelationships
```

创建带默认命名空间的关系集合。

### NewXMLRelationship

```go
func NewXMLRelationship(id, relType, target string) XMLRelationship
```

创建单个关系。

### NewXMLRelationshipExternal

```go
func NewXMLRelationshipExternal(id, relType, target string) XMLRelationship
```

创建外部关系（`TargetMode=External`）。

---

## XMLRelationships 方法

### Add

```go
func (rs *XMLRelationships) Add(rel XMLRelationship)
```

添加关系到集合。

### AddNew

```go
func (rs *XMLRelationships) AddNew(id, relType, target string)
```

创建并添加新关系。

### GetByID

```go
func (rs *XMLRelationships) GetByID(id string) *XMLRelationship
```

根据 ID 获取关系。

### GetByType

```go
func (rs *XMLRelationships) GetByType(relType string) []XMLRelationship
```

根据类型获取所有关系。

### GetByTarget

```go
func (rs *XMLRelationships) GetByTarget(target string) *XMLRelationship
```

根据目标路径获取关系。

### GetByType

```go
func (rs *XMLRelationships) GetByType(relType string) []XMLRelationship
```

根据类型获取所有关系。

### Count

```go
func (rs *XMLRelationships) Count() int
```

返回关系数量。

---

## XML 序列化/反序列化

### ToXML

```go
func (rs *XMLRelationships) ToXML() ([]byte, error)
```

将关系集合序列化为 XML 字节。

### ParseRelationships

```go
func ParseRelationships(data []byte) (*XMLRelationships, error)
```

从 XML 字节解析关系集合。
