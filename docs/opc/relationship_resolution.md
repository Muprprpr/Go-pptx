# 关系解析

OPC 包中的部件通过关系（Relationship）相互引用。本文档描述关系解析机制及其使用方法。

## 概述

在 OPC 包中，关系定义了部件之间的引用关系：

```
presentation.xml ──关系──> slides/slide1.xml
                   │
                   └──关系──> slideMasters/slideMaster1.xml
```

关系存储在 `_rels` 目录下的 `.rels` 文件中，目标路径可以是：
- **绝对路径**：以 `/` 开头，如 `/ppt/slides/slide1.xml`
- **相对路径**：相对于源部件，如 `slides/slide1.xml`

## 相对路径解析

### 问题

当关系目标使用相对路径时（如 `slides/slide1.xml`），需要基于源部件的位置来解析为绝对路径。

```xml
<!-- presentation.xml.rels -->
<Relationship Id="rId2" Type="..." Target="slides/slide1.xml"/>
```

对于 `/ppt/presentation.xml`，目标应解析为 `/ppt/slides/slide1.xml`。

### 解决方案

`NewRelationship` 函数会自动处理相对路径解析：

```go
func NewRelationship(rID, relType, targetURI string, isExternal bool, source *PackURI) *Relationship {
    // ...
    if source != nil && !strings.HasPrefix(targetURI, "/") {
        // 相对路径：使用 source 的目录来解析
        rel.target = source.Join(targetURI)
    } else {
        // 绝对路径：直接创建
        rel.target = NewPackURI(targetURI)
    }
    // ...
}
```

### PackURI.Join 方法

`Join` 方法基于 URI 的目录解析相对路径：

```go
source := opc.NewPackURI("/ppt/presentation.xml")

source.Join("slides/slide1.xml")        // → /ppt/slides/slide1.xml
source.Join("../theme/theme1.xml")      // → /theme/theme1.xml
source.Join("./slides/slide1.xml")      // → /ppt/slides/slide1.xml
source.Join("/ppt/slides/slide1.xml")   // → /ppt/slides/slide1.xml (绝对路径保持不变)
```

## 关系 API

### 创建关系

```go
// 通过 Part 的方法创建
rel, err := part.AddRelationship(relType, targetURI, isExternal)

// 直接创建
rel := opc.NewRelationship(rID, relType, targetURI, isExternal, sourceURI)
```

### 解析关系

```go
// 通过 Package 解析到目标部件
targetPart := pkg.ResolveRelationship(sourcePart, relType)

// 通过 Part 获取关系目标 URI
targetURI := sourcePart.GetRelatedPart(rID)
```

### 关系属性

```go
rel.RID()         // 关系 ID，如 "rId1"
rel.Type()        // 关系类型 URI
rel.TargetURI()   // 目标的绝对路径（解析后）
rel.TargetRef()   // 目标的相对引用（用于序列化）
rel.IsExternal()  // 是否为外部关系
rel.SourceURI()   // 源部件 URI
```

## 常见关系类型

| 源部件 | 目标部件 | 关系类型 |
|--------|----------|----------|
| Package | Presentation | `officeDocument` |
| Presentation | Slide | `slide` |
| Presentation | SlideMaster | `slideMaster` |
| Slide | SlideLayout | `slideLayout` |
| SlideLayout | SlideMaster | `slideMaster` |
| SlideMaster | Theme | `theme` |
| Slide | Image | `image` |

## 示例：完整的关系链

```go
pkg := opc.NewPackage()

// 1. 创建部件
presPart, _ := pkg.CreatePart(opc.NewPackURI("/ppt/presentation.xml"), contentType, data)
slidePart, _ := pkg.CreatePart(opc.NewPackURI("/ppt/slides/slide1.xml"), contentType, slideData)

// 2. 添加关系（使用相对路径）
presPart.AddRelationship(
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
    "slides/slide1.xml",  // 相对路径
    false,
)

// 3. 解析关系
resolved := pkg.ResolveRelationship(presPart, "http://.../slide")
// resolved == slidePart
```

## 注意事项

1. **相对路径格式**：使用正斜杠 `/`，不要使用反斜杠
2. **路径解析**：相对路径基于源部件的目录解析
3. **序列化**：`TargetRef()` 返回适合写入 XML 的相对路径
4. **外部关系**：外部目标的 `isExternal` 应设为 `true`
