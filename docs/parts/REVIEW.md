# parts 包功能概览

## 1. Slide 模块 (slide.go, slide_types.go)

**核心职责**：幻灯片 XML 结构的生成和解析

| 类型/函数 | 说明 |
|----------|------|
| SlidePart | 幻灯片部件，对应 /ppt/slides/slideN.xml |
| SlideLayoutPart | 版式部件，对应 /ppt/slideLayouts/slideLayoutN.xml |
| ShapeIDAllocator | 形状 ID 分配器（单线程） |
| ShapeIDAllocatorSync | 形状 ID 分配器（线程安全） |
| XMLWriter / XMLWriterPool | 流式 XML 写入辅助 |
| XML 结构类型：XSlide, XSpTree, XSp, XPicture, XGraphicFrame, XTable, XTextBody 等 |

**注意**： `SlidePart` 的关係管理使用 `opc.Relationships`，去重逻辑通过 `AddImageRel`/`AddMediaRel`/`AddChartRel`/`AddTableRel` 方法封装。

## 2. Master 模块 (master.go, master_types.go, master_parser.go, master_cache.go)
**核心职责**： 母版/版式的只读数据结构、解析和缓存

| 类型/函数 | 说明 |
|----------|------|
| MasterManager | 母版管理器（门面模式） |
| MasterCache | 母版/版式缓存（并发安全读取） |
| SlideMasterData | 母版只读数据 |
| SlideLayoutData | 版式只读数据 |
| Placeholder | 占位符定义 |
| Background | 背景定义 |
| ParseLayout() | 解析版式 XML |
| ParseMaster() | 解析母版 XML |
| 枽数类型：PlaceholderType, BackgroundType, SlideLayoutType | |

## 3. Presentation 模块 (presentation.go)
**核心职责**: 演示文稿根节点

| 类型/函数 | 说明 |
| ---------- | ------ |
| PresentationPart | 演示文稿部件，对应 /ppt/presentation.xml |
| SlideSize | 幻灯片尺寸（EMU 单位） |
| StandardSlideSizes | 标准尺寸（16:9, 4:3） |
| EMUFromPoints() 等 | EMU 单位转换函数 |
| XML 结构类型：XPresentation, XSldIdLst, XSldMasterIdLst | |

## 4. Media 模块 (media.go, media_manager.go)
**核心职责**: 媒体资源管理

| 类型/函数 | 说明 |
| ---------- | ------ |
| MediaResource | 媒体资源（图片/音频/视频） |
| MediaManager | 媒体资源管理器（并发安全缓存） |
| MediaType | 媒体类型枚举 |
| NewMediaResourceFromBytes() | 从字节创建资源 |
| NewMediaResourceFromReader() | 从 Reader 创建资源 |
| 功能：去重、多索引（rID/fileName/target/hash）、MIME 类型推断 |

## 5. Relationship 模块 (relationship.go)
**核心职责**: OPC 关系 XML 结构定义（纯 DTO)

| 类型/函数 | 说明 |
| ---------- | ------ |
| XMLRelationships | 关系集合（用于 .rels 文件的序列化/反序列化） |
| XMLRelationship | 单个关系 |
| ParseRelationships() | 解析关系 XML |
| 常量：RelTypeImage, RelTypeSlide, RelTypeSlideLayout 等 |

**注意**： 关系管理逻辑已移至 `opc` 层的 `opc.Relationships`，此模块仅保留 XML DTO 用于 .rels 文件的读写。

## 6. CoreProps 模块 (coreprops.go)
**核心职责**: 核心属性

| 类型/函数 | 说明 |
| ---------- | ------ |
| XMLCoreProperties | 核心属性结构 |
| XMLW3CDTFDate | W3CDTF 日期格式 |
| ParseCoreProperties() | 解析核心属性 |

## 7. XML 工具模块 (xmlutils.go)
**核心职责**: XML 处理工具

| 函数/常量 | 说明 |
|----------|------|
| XMLDeclaration | XML 声明头常量 |
| StripNamespacePrefixes() | 去除命名空间前缀 |

## 8. XML Master Models (xml_master_models.go)
**核心职责**: 母版/版式 XML 解析用到的中间结构

| 类型 | 说明 |
|------|------|
| XMLOffset, XMLExtents, XMLTransform | 位置/尺寸/变换 |
| XMLPlaceholder | 占位符 |
| XMLShape, XMLShapeTree | 形状/形状树 |
| XMLBackground, XMLFillProperties | 背景/填充 |
| XMLSlideLayout, XMLSlideMaster | 版式/母版 |

---

# 架构分层图

```
┌─────────────────────────────────────────────────────────────┐
│                        opc 包                              │
│  职责：通用 OPC 规范实现（包、部件、关系管理）          │
│  - PackURI: 路径处理                                        │
│  - Relationships: 线程安全的关系管理 + 原子 ID 分配      │
│  - Part/Package: 部件和包的基础结构                     │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                        parts 包                            │
│  职责：PPTX 特定的 XML 结构定义 + 序列化/反序列化         │
├─────────────────────────────────────────────────────────────┤
│                                                            │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐     │
│  │   Slide      │  │   Master     │  │ Presentation │     │
│  │  幻灯片 XML  │  │ 母版/版式    │  │  演示文稿    │     │
│  └──────────────┘  └──────────────┘  └──────────────┘     │
│                                                            │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐     │
│  │    Media     │  │ Relationship │  │  CoreProps   │     │
│  │  媒体资源    │  │   XML DTO    │  │  核心属性    │     │
│  └──────────────┘  └──────────────┘  └──────────────┘     │
│                                                            │
│  ┌──────────────┐  ┌──────────────┐                       │
│  │   XMLUtils   │  │ XMLWriter    │                       │
│  │  XML 工具    │  │  流式写入    │                       │
│  └──────────────┘  └──────────────┘                       │
│                                                            │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                       slide 包                               │
│  职责：高层业务逻辑（SlideBuilder、MediaManager）            │
└─────────────────────────────────────────────────────────────┘
```

---

# 文件清单

| 文件 | 行数 | 主要内容 |
|------|------|----------|
| slide.go | ~1700 | SlidePart、XMLWriter、WriteXML 方法 |
| slide_types.go | ~450 | XML 结构类型定义 |
| media_manager.go | 460 | MediaManager 媒体管理器 |
| presentation.go | 393 | PresentationPart |
| master_parser.go | 344 | 母版/版式解析器 |
| master_types.go | 358 | 母版数据结构 |
| master_cache.go | 275 | MasterCache 缓存 |
| xml_master_models.go | 272 | 母版 XML 中间结构 |
| master.go | 255 | MasterManager |
| media.go | 244 | MediaResource |
| relationship.go | 179 | XMLRelationships（纯 DTO） |
| coreprops.go | 161 | XMLCoreProperties |
| xmlutils.go | 89 | XML 工具函数 |
