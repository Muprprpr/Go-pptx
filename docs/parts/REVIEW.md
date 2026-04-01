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

## 6. Theme 模块 (theme.go, theme_types.go, theme_default.go)
**核心职责**: 主题模板管理（最小化处理）

| 类型/函数 | 说明 |
| ---------- | ------ |
| ThemePart | 主题部件，对应 /ppt/theme/themeN.xml |
| XTheme, XThemeElements | 主题 XML 结构 |
| XColorScheme, XColorVariant | 颜色方案 |
| XFontScheme, XFontCollection | 字体方案 |
| XFmtScheme | 格式方案（通过 InnerXML 保留原始数据） |
| DefaultThemeXML | 完整的 Office 主题模板常量 |
| DefaultTheme() | 获取默认主题（单例懒加载） |
| CloneTheme() | 克隆主题（深拷贝） |
| GetThemeColor/GetThemeColorRGB/GetThemeColorType | 颜色访问方法 |
| SetThemeColorRGB/SetThemeColorSystem | 颜色设置方法 |
| SetThemeMajorFont/SetThemeMinorFont/SetThemeScriptFont | 字体设置方法 |

**设计原则**：
1. **预留入口，非主要依据**：提供了颜色/字体的读写方法，但主题复杂度高，不建议深度定制
2. **模板优先**：通过 `DefaultThemeXML`（完整 Office 主题）+ `CloneTheme()` 保证生成的 PPTX 结构完整
3. **数据保留**：FmtScheme 使用 `InnerXML` 保留原始 XML，避免解析丢失

## 7. AppProps 模块 (appprops.go, appprops_types.go)
**核心职责**: 应用程序属性（公司、管理者等）

| 类型/函数 | 说明 |
| ---------- | ------ |
| AppPropsPart | 应用程序属性部件，对应 /docProps/app.xml |
| XMLAppProps | 应用属性 XML 结构 |
| GetAppCompany/SetAppCompany | 公司名称读写 |
| GetAppManager/SetAppManager | 管理者读写 |
| GetAppSlideCount/SetAppSlideCount | 幻灯片数量读写 |
| SetAppWordCount/SetAppTotalTime | 字数/编辑时间设置 |
| HeadingPairs/TitlesOfParts | 标题对/部件标题（InnerXML 保留） |

**设计说明**：
- OOXML 规定公司、管理者等元数据必须写在 `/docProps/app.xml`
- 方法统一使用 `App` 前缀避免与 Go 原生关键字冲突
- HeadingPairs 和 TitlesOfParts 使用 InnerXML 保留原始结构，避免复杂解析

## 8. CoreProps 模块 (coreprops.go)
**核心职责**: 核心属性

| 类型/函数 | 说明 |
| ---------- | ------ |
| XMLCoreProperties | 核心属性结构 |
| XMLW3CDTFDate | W3CDTF 日期格式 |
| ParseCoreProperties() | 解析核心属性 |

## 9. Chart 模块 (chart.go, chart_types.go)
**核心职责**: 图表部件（模板 + 占位符策略）

| 类型/函数 | 说明 |
| ---------- | ------ |
| ChartPart | 图表部件，对应 /ppt/charts/chartN.xml |
| ChartType | 图表类型枚举（Bar/Pie/Line/Area/Scatter/Doughnut） |
| ChartTemplateBar/Pie/Line... | 预定义图表模板常量 |
| SetTemplate/SetRawXML | 设置图表模板/原始 XML |
| ReplacePlaceholder | 替换单个占位符 |
| ReplacePlaceholders | 批量替换占位符 |
| SetExternalDataRID/GetExternalDataRID | 外部 Excel 数据引用 |
| HasExternalData | 检查是否有外部数据引用 |

**设计策略**：
- **模板 + 占位符**：不尝试用 Go Struct 映射复杂图表 XML（几百种元素组合）
- **预定义模板**：提供常见图表类型（柱状图、饼图、折线图等）
- **占位符替换**：`{{CHART_TITLE}}`、`{{CATEGORIES}}`、`{{SERIES_VALUES}}` 等
- **两种路线**：
  - 路线 C（无 Excel）：数据直接嵌入 `strCache`/`numCache`，无外部依赖
  - 路线 A/B（有 Excel）：通过 `externalData` 引用嵌入的 Excel 文件

**常用占位符**：
| 占位符 | 说明 |
|--------|------|
| `{{CHART_TITLE}}` | 图表标题 |
| `{{SERIES_NAME}}` | 系列名称 |
| `{{CATEGORIES}}` | 分类标签 XML 片段 |
| `{{SERIES_VALUES}}` | 数值 XML 片段 |
| `{{CAT_COUNT}}` | 分类数量 |
| `{{CAT_COUNT_PLUS_1}}` | 分类数量+1（用于 Excel 公式） |

## 10. Embedding 模块 (embedding.go)
**核心职责**: 嵌入数据部件

| 类型/函数 | 说明 |
| ---------- | ------ |
| EmbeddingPart | 嵌入部件，对应 /ppt/embeddings/*.xlsx |
| EmbeddingType | 嵌入类型枚举（Excel/Word/Other） |
| Data/SetData | 二进制数据读写 |
| SetDataReader | 从 Reader 设置数据 |
| DetectEmbeddingType | 从文件名检测类型 |

**设计说明**：
- 嵌入数据是二进制文件（如 Excel），不进行 XML 解析
- 提供 Reader/Writer 接口便于流式处理

## 11. XML 工具模块 (xmlutils.go)
**核心职责**: XML 处理工具

| 函数/常量 | 说明 |
|----------|------|
| XMLDeclaration | XML 声明头常量 |
| StripNamespacePrefixes() | 去除命名空间前缀 |

## 12. XML Master Models (xml_master_models.go)
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
│  │    Theme     │  │    Media     │  │ Relationship │     │
│  │  主题模板    │  │  媒体资源    │  │   XML DTO    │     │
│  └──────────────┘  └──────────────┘  └──────────────┘     │
│                                                            │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐     │
│  │  CoreProps   │  │     App      │  │   XMLUtils   │     │
│  │  核心属性    │  │  应用属性    │  │  XML 工具    │     │
│  └──────────────┘  └──────────────┘  └──────────────┘     │
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
| theme.go | ~490 | ThemePart、主题读写方法 |
| theme_types.go | ~180 | 主题 XML 结构类型 |
| theme_default.go | ~260 | 默认主题模板、CloneTheme |
| appprops.go | ~270 | AppPropsPart、应用属性读写方法 |
| appprops_types.go | ~100 | 应用属性 XML 结构类型 |
| chart.go | ~130 | ChartPart、图表读写方法 |
| chart_types.go | ~120 | 图表 XML 结构类型 |
| embedding.go | ~130 | EmbeddingPart、嵌入数据读写 |
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
