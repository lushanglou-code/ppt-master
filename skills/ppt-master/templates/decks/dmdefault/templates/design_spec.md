---
deck_id: dmdefault
kind: deck
category: brand
summary: 达梦数据库官方经典科技汇报模板，蓝白主色，简洁大气，适用于技术交流与售前演示。
keywords: [达梦数据库, 达梦数据, DM, 官方模板, 科技汇报, 解决方案]
primary_color: "#0000D2"
canvas_format: ppt169
canvas_width: 1280
canvas_height: 720
canvas_viewbox: "0 0 1280 720"
source_canvas_width: 1280
source_canvas_height: 720
source_viewbox: "0 0 1280 720"
replication_mode: fidelity
native_structure_mode: structured
page_count: 5
---

# 达梦数据库官方默认模板 (dmdefault) — Design Specification

## I. 模板概述与应用场景 (Template Overview)

| 属性 | 定义与规范 |
| --- | --- |
| **模板代号** | `dmdefault` (达梦官方经典科技汇报模板) |
| **适用场景** | 达梦数据库售前交流、央企与行业解决方案汇报、信创技术交流、产品发布与答辩演示 |
| **设计核心准则** | **固化装饰母版，解放内容排版**：严格规范装饰转场页（封面、目录、章节过渡、封底）与全局设计元素（色彩、字体、标题规范、页脚标语）；对正文内容页内部排版（卡片、架构拓扑、跑分对比、流程图等）保持完全开放与自由度。 |
| **设计基调** | 权威、稳健、严谨、科技感、数据驱动 |

## II. 色彩体系与设计元素 (Color & Design Elements)

### 1. 色彩规范 (Color Tokens)

| 角色 | 色值 HEX | 业务意图与用法 |
| --- | --- | --- |
| **Primary (主色深蓝)** | `#0000D2` | 官方标准深蓝，用于封面大Slogan、目录项序号、过渡页大字号章节号 |
| **Secondary Accent (科技蓝)** | `#265EFA` | 科技亮蓝，用于红蓝双色微饰条、流程连线、重点高光 |
| **Accent Red (达梦正红)** | `#FF0000` / `#EC1B34` | 达梦点睛红，与科技蓝组成标志性垂直双色条；用于核心突破指标与高光徽章 |
| **Body Title (正文大标题)** | `#000000` / `#020202` | 内容页标准标题与封面主标题，高对比度保证大屏清晰 |
| **Body Text (正文文字)** | `#1E293B` / `#334155` | 正文卡片段落与描述说明 |
| **Muted / Footer (页脚与次要文字)** | `#808080` / `#A6A6A6` | 统一页脚标语、官方网址与次要注释 |

### 2. 字体体系 (Typography)

* **中文字体栈**：`HarmonyOS Sans SC`, `PingFang SC`, `Microsoft YaHei`, sans-serif
* **英文字体栈**：`Arial`, `Segoe UI`, `HarmonyOS Sans SC`, sans-serif
* **数据字体**：`DIN`, `Segoe UI`, Arial, sans-serif
* **字号梯度**：
  * 封面主标题：54px Bold
  * 封面副标题：24px Regular
  * 章节号：106px Bold
  * 章节名：80px Bold
  * 内容页主标题：40px Bold (统一 Y 轴锚点 y=69, 左侧固定配置红蓝微条)
  * 卡片/模块标题：18~22px Bold
  * 正文正文字号：14~16px Regular
  * 页眉页脚辅助字：16px / 21px Regular

## III. 母版固化契约 (Fixed Decorative Masters)

本模板的 4 类装饰页母版完全固化，不随具体内容被随意改动：
1. **封面 (`01_cover.svg`)**：
   - 包含官方科技蓝母版底纹与左上角彩色达梦 Logo；
   - 固化口号“打造中国的世界级数据库品牌”与企业署名“武汉达梦数据库股份有限公司”；
   - 开放插槽：`{{TITLE}}`（主标题）、`{{SUBTITLE}}`（副标题）、`{{DATE}}`（汇报日期）。
2. **目录页 (`02_toc.svg`)**：
   - 固化官方背景、右上角 Logo、左侧大号浅灰 `CONTENT` 字母装饰与“目录”主标题；
   - 固化页脚标语与官网网址；
   - 提供 4 组经典章节条目导航槽位 (`{{TOC_ITEM_1}}` ~ `{{TOC_ITEM_4}}`)。
3. **章节过渡页 (`03_chapter.svg`)**：
   - 固化横向微光渐变光带与左侧红蓝垂直双色微饰条；
   - 固化页脚标语与官网网址；
   - 开放插槽：`{{CHAPTER_NUM}}`（超大序号）、`{{CHAPTER_TITLE}}`（章节名称）。
4. **封底 (`05_ending.svg`)**：
   - 固化封面同款科技蓝底纹、左上角达梦 Logo；
   - 固化“感谢聆听”大标题与“打造中国的世界级数据库品牌”标语；
   - 固化官方二维码与服务联系方式（公司名、网址、400服务热线）。

## IV. 开放式内容页排版规范 (Open Content Principle)

* **内容页母版 (`04_content.svg`)**：
  - **固定母版边界**：
    - 浅色纯净底纹；
    - 标题左侧标准红蓝垂直微饰条（红 11.67×42.88，蓝 11.67×29.13）；
    - 标题标准格式：`x=80, y=69, font-size=40, font-weight=700, fill="#000000"`；
    - 右上角达梦标准 Logo（220.83×34.81）；
    - 页脚：标语“打造中国的世界级数据库品牌” + 网址“www.dameng.com”。
  - **正文内容区域（Body Viewport）完全开放**：
    - 视口范围：`x: 70 ~ 1210`，`y: 110 ~ 660`（宽 1140px，高 550px）；
    - **禁止任何强制性卡片锁定框或死板网格**；
    - 执行器（Executor）拥有完全的排版自由度，可根据具体的沟通诉求绘制：
      - 2~4 列模块化指标卡片；
      - “两地三中心”或分布式高可用架构拓扑图；
      - 性能跑分对比表与雷达图；
      - 迁移割接四步法时间轴与流水线；
      - 央企标杆案例成果看板。
