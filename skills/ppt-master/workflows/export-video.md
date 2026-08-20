---
description: Export narrated PPTX to MP4 presentation video with automatic page turns via PowerPoint COM API and Quality Gates
---

# Export Video Standardized SOP & Quality Gates Workflow

> 权威 SOP 规范：在 Windows 环境下将带有 TTS 演讲语音与转场 Timing 的 PPTX 渲染导出为 MP4 高清带声视频的标准流程与质量门禁指南。

---

## 核心 SOP 标准化四步流水线 (Standardized Pipeline)

```
[源文档/演讲备注] 
     │
     ▼ Step 1: TTS 语音合成 (notes_to_audio.py)
[audio/01.mp3 ~ 66.mp3]
     │
     ▼ Step 2: 基准文件选择 & 无损音轨同步 (inject/sync_narration)
[exports/*_narrated.pptx]  ──►  【质量门禁 1 & 2: XML 语法解析门禁 + Shape 零丢失门禁】
     │
     ▼ Step 3: PowerPoint COM 静默 1080p 视频渲染 (export_video_com.py)
[exports/*_1080p.mp4]     ──►  【质量门禁 3 & 4: 防挂起弹窗拦截 + 物理音轨与时长门禁】
```

---

## 标准执行命令 SOP (Standard Commands)

### Step 1: 生成单页旁白语音
```bash
python skills/ppt-master/scripts/notes_to_audio.py <project_path> --voice zh-CN-YunjianNeural
```

### Step 2: 基准文件无损音轨与 Timing 强同步 (注入自动播放节点与 1.0s 留白)
*规则*：必须以保留 100% 原始排版与 DrawingML 矢量 Shapes 的原装文件 (`with_notes.pptx`) 为唯一基准，禁止使用中间裁剪临时文件。
```bash
# 自动强同步旁白 MP3 字节流与切换时间 (advance_after = dur + 1.0s 留白 + 1.5s 淡出)
python skills/ppt-master/scripts/sync_narration_to_pptx.py <project_path> --input-pptx <project_path>/exports/<base_with_notes>.pptx
```

### Step 3: 触发 PowerPoint COM 高清视频渲染 (显式 1080p)
```bash
# 强制指定 1080p 高清分辨率导出
python skills/ppt-master/scripts/export_video_com.py <project_path> --input-pptx <project_path>/exports/<narrated_pptx>.pptx --resolution 1080
```

---

## 四大强制质量门禁 (Mandatory Quality Gates)

为彻底解决“页面元素丢失、视频静音无声、渲染路径锁挂起”等反复修改问题，必须严格通过以下四道质量门禁：

### 1. XML 语法解析门禁 (XML Parse Gate)
- **校验原则**：在音轨注入或打包后，必须自动运行 `xml.etree.ElementTree` 扫描全量 Slide XML。
- **合格标准**：`Parse Errors == 0`。严禁任何 `<p:transition>` 正则误截断自闭合符号 `/>` 的非 well-formed XML 流向导出环节。

### 2. 画面 Shape 零丢失门禁 (Shape Preservation Gate)
- **校验原则**：对比注入前后 Slide 1 ~ N 的原生 `sp` (Shape) / `pic` (Picture) 数量。
- **合格标准**：`Shapes Mismatch Count == 0`。基准 Shapes 必须 100% 完全相同。

### 3. COM 渲染环境防挂起门禁 (Process & Modal Gate)
- **校验原则**：
  - 启动 COM 前强行清理 `POWERPNT.EXE` 遗留进程锁；
  - 执行 `Unblock-File` 解除 Windows MOTW 区域锁定；
  - 显式置位 `ppt_app.DisplayAlerts = 1` (ppAlertsNone) 拦截所有字体/外部资源缺失模态弹窗。

### 4. 产物物理音质与时长校验门禁 (Video Verification Gate)
- **校验原则**：导出完成后自动探查 MP4 元数据：
  - `VertResolution == 1080` (1920x1080 分辨率)
  - `Audio Channels >= 1` (确保视频包含声音，严禁 0 Channels 静音 MP4)
  - `Exact Video Duration >= Total Narration Duration` (确保语音未被截断)
