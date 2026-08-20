# PPT 导出旁白 MP4 视频标准化流程 (SOP) 与避坑知识库

## 1. 核心经验与避坑规约 (Lessons Learned)

### 避坑点 A: 基准文件错选导致页面元素丢失
- **根因**：若误选裁剪过/清洗过的中间 PPTX 临时文件，或正则表达式在改写 `<p:transition>` 标签时破坏了自闭合符号 `/>`，会导致 XML Parse Error。PowerPoint 打开损坏的 XML 时会弹窗并自动剔除修复失败的 Shapes，造成元素丢失。
- **SOP 规约**：
  1. 统一以 `with_notes.pptx` (包含 100% 原生排版、图标、表格与矢量图形) 为唯一单点真实源 (SSOT) 基准；
  2. 正则修改 transition 标签时必须保留自闭合 `/` 属性，且必须包含 XML 语法 Parse Error = 0 自动校验；
  3. 必须进行 Shapes 数量自动对比（`Shapes Mismatch == 0`）。

### 避坑点 B: 导出静音视频 (0 Channels)
- **根因**：仅仅在 XML 中配置切片时间 `advTm` 不足以让 PowerPoint 播放声音。缺失媒体 `<p:pic>` Shape、Relationships `rId` 以及 `<p:audio>` Timing 节点会导致 PowerPoint COM 认为该 Slide 为无声页面。
- **SOP 规约**：
  1. 音轨注入时必须补齐隐藏音频 Shape、Relationships 和 `<p:audio>` 页面载入即播放节点；
  2. 切换时间公式固化为：`advance_after = audio_duration + 1.0s (留白) + 1.5s (淡出)`，且显式置位 `advClick="0"`；
  3. 视频导出完成后必须自动调用 `mutagen` / `pymediainfo` 探查 MP4 元数据，确保 `channels >= 1`。

### 避坑点 C: PowerPoint COM 卡死与打不开文件
- **根因**：同名文档在 PowerPoint 中重复打开冲突、MOTW Web 区域锁定以及缺失字体的 Modal 提示弹窗。
- **SOP 规约**：
  1. 显式置位 `ppt_app.DisplayAlerts = 1` (ppAlertsNone) 拦截所有弹窗阻塞；
  2. 执行 `powershell Unblock-File` 自动解除 Web 区域锁；
  3. 打开前扫描 `ppt_app.Presentations` 句柄，自适应复用已打开的 Presentation。

---

## 2. 标准三步执行流水线

```bash
# Step 1: TTS 生成单页音频
python skills/ppt-master/scripts/notes_to_audio.py <project_path> --voice zh-CN-YunjianNeural

# Step 2: 基准无损音轨强同步
python skills/ppt-master/scripts/sync_narration_to_pptx.py <project_path> --input-pptx <project_path>/exports/<base_with_notes>.pptx

# Step 3: PowerPoint COM 1080p 超清导出
python skills/ppt-master/scripts/export_video_com.py <project_path> --input-pptx <project_path>/exports/<narrated_pptx>.pptx --resolution 1080
```
