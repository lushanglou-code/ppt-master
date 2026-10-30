# PPT 视频配音默认音色规范 (Voice Defaults)

在后续所有 PPT 演讲视频生成、PPTX 录音旁白嵌入与解说音频合成任务中，默认采用以下音色规范：

## 默认音色配置
- **默认女声 (Default Female Voice)**:
  - 音色名称: `Ava`
  - ShortName / Model: `en-US-AvaMultilingualNeural`
  - 特点: 微软新一代多语言大模型女声，现代科技感、专业端庄、知性自然、中英文混读发音地道，彻底杜绝低幼感。
- **默认男声 (Default Male Voice)**:
  - 音色名称: `Brian`
  - ShortName / Model: `en-US-BrianMultilingualNeural`
  - 特点: 微软新一代多语言大模型成熟男声（与 Ava 属于同一多语言大模型家族），深沉稳健、成熟专业、商业顾问风范，英文技术术语发音原生自然。

## 执行要求
- 当用户要求“女声”或未明确指定具体音色时，一律默认选用 `en-US-AvaMultilingualNeural`。
- 当用户要求“男声”时，一律默认选用 `en-US-BrianMultilingualNeural`。
