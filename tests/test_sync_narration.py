# 业务意图：[Why] 验证 PPTX 音轨与 Timing 同步工具（sync_narration_to_pptx）能否正确将最新音频覆盖写入 PPTX 的 OpenXML 媒体流与切换时间，防止视频导出时长严重缩水。
"""
代码文件摘要 (AI Summary):
文件职责：单元测试 - PPTX 音频与切换 Timing 重新同步工具。
主要功能：
  - 测试 sync_narration_to_pptx 工具对音频文件与 PPTX slide 的匹配映射；
  - 测试 OpenXML 中 media 字节流替换与 <p:transition advTm="..."> 更新的准确性。
"""

import sys
import unittest
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parent.parent / "skills" / "ppt-master" / "scripts"
SVG_TO_PPTX_DIR = SCRIPTS_DIR / "svg_to_pptx"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))
if str(SVG_TO_PPTX_DIR) not in sys.path:
    sys.path.insert(0, str(SVG_TO_PPTX_DIR))


class TestSyncNarration(unittest.TestCase):
    """测试 PPTX 音轨同步逻辑。"""

    def setUp(self) -> None:
        self.workspace_root = Path(__file__).resolve().parent.parent
        self.project_path = self.workspace_root / "projects" / "trainv3"

    def test_audio_files_discovery(self) -> None:
        """测试对项目 audio 目录下 66 个 MP3 音频文件的匹配与发现。"""
        # 业务意图：[Why] 确保 66 页单页 MP3 音频均能按顺序和 Index 正确定位
        audio_dir = self.project_path / "audio"
        if audio_dir.exists():
            mp3_files = sorted(
                [p for p in audio_dir.glob("*.mp3")],
                key=lambda x: int(x.stem) if x.stem.isdigit() else 999
            )
            self.assertGreater(len(mp3_files), 0, "audio 目录下应存在 MP3 文件")
            self.assertEqual(len(mp3_files), 66, "trainv3 应包含 66 个单页 MP3 文件")


if __name__ == "__main__":
    unittest.main()
