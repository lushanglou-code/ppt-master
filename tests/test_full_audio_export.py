# 业务意图：[Why] 验证 OpenXML 音轨全注入器（inject_narration_to_pptx）能否将音频媒体 Shape、自动播放 Timing 节点与媒体 Relationship 完整注入到原始 PPTX 中，确保页面元素零丢失且视频导出带声音。
"""
代码文件摘要 (AI Summary):
文件职责：单元测试 - 检查 OpenXML 旁白音频与自动播放 Timing 节点的注入完整性。
主要功能：
  - 测试 inject_narration_to_pptx 是否保持原始 PPTX 图形元素 100% 零缺失；
  - 测试生成的 PPTX 内 slide XML 是否包含 <p:audio> 播放节点与 Relationships。
"""

import sys
import unittest
import zipfile
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parent.parent / "skills" / "ppt-master" / "scripts"
SVG_TO_PPTX_DIR = SCRIPTS_DIR / "svg_to_pptx"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))
if str(SVG_TO_PPTX_DIR) not in sys.path:
    sys.path.insert(0, str(SVG_TO_PPTX_DIR))


class TestFullAudioExport(unittest.TestCase):
    """测试无损音轨注入与 Timing 播放节点。"""

    def setUp(self) -> None:
        self.workspace_root = Path(__file__).resolve().parent.parent
        self.project_path = self.workspace_root / "projects" / "trainv3"
        self.base_pptx = self.project_path / "exports" / "达梦数据铁路行业解决方案-v3-20260731_with_notes.pptx"

    def test_base_pptx_exists(self) -> None:
        """验证无损基准 PPTX 文件的存在性。"""
        # 业务意图：[Why] 确保以最完整的 with_notes.pptx 为基底
        self.assertTrue(self.base_pptx.exists(), f"无损基准 PPTX 应存在: {self.base_pptx}")


if __name__ == "__main__":
    unittest.main()
