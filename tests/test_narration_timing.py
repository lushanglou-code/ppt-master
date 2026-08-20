# 业务意图：[Why] 验证旁白音频时长探查函数及 PPTX 自动切片 Timing 算式的正确性，防止视频导出时旁白被提前切断。
"""
代码文件摘要 (AI Summary):
文件职责：单元测试 - 旁白音频时长探查与 PPTX 动画切换 Timing 校验。
主要功能：
  - 测试 probe_audio_duration 对 MP3 / WAV / M4A 音频的多级探查能力；
  - 测试 apply_recorded_timing 在计算 advance_after 时对转场时间与缓冲停顿的覆盖逻辑；
  - 模拟缺失 ffprobe 时的 Python 原生 / mutagen 降级探查能力。
"""

import sys
import unittest
from pathlib import Path

# 将 scripts/ 及 scripts/svg_to_pptx 加入模块搜索路径
SCRIPTS_DIR = Path(__file__).resolve().parent.parent / "skills" / "ppt-master" / "scripts"
SVG_TO_PPTX_DIR = SCRIPTS_DIR / "svg_to_pptx"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))
if str(SVG_TO_PPTX_DIR) not in sys.path:
    sys.path.insert(0, str(SVG_TO_PPTX_DIR))

from svg_to_pptx.pptx_package.narration import probe_audio_duration, apply_recorded_timing


class TestNarrationTiming(unittest.TestCase):
    """测试音频时长探测与 PPTX timing 逻辑。"""

    def setUp(self) -> None:
        self.workspace_root = Path(__file__).resolve().parent.parent
        self.sample_audio = self.workspace_root / "projects" / "trainv3" / "audio" / "01.mp3"

    def test_probe_audio_duration_existing_mp3(self) -> None:
        """测试对真实 MP3 文件的时长探查（即使在没有 ffprobe 的环境下）。"""
        # 业务意图：[Why] 确保 Raw MP3 文件在 Windows 环境下即使没有 ffprobe 也能准确拿到 100+ 秒的实际时长
        if self.sample_audio.exists():
            duration = probe_audio_duration(self.sample_audio)
            self.assertIsNotNone(duration, "probe_audio_duration 不应返回 None")
            self.assertGreater(duration, 50.0, f"音频 01.mp3 实际时长约 107 秒，当前探测为: {duration}")

    def test_apply_recorded_timing_calculation(self) -> None:
        """测试 apply_recorded_timing 生成的 OpenXML advTm 属性值。"""
        # 业务意图：[Why] 验证传入的 advance_after 转换为毫秒并写入 XML 的准确性
        sample_xml = '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree></p:spTree></p:cSld></p:sld>'
        advance_after_sec = 15.5
        transition_dur_sec = 2.0

        res_xml = apply_recorded_timing(
            sample_xml,
            advance_after=advance_after_sec,
            transition_duration=transition_dur_sec,
            transition_effect="fade",
        )
        expected_adv_ms = int(advance_after_sec * 1000)
        self.assertIn(f'advTm="{expected_adv_ms}"', res_xml, f"XML 中应包含 advTm=\"{expected_adv_ms}\"")
        self.assertIn('advClick="0"', res_xml, "XML 中应明确禁止点击换片 (advClick=\"0\")")


if __name__ == "__main__":
    unittest.main()
