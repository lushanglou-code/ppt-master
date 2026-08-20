#!/usr/bin/env python3
"""
代码文件摘要 (AI Summary):
文件职责：将项目 audio/ 目录下的最新旁白音频二进制数据与时长重新同步/替换写入 PPTX 文件中。
主要功能：
  - 自动定位项目 audio/ 目录下的最新旁白音频 (01.mp3, 02.mp3, ...)
  - 解包与更新 PPTX OpenXML 中的 ppt/media/ 媒体流数据
  - 重新计算并注入每页 Slide 的 <p:transition advClick="0" advTm="..."> 定时切片时间
  - 解决音频文件更新后 PPTX 内部依然残留旧短媒体导致的视频时长严重缩水问题
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import sys
import zipfile
from pathlib import Path

# 将 scripts/ 及 scripts/svg_to_pptx 加入 sys.path
SCRIPTS_DIR = Path(__file__).resolve().parent
SVG_TO_PPTX_DIR = SCRIPTS_DIR / "svg_to_pptx"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))
if str(SVG_TO_PPTX_DIR) not in sys.path:
    sys.path.insert(0, str(SVG_TO_PPTX_DIR))

from svg_to_pptx.pptx_package.narration import (
    NARRATION_EXTENSIONS,
    apply_recorded_timing,
    find_narration_files,
    inject_narration,
    next_shape_id,
    probe_audio_duration,
)


def _leading_number(text: str) -> int | None:
    """提取文件名开头的页码数字"""
    match = re.match(r"^(\d{1,3})", text.strip())
    return int(match.group(1)) if match else None


def sync_narration_to_pptx(
    project_path: Path,
    input_pptx_path: Path | None = None,
    output_pptx_path: Path | None = None,
    narration_padding: float = 1.0,
    transition_duration: float = 1.5,
) -> Path:
    """把 project_path/audio/ 里的最新音频同步注入替换到 input_pptx_path 中"""
    # 业务意图：[Why] 彻底保证 PPTX 内部媒体二进制流与切换 Timing 与硬盘上的最新 MP3 文件 100% 保持同步
    project_path = Path(project_path).resolve()
    audio_dir = project_path / "audio"
    if not audio_dir.exists() or not audio_dir.is_dir():
        raise FileNotFoundError(f"项目音频目录不存在: {audio_dir}")

    audio_files = sorted(
        [p for p in audio_dir.iterdir() if p.is_file() and p.suffix.lower() in NARRATION_EXTENSIONS],
        key=lambda x: _leading_number(x.stem) or 999,
    )
    if not audio_files:
        raise FileNotFoundError(f"在 {audio_dir} 中未找到任何旁白音频文件")

    # 寻找 input pptx
    if input_pptx_path:
        pptx_in = Path(input_pptx_path)
        if not pptx_in.exists():
            candidate = (project_path / input_pptx_path).resolve()
            if candidate.exists():
                pptx_in = candidate
            else:
                candidate_sub = (project_path / "exports" / Path(input_pptx_path).name).resolve()
                if candidate_sub.exists():
                    pptx_in = candidate_sub
    else:
        exports_dir = project_path / "exports"
        candidates = list(exports_dir.glob("*.pptx")) if exports_dir.is_dir() else []
        if not candidates:
            candidates = list(project_path.glob("*.pptx"))
        if not candidates:
            raise FileNotFoundError(f"在项目 {project_path} 中未找到任何 .pptx 文件")

        # 过滤掉带有 ~$ 临时文件，优先选 clean 或原始 PPTX
        candidates = [c for c in candidates if not c.name.startswith("~$")]
        candidates.sort(key=lambda x: x.stat().st_mtime, reverse=True)
        pptx_in = candidates[0]

    if not pptx_in.exists():
        raise FileNotFoundError(f"输入的 PPTX 文件不存在: {pptx_in}")

    if output_pptx_path:
        pptx_out = Path(output_pptx_path)
    else:
        stem = pptx_in.stem
        base_stem = re.sub(r"(_clean|_narrated|_synced)+$", "", stem)
        out_name = f"{base_stem}_clean_narrated.pptx"
        pptx_out = project_path / "exports" / out_name

    pptx_out.parent.mkdir(parents=True, exist_ok=True)

    print(f"==================================================")
    print(f"开始 PPTX 音轨与 Timing 强同步注入流程")
    print(f"项目目录: {project_path}")
    print(f"输入 PPTX: {pptx_in}")
    print(f"输出 PPTX: {pptx_out}")
    print(f"找到旁白音频: {len(audio_files)} 个")
    print(f"==================================================")

    # 将音频文件按页码序号索引映射 {slide_index: audio_path}
    audio_by_index: dict[int, Path] = {}
    for idx, a_path in enumerate(audio_files, 1):
        num = _leading_number(a_path.stem)
        slide_num = num if num is not None else idx
        audio_by_index[slide_num] = a_path

    # 读取并重构 ZIP 内部文件
    tmp_out = pptx_out.with_suffix(".tmp.pptx")
    total_audio_time = 0.0

    with zipfile.ZipFile(pptx_in, "r") as z_in, zipfile.ZipFile(tmp_out, "w", compression=zipfile.ZIP_DEFLATED) as z_out:
        # 1. 扫描 XML 与 Relationship 构建 slide 旁白音频媒体映射
        rel_media_map: dict[int, list[str]] = {}
        for item in z_in.infolist():
            name = item.filename
            if name.startswith("ppt/slides/_rels/slide") and name.endswith(".xml.rels"):
                slide_num_str = re.search(r"slide(\d+)\.xml\.rels$", name)
                if slide_num_str:
                    s_num = int(slide_num_str.group(1))
                    rels_content = z_in.read(name).decode("utf-8")
                    # 严禁匹配 image 关系！只提取音频/媒体（audio/media/narration/mp3/m4a/wav）Target
                    try:
                        import xml.etree.ElementTree as ET
                        root_rels = ET.fromstring(rels_content)
                        audio_targets = []
                        for rel in root_rels.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                            r_type = rel.get("Type", "")
                            r_target = rel.get("Target", "")
                            # 仅保留音频或媒体关系，严禁包含 image 关系
                            if ("audio" in r_type or "media" in r_type or any(r_target.lower().endswith(ext) for ext in NARRATION_EXTENSIONS)) and "image" not in r_type:
                                if "../media/" in r_target:
                                    t_name = r_target.split("../media/")[-1]
                                    audio_targets.append(f"ppt/media/{t_name}")
                        if audio_targets:
                            rel_media_map[s_num] = audio_targets
                    except Exception:
                        pass

        # 2. 遍历 ZIP 入库，执行替换与 Timing 调整
        media_written: set[str] = set()

        for item in z_in.infolist():
            name = item.filename
            content = z_in.read(name)

            # A. 处理 slide{i}.xml 的 transition advTm
            if name.startswith("ppt/slides/slide") and name.endswith(".xml") and not "_rels" in name:
                s_match = re.search(r"slide(\d+)\.xml$", name)
                if s_match:
                    s_num = int(s_match.group(1))
                    audio_path = audio_by_index.get(s_num)
                    if audio_path and audio_path.exists():
                        duration = probe_audio_duration(audio_path)
                        if duration and duration > 0:
                            total_audio_time += duration
                            adv_total = duration + narration_padding + transition_duration
                            slide_xml_str = content.decode("utf-8")
                            updated_xml = apply_recorded_timing(
                                slide_xml_str,
                                advance_after=adv_total,
                                transition_duration=transition_duration,
                                transition_effect="fade",
                            )
                            content = updated_xml.encode("utf-8")

            # B. 处理 ppt/media/ 媒体流替换（只限旁白音频文件，绝对不覆盖图片文件）
            if name.startswith("ppt/media/"):
                is_audio_file = any(name.lower().endswith(ext) for ext in NARRATION_EXTENSIONS) or "narration" in name.lower()
                if is_audio_file:
                    for s_num, m_paths in rel_media_map.items():
                        if name in m_paths:
                            audio_path = audio_by_index.get(s_num)
                            if audio_path and audio_path.exists():
                                content = audio_path.read_bytes()
                                media_written.add(name)
                                break

            z_out.writestr(name, content, compress_type=zipfile.ZIP_DEFLATED)

    if tmp_out.exists():
        if pptx_out.exists():
            try:
                pptx_out.unlink()
            except Exception:
                pass
        shutil.move(str(tmp_out), str(pptx_out))

    print(f"[成功] PPTX 音轨与 Timing 重新同步完成！")
    print(f"已同步处理音轨总时长: {total_audio_time:.2f} 秒 ({total_audio_time/60.0:.2f} 分钟)")
    print(f"输出 PPTX: {pptx_out}")

    return pptx_out


def main() -> int:
    parser = argparse.ArgumentParser(description="PPTX 旁白音频与 Timing 强同步注入工具")
    parser.add_argument("project_path", help="项目目录路径")
    parser.add_argument("--input-pptx", help="指定输入的 .pptx 文件")
    parser.add_argument("--output-pptx", help="指定输出的 .pptx 文件")
    parser.add_argument("--padding", type=float, default=1.0, help="旁白播放完后的静止留白秒数 (默认: 1.0)")
    parser.add_argument("--transition-dur", type=float, default=1.5, help="转场动画秒数 (默认: 1.5)")

    args = parser.parse_args()
    proj_path = Path(args.project_path)

    import shutil  # 显式导入保证全局可用

    sync_narration_to_pptx(
        project_path=proj_path,
        input_pptx_path=Path(args.input_pptx) if args.input_pptx else None,
        output_pptx_path=Path(args.output_pptx) if args.output_pptx else None,
        narration_padding=args.padding,
        transition_duration=args.transition_dur,
    )
    return 0


if __name__ == "__main__":
    import shutil
    sys.exit(main())
