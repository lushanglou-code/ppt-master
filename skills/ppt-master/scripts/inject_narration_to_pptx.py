#!/usr/bin/env python3
"""
代码文件摘要 (AI Summary):
文件职责：以 100% 保持元素完整的原始 PPTX 为基底，无损注入旁白音频 Shape、Relationships 与自动播放 Timing 节点。
主要功能：
  - 严格保持基准 PPTX (with_notes.pptx) 的原始 DrawingML 矢量 Shapes、排版与样式 100% 零丢失
  - 注册 [Content_Types].xml 音频格式
  - 向 slide XML 注入隐藏音频 Shape (<p:pic>) 与页面载入即触发播放的 <p:audio> Timing 节点
  - 写入 ppt/slides/_rels/slide{i}.xml.rels 媒体 Relationship 关联
  - 将 audio/ 目录下 66 个 MP3 写入 ppt/media/ 目录
  - 精确计算并注入 <p:transition advClick="0" advTm="..."> 自动切换时间
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

SCRIPTS_DIR = Path(__file__).resolve().parent
SVG_TO_PPTX_DIR = SCRIPTS_DIR / "svg_to_pptx"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))
if str(SVG_TO_PPTX_DIR) not in sys.path:
    sys.path.insert(0, str(SVG_TO_PPTX_DIR))

from svg_to_pptx.pptx_package.narration import (
    AUDIO_MARKER_PNG_BYTES as TRANSPARENT_PNG_BYTES,
    AUDIO_REL_TYPE,
    IMAGE_REL_TYPE,
    MEDIA_REL_TYPE,
    NARRATION_EXTENSIONS,
    apply_recorded_timing,
    inject_narration,
    next_shape_id,
    probe_audio_duration,
)


def _leading_number(text: str) -> int | None:
    match = re.match(r"^(\d{1,3})", text.strip())
    return int(match.group(1)) if match else None


def _ensure_content_types(content_types_xml: str) -> str:
    """确保 [Content_Types].xml 包含 mp3 类型的映射"""
    if 'Extension="mp3"' in content_types_xml or 'Extension="MP3"' in content_types_xml:
        return content_types_xml

    mp3_entry = '<Default Extension="mp3" ContentType="audio/mpeg"/>'
    png_entry = '<Default Extension="png" ContentType="image/png"/>'

    if mp3_entry not in content_types_xml:
        content_types_xml = content_types_xml.replace("</Types>", f"  {mp3_entry}\n</Types>", 1)
    if 'Extension="png"' not in content_types_xml and png_entry not in content_types_xml:
        content_types_xml = content_types_xml.replace("</Types>", f"  {png_entry}\n</Types>", 1)

    return content_types_xml


def inject_narration_to_pptx(
    project_path: Path,
    input_pptx_path: Path | None = None,
    output_pptx_path: Path | None = None,
    narration_padding: float = 1.0,
    transition_duration: float = 1.5,
) -> Path:
    """无损注入旁白音频 Shape 与播放节点到 PPTX 中"""
    # 业务意图：[Why] 以保留 100% 原始画面的 PPTX 为基底，补全音频 Shape 与 <p:audio> 自动播放节点，解决元素丢失和导出视频无声音的问题
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

    # 优先选取无损的 with_notes.pptx 原始模板
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
        candidates = list(exports_dir.glob("*with_notes*.pptx")) if exports_dir.is_dir() else []
        if not candidates and exports_dir.is_dir():
            candidates = list(exports_dir.glob("*.pptx"))
        if not candidates:
            candidates = list(project_path.glob("*.pptx"))

        candidates = [c for c in candidates if not c.name.startswith("~$") and "_narrated" not in c.name]
        candidates.sort(key=lambda x: x.stat().st_mtime, reverse=True)
        if not candidates:
            raise FileNotFoundError(f"未找到可用的基准 PPTX 文件")
        pptx_in = candidates[0]

    if not pptx_in.exists():
        raise FileNotFoundError(f"输入的基准 PPTX 文件不存在: {pptx_in}")

    if output_pptx_path:
        pptx_out = Path(output_pptx_path)
    else:
        stem = pptx_in.stem
        base_stem = re.sub(r"(_with_notes|_clean|_narrated|_synced)+$", "", stem)
        out_name = f"{base_stem}_1080p_narrated.pptx"
        pptx_out = project_path / "exports" / out_name

    pptx_out.parent.mkdir(parents=True, exist_ok=True)

    print(f"==================================================")
    print(f"开始无损旁白音轨与 Timing 全节点注入流程")
    print(f"基准 PPTX: {pptx_in}")
    print(f"目标 PPTX: {pptx_out}")
    print(f"音频数量: {len(audio_files)} 个")
    print(f"==================================================")

    # 索引音频映射
    audio_by_index: dict[int, Path] = {}
    for idx, a_path in enumerate(audio_files, 1):
        num = _leading_number(a_path.stem)
        slide_num = num if num is not None else idx
        audio_by_index[slide_num] = a_path

    tmp_out = pptx_out.with_suffix(".tmp.pptx")
    total_audio_time = 0.0
    poster_name = "image_poster.png"

    # 先预处理并精准生成每一页 Slide 的全新 Relationships 与 rId
    slide_rids_map: dict[int, tuple[str, str, str, str]] = {} # s_num -> (audio_rid, media_rid, poster_rid, updated_rels_str)

    with zipfile.ZipFile(pptx_in, "r") as z_in:
        for name in z_in.namelist():
            if name.startswith("ppt/slides/_rels/slide") and name.endswith(".xml.rels"):
                s_match = re.search(r"slide(\d+)\.xml\.rels$", name)
                if s_match:
                    s_num = int(s_match.group(1))
                    audio_path = audio_by_index.get(s_num)
                    if audio_path and audio_path.exists():
                        rels_str = z_in.read(name).decode("utf-8")
                        audio_filename = f"narration_{s_num:03d}{audio_path.suffix.lower()}"

                        rids = [int(m) for m in re.findall(r'rId(\d+)', rels_str)]
                        max_rid = max(rids, default=1)
                        audio_rid = f"rId{max_rid + 1}"
                        media_rid = f"rId{max_rid + 2}"
                        poster_rid = f"rId{max_rid + 3}"

                        new_rels = f'''  <Relationship Id="{audio_rid}" Type="{AUDIO_REL_TYPE}" Target="../media/{audio_filename}"/>
  <Relationship Id="{media_rid}" Type="{MEDIA_REL_TYPE}" Target="../media/{audio_filename}"/>
  <Relationship Id="{poster_rid}" Type="{IMAGE_REL_TYPE}" Target="../media/{poster_name}"/>
</Relationships>'''
                        updated_rels_str = rels_str.replace("</Relationships>", new_rels, 1)
                        slide_rids_map[s_num] = (audio_rid, media_rid, poster_rid, updated_rels_str)

    with zipfile.ZipFile(pptx_in, "r") as z_in, zipfile.ZipFile(tmp_out, "w", compression=zipfile.ZIP_DEFLATED) as z_out:
        # 1. 处理 [Content_Types].xml
        content_types = z_in.read("[Content_Types].xml").decode("utf-8")
        content_types = _ensure_content_types(content_types)
        z_out.writestr("[Content_Types].xml", content_types.encode("utf-8"), compress_type=zipfile.ZIP_DEFLATED)

        # 写入透明 PNG 海报图
        poster_path_in_zip = f"ppt/media/{poster_name}"
        z_out.writestr(poster_path_in_zip, TRANSPARENT_PNG_BYTES, compress_type=zipfile.ZIP_DEFLATED)

        # 2. 遍历文件打包
        for item in z_in.infolist():
            name = item.filename
            if name == "[Content_Types].xml":
                continue

            content = z_in.read(name)

            # 处理 slide{i}.xml.rels
            if name.startswith("ppt/slides/_rels/slide") and name.endswith(".xml.rels"):
                s_match = re.search(r"slide(\d+)\.xml\.rels$", name)
                if s_match:
                    s_num = int(s_match.group(1))
                    if s_num in slide_rids_map:
                        content = slide_rids_map[s_num][3].encode("utf-8")

            # 处理 slide{i}.xml
            elif name.startswith("ppt/slides/slide") and name.endswith(".xml") and not "_rels" in name:
                s_match = re.search(r"slide(\d+)\.xml$", name)
                if s_match:
                    s_num = int(s_match.group(1))
                    audio_path = audio_by_index.get(s_num)
                    if audio_path and audio_path.exists() and s_num in slide_rids_map:
                        duration = probe_audio_duration(audio_path)
                        if duration and duration > 0:
                            total_audio_time += duration
                            adv_total = duration + narration_padding + transition_duration

                            slide_xml_str = content.decode("utf-8")
                            audio_filename = f"narration_{s_num:03d}{audio_path.suffix.lower()}"

                            audio_rid, media_rid, poster_rid, _ = slide_rids_map[s_num]

                            # 注入隐藏音频 Shape 与自动播放 Timing 节点
                            shape_id = next_shape_id(slide_xml_str)
                            slide_xml_str = inject_narration(
                                slide_xml_str,
                                shape_id=shape_id,
                                shape_name=audio_filename,
                                audio_rid=audio_rid,
                                media_rid=media_rid,
                                poster_rid=poster_rid,
                            )

                            # 注入切换时间
                            slide_xml_str = apply_recorded_timing(
                                slide_xml_str,
                                advance_after=adv_total,
                                transition_duration=transition_duration,
                                transition_effect="fade",
                            )
                            content = slide_xml_str.encode("utf-8")

                            # 写入 MP3 文件到 ppt/media/
                            media_in_zip_path = f"ppt/media/{audio_filename}"
                            z_out.writestr(media_in_zip_path, audio_path.read_bytes(), compress_type=zipfile.ZIP_DEFLATED)

            z_out.writestr(name, content, compress_type=zipfile.ZIP_DEFLATED)

    if tmp_out.exists():
        if pptx_out.exists():
            try:
                pptx_out.unlink()
            except Exception:
                pass
        shutil.move(str(tmp_out), str(pptx_out))

    print(f"[成功] 无损旁白音轨与 Timing 节点注入完成！")
    print(f"已处理音轨总时长: {total_audio_time:.2f} 秒 ({total_audio_time/60.0:.2f} 分钟)")
    print(f"输出 PPTX: {pptx_out}")

    return pptx_out


def main() -> int:
    parser = argparse.ArgumentParser(description="无损 PPTX 旁白音轨与 Timing 全节点注入工具")
    parser.add_argument("project_path", help="项目目录路径")
    parser.add_argument("--input-pptx", help="指定输入的基准 .pptx 文件")
    parser.add_argument("--output-pptx", help="指定输出的 .pptx 文件")
    parser.add_argument("--padding", type=float, default=1.0, help="旁白播放完后的静止留白秒数 (默认: 1.0)")
    parser.add_argument("--transition-dur", type=float, default=1.5, help="转场动画秒数 (默认: 1.5)")

    args = parser.parse_args()
    proj_path = Path(args.project_path)

    inject_narration_to_pptx(
        project_path=proj_path,
        input_pptx_path=Path(args.input_pptx) if args.input_pptx else None,
        output_pptx_path=Path(args.output_pptx) if args.output_pptx else None,
        narration_padding=args.padding,
        transition_duration=args.transition_dur,
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
