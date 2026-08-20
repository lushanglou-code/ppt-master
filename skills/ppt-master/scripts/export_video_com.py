#!/usr/bin/env python3
"""
代码文件摘要 (AI Summary):
文件职责：使用 Windows PowerPoint COM 接口 (win32com) 将包含旁白与时间轴的 PPTX 自动化导出为 MP4 演讲视频。
主要功能：
  - 自动定位项目 exports/ 目录下的最新 PPTX (优先 _narrated 结尾的文件)
  - 自动开启每页幻灯片的 AdvanceOnTime 与关闭 AdvanceOnClick
  - 通过 PowerPoint COM 接口调用 Presentation.CreateVideo
  - 支持设置分辨率 (720p/1080p/4K)、画质 (Quality) 及帧率 (FPS)
  - 实时轮询导出进度与超时管理 (兼容处理 PowerPoint COM 写完媒体流后的句柄释放状态)
  - 导出结果自动存入 <project_path>/exports/<name>_<resolution>p.mp4
"""

import sys
import os
import time
import argparse
import re
import shutil
import subprocess
from pathlib import Path

def find_target_pptx(project_path: Path, input_pptx: str = None) -> Path:
    """寻找要导出的 PPTX 文件，并在音频不同步时自动触发 sync_narration_to_pptx"""
    # 业务意图：[Why] 自动检测硬盘 audio/ 最新音频与 PPTX 内时长的差异，若 PPTX 属于旧缩水版本则自动重新同步
    if input_pptx:
        p = Path(input_pptx)
        if p.exists():
            return p.resolve()
        p_sub = (project_path / input_pptx).resolve()
        if p_sub.exists():
            return p_sub
        p_exp = (project_path / "exports" / Path(input_pptx).name).resolve()
        if p_exp.exists():
            return p_exp
        raise FileNotFoundError(f"指定的 PPTX 文件不存在: {input_pptx}")
    else:
        exports_dir = project_path / "exports"
        if not exports_dir.is_dir():
            raise FileNotFoundError(f"项目导出目录不存在: {exports_dir}")

        pptx_files = [f for f in exports_dir.glob("*.pptx") if not f.name.startswith("~$")]
        if not pptx_files:
            raise FileNotFoundError(f"在 {exports_dir} 中未找到任何 .pptx 文件")

        narrated_files = [f for f in pptx_files if "_narrated" in f.stem]
        raw_files = [f for f in pptx_files if "_narrated" not in f.stem]
        
        if narrated_files:
            narrated_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
            target = narrated_files[0]
        else:
            pptx_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
            target = pptx_files[0]

    # 检查 audio/ 目录是否更新
    audio_dir = project_path / "audio"
    if audio_dir.is_dir():
        try:
            from sync_narration_to_pptx import sync_narration_to_pptx
            from pptx_narration import probe_audio_duration

            audio_files = list(audio_dir.glob("*.mp3")) + list(audio_dir.glob("*.m4a")) + list(audio_dir.glob("*.wav"))
            if audio_files:
                total_disk_audio = sum((probe_audio_duration(a) or 0.0) for a in audio_files)

                # 检查 target PPTX 内的 transition advTm
                import zipfile
                total_pptx_adv = 0.0
                if target.exists():
                    try:
                        with zipfile.ZipFile(target, "r") as z:
                            for name in z.namelist():
                                if name.startswith("ppt/slides/slide") and name.endswith(".xml") and not "_rels" in name:
                                    content = z.read(name).decode("utf-8")
                                    match = re.search(r'advTm="(\d+)"', content)
                                    if match:
                                        total_pptx_adv += float(match.group(1)) / 1000.0
                    except Exception:
                        pass

                # 若 PPTX 内部时长落后于磁盘音频超过 10 秒，自动先触发强同步
                if total_disk_audio > 0 and total_pptx_adv < (total_disk_audio - 10.0):
                    print(f"[提示] 检测到 PPTX 时长 ({total_pptx_adv:.1f}s) 显著小于最新音频总长 ({total_disk_audio:.1f}s)，自动触发音轨重新同步...")
                    synced_pptx = sync_narration_to_pptx(project_path=project_path, input_pptx_path=target.resolve())
                    return synced_pptx
        except Exception as e:
            print(f"[警告] 音频自动同步检测跳过: {e}", file=sys.stderr)

    return target

def export_ppt_to_video(
    pptx_path: Path,
    output_mp4_path: Path,
    resolution: int = 720,
    quality: int = 85,
    fps: int = 30,
    timeout_seconds: int = 1800,
    use_timings_and_narrations: bool = True
) -> Path:
    """使用 win32com 接口导出视频"""
    try:
        import win32com.client
    except ImportError:
        print("[错误] 未安装 pywin32 库。请运行: pip install pywin32", file=sys.stderr)
        sys.exit(1)

    abs_pptx = str(pptx_path.resolve())
    abs_mp4 = str(output_mp4_path.resolve())

    output_mp4_path.parent.mkdir(parents=True, exist_ok=True)

    print(f"==================================================")
    print(f"开始 PowerPoint COM 视频导出流程")
    print(f"输入 PPTX: {abs_pptx}")
    print(f"输出 MP4 : {abs_mp4}")
    print(f"目标分辨率: {resolution}p | 画质: {quality} | 帧率: {fps} fps")
    print(f"使用旁白/时间轴: {use_timings_and_narrations}")
    print(f"==================================================")

    ppt_app = None
    presentation = None

    try:
        # 强制清理挂起的 PowerPoint 遗留进程，释放文件锁
        try:
            subprocess.run(["taskkill", "/F", "/IM", "POWERPNT.EXE"], capture_output=True)
        except Exception:
            pass

        print("[1/3] 启动 PowerPoint COM 应用程序...")
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        ppt_app.Visible = True
        try:
            ppt_app.DisplayAlerts = 1  # ppAlertsNone: 禁用一切弹窗阻塞
        except Exception:
            pass

        print("[2/3] 准备演示文稿并配置转场属性...")
        abs_pptx = str(pptx_path.resolve())
        # 解除文件锁与 Web 区域限制
        try:
            subprocess.run(["powershell", "-Command", f'Unblock-File -Path "{abs_pptx}"'], capture_output=True)
        except Exception:
            pass

        presentation = None
        target_name = Path(abs_pptx).name
        for p in ppt_app.Presentations:
            try:
                if p.Name == target_name:
                    presentation = p
                    print(f"  [提示] 演示文稿 '{target_name}' 已在 PowerPoint 中打开，直接复用当前句柄。")
                    break
            except Exception:
                pass

        if presentation is None:
            for retry in range(3):
                try:
                    presentation = ppt_app.Presentations.Open(
                        abs_pptx, ReadOnly=True, WithWindow=True
                    )
                    if presentation:
                        break
                except Exception as e:
                    print(f"  [重试 {retry+1}/3] 打开演示文稿暂未就绪, 等待 2 秒...")
                    time.sleep(2)

        if presentation is None:
            raise RuntimeError(f"PowerPoint COM 无法打开演示文稿文件: {abs_pptx}")

        # 业务意图：[Why] 动态校验并强行保护每页 AdvanceTime，确保自动换片时间必定大于音频时长 + 2.5 秒缓冲，防止视频提前翻页切断音频
        for slide in presentation.Slides:
            try:
                slide.SlideShowTransition.AdvanceOnClick = False
                slide.SlideShowTransition.AdvanceOnTime = True

                max_media_len_sec = 0.0
                for shape in slide.Shapes:
                    try:
                        # 检查音频/媒体 shape (msoMedia = 16)
                        if shape.Type in (16, 14) or getattr(shape, "MediaType", None) in (1, 3):
                            mf = getattr(shape, "MediaFormat", None)
                            if mf and hasattr(mf, "Length"):
                                length_ms = float(mf.Length)
                                if length_ms > 0:
                                    sec = length_ms / 1000.0
                                    if sec > max_media_len_sec:
                                        max_media_len_sec = sec
                    except Exception:
                        pass

                if max_media_len_sec > 0:
                    safe_duration = max_media_len_sec + 2.5
                    current_adv = float(slide.SlideShowTransition.AdvanceTime)
                    if current_adv < safe_duration:
                        slide.SlideShowTransition.AdvanceTime = safe_duration
            except Exception:
                pass

        print("[3/3] 触发 CreateVideo 渲染线程...")
        presentation.CreateVideo(
            FileName=abs_mp4,
            UseTimingsAndNarrations=use_timings_and_narrations,
            DefaultSlideDuration=5,
            VertResolution=resolution,
            FramesPerSecond=fps,
            Quality=quality
        )

        start_time = time.time()
        last_reported = 0
        last_size = 0
        stable_count = 0

        while True:
            status = presentation.CreateVideoStatus
            elapsed = int(time.time() - start_time)

            current_size = os.path.getsize(abs_mp4) if os.path.exists(abs_mp4) else 0

            # 校验导出完成：
            # 1. 状态为 Done (2) 或 极短处理后
            # 2. 或是文件已生成 (>1MB) 且尺寸在多次检测中保持稳定/写出完毕
            if current_size > 1024 * 1024:  # 大于 1MB
                if current_size == last_size:
                    stable_count += 1
                else:
                    stable_count = 0

                if status == 2 or stable_count >= 3 or status == 3:
                    time.sleep(1)
                    size_mb = current_size / (1024 * 1024)
                    print(f"\n[成功] 视频渲染导出完成! 耗时: {elapsed} 秒")
                    print(f"输出路径: {abs_mp4}")
                    print(f"视频文件大小: {size_mb:.2f} MB")

                    # 自动质量门禁校验 (Video Output Quality Gate)
                    try:
                        from mutagen.mp4 import MP4
                        v = MP4(abs_mp4)
                        dur = getattr(v.info, "length", 0)
                        ch = getattr(v.info, "channels", 0)
                        print(f"[产物质量校验] 视频精确时长: {dur:.2f}s ({dur/60.0:.2f}min) | 音频通道数: {ch}")
                        if ch == 0:
                            print(f"[警告] 检测到导出的 MP4 音频通道为 0 (可能静音)，请核查 PPTX 音频节点！", file=sys.stderr)
                    except Exception as ve:
                        print(f"[提示] 视频元数据门禁扫描: {ve}")

                    return output_mp4_path

            last_size = current_size

            if elapsed - last_reported >= 5:
                size_str = f" (生成中: {current_size / (1024 * 1024):.2f} MB)" if current_size > 0 else ""
                print(f"  --> 视频渲染中... 已耗时 {elapsed} 秒{size_str}", end="\r", flush=True)
                last_reported = elapsed

            if elapsed > timeout_seconds:
                if current_size > 1024 * 1024:
                    print(f"\n[警告] 达到超时限制，但视频已成功写入 ({current_size / (1024 * 1024):.2f} MB)")
                    return output_mp4_path
                raise TimeoutError(f"渲染视频超时（超过 {timeout_seconds} 秒）！")

            time.sleep(2)

    except Exception as e:
        print(f"\n[错误] 导出视频失败: {e}", file=sys.stderr)
        raise
    finally:
        if presentation:
            try:
                presentation.Close()
            except Exception:
                pass
        if ppt_app:
            try:
                ppt_app.Quit()
            except Exception:
                pass

def main():
    parser = argparse.ArgumentParser(description="PowerPoint COM 视频自动导出工具")
    parser.add_argument("project_path", nargs="?", help="项目目录路径")
    parser.add_argument("--input-pptx", help="显式指定输入的 .pptx 文件路径")
    parser.add_argument("--output-mp4", help="显式指定输出的 .mp4 文件路径")
    parser.add_argument("--resolution", type=int, default=720, choices=[720, 1080, 2160], help="输出视频垂直分辨率 (默认: 720p)")
    parser.add_argument("--quality", type=int, default=85, help="视频画质质量 1-100 (默认: 85)")
    parser.add_argument("--fps", type=int, default=30, help="视频帧率 (默认: 30)")
    parser.add_argument("--timeout", type=int, default=1800, help="渲染最大超时时间(秒) (默认: 1800)")
    parser.add_argument("--no-timings", action="store_true", help="不使用旁白与翻页时间轴 (强制定长切换)")

    args = parser.parse_args()

    if not args.project_path and not args.input_pptx:
        parser.error("必须提供 project_path 或 --input-pptx 参数")

    proj_path = Path(args.project_path) if args.project_path else Path(args.input_pptx).parent.parent
    target_pptx = find_target_pptx(proj_path, args.input_pptx)

    if args.output_mp4:
        output_mp4 = Path(args.output_mp4)
    else:
        out_name = f"{target_pptx.stem}_{args.resolution}p.mp4"
        output_mp4 = proj_path / "exports" / out_name

    export_ppt_to_video(
        pptx_path=target_pptx,
        output_mp4_path=output_mp4,
        resolution=args.resolution,
        quality=args.quality,
        fps=args.fps,
        timeout_seconds=args.timeout,
        use_timings_and_narrations=not args.no_timings
    )

if __name__ == "__main__":
    main()
