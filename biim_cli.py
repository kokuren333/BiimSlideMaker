#!/usr/bin/env python3
"""Agent-friendly CLI for authoring and rendering Biim-style videos."""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any

import requests
import yaml
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent
ACTIONS = {"idle", "wave", "nod", "think", "point", "cheer", "walk", "surprise"}
DEFAULT_VOICE = {
    "name": "kokuren_3rd",
    "speaker_uuid": "38d7216c-e595-4d8f-b06c-1fc376e47c0a",
    "style_name": "ノーマル",
    "style_id": 1069147200,
}


def split_script(value: str) -> list[str]:
    value = (value or "").replace("\r", "").replace("\n", "")
    chunks = [part.strip() for part in re.split(r"(?<=[。！？!?])", value) if part.strip()]
    if not chunks and value.strip():
        chunks = [value.strip()]
    return chunks


def resolve(base: Path, value: str | None) -> Path | None:
    if not value:
        return None
    path = Path(value).expanduser()
    return path if path.is_absolute() else (base / path).resolve()


def load_project(path: Path) -> tuple[dict[str, Any], Path]:
    project_file = path.resolve()
    if project_file.is_dir():
        project_file = project_file / "project.yaml"
    if not project_file.is_file():
        raise FileNotFoundError(f"Project file not found: {project_file}")
    data = yaml.safe_load(project_file.read_text(encoding="utf-8")) or {}
    if not isinstance(data, dict):
        raise ValueError("Project root must be a YAML mapping")
    return data, project_file.parent


def validate_project(data: dict[str, Any], base: Path) -> list[str]:
    errors: list[str] = []
    if int(data.get("version", 1)) != 1:
        errors.append("version must be 1")
    canvas = data.get("canvas", {})
    width, height = int(canvas.get("width", 1920)), int(canvas.get("height", 1080))
    if width <= 0 or height <= 0 or width * 9 != height * 16:
        errors.append("canvas must be a positive 16:9 size (default 1920x1080)")
    slides = data.get("slides")
    if not isinstance(slides, list) or not slides:
        errors.append("slides must be a non-empty list")
        return errors
    seen: set[str] = set()
    for index, slide in enumerate(slides, 1):
        label = f"slides[{index}]"
        if not isinstance(slide, dict):
            errors.append(f"{label} must be a mapping")
            continue
        sid = str(slide.get("id", index))
        if sid in seen:
            errors.append(f"{label}.id duplicates {sid}")
        seen.add(sid)
        image = resolve(base, slide.get("image"))
        if image is None or not image.is_file():
            errors.append(f"{label}.image does not exist: {slide.get('image')}")
        else:
            suffix = image.suffix.lower()
            if suffix not in {".png", ".jpg", ".jpeg", ".webp", ".svg"}:
                errors.append(f"{label}.image must be PNG, JPEG, WebP, or SVG")
            if suffix == ".svg":
                try:
                    import cairosvg  # noqa: F401
                except ImportError:
                    errors.append("SVG slides require cairosvg; install requirements.txt")
        utterances = split_script(str(slide.get("script", "")))
        if not utterances:
            errors.append(f"{label}.script is empty")
        tts_texts = slide.get("tts_texts", []) or []
        if not isinstance(tts_texts, list):
            errors.append(f"{label}.tts_texts must be a list of per-sentence readings")
        else:
            if len(tts_texts) > len(utterances):
                errors.append(f"{label}.tts_texts has more entries than script utterances")
            for reading in tts_texts:
                if not isinstance(reading, str):
                    errors.append(f"{label}.tts_texts entries must be strings")
        motions = slide.get("motions", []) or []
        if not isinstance(motions, list):
            errors.append(f"{label}.motions must be a list")
        else:
            for action in motions:
                if not isinstance(action, str) or action not in ACTIONS:
                    errors.append(f"{label}.motions has unknown action {action!r}")
            if len(motions) > len(utterances):
                errors.append(f"{label}.motions has more entries than script utterances")
    assets = data.get("assets", {}) or {}
    if not isinstance(assets, dict):
        errors.append("assets must be a mapping")
        assets = {}
    background = resolve(base, assets.get("background"))
    if background is None or not background.is_file():
        errors.append("assets.background must point to an existing frame image")
    animations = resolve(base, assets.get("animations", "assets/animations"))
    if animations is None or not (animations / "manifest.json").is_file():
        errors.append("assets.animations must contain manifest.json")
    elif isinstance(slides, list):
        try:
            action_manifest = json.loads((animations / "manifest.json").read_text(encoding="utf-8"))
            for slide in slides:
                if not isinstance(slide, dict):
                    continue
                requested = {action for action in (slide.get("motions", []) or []) if isinstance(action, str)} | {"idle"}
                for action in requested:
                    entry = action_manifest.get("actions", {}).get(action)
                    preview = animations / (entry.get("preview") if entry else f"{action}.gif")
                    if action not in ACTIONS or not preview.is_file():
                        errors.append(f"Missing animation for action {action!r}: {preview}")
        except (OSError, ValueError, TypeError) as exc:
            errors.append(f"assets.animations manifest is invalid: {exc}")
    for asset_name in ("bgm", "font"):
        value = assets.get(asset_name)
        if value:
            asset_path = resolve(base, value)
            if asset_path is None or not asset_path.is_file():
                errors.append(f"assets.{asset_name} does not exist: {value}")
    voice = data.get("voice", {}) or {}
    try:
        if int(voice.get("style_id", DEFAULT_VOICE["style_id"])) <= 0:
            errors.append("voice.style_id must be a positive integer")
    except (TypeError, ValueError):
        errors.append("voice.style_id must be a positive integer")
    if width > 0 and height > 0:
        layout = data.get("layout", {}) or {}
        boxes = {
            "slide": [40, 28, 1280, 720],
            "subtitle": [330, 847, 1520, 163],
            "notes_top": [1413, 66, 444, 324],
            "notes_bottom": [1413, 410, 444, 310],
            "character": [35, 795, 250, 250],
        }
        scale = height / 1080
        for name, default in boxes.items():
            raw = layout.get(name)
            try:
                raw_box = [round(n * scale) for n in default] if raw is None else list(map(int, raw))
                if len(raw_box) != 4:
                    raise ValueError
                x, y, box_width, box_height = raw_box
                x0, y0, x1, y1 = x, y, x + box_width, y + box_height
            except (TypeError, ValueError):
                errors.append(f"layout.{name} must be [x, y, width, height]")
                continue
            if x0 < 0 or y0 < 0 or x1 > width or y1 > height or x1 <= x0 or y1 <= y0:
                errors.append(f"layout.{name} must be a non-empty box inside the canvas")
        try:
            char = layout.get("character", [round(n * scale) for n in boxes["character"]])
            sub = layout.get("subtitle", [round(n * scale) for n in boxes["subtitle"]])
            cx, cy, cw, ch = map(int, char)
            sx, sy, sw, sh = map(int, sub)
            if sx < cx + cw and sx + sw > cx and sy < cy + ch and sy + sh > cy:
                errors.append("layout.subtitle overlaps layout.character")
        except (TypeError, ValueError):
            pass
    return errors


def action_gif(actions_dir: Path, action: str) -> Path:
    manifest = json.loads((actions_dir / "manifest.json").read_text(encoding="utf-8"))
    entry = manifest.get("actions", {}).get(action)
    candidate = actions_dir / (entry.get("preview") if entry else f"{action}.gif")
    if action not in ACTIONS or not candidate.is_file():
        raise FileNotFoundError(f"Animation GIF not found for action {action}: {candidate}")
    return candidate


def get_font(path: Path | None, size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = [path] if path else []
    candidates += [
        Path("C:/Windows/Fonts/meiryo.ttc"), Path("C:/Windows/Fonts/YuGothR.ttc"),
        Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"),
        Path("/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"),
    ]
    for candidate in candidates:
        if candidate and candidate.is_file():
            return ImageFont.truetype(str(candidate), size=size)
    return ImageFont.load_default()


def wrap_to_width(draw: ImageDraw.ImageDraw, value: str, font: ImageFont.ImageFont, width: int) -> list[str]:
    lines: list[str] = []
    for paragraph in value.splitlines() or [value]:
        current = ""
        for char in paragraph:
            if current and draw.textlength(current + char, font=font) > width:
                lines.append(current)
                current = char
            else:
                current += char
        lines.append(current)
    return lines


def draw_block(draw: ImageDraw.ImageDraw, value: str, box: list[int], font: ImageFont.ImageFont,
               color: tuple[int, int, int], centered: bool = False) -> None:
    x0, y0, x1, y1 = box
    lines = wrap_to_width(draw, value, font, x1 - x0)
    line_height = max(font.getbbox("国Ag")[3] - font.getbbox("国Ag")[1] + 5, 1)
    max_lines = max(1, (y1 - y0) // line_height)
    lines = lines[:max_lines]
    y = y0 + max(0, (y1 - y0 - len(lines) * line_height) // 2) if centered else y0
    for line in lines:
        x = x0 + max(0, (x1 - x0 - draw.textlength(line, font=font)) / 2) if centered else x0
        draw.text((x, y), line, fill=color, font=font, stroke_width=2, stroke_fill=(10, 10, 15))
        y += line_height


def get_box(layout: dict[str, Any], name: str, default: list[int], width: int, height: int) -> list[int]:
    value = layout.get(name)
    if value is not None:
        return list(map(int, value))
    scale = height / 1080
    return [round(item * scale) for item in default]


def slide_to_image(path: Path, size: tuple[int, int]) -> Image.Image:
    if path.suffix.lower() == ".svg":
        try:
            import cairosvg
        except ImportError as exc:
            raise RuntimeError("SVG rendering needs cairosvg. Install requirements.txt") from exc
        import io
        png = cairosvg.svg2png(url=str(path))
        source = Image.open(io.BytesIO(png)).convert("RGBA")
    else:
        source = Image.open(path).convert("RGBA")
    source.thumbnail(size, Image.Resampling.LANCZOS)
    canvas = Image.new("RGBA", size, (0, 0, 0, 0))
    canvas.alpha_composite(source, ((size[0] - source.width) // 2, (size[1] - source.height) // 2))
    return canvas


def render_frame(data: dict[str, Any], base: Path, slide: dict[str, Any], utterance: str,
                 target: Path) -> Path:
    canvas = data.get("canvas", {})
    width, height = int(canvas.get("width", 1920)), int(canvas.get("height", 1080))
    layout = data.get("layout", {})
    bg = resolve(base, data["assets"]["background"])
    assert bg is not None
    frame = Image.open(bg).convert("RGBA").resize((width, height), Image.Resampling.LANCZOS)
    draw = ImageDraw.Draw(frame)
    slide_box = get_box(layout, "slide", [40, 28, 1280, 720], width, height)
    sx, sy, sw, sh = map(int, slide_box)
    slide_image = slide_to_image(resolve(base, slide["image"]), (sw, sh))  # type: ignore[arg-type]
    frame.alpha_composite(slide_image, (sx, sy))
    subtitle_raw = get_box(layout, "subtitle", [330, 847, 1520, 163], width, height)
    subtitle = [subtitle_raw[0], subtitle_raw[1], subtitle_raw[0] + subtitle_raw[2], subtitle_raw[1] + subtitle_raw[3]]
    top_raw = get_box(layout, "notes_top", [1413, 66, 444, 324], width, height)
    notes_top = [top_raw[0], top_raw[1], top_raw[0] + top_raw[2], top_raw[1] + top_raw[3]]
    bottom_raw = get_box(layout, "notes_bottom", [1413, 410, 444, 310], width, height)
    notes_bottom = [bottom_raw[0], bottom_raw[1], bottom_raw[0] + bottom_raw[2], bottom_raw[1] + bottom_raw[3]]
    assets = data.get("assets", {})
    font_path = resolve(base, assets.get("font"))
    scale = height / 1080
    script_font = get_font(font_path, max(1, round(int(layout.get("subtitle_font_size", 54)) * scale)))
    note_font = get_font(font_path, max(1, round(int(layout.get("note_font_size", 35)) * scale)))
    draw_block(draw, utterance, subtitle, script_font, tuple(layout.get("subtitle_color", [255, 255, 255])), True)
    draw_block(draw, str(slide.get("note_top", "")), notes_top, note_font,
               tuple(layout.get("note_color", [238, 244, 255])))
    draw_block(draw, str(slide.get("note_bottom", "")), notes_bottom, note_font,
               tuple(layout.get("note_color", [238, 244, 255])))
    target.parent.mkdir(parents=True, exist_ok=True)
    frame.convert("RGB").save(target, "PNG", optimize=True)
    return target


def synthesize(text: str, dest: Path, voice: dict[str, Any], engine_url: str) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.is_file() and dest.stat().st_size:
        return
    style_id = int(voice.get("style_id", DEFAULT_VOICE["style_id"]))
    response = requests.post(
        f"{engine_url.rstrip('/')}/audio_query", params={"text": text, "speaker": style_id}, timeout=45
    )
    response.raise_for_status()
    query = response.json()
    response = requests.post(
        f"{engine_url.rstrip('/')}/synthesis", params={"speaker": style_id}, json=query, timeout=180
    )
    response.raise_for_status()
    dest.write_bytes(response.content)


def ffmpeg(command: list[str]) -> None:
    result = subprocess.run(command, text=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    if result.returncode:
        raise RuntimeError(f"ffmpeg exited with {result.returncode}:\n{result.stdout[-6000:]}")


def build_video(data: dict[str, Any], base: Path) -> Path:
    errors = validate_project(data, base)
    if errors:
        raise ValueError("Project validation failed:\n- " + "\n- ".join(errors))
    output = resolve(base, data.get("output", "output/final.mp4"))
    assert output is not None
    work = output.parent / (output.stem + "_work")
    frame_dir, audio_dir, segment_dir = work / "frames", work / "audio", work / "segments"
    for directory in (frame_dir, audio_dir, segment_dir):
        directory.mkdir(parents=True, exist_ok=True)
    fps = int(data.get("fps", 30))
    voice = data.get("voice", {})
    engine_url = str(voice.get("engine_url", "http://127.0.0.1:10101"))
    assets = data["assets"]
    animations = resolve(base, assets.get("animations", "assets/animations"))
    assert animations is not None
    entries: list[Path] = []
    sequence = 0
    for slide_index, slide in enumerate(data["slides"], 1):
        sentences = split_script(str(slide.get("script", "")))
        motions = slide.get("motions", []) or []
        tts_texts = slide.get("tts_texts", []) or []
        for chunk_index, sentence in enumerate(sentences, 1):
            action = motions[chunk_index - 1] if chunk_index <= len(motions) else "idle"
            tts_text = tts_texts[chunk_index - 1] if chunk_index <= len(tts_texts) and tts_texts[chunk_index - 1] else sentence
            frame_path = frame_dir / f"{slide_index:03d}_{chunk_index:03d}.png"
            voice_key = f"{tts_text}\0{voice.get('speaker_uuid', '')}\0{voice.get('style_id', DEFAULT_VOICE['style_id'])}\0{engine_url}"
            cache_key = hashlib.sha256(voice_key.encode("utf-8")).hexdigest()[:12]
            audio_path = audio_dir / f"{slide_index:03d}_{chunk_index:03d}_{cache_key}.wav"
            segment_path = segment_dir / f"{slide_index:03d}_{chunk_index:03d}.mp4"
            render_frame(data, base, slide, sentence, frame_path)
            synthesize(tts_text, audio_path, voice, engine_url)
            action_file = action_gif(animations, action)
            canvas = data.get("canvas", {})
            width, height = int(canvas.get("width", 1920)), int(canvas.get("height", 1080))
            char = get_box(data.get("layout", {}), "character", [35, 795, 250, 250], width, height)
            ffmpeg([
                "ffmpeg", "-y", "-loop", "1", "-framerate", str(fps), "-i", str(frame_path),
                "-stream_loop", "-1", "-i", str(action_file), "-i", str(audio_path),
                "-filter_complex", f"[1:v]scale={char[2]}:{char[3]}:force_original_aspect_ratio=decrease[char];"
                f"[0:v][char]overlay={char[0]}:{char[1]}:eof_action=repeat[v]",
                "-map", "[v]", "-map", "2:a:0", "-c:v", "libx264", "-preset", "medium",
                "-crf", "18", "-r", str(fps), "-vsync", "cfr", "-pix_fmt", "yuv420p",
                "-c:a", "aac", "-b:a", "192k", "-ar", "48000", "-shortest",
                "-movflags", "+faststart", str(segment_path),
            ])
            sequence += 1
            entries.append(segment_path)
            print(f"[{sequence}] slide={slide.get('id', slide_index)} motion={action}: {sentence}")
    concat_file = work / "concat.txt"
    concat_lines = []
    for path in entries:
        escaped = path.resolve().as_posix().replace("'", "'\\''")
        concat_lines.append(f"file '{escaped}'\n")
    concat_file.write_text("".join(concat_lines), encoding="utf-8")
    silent_bgm = output.with_name(output.stem + "_narration.mp4")
    output.parent.mkdir(parents=True, exist_ok=True)
    ffmpeg(["ffmpeg", "-y", "-f", "concat", "-safe", "0", "-i", str(concat_file),
            "-c:v", "libx264", "-preset", "medium", "-crf", "18", "-pix_fmt", "yuv420p",
            "-r", str(fps), "-c:a", "aac", "-b:a", "192k", "-ar", "48000", "-movflags", "+faststart",
            str(silent_bgm)])
    bgm = resolve(base, assets.get("bgm"))
    if bgm and bgm.is_file():
        volume = float(data.get("audio", {}).get("bgm_volume", 0.2))
        ffmpeg(["ffmpeg", "-y", "-i", str(silent_bgm), "-stream_loop", "-1", "-i", str(bgm),
                "-filter_complex", f"[1:a]volume={volume}[bgm];[0:a][bgm]amix=inputs=2:duration=first:dropout_transition=2[a]",
                "-map", "0:v:0", "-map", "[a]", "-c:v", "copy", "-c:a", "aac", "-b:a", "192k",
                "-ar", "48000", "-movflags", "+faststart", "-shortest", str(output)])
    else:
        shutil.copy2(silent_bgm, output)
    print(f"Wrote {output}")
    return output


def init_project(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    assets = path / "assets"
    assets.mkdir(exist_ok=True)
    shutil.copy2(ROOT / "biimslide_1920x1080.png", assets / "frame.png")
    shutil.copytree(ROOT / "animations", assets / "animations", dirs_exist_ok=True)
    (path / "slides").mkdir(exist_ok=True)
    sample_svg = '''<svg xmlns="http://www.w3.org/2000/svg" width="1280" height="720" viewBox="0 0 1280 720"><rect width="1280" height="720" fill="#101827"/><text x="640" y="300" text-anchor="middle" font-family="sans-serif" font-size="64" fill="white">タイトル</text><text x="640" y="390" text-anchor="middle" font-family="sans-serif" font-size="32" fill="#b9d7ff">内容に合わせてこのSVGを編集</text></svg>'''
    (path / "slides" / "001.svg").write_text(sample_svg, encoding="utf-8")
    project = {
        "version": 1, "title": path.name, "canvas": {"width": 1920, "height": 1080}, "fps": 30,
        "output": "output/final.mp4",
        "assets": {"background": "assets/frame.png", "animations": "assets/animations", "bgm": ""},
        "voice": {**DEFAULT_VOICE, "engine_url": "http://127.0.0.1:10101"},
        "layout": {"slide": [40, 28, 1280, 720], "subtitle": [330, 847, 1520, 163],
                   "notes_top": [1413, 66, 444, 324], "notes_bottom": [1413, 410, 444, 310],
                   "subtitle_font_size": 54, "note_font_size": 35},
        "slides": [{"id": 1, "image": "slides/001.svg", "script": "こんにちは。ここにナレーションを書きます。",
                    "motions": ["wave", "idle"],
                    "note_top": "ポイント", "note_bottom": "補足説明"}],
    }
    target = path / "project.yaml"
    target.write_text(yaml.safe_dump(project, allow_unicode=True, sort_keys=False), encoding="utf-8")
    return target


def main() -> int:
    parser = argparse.ArgumentParser(description="Create and render Agent-authored 16:9 Biim-style video projects")
    sub = parser.add_subparsers(dest="command", required=True)
    create = sub.add_parser("init", help="create a project with editable SVG slide and bundled assets")
    create.add_argument("directory", type=Path)
    check = sub.add_parser("validate", help="validate project schema, assets, slides, and motions")
    check.add_argument("project", type=Path)
    build = sub.add_parser("build", help="synthesize narration and export the final MP4")
    build.add_argument("project", type=Path)
    args = parser.parse_args()
    try:
        if args.command == "init":
            print(init_project(args.directory.resolve()))
            return 0
        data, base = load_project(args.project)
        problems = validate_project(data, base)
        if problems:
            print("\n".join(f"ERROR: {item}" for item in problems), file=sys.stderr)
            return 2
        if args.command == "validate":
            print(f"Valid project: {base / 'project.yaml'} ({len(data['slides'])} slides)")
            return 0
        build_video(data, base)
        return 0
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
