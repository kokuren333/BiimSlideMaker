#!/usr/bin/env python3
"""
MovieMaker GUI tool
===================

This tkinter-based utility automates the workflow described in 仕様書.md:

1. Convert PDF slides into 1280x720 PNG images that stay aligned with YAML ids.
2. Split YAML scripts by "。" and synthesize narration with AivisSpeech Engine.
3. Composite slides, notes, and subtitles on top of a 1920x1080 template and
   stitch everything into an mp4 video via ffmpeg while adding low-volume BGM.

Prerequisites
-------------
- Python 3.9+
- pip install pymupdf pillow pyyaml requests
- ffmpeg available on PATH (or specify full path in the GUI)
- AivisSpeech Engine running locally (default http://127.0.0.1:10101)

Fonts, background, and BGM default to the paths listed in the specification but
can be overridden from the GUI.
"""

from __future__ import annotations

import concurrent.futures
import io
import json
import os
import queue
import re
import shutil
import subprocess
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

try:
    import fitz  # type: ignore
except ImportError as exc:  # pragma: no cover
    raise SystemExit(
        "PyMuPDF (pymupdf) is required. Install it with: pip install pymupdf"
    ) from exc

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError as exc:  # pragma: no cover
    raise SystemExit("Pillow is required. Install it with: pip install pillow") from exc

try:
    import requests
except ImportError as exc:  # pragma: no cover
    raise SystemExit(
        "The requests package is required. Install it with: pip install requests"
    ) from exc

try:
    import yaml
except ImportError as exc:  # pragma: no cover
    raise SystemExit("PyYAML is required. Install it with: pip install pyyaml") from exc

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk


CANVAS_SIZE = (1920, 1080)
SLIDE_SIZE = (1280, 720)
SLIDE_TOP_LEFT = (40, 28)
SLIDE_CORNER_RADIUS = 32
SCRIPT_BOX = (88, 847, 1850, 1009)
NOTE_BOX = (1413, 66, 1857, 721)
SCRIPT_COLOR = (255, 255, 255)
NOTE_COLOR = (238, 244, 255)

DEFAULT_BACKGROUND = Path(r"biimslide_1920x1080.png")
DEFAULT_SCRIPT_FONT = Path(
    r"下側字幕のフォントパス"
)
DEFAULT_NOTE_FONT = Path(
    r"右側ノートのフォントパス"
)
DEFAULT_BGM = Path(r"(Glass Weather).mp3")


def ensure_directory(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def read_yaml_any(path: Path) -> Dict:
    """Try multiple encodings so cp932 YAML files also work."""
    for encoding in ("utf-8", "utf-8-sig", "cp932", "shift_jis"):
        try:
            text = path.read_text(encoding=encoding)
            return yaml.safe_load(text) or {}
        except UnicodeDecodeError:
            continue
    raise ValueError(f"YAMLの読み込みに失敗しました: {path}")


def split_script(text: str) -> List[str]:
    """Split script by the Japanese period '。', preserving the delimiter."""
    cleaned = (text or "").replace("\r", "")
    parts: List[str] = []
    for chunk in cleaned.split("。"):
        chunk = chunk.strip()
        if not chunk:
            continue
        suffix = "。" if not chunk.endswith(("。", "！", "？", "!", ".")) else ""
        parts.append(chunk + suffix)
    return parts or [cleaned.strip()]


def run_ffmpeg(command: Sequence[str]) -> None:
    proc = subprocess.run(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        check=False,
    )
    if proc.returncode != 0:
        raise RuntimeError(
            f"ffmpeg失敗 (exit {proc.returncode}):\n{proc.stdout.strip()}"
        )


def wrap_text_lines(
    text: str,
    font: ImageFont.FreeTypeFont,
    max_width: int,
    draw: ImageDraw.ImageDraw,
) -> List[str]:
    """Greedy wrapping that also works for Japanese text."""
    if not text:
        return []

    lines: List[str] = []
    for paragraph in re.split(r"\n+", text.strip()):
        if not paragraph:
            continue
        current = ""
        for char in paragraph:
            attempt = current + char
            width = draw.textlength(attempt, font=font)
            if width <= max_width or not current:
                current = attempt
            else:
                lines.append(current)
                current = char
        if current:
            lines.append(current)
    return lines or [""]


def fit_text_to_box(
    text: str,
    font_path: Path,
    box: Tuple[int, int, int, int],
    max_size: int,
    min_size: int,
) -> Tuple[ImageFont.FreeTypeFont, List[str], int, int]:
    """Return (font, lines, line_height, line_spacing)."""
    temp_img = Image.new("RGB", (10, 10))
    draw = ImageDraw.Draw(temp_img)
    box_width = box[2] - box[0]
    box_height = box[3] - box[1]

    fallback: Optional[Tuple[ImageFont.FreeTypeFont, List[str], int, int]] = None

    for size in range(max_size, min_size - 1, -2):
        font = ImageFont.truetype(str(font_path), size)
        lines = wrap_text_lines(text, font, box_width, draw)
        ascent, descent = font.getmetrics()
        line_height = ascent + descent
        line_spacing = max(4, int(size * 0.2))
        total_height = line_height * len(lines) + line_spacing * max(0, len(lines) - 1)
        fallback = (font, lines, line_height, line_spacing)
        if total_height <= box_height:
            return fallback

    if not fallback:
        raise RuntimeError("フォントの計測に失敗しました。")
    return fallback


def draw_text_block(
    draw: ImageDraw.ImageDraw,
    lines: List[str],
    font: ImageFont.FreeTypeFont,
    box: Tuple[int, int, int, int],
    color: Tuple[int, int, int],
    align: str,
    line_height: int,
    line_spacing: int,
) -> None:
    if not lines:
        return

    x0, y0, x1, y1 = box
    box_width = x1 - x0
    box_height = y1 - y0
    total_height = line_height * len(lines) + line_spacing * max(0, len(lines) - 1)
    if align == "center":
        y = y0 + (box_height - total_height) / 2
    else:
        y = y0

    for line in lines:
        if align == "center":
            text_width = draw.textlength(line, font=font)
            x = x0 + (box_width - text_width) / 2
        else:
            x = x0
        draw.text((x, y), line, fill=color, font=font)
        y += line_height + line_spacing


def compose_frame(
    slide_path: Path,
    background_path: Path,
    script_font: Path,
    note_font: Path,
    script_text: str,
    note_text: str,
    dest_path: Path,
) -> None:
    base = Image.open(background_path).convert("RGBA")
    slide = Image.open(slide_path).convert("RGBA")
    if slide.size != SLIDE_SIZE:
        slide = slide.resize(SLIDE_SIZE, Image.LANCZOS)
    mask = Image.new("L", SLIDE_SIZE, 0)
    mask_draw = ImageDraw.Draw(mask)
    mask_draw.rounded_rectangle(
        (0, 0, SLIDE_SIZE[0], SLIDE_SIZE[1]),
        radius=SLIDE_CORNER_RADIUS,
        fill=255,
    )
    slide.putalpha(mask)
    base.paste(slide, SLIDE_TOP_LEFT, slide)
    draw = ImageDraw.Draw(base)

    if script_text.strip():
        font, lines, line_height, line_spacing = fit_text_to_box(
            script_text.strip(), script_font, SCRIPT_BOX, max_size=70, min_size=28
        )
        draw_text_block(
            draw,
            lines,
            font,
            SCRIPT_BOX,
            SCRIPT_COLOR,
            "center",
            line_height,
            line_spacing,
        )

    if note_text.strip():
        font, lines, line_height, line_spacing = fit_text_to_box(
            note_text.strip(), note_font, NOTE_BOX, max_size=46, min_size=18
        )
        draw_text_block(
            draw,
            lines,
            font,
            NOTE_BOX,
            NOTE_COLOR,
            "left",
            line_height,
            line_spacing,
        )

    ensure_directory(dest_path.parent)
    base.convert("RGB").save(dest_path, "PNG", optimize=True)


class AivisSpeechClient:
    def __init__(self, base_url: str) -> None:
        self.base_url = base_url.rstrip("/")
        self.session = requests.Session()

    def list_speakers(self) -> List[Dict]:
        resp = self.session.get(f"{self.base_url}/speakers", timeout=15)
        resp.raise_for_status()
        return resp.json()

    def initialize_speaker(self, speaker_id: str, skip_reinit: bool = False) -> None:
        resp = self.session.post(
            f"{self.base_url}/initialize_speaker",
            params={"speaker": speaker_id, "skip_reinit": str(skip_reinit).lower()},
            timeout=60,
        )
        resp.raise_for_status()

    def synthesize(self, speaker_id: str, text: str, dest: Path) -> None:
        query_resp = self.session.post(
            f"{self.base_url}/audio_query",
            params={"text": text, "speaker": speaker_id},
            timeout=30,
        )
        query_resp.raise_for_status()
        query = query_resp.json()
        synth_resp = self.session.post(
            f"{self.base_url}/synthesis",
            params={"speaker": speaker_id, "enable_interrogative_upspeak": "true"},
            json=query,
            timeout=120,
        )
        synth_resp.raise_for_status()
        ensure_directory(dest.parent)
        dest.write_bytes(synth_resp.content)


@dataclass
class Segment:
    slide_id: int
    chunk_index: int
    sequence: int
    script_text: str
    note_bottom: str
    slide_image: str
    audio_path: str

    @property
    def chunk_name(self) -> str:
        return f"{self.slide_id:03d}_{self.chunk_index:02d}"


class UILogger:
    def __init__(self, widget: scrolledtext.ScrolledText) -> None:
        self.widget = widget
        self.queue: "queue.Queue[str]" = queue.Queue()
        self.widget.after(80, self.flush)

    def log(self, message: str) -> None:
        timestamp = time.strftime("[%H:%M:%S]")
        self.queue.put(f"{timestamp} {message}\n")

    def flush(self) -> None:
        try:
            self.widget.configure(state="normal")
            while not self.queue.empty():
                msg = self.queue.get()
                self.widget.insert(tk.END, msg)
                self.widget.see(tk.END)
        finally:
            self.widget.configure(state="disabled")
            self.widget.after(80, self.flush)


class MovieMakerApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("MovieMaker GUI (AivisSpeech)")
        self.root.geometry("1150x780")
        self._manifest_cache: List[Segment] = []

        self._init_vars()
        self._build_layout()

    def _init_vars(self) -> None:
        cwd = Path.cwd()
        self.pdf_var = tk.StringVar()
        self.yaml_var = tk.StringVar(value=str((cwd / "test.yaml").resolve()))
        self.slide_dir_var = tk.StringVar(value=str((cwd / "slides").resolve()))
        self.audio_dir_var = tk.StringVar(value=str((cwd / "audio").resolve()))
        self.frame_dir_var = tk.StringVar(value=str((cwd / "frames").resolve()))
        self.segment_dir_var = tk.StringVar(value=str((cwd / "segments").resolve()))
        self.output_video_var = tk.StringVar(value=str((cwd / "final.mp4").resolve()))
        self.manifest_path_var = tk.StringVar(
            value=str((cwd / "movie_manifest.json").resolve())
        )
        self.concat_list_var = tk.StringVar(
            value=str((cwd / "concat_list.txt").resolve())
        )

        self.background_var = tk.StringVar(value=str(DEFAULT_BACKGROUND))
        self.script_font_var = tk.StringVar(value=str(DEFAULT_SCRIPT_FONT))
        self.note_font_var = tk.StringVar(value=str(DEFAULT_NOTE_FONT))
        self.bgm_var = tk.StringVar(value=str(DEFAULT_BGM))
        self.ffmpeg_var = tk.StringVar(value="ffmpeg")
        self.aivis_url_var = tk.StringVar(value="http://127.0.0.1:10101")
        self.speaker_id_var = tk.StringVar(value="888753760")
        self.worker_var = tk.IntVar(value=max(2, os.cpu_count() or 4))
        self.prewarm_var = tk.BooleanVar(value=True)

        self.slide_progress = tk.DoubleVar(value=0.0)
        self.audio_progress = tk.DoubleVar(value=0.0)
        self.video_progress = tk.DoubleVar(value=0.0)

    def _build_layout(self) -> None:
        container = ttk.Frame(self.root, padding=12)
        container.pack(fill=tk.BOTH, expand=True)

        inputs = ttk.LabelFrame(container, text="入力ファイル")
        inputs.pack(fill=tk.X, expand=False, pady=(0, 10))
        self._path_row(inputs, "PDF:", self.pdf_var, self._browse_pdf, 0)
        self._path_row(inputs, "YAML:", self.yaml_var, self._browse_yaml, 1)

        dirs = ttk.LabelFrame(container, text="出力ディレクトリ")
        dirs.pack(fill=tk.X, pady=(0, 10))
        self._path_row(
            dirs, "スライドPNG:", self.slide_dir_var, lambda: self._browse_dir(self.slide_dir_var), 0
        )
        self._path_row(
            dirs, "音声WAV:", self.audio_dir_var, lambda: self._browse_dir(self.audio_dir_var), 1
        )
        self._path_row(
            dirs, "字幕フレーム:", self.frame_dir_var, lambda: self._browse_dir(self.frame_dir_var), 2
        )
        self._path_row(
            dirs, "チャンク動画:", self.segment_dir_var, lambda: self._browse_dir(self.segment_dir_var), 3
        )
        self._path_row(
            dirs,
            "マニフェスト:",
            self.manifest_path_var,
            lambda: self._choose_file(self.manifest_path_var, save=True),
            4,
            button_label="保存先",
        )
        self._path_row(
            dirs,
            "最終MP4:",
            self.output_video_var,
            lambda: self._choose_file(self.output_video_var, save=True),
            5,
            button_label="保存先",
        )

        aivis = ttk.LabelFrame(container, text="AivisSpeech 設定")
        aivis.pack(fill=tk.X, pady=(0, 10))
        self._path_row(aivis, "Engine URL:", self.aivis_url_var, None, 0, browse=False)
        self._path_row(
            aivis,
            "Speaker ID:",
            self.speaker_id_var,
            self._fetch_speakers,
            1,
            button_label="話者一覧",
        )
        worker_row = ttk.Frame(aivis)
        worker_row.grid(row=2, column=0, columnspan=3, sticky="we", pady=2)
        ttk.Label(worker_row, text="並列ワーカー数:").pack(side=tk.LEFT)
        ttk.Spinbox(
            worker_row,
            from_=1,
            to=max(32, (os.cpu_count() or 4) * 2),
            textvariable=self.worker_var,
            width=6,
        ).pack(side=tk.LEFT, padx=(5, 20))
        ttk.Checkbutton(
            worker_row, text="事前に /initialize_speaker を叩く", variable=self.prewarm_var
        ).pack(side=tk.LEFT)

        video = ttk.LabelFrame(container, text="合成設定")
        video.pack(fill=tk.X, pady=(0, 10))
        self._path_row(video, "背景PNG:", self.background_var, lambda: self._choose_file(self.background_var), 0)
        self._path_row(
            video, "字幕フォント:", self.script_font_var, lambda: self._choose_file(self.script_font_var), 1
        )
        self._path_row(
            video, "ノートフォント:", self.note_font_var, lambda: self._choose_file(self.note_font_var), 2
        )
        self._path_row(video, "BGM:", self.bgm_var, lambda: self._choose_file(self.bgm_var), 3)
        self._path_row(
            video, "ffmpeg実行ファイル:", self.ffmpeg_var, lambda: self._choose_file(self.ffmpeg_var), 4
        )

        actions = ttk.Frame(container)
        actions.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(
            actions, text="1. スライド生成", command=lambda: self._run_async(self._generate_slides)
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            actions, text="2. 音声合成", command=lambda: self._run_async(self._generate_audio)
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            actions, text="3. 動画出力", command=lambda: self._run_async(self._assemble_video)
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(actions, text="出力リフレッシュ", command=self._refresh_outputs).pack(
            side=tk.RIGHT, padx=5
        )
        ttk.Button(actions, text="一括実行", command=lambda: self._run_async(self._run_all)).pack(
            side=tk.RIGHT, padx=5
        )

        progress_frame = ttk.Frame(container)
        progress_frame.pack(fill=tk.X, pady=(0, 8))
        self._progress_row(progress_frame, "Slides", self.slide_progress, 0)
        self._progress_row(progress_frame, "Audio", self.audio_progress, 1)
        self._progress_row(progress_frame, "Video", self.video_progress, 2)

        log_frame = ttk.LabelFrame(container, text="ログ")
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_widget = scrolledtext.ScrolledText(
            log_frame, height=12, state="disabled", font=("Consolas", 10)
        )
        self.log_widget.pack(fill=tk.BOTH, expand=True)
        self.logger = UILogger(self.log_widget)

    def _path_row(
        self,
        parent: ttk.LabelFrame,
        label: str,
        var: tk.StringVar,
        callback,
        row: int,
        button_label: str = "参照",
        browse: bool = True,
    ) -> None:
        ttk.Label(parent, text=label, width=16).grid(
            row=row, column=0, sticky="w", padx=(4, 2), pady=2
        )
        entry = ttk.Entry(parent, textvariable=var)
        entry.grid(row=row, column=1, sticky="we", pady=2)
        parent.columnconfigure(1, weight=1)
        if browse and callback:
            ttk.Button(parent, text=button_label, command=callback, width=10).grid(
                row=row, column=2, padx=(6, 2), pady=2
            )
        elif browse:
            ttk.Button(parent, text=button_label, command=lambda: self._choose_file(var)).grid(
                row=row, column=2, padx=(6, 2), pady=2
            )

    def _progress_row(self, parent: ttk.Frame, label: str, var: tk.DoubleVar, row: int) -> None:
        ttk.Label(parent, text=label, width=10).grid(row=row, column=0, sticky="w")
        bar = ttk.Progressbar(parent, maximum=1.0, variable=var)
        bar.grid(row=row, column=1, sticky="we", padx=6, pady=2)
        parent.columnconfigure(1, weight=1)

    # ---------- file dialogs ----------
    def _choose_file(self, var: tk.StringVar, save: bool = False) -> None:
        path = (
            filedialog.asksaveasfilename(title="保存先を選択")
            if save
            else filedialog.askopenfilename(title="ファイルを選択")
        )
        if path:
            var.set(path)

    def _browse_pdf(self) -> None:
        path = filedialog.askopenfilename(
            title="PDFを選択", filetypes=[("PDF", "*.pdf"), ("All files", "*.*")]
        )
        if path:
            self.pdf_var.set(path)

    def _browse_yaml(self) -> None:
        path = filedialog.askopenfilename(
            title="YAMLを選択", filetypes=[("YAML", "*.yaml;*.yml"), ("All files", "*.*")]
        )
        if path:
            self.yaml_var.set(path)

    def _browse_dir(self, var: tk.StringVar) -> None:
        path = filedialog.askdirectory(title="フォルダを選択")
        if path:
            var.set(path)

    # ---------- background execution ----------
    def _run_async(self, func) -> None:
        threading.Thread(target=lambda: self._safe_run(func), daemon=True).start()

    def _safe_run(self, func) -> None:
        try:
            func()
        except Exception as exc:  # pragma: no cover
            self.logger.log(f"ERROR: {exc}")
            self.root.after(0, lambda: messagebox.showerror("エラー", str(exc)))

    def log(self, message: str) -> None:
        self.logger.log(message)

    def _set_progress(self, var: tk.DoubleVar, value: float) -> None:
        self.root.after(0, lambda: var.set(min(max(value, 0.0), 1.0)))

    # ---------- pipeline orchestrations ----------
    def _refresh_outputs(self) -> None:
        targets = [
            ("スライドPNG", Path(self.slide_dir_var.get())),
            ("音声WAV", Path(self.audio_dir_var.get())),
            ("字幕フレーム", Path(self.frame_dir_var.get())),
            ("チャンク動画", Path(self.segment_dir_var.get())),
        ]
        file_targets = [
            ("マニフェスト", Path(self.manifest_path_var.get())),
            ("concatリスト", Path(self.concat_list_var.get())),
            ("最終MP4", Path(self.output_video_var.get())),
            (
                "ナレーションMP4",
                Path(self.output_video_var.get()).with_name(
                    Path(self.output_video_var.get()).stem + "_narration.mp4"
                ),
            ),
        ]

        if not messagebox.askyesno(
            "確認", "選択中の出力ディレクトリと関連ファイルを削除します。よろしいですか？"
        ):
            return

        errors = []
        for label, path in targets:
            try:
                if path.exists():
                    shutil.rmtree(path)
                    self.log(f"{label} を削除しました: {path}")
            except Exception as exc:
                errors.append(f"{label}: {exc}")

        for label, path in file_targets:
            try:
                if path.exists():
                    path.unlink()
                    self.log(f"{label} を削除しました: {path}")
            except Exception as exc:
                errors.append(f"{label}: {exc}")

        if errors:
            messagebox.showerror("削除エラー", "\n".join(errors))
        else:
            messagebox.showinfo("完了", "出力ディレクトリをリフレッシュしました。")
            self.log("出力ディレクトリをリフレッシュしました。")

    def _run_all(self) -> None:
        self._set_progress(self.slide_progress, 0)
        self._set_progress(self.audio_progress, 0)
        self._set_progress(self.video_progress, 0)
        self.log("=== 全工程を開始します ===")
        self._generate_slides()
        self._generate_audio()
        self._assemble_video()
        self.log("=== 全工程が完了しました ===")

    def _generate_slides(self) -> None:
        pdf_path = Path(self.pdf_var.get())
        if not pdf_path.is_file():
            raise FileNotFoundError(f"PDFが見つかりません: {pdf_path}")
        slides_dir = ensure_directory(Path(self.slide_dir_var.get()))

        yaml_data = []
        yaml_path = Path(self.yaml_var.get())
        if yaml_path.is_file():
            try:
                loaded = read_yaml_any(yaml_path)
                yaml_data = loaded.get("slides", loaded) or []
            except Exception as exc:
                self.log(f"YAMLの読み込みをスキップしました: {exc}")

        doc = fitz.open(pdf_path)
        total = len(doc)
        self.log(f"PDFを読み込みました: {pdf_path} ({total} ページ)")

        for index, page in enumerate(doc, start=1):
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            image = Image.open(io.BytesIO(pix.pil_tobytes("png")))
            image = image.convert("RGB").resize(SLIDE_SIZE, Image.LANCZOS)
            slide_id = (
                int(yaml_data[index - 1].get("id", index))
                if index - 1 < len(yaml_data)
                else index
            )
            filename = f"slide_{slide_id:03d}.png"
            dest = slides_dir / filename
            image.save(dest, "PNG", optimize=True)
            self.log(f"作成: {dest.name}")
            self._set_progress(self.slide_progress, index / total)

        self.log("スライドPNG生成が完了しました。")
        self._set_progress(self.slide_progress, 1.0)

    def _generate_audio(self) -> None:
        yaml_path = Path(self.yaml_var.get())
        if not yaml_path.is_file():
            raise FileNotFoundError(f"YAMLが見つかりません: {yaml_path}")
        slides_dir = Path(self.slide_dir_var.get())
        if not slides_dir.exists():
            raise FileNotFoundError("スライドPNGが存在しません。先にステップ1を実行してください。")
        audio_dir = ensure_directory(Path(self.audio_dir_var.get()))

        data = read_yaml_any(yaml_path)
        slides = data.get("slides", [])
        if not isinstance(slides, list) or not slides:
            raise ValueError("YAML内に slides 配列が見つかりません。")

        slide_images = {
            int(match.group(1)): path
            for path in slides_dir.glob("slide_*.png")
            if (match := re.search(r"(\d+)", path.stem))
        }
        if not slide_images:
            raise RuntimeError("スライドPNGが見つかりません。")

        segments: List[Segment] = []
        sequence = 1
        for slide in slides:
            slide_id = int(slide.get("id", sequence))
            script = slide.get("script", "")
            note = slide.get("note_bottom", "")
            sentences = split_script(script)
            for idx, sentence in enumerate(sentences, start=1):
                audio_path = audio_dir / f"chunk_{slide_id:03d}_{idx:02d}.wav"
                segment = Segment(
                    slide_id=slide_id,
                    chunk_index=idx,
                    sequence=sequence,
                    script_text=sentence.strip(),
                    note_bottom=str(note or ""),
                    slide_image=str(slide_images.get(slide_id, "")),
                    audio_path=str(audio_path),
                )
                if not segment.slide_image:
                    raise FileNotFoundError(
                        f"スライドID {slide_id} に対応するPNGが見つかりません。"
                    )
                segments.append(segment)
                sequence += 1

        if not segments:
            raise ValueError("音声化する script が見つかりません。")

        client = AivisSpeechClient(self.aivis_url_var.get())
        speaker_id = self.speaker_id_var.get().strip()
        if self.prewarm_var.get():
            self.log(f"/initialize_speaker を実行中... (speaker={speaker_id})")
            client.initialize_speaker(speaker_id)

        total = len(segments)
        self.log(f"音声合成ジョブ: {total} 件")
        workers = max(1, int(self.worker_var.get()))
        completed = 0
        lock = threading.Lock()

        def synthesize(segment: Segment) -> None:
            if Path(segment.audio_path).is_file():
                return
            client.synthesize(speaker_id, segment.script_text, Path(segment.audio_path))

        with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as executor:
            future_map = {executor.submit(synthesize, seg): seg for seg in segments}
            for future in concurrent.futures.as_completed(future_map):
                future.result()
                with lock:
                    completed += 1
                    seg = future_map[future]
                    self._set_progress(self.audio_progress, completed / total)
                    self.log(
                        f"[{completed}/{total}] 音声完了: slide {seg.slide_id} chunk {seg.chunk_index}"
                    )

        manifest_path = Path(self.manifest_path_var.get())
        ensure_directory(manifest_path.parent)
        payload = {
            "generated_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
            "pdf": self.pdf_var.get(),
            "yaml": self.yaml_var.get(),
            "segments": [segment.__dict__ for segment in segments],
        }
        manifest_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        self._manifest_cache = segments
        self.log(f"音声合成が完了しました。マニフェストを保存: {manifest_path}")
        self._set_progress(self.audio_progress, 1.0)

    def _assemble_video(self) -> None:
        segments = self._manifest_cache or self._load_manifest_from_disk()
        if not segments:
            raise RuntimeError("マニフェストが読み込まれていません。先にステップ2を実行してください。")

        frame_dir = ensure_directory(Path(self.frame_dir_var.get()))
        segment_dir = ensure_directory(Path(self.segment_dir_var.get()))
        concat_list = Path(self.concat_list_var.get())
        background = Path(self.background_var.get())
        script_font = Path(self.script_font_var.get())
        note_font = Path(self.note_font_var.get())
        bgm_path = Path(self.bgm_var.get())
        ffmpeg_bin = self.ffmpeg_var.get().strip() or "ffmpeg"
        final_output = Path(self.output_video_var.get())
        narrator_only = final_output.with_name(final_output.stem + "_narration.mp4")

        total = len(segments)
        self.log(f"動画チャンク生成: {total} 本")
        for idx, segment in enumerate(segments, start=1):
            chunk = segment.chunk_name
            frame_path = frame_dir / f"{chunk}.png"
            video_path = segment_dir / f"{chunk}.mp4"
            compose_frame(
                slide_path=Path(segment.slide_image),
                background_path=background,
                script_font=script_font,
                note_font=note_font,
                script_text=segment.script_text,
                note_text=segment.note_bottom,
                dest_path=frame_path,
            )
            if not Path(segment.audio_path).is_file():
                raise FileNotFoundError(f"音声が見つかりません: {segment.audio_path}")
            run_ffmpeg(
                [
                    ffmpeg_bin,
                    "-y",
                    "-loop",
                    "1",
                    "-i",
                    str(frame_path),
                    "-i",
                    segment.audio_path,
                    "-c:v",
                    "libx264",
                    "-tune",
                    "stillimage",
                    "-preset",
                    "medium",
                    "-crf",
                    "18",
                    "-c:a",
                    "aac",
                    "-b:a",
                    "192k",
                    "-shortest",
                    "-pix_fmt",
                    "yuv420p",
                    str(video_path),
                ]
            )
            self.log(f"チャンク動画完成: {video_path.name}")
            self._set_progress(self.video_progress, idx / (total * 2))

        ensure_directory(concat_list.parent)
        with concat_list.open("w", encoding="utf-8") as fp:
            for segment in segments:
                path = segment_dir / f"{segment.chunk_name}.mp4"
                fp.write(f"file '{path.as_posix()}'\n")
        self.log(f"concatリストを出力: {concat_list}")

        run_ffmpeg(
            [
                ffmpeg_bin,
                "-y",
                "-f",
                "concat",
                "-safe",
                "0",
                "-i",
                str(concat_list),
                "-c",
                "copy",
                str(narrator_only),
            ]
        )
        self.log(f"ナレーション動画を作成: {narrator_only}")
        self._set_progress(self.video_progress, 0.7)

        if not bgm_path.is_file():
            raise FileNotFoundError(f"BGMファイルが見つかりません: {bgm_path}")

        run_ffmpeg(
            [
                ffmpeg_bin,
                "-y",
                "-i",
                str(narrator_only),
                "-stream_loop",
                "-1",
                "-i",
                str(bgm_path),
                "-filter_complex",
                "[1:a]volume=0.2[a_bgm];[0:a][a_bgm]amix=inputs=2:duration=first[a_mix]",
                "-map",
                "0:v",
                "-map",
                "[a_mix]",
                "-c:v",
                "copy",
                "-c:a",
                "aac",
                "-shortest",
                str(final_output),
            ]
        )
        self.log(f"最終動画を書き出しました: {final_output}")
        self._set_progress(self.video_progress, 1.0)

    def _load_manifest_from_disk(self) -> List[Segment]:
        manifest_path = Path(self.manifest_path_var.get())
        if not manifest_path.is_file():
            raise FileNotFoundError(f"マニフェストが見つかりません: {manifest_path}")
        data = json.loads(manifest_path.read_text(encoding="utf-8"))
        segments = [
            Segment(
                slide_id=int(seg["slide_id"]),
                chunk_index=int(seg["chunk_index"]),
                sequence=int(seg.get("sequence", idx + 1)),
                script_text=seg["script_text"],
                note_bottom=seg.get("note_bottom", ""),
                slide_image=seg["slide_image"],
                audio_path=seg["audio_path"],
            )
            for idx, seg in enumerate(data.get("segments", []))
        ]
        self._manifest_cache = segments
        self.log(f"マニフェストを再読み込みしました: {manifest_path}")
        return segments

    # ---------- speaker picker ----------
    def _fetch_speakers(self) -> None:
        def worker() -> None:
            try:
                client = AivisSpeechClient(self.aivis_url_var.get())
                speakers = client.list_speakers()
                self.root.after(0, lambda: self._show_speaker_picker(speakers))
            except Exception as exc:
                self.root.after(0, lambda: messagebox.showerror("取得失敗", str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def _show_speaker_picker(self, speakers: List[Dict]) -> None:
        win = tk.Toplevel(self.root)
        win.title("話者一覧 (/speakers)")
        tree = ttk.Treeview(win, columns=("name", "style", "id"), show="headings")
        tree.heading("name", text="話者")
        tree.heading("style", text="スタイル")
        tree.heading("id", text="ID")
        tree.column("name", width=180)
        tree.column("style", width=140)
        tree.column("id", width=120)
        tree.pack(fill=tk.BOTH, expand=True)

        for speaker in speakers:
            name = speaker.get("name")
            for style in speaker.get("styles", []):
                tree.insert(
                    "",
                    tk.END,
                    values=(name, style.get("name"), style.get("id")),
                )

        def on_pick(event) -> None:
            selected = tree.selection()
            if selected:
                values = tree.item(selected[0], "values")
                self.speaker_id_var.set(str(values[2]))
                win.destroy()

        tree.bind("<Double-1>", on_pick)


def main() -> None:
    app = MovieMakerApp()
    app.root.mainloop()


if __name__ == "__main__":
    main()
