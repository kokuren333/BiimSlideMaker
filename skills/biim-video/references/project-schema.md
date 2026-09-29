# Portable PowerShell project schema

The Agent-facing CLI reads `project.json`. Author each slide as a standalone HTML document or use PNG/JPEG/BMP/GIF artwork. HTML uses Microsoft Edge or Chrome in headless mode; the fonts and KaTeX distribution are bundled in each initialized project so slide rendering works offline.

## Example

```json
{
  "version": 1,
  "title": "Example explainer",
  "canvas": { "width": 1920, "height": 1080 },
  "fps": 30,
  "output": "output/final.mp4",
  "assets": {
    "background": "assets/frame.png",
    "animations": "assets/animations",
    "bgm": ""
  },
  "voice": {
    "name": "kokuren_3rd",
    "speaker_uuid": "38d7216c-e595-4d8f-b06c-1fc376e47c0a",
    "style_name": "ノーマル",
    "style_id": 1069147200,
    "engine_url": "http://127.0.0.1:10101"
  },
  "audio": { "bgm_volume": 0.2 },
  "layout": {
    "slide": [40, 28, 1280, 720],
    "subtitle": [385, 847, 1465, 163],
    "notes_top": [1413, 66, 444, 164],
    "notes_bottom": [1413, 260, 444, 460],
    "character": [35, 755, 300, 300],
    "subtitle_font_size": 54,
    "note_top_font_size": 38,
    "note_font_size": 32,
    "fonts": { "slide": "Noto Sans JP", "subtitle": "Noto Sans JP", "notes": "Noto Sans JP" }
  },
  "slides": [
    {
      "id": 1,
      "html": "slides/001.html",
      "script": "こんにちは。AivisSpeech APIを使います。",
      "motions": ["wave", "point"],
      "tts_texts": ["", "エイビススピーチ エーピーアイを使います。"],
      "note_top": "要点",
      "note_bottom": "用語や背景を簡潔に補足します。"
    }
  ]
}
```

`slides[].html` is the recommended format. The HTML viewport is 1280×720; author the page at that size and set `margin:0`, `overflow:hidden`, and explicit dimensions. Local images can be embedded with relative URLs and responsive CSS such as `max-width:100%; max-height:100%; object-fit:contain`. An intentional crop can use `object-fit:cover` and `object-position`. Use local project files and provide alt text. `slides[].image` still accepts PNG, JPEG, BMP, and GIF.

The `init` template loads bundled KaTeX. Use `\(...\)` for inline formulas and `\[...\]` or `$$...$$` for display formulas. Keep the KaTeX stylesheet, scripts, and fonts under `assets/katex/` when you copy the template references.

## Layout and text safety

- `canvas`: positive 16:9 pixel size. Default 1920×1080.
- `layout`: `[x, y, width, height]` boxes in output pixels. Defaults scale for other 16:9 sizes. Character and subtitle boxes are validated for overlap.
- Character default: 300×300 at `[35,755]`; the subtitle starts at x=385 to reserve room beside it.
- `notes_top` uses about one quarter of the right-panel height for a left-aligned heading. `notes_bottom` uses the lower three quarters for left-aligned supporting text. Defaults leave a visible gap between the two areas.
- Subtitle text uses the bold subtitle font and begins at 54 px. It shrinks to fit, then wraps. Note text also shrinks and wraps inside its own area; the top heading uses a larger bold size than the bottom note. If content still cannot fit at the minimum safe size, the last line ends with an ellipsis instead of drawing outside the box.
- `layout.fonts` provides independent `slide`, `subtitle`, and `notes` family names for HTML CSS/role selection. The bundled font is Noto Sans JP variable TTF with its SIL Open Font License at `assets/fonts/OFL.txt`.
- Keep slides to one main idea with readable type and contrast. Split dense material across slides rather than forcing it into the frame.

## Narration and notes

- `slides[].script`: visible subtitle and default spoken text, split on `。！？!?`; each sentence becomes one audio/video segment.
- `slides[].motions`: optional action per spoken sentence: `idle`, `wave`, `nod`, `think`, `point`, `cheer`, `walk`, `surprise`. Missing entries use `idle`.
- `slides[].tts_texts`: optional TTS-only pronunciation overrides aligned by sentence index. Empty or omitted entries send `script` unchanged.
- `slides[].note_top` and `slides[].note_bottom`: heading and supporting text on the right panel. Keep the heading concise. Use the large bottom area for a meaningful 2–4 sentence explanation (typically about 60–140 Japanese characters), such as context, reasoning, an example, or an important caveat; avoid repeating the slide verbatim or adding filler.
- `renderer.browser`: optional Edge/Chrome executable path, absolute or relative to the project. Edge is discovered from its Windows application registration by default.

The default voice is kokuren_3rd, UUID `38d7216c-e595-4d8f-b06c-1fc376e47c0a`, style ノーマル, style ID `1069147200`. The Engine API receives the style ID as `speaker`.

## Commands

```powershell
.\biim-video.ps1 init projects/my-video
.\biim-video.ps1 validate projects/my-video
.\biim-video.ps1 preview projects/my-video
.\biim-video.ps1 build projects/my-video
```

`preview` creates a static composite for each narration sentence, including the selected character pose, under `output/preview/`. It does not call AivisSpeech or FFmpeg.

Video output remains H.264/AAC, 1920×1080 at 30 fps by default, `yuv420p`, and MP4 faststart. FFmpeg and the local AivisSpeech Engine are required for the final build; HTML slide rendering also needs Edge or Chrome. Python, pip, Node.js, and external slide apps are not required.
