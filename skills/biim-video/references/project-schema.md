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
    "slide": [16, 16, 1440, 810],
    "subtitle": [350, 870, 1528, 178],
    "notes_top": [1498, 34, 388, 182],
    "notes_bottom": [1498, 274, 388, 532],
    "character": [0, 740, 330, 332],
    "subtitle_font_size": 64,
    "note_top_font_size": 44,
    "note_font_size": 34,
    "subtitle_color": "#ff3434",
    "subtitle_stroke": 4,
    "subtitle_outer_stroke": 9,
    "character_crop": [36, 57, 184, 148],
    "fonts": { "slide": "Noto Sans JP", "subtitle": "M PLUS Rounded 1c", "notes": "Noto Sans JP" }
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

`slides[].html` is the recommended format. The HTML CSS viewport is 1280×720, captured at 2560×1440; author the page at that size and set `margin:0`, `overflow:hidden`, and explicit dimensions. Local images can be embedded with relative URLs and responsive CSS such as `max-width:100%; max-height:100%; object-fit:contain`. An intentional crop can use `object-fit:cover` and `object-position`. Use local project files and provide alt text. `slides[].image` still accepts PNG, JPEG, BMP, and GIF.

The `init` template loads bundled KaTeX. Use `\(...\)` for inline formulas and `\[...\]` or `$$...$$` for display formulas. Keep the KaTeX stylesheet, scripts, and fonts under `assets/katex/` when you copy the template references.

## Layout and text safety

- `canvas`: positive 16:9 pixel size. Default 1920×1080.
- `layout`: `[x, y, width, height]` boxes in output pixels. Defaults scale for other 16:9 sizes. Character and subtitle boxes are validated for overlap.
- Character default: `[0,740,330,332]`; bottom-center proportional fit, rendered above every layer. Slight slide overlap is allowed. Captions begin at x=350, clear of the character.
- `character_crop`: optional `[x,y,width,height]` in source GIF pixels, applied consistently in previews and FFmpeg. `[36,57,184,148]` covers every frame of the bundled animations. Remove or recompute when replacing assets.
- Notes occupy separate upper/lower frame openings with a visible gap, both top-left aligned. Heading is larger and bolder than body.
- Caption uses M PLUS Rounded 1c ExtraBold 800, starts at 64px, and wraps/shrinks to fit. Notes use Noto Sans JP heading 800/44px and body 600/34px. Sizes are 1080p reference sizes scaled by canvas height.
- `subtitle_color`: `#RRGGBB` fill, default `#ff3434`. `subtitle_stroke`: black inner stroke width (4px); `subtitle_outer_stroke`: white outer stroke (9px). Set both to 0 to remove strokes. Adjust caption box padding for larger strokes.
- `layout.fonts` provides independent `slide`, `subtitle`, `notes` family names. Local Noto Sans JP and M PLUS Rounded 1c font faces are loaded directly. System font families may also be selected. Keep the fonts and both OFL files with the project.
- Overlay text renders at 2× through Edge and downsamples. Remaining caption/note overflow is a render error; shorten content or adjust the box rather than truncating.
- `renderer.audit_slides: true` enables browser checks for slide text below 24 CSS px, text outside the canvas/`.art` diagram region, and heading/diagram overlap. Defaults enabled for new projects; opt in for existing ones. These checks supplement semantic and visual inspection.
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

Preview a selected range by array position: `preview projects/my-video -PreviewFrom 14 -PreviewTo 18`. Generated frame HTML beside each preview is an inspectable rendering artifact; edit project.json and slide HTML as the source of truth.
