# Portable PowerShell project schema

The Agent-facing CLI uses `project.json`, parsed by built-in PowerShell/.NET JSON support. Slide artwork is a separate PNG, JPEG, BMP, or GIF image; no presentation framework is required. The CLI itself has no Python or pip dependencies. FFmpeg is used for encoding, and AivisSpeech Engine is the local speech service.

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
    "subtitle": [330, 847, 1520, 163],
    "notes_top": [1413, 66, 444, 324],
    "notes_bottom": [1413, 410, 444, 310],
    "character": [35, 795, 250, 250],
    "subtitle_font_size": 54,
    "note_font_size": 35,
    "font_family": "Meiryo"
  },
  "slides": [
    {
      "id": 1,
      "image": "slides/001.png",
      "script": "こんにちは。AivisSpeech APIを使います。",
      "motions": ["wave", "point"],
      "tts_texts": ["", "エイビススピーチ エーピーアイを使います。"],
      "note_top": "要点",
      "note_bottom": "用語や背景を簡潔に補足します。"
    }
  ]
}
```

## Fields

- `canvas`: positive 16:9 pixel size. Default 1920x1080.
- `fps`: integer 1..60; default 30.
- `output`: MP4 destination relative to the project directory; default `output/final.mp4`.
- `assets.background`: existing frame image.
- `assets.animations`: directory containing `manifest.json`, action preview GIFs, and optional sprite sheets. `init` copies the bundled animation set.
- `assets.bgm`: optional soundtrack path; empty string disables it.
- `voice.style_id`: AivisSpeech Engine style id passed as the compatible API's `speaker` parameter. The voice name and speaker UUID stay in the project as model metadata.
- `layout`: `[x, y, width, height]` boxes in output pixels. Defaults scale for other 16:9 sizes. Validation rejects boxes outside the canvas and subtitle/character overlap.
- `slides[].image`: relative or absolute PNG, JPEG, BMP, or GIF art.
- `slides[].script`: visible subtitle and base spoken text, split on `。！？!?`.
- `slides[].motions`: optional action name per split sentence. Allowed: `idle`, `wave`, `nod`, `think`, `point`, `cheer`, `walk`, `surprise`. Missing entries default to `idle`.
- `slides[].tts_texts`: optional TTS-only pronunciation overrides by sentence index. Empty or omitted entries send the corresponding `script` sentence unchanged; the visible subtitle never changes.
- `slides[].note_top` and `slides[].note_bottom`: right-column label and supporting notes.

The script caches speech audio using a fingerprint of the TTS text and voice settings. Intermediates are under `<output-stem>_work/`.

## Commands

```powershell
.\biim-video.ps1 init projects/my-video
.\biim-video.ps1 validate projects/my-video
.\biim-video.ps1 build projects/my-video
```
