# Project file schema

Projects use a YAML manifest, but slide artwork is deliberately not locked to a presentation framework. `biim_cli.py` currently accepts one static asset per slide in SVG, PNG, JPEG, or WebP format. SVG rendering requires the `cairosvg` requirement.

## Example

```yaml
version: 1
title: Example explainer
canvas: {width: 1920, height: 1080}
fps: 30
output: output/final.mp4
assets:
  background: assets/frame.png
  animations: assets/animations
  bgm: assets/music.mp3 # optional; use "" for no music
  font: assets/font.ttf # optional; system Japanese fonts are auto-detected
voice:
  name: kokuren_3rd
  speaker_uuid: 38d7216c-e595-4d8f-b06c-1fc376e47c0a
  style_name: ノーマル
  style_id: 1069147200
  engine_url: http://127.0.0.1:10101
audio:
  bgm_volume: 0.2
layout:
  slide: [40, 28, 1280, 720]
  subtitle: [330, 847, 1520, 163]
  notes_top: [1413, 66, 444, 324]
  notes_bottom: [1413, 410, 444, 310]
  character: [35, 795, 250, 250]
  subtitle_font_size: 54
  note_font_size: 35
slides:
  - id: 1
    image: slides/001.svg
    script: |
      こんにちは。
      AivisSpeech APIを使います。
    motions: [wave, point]
    tts_texts: ["", "エイビススピーチ エーピーアイを使います。"]
    note_top: 要点
    note_bottom: |
      用語や補足を簡潔に書きます。
```

## Fields and behavior

- `version`: currently `1`.
- `canvas`: positive 16:9 dimensions. Defaults to 1920x1080.
- `fps`: output frame rate; defaults to 30.
- `output`: MP4 path relative to `project.yaml` directory; defaults to `output/final.mp4`.
- `assets.background`: frame/background image. `biim_cli.py init` copies the repository's default frame into the project.
- `assets.animations`: directory containing `manifest.json` plus each action's preview GIF. The initializer bundles all actions.
- `assets.bgm`: optional audio file mixed under narration; blank disables music.
- `assets.font`: optional Japanese font file. If omitted, common system fonts are searched.
- `voice.style_id`: AivisSpeech Engine global style ID passed to the compatible `speaker` query parameter. Keep the UUID and voice name as provenance metadata.
- `layout`: `[x, y, width, height]` boxes in output pixels. Defaults scale with 16:9 canvas size. Tune for alternate frame images. Keep subtitle box clear of the bottom-left character; validation rejects an overlap.
- `layout.character`: optional character overlay box; default `[35, 795, 250, 250]` at 1920x1080.
- `slides[].image`: relative or absolute static artwork path. Supported: `.svg`, `.png`, `.jpg`, `.jpeg`, `.webp`.
- `slides[].script`: narration. Split on `。！？!?`; line breaks are joined before splitting.
- `slides[].tts_texts`: optional TTS-only pronunciation override for each split sentence, by index. The subtitle always uses the original `script` text. Use katakana only for terms likely to be misread; leave the entry empty or omit it to send the displayed sentence unchanged. Excess entries or non-string values fail validation.
- `slides[].motions`: optional action list corresponding to split spoken sentences. Allowed names: `idle`, `wave`, `nod`, `think`, `point`, `cheer`, `walk`, `surprise`. Omitted items default to `idle`; excess items are validation errors.
- `slides[].note_top` and `slides[].note_bottom`: visible right-column heading and explanation.

The generator caches WAVs under `<output-stem>_work/audio/` using a fingerprint of TTS text and voice settings. Changing text or voice settings creates a new cache entry.

## Commands

```powershell
python biim_cli.py init projects/my-video
python biim_cli.py validate projects/my-video
python biim_cli.py build projects/my-video
```

The project can live elsewhere; the CLI resolves asset paths relative to its `project.yaml`.
