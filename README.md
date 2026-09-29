# BiimSlideMaker

Agent-first toolkit for authoring editable 16:9 narrated videos. Give the Agent a topic and creative brief; it creates the project files directly, validates them, and renders an upload-ready MP4 through the local CLI. No Marp, GUI prompt transfer, or Gemini web workflow is required.

## Agent workflow

At the start of a new video, the Agent checks for any missing details:

- Topic and intended audience
- Approximate video duration or slide count
- Preferred visual design
- Reference sources, files, or URLs
- Special requirements, required points, or topics to avoid

It then creates and edits project files in the workspace. The repository's [`AGENTS.md`](AGENTS.md) gives project instructions. [`skills/biim-video/SKILL.md`](skills/biim-video/SKILL.md) explains the video workflow and links to the project schema and researched Biim layout guide.

## Requirements

- Python 3.10+
- FFmpeg on `PATH`
- AivisSpeech Engine running locally at `http://127.0.0.1:10101` for narration synthesis
- Install Python libraries:

```powershell
python -m pip install -r requirements.txt
```

## Start a project

```powershell
python biim_cli.py init projects/my-video
```

The initializer creates `project.yaml`, a sample SVG slide, the default frame, and all character animation GIFs under the project. Edit `project.yaml` and the slide assets directly, or replace the SVG with PNG, JPEG, or WebP artwork.

Check the project before synthesis:

```powershell
python biim_cli.py validate projects/my-video
```

Render the video (AivisSpeech Engine and FFmpeg must be available):

```powershell
python biim_cli.py build projects/my-video
```

By default the final file is `projects/my-video/output/final.mp4`. Intermediate frames, speech WAVs, and sentence-length video segments are kept in `output/final_work/` so the Agent can revise or inspect individual parts. Speech cache names include a fingerprint of the synthesis text and voice settings.

## Project format

Each project is a directory with a `project.yaml` manifest and editable slide assets. YAML is used for metadata and narration; slide artwork may use SVG, PNG, JPEG, or WebP. SVG is a convenient text format for Agent-authored diagrams and layouts, but Marp and Markdown slides are not required. The frame, layout boxes, font sizes, voice, soundtrack, and output location can be configured per project.

The default manifest uses this core structure:

```yaml
version: 1
title: Example video
canvas: {width: 1920, height: 1080}
fps: 30
output: output/final.mp4
assets:
  background: assets/frame.png
  animations: assets/animations
  bgm: ""
voice:
  name: kokuren_3rd
  speaker_uuid: 38d7216c-e595-4d8f-b06c-1fc376e47c0a
  style_name: ノーマル
  style_id: 1069147200
  engine_url: http://127.0.0.1:10101
slides:
  - id: 1
    image: slides/001.svg
    script: |
      こんにちは。
      AivisSpeech APIを使います。
    motions: [wave, point]
    tts_texts: ["", "エイビススピーチ エーピーアイを使います。"]
    note_top: 要点
    note_bottom: 補足説明
```

`script` is split at `。！？!?`; each sentence gets one audio/video segment and one `motions` entry. Motion names are `idle`, `wave`, `nod`, `think`, `point`, `cheer`, `walk`, and `surprise`. Omitted motion entries default to `idle`. Optional `tts_texts` entries override only the text sent to speech synthesis; the subtitle remains the original `script`. Leave an entry empty to use the script as-is. The `aivis-pronunciation` skill guides selective readings for hard-to-pronounce kanji and English. See [`skills/biim-video/references/project-schema.md`](skills/biim-video/references/project-schema.md) for all fields.

## Layout and output

The default frame follows Biim's functional layout: primary visual at upper left, supplementary explanation at right, spoken subtitle along the bottom, and character at lower left. The character occupies roughly x=35–285; subtitles start at x=330 so they do not overlap. The `layout` fields in `project.yaml` can be tuned for alternate backgrounds and content. Boxes use `[x, y, width, height]` and scale from the 1920×1080 defaults for other 16:9 canvases; validation rejects subtitle-character overlaps.

Default output is 1920×1080, 16:9, 30 fps, H.264 video (CRF 18), AAC audio (192 kbps, 48 kHz), `yuv420p`, and MP4 faststart. BGM is optional and mixed beneath narration. These broadly compatible settings are suitable for common video platforms.

The default AivisSpeech voice is `kokuren_3rd` (UUID `38d7216c-e595-4d8f-b06c-1fc376e47c0a`), style `ノーマル`, ID `1069147200`. The CLI passes the style ID as the Engine API's `speaker` parameter. The UUID and name remain in the manifest as model metadata. The local model must be installed in AivisSpeech Engine.

## Legacy GUI

`movie_maker_gui.py` remains available for existing PDF plus YAML projects. It is a legacy compatibility workflow; new Agent-authored projects should use `biim_cli.py`.

## Biim references

The design notes in [`skills/biim-video/references/biim-conventions.md`](skills/biim-video/references/biim-conventions.md) are based on [biim's interview about the system's origin](https://denfaminicogamer.jp/interview/190514c) and a [community overview](https://w.atwiki.jp/cookie_kaisetu/pages/552.html). The layout guide treats the original as a practical information hierarchy and adapts it to the user's Kokuren character and content.
