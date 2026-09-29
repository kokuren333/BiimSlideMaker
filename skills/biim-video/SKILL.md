---
name: biim-video
description: Create or edit Agent-authored Biim-style explainer video projects in this repository, including slide artwork, narration, frame notes, subtitle layout, character motion, and MP4 output.
---

# Biim style video production

Use this skill when the user asks to plan or create a narrated slide or explainer video in this repository. Read [project-schema.md](references/project-schema.md) before editing project files. For Biim conventions and respectful adaptation, use [biim-conventions.md](references/biim-conventions.md).

When pronunciation may be unreliable, read and apply the repository skill `skills/aivis-pronunciation/SKILL.md` and put selective reading overrides in `slides[].tts_texts`.

## Gather the brief

Before creating a new video, ask together for any missing essentials: topic and audience, approximate duration or slide count, preferred visual style, source material or URLs, and special requirements. Reuse information already provided; accept "you decide" and record reasonable assumptions in the project notes.

## Build an editable project

- Create a project directory with `.\biim-video.ps1 init <project-dir>`.
- Keep the project editable by authoring `project.json` and slide assets directly. Do not route prompts or scripts through a GUI or an external LLM interface.
- Use PNG, JPEG, BMP, or GIF per-slide artwork. Build custom layouts in the slide artwork and keep each source at 16:9; the renderer places it in the main frame. Marp is optional and not required.
- Write the displayed subtitle and narration in `slides[].script`. `biim-video.ps1` splits it at Japanese sentence punctuation (`。！？!?`); provide a `motions` entry per spoken sentence only when an expressive gesture is useful. The default is `idle`.
- Preserve the subtitle exactly in `script`. If AivisSpeech may misread a particular term, keep the spoken rendering in the matching `tts_texts` entry; do not rewrite the visible caption to match phonetic spelling.
- Give each frame a concise `note_top` heading and useful `note_bottom` explanation. Keep facts and source links in the project notes or visible notes when needed.
- Choose motion with intent: `wave` for greeting, `nod` for agreement or confirmation, `think` for consideration, `point` for explanation or emphasis, `cheer` for a positive result, `surprise` for a genuine surprise, and `walk` only when movement is called for. Avoid changing motion every sentence without a reason.
- Preserve the reserved bottom-left character area. At 1920x1080 the CLI places the character at x=35, y=795, 250px square, and starts subtitle text at x=330 so the two do not overlap. Keep subtitles concise and readable in that narrower region. Layout boxes scale for other 16:9 sizes, and validation rejects overlap.
- Keep output at 16:9. Default delivery is 1920x1080, 30fps, H.264/AAC, `yuv420p`, MP4 faststart. These settings are defined by the CLI.
- The default AivisSpeech voice metadata is kokuren_3rd, UUID `38d7216c-e595-4d8f-b06c-1fc376e47c0a`, style ノーマル, style ID `1069147200`. Use the style ID for the Engine API `speaker` parameter; the UUID identifies the voice model.

## Validate and deliver

1. Run `.\biim-video.ps1 validate <project-dir>` and fix all errors.
2. When the local AivisSpeech Engine and ffmpeg are available, run `.\biim-video.ps1 build <project-dir>` to create the MP4. Build also writes generated frames, WAVs, and segment videos under `<output-stem>_work/` for review.
3. Summarize the project path, output path, the validation/build performed, and any unavailable local service that prevented rendering.
