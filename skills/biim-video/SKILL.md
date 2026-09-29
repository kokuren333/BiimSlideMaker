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
- Keep the project editable by authoring `project.json`, HTML slide files, and image assets directly. Do not route prompts or scripts through a GUI or an external LLM interface.
- Prefer a standalone 16:9 HTML file for each slide (`slides/001.html`). The renderer captures HTML in Edge/Chrome at 1280×720 and scales it into the main frame. Put reusable images in the project, reference them with relative URLs, and size them with CSS (`max-width`, `max-height`, `object-fit: contain` or `cover`). Use `alt` text and preserve aspect ratio unless cropping is intentional. An attached image may be copied into the project and incorporated as evidence, a diagram, or a background element; crop and scale it to the slide composition rather than simply shrinking the whole screenshot.
- Use the bundled local KaTeX assets for math (`\(...\)` inline, `\[...\]` or `$$...$$` display); see the generated HTML slide for its CSS/JS references. Keep slide HTML and images local so rendering does not depend on CDN access.
- Design for the Biim frame: show one key idea per slide, use a clear visual hierarchy and strong contrast, and avoid packing paragraphs or many competing cards into the 16:9 content area. Use medium or bold weights for important text and a readable body size; split dense explanations across slides rather than shrinking everything.
- Write the displayed subtitle and narration in `slides[].script`. `biim-video.ps1` splits it at Japanese sentence punctuation (`。！？!?`); provide a `motions` entry per spoken sentence only when an expressive gesture is useful. The default is `idle`.
- Preserve the subtitle exactly in `script`. If AivisSpeech may misread a particular term, keep the spoken rendering in the matching `tts_texts` entry; do not rewrite the visible caption to match phonetic spelling.
- Give each frame a concise `note_top` heading and useful `note_bottom` explanation. Keep facts and source links in the project notes or visible notes when needed.
- Choose motion with intent: `wave` for greeting, `nod` for agreement or confirmation, `think` for consideration, `point` for explanation or emphasis, `cheer` for a positive result, `surprise` for a genuine surprise, and `walk` only when movement is called for. Avoid changing motion every sentence without a reason.
- Preserve the reserved bottom-left character area. At 1920×1080 the larger character defaults to x=35, y=755, 300×300 px; subtitles begin at x=385. Narration is split at `。！？!?`; each subtitle starts at a bold 54 px and shrinks to fit, then wraps naturally, with a final ellipsis if the text still exceeds its box. Keep sentences concise so this safety net is rarely needed.
- The right note panel is split into a top quarter heading area and a lower three-quarter explanation area with a clear gap. Both align left at the top of their regions. `note_top` uses a larger bold font than `note_bottom`; both shrink and wrap to stay inside their assigned boxes, with ellipsis only as a last resort. Keep the heading short. Give `note_bottom` enough substance to serve as real supporting explanation: usually 2–4 concise sentences (roughly 60–140 Japanese characters), adding useful context such as why it matters, how it works, a concrete example, or a caveat. Do not merely repeat the slide title or subtitle, and do not pad the note with irrelevant filler.
- Slide, subtitle, and note font roles are separately configurable in `layout.fonts`. Noto Sans JP is bundled under SIL Open Font License and used by default; the subtitle is bold, the note heading is bold and larger, and note body is regular. Keep its OFL license with the font when redistributing.
- Keep output at 16:9. Default delivery is 1920x1080, 30fps, H.264/AAC, `yuv420p`, MP4 faststart. These settings are defined by the CLI.
- The default AivisSpeech voice metadata is kokuren_3rd, UUID `38d7216c-e595-4d8f-b06c-1fc376e47c0a`, style ノーマル, style ID `1069147200`. Use the style ID for the Engine API `speaker` parameter; the UUID identifies the voice model.

## Validate and deliver

Before final encoding, render the static preview for the project and inspect the slide, subtitle, notes, and character placement. Preview does not call AivisSpeech or FFmpeg.

1. Run `.\biim-video.ps1 validate <project-dir>` and fix all errors.
2. When the local AivisSpeech Engine and ffmpeg are available, run `.\biim-video.ps1 build <project-dir>` to create the MP4. Build also writes generated frames, WAVs, and segment videos under `<output-stem>_work/` for review.
3. Summarize the project path, output path, the validation/build performed, and any unavailable local service that prevented rendering.
