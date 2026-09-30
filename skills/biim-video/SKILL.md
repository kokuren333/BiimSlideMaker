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
- Prefer a standalone 16:9 HTML file for each slide (`slides/001.html`). The renderer captures HTML in Edge/Chrome at a 1280×720 CSS viewport with 2× device scale (2560×1440), then fits it into the 1440×810 main frame. Put reusable images in the project, reference them with relative URLs, and size them with CSS (`max-width`, `max-height`, `object-fit: contain` or `cover`). Use `alt` text and preserve aspect ratio unless cropping is intentional. An attached image may be copied into the project and incorporated as evidence, a diagram, or a background element; crop and scale it to the slide composition rather than simply shrinking the whole screenshot.
- Use the bundled local KaTeX assets for math (`\(...\)` inline, `\[...\]` or `$$...$$` display); see the generated HTML slide for its CSS/JS references. Keep slide HTML and images local so rendering does not depend on CDN access.
- Design for the Biim frame: show one key idea per slide, use a clear visual hierarchy and strong contrast, and avoid packing paragraphs or many competing cards into the 16:9 content area. Use medium or bold weights for important text and a readable body size; split dense explanations across slides rather than shrinking everything.
- Write the displayed subtitle and narration in `slides[].script`. `biim-video.ps1` splits it at Japanese sentence punctuation (`。！？!?`); provide a `motions` entry per spoken sentence only when an expressive gesture is useful. The default is `idle`.
- Preserve the subtitle exactly in `script`. If AivisSpeech may misread a particular term, keep the spoken rendering in the matching `tts_texts` entry; do not rewrite the visible caption to match phonetic spelling.
- Give each frame a concise `note_top` heading and useful `note_bottom` explanation. Keep facts and source links in the project notes or visible notes when needed.
- Choose motion with intent: `wave` for greeting, `nod` for agreement or confirmation, `think` for consideration, `point` for explanation or emphasis, `cheer` for a positive result, `surprise` for a genuine surprise, and `walk` only when movement is called for. Avoid changing motion every sentence without a reason.
- Generate an original simple rectangular frame by default. If the ignored local assets/frame-nc293888.png exists, init uses it instead. Do not include this user-supplied image or local production projects in commits unless redistribution rights have been confirmed. At 1920×1080 the main slide is `[16,16,1440,810]`, note heading `[1498,34,388,182]`, note body `[1498,274,388,532]`, caption `[350,870,1528,178]`, and character `[0,740,330,332]`. Tune project boxes to the actual supplied frame. Character renders last, above all other layers; light overlap with the main slide is allowed. Keep captions clear of its box.
- The bundled character crop `[36,57,184,148]` covers the union of visible pixels across every animation frame. It removes transparent margins, scales proportionally, and anchors bottom-center in the character box. Recompute or remove the crop for different assets. Preview and video must agree.
- Caption defaults: local M PLUS Rounded 1c ExtraBold 800, 64px, red fill with black 4px and white 9px strokes. Notes use local Noto Sans JP, heading 800/44px and body 600/34px. Roles remain configurable in `layout.fonts`. Preserve both OFL licenses and KaTeX MIT.
- Both right fields align top-left inside the actual upper/lower frame openings. Use a short heading and 2–4 meaningful supporting sentences, usually 60–140 Japanese characters. Do not duplicate note elements inside slide HTML.
- All overlay text renders in Edge with local font faces at 2× and downsamples to final resolution. Fit text with natural wrapping and size reduction; overflow is a render error, never silently ellipsize narration.
- New projects enable `renderer.audit_slides`: reject text below 24 CSS px, text outside slide/diagram regions, or heading/diagram overlap. Existing projects can opt in. Use about 52–60px titles, 28–34px leads, and at least 26px diagram labels at the 1280×720 authoring size. Check at final composite scale, not only in standalone HTML.
- Audit diagram semantics as well as geometry: label numerator/denominator and units, keep percentage bars proportional to their values, avoid unlabeled circles/meters implying unsupported relationships, and use arrows only for actual sequence or transformation. Use semantic grouping (fraction, comparison, timeline, part-to-whole); split rather than cram. Inspect every distinct slide visually.
- Keep output at 16:9. Default delivery is 1920x1080, 30fps, H.264/AAC, `yuv420p`, MP4 faststart. These settings are defined by the CLI.
- The default AivisSpeech voice metadata is kokuren_3rd, UUID `38d7216c-e595-4d8f-b06c-1fc376e47c0a`, style ノーマル, style ID `1069147200`. Use the style ID for the Engine API `speaker` parameter; the UUID identifies the voice model.

## Validate and deliver

Before final encoding, render the static preview for the project and inspect the slide, subtitle, notes, and character placement. Preview does not call AivisSpeech or FFmpeg.

1. Run `.\biim-video.ps1 validate <project-dir>` and fix all errors.
2. When the local AivisSpeech Engine and ffmpeg are available, run `.\biim-video.ps1 build <project-dir>` to create the MP4. Build also writes generated frames, WAVs, and segment videos under `<output-stem>_work/` for review.
3. Summarize the project path, output path, the validation/build performed, and any unavailable local service that prevented rendering.
