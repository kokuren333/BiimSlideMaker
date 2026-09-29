# Biim-style layout and writing notes

## What the reference describes

The original biim system was devised for RTA/gameplay commentary as a way to keep gameplay and explanation visible together. In the creator interview, biim describes the arrangement as main footage at upper left, explanation at right, comment/caption area along the bottom, and a Yukkuri character at lower left. The rationale was to avoid making viewers repeatedly shift attention between the gameplay and explanation.

This project adapts those functional zones to an explainer canvas: main slide at upper left, notes at right, synchronized narration subtitles along the bottom, and Kokuren at lower left. Do not imitate specific characters, jokes, voice styles, or copyrighted assets from reference videos. Keep the user's own content and supplied assets.

## Writing for the regions

- **Main slide**: show the current claim, object, comparison, diagram, or evidence. Make the visual useful even when viewed briefly; avoid duplicating every narration sentence as slide text.
- **Right notes**: `note_top` is a short label or current point. `note_bottom` contains supporting details, definitions, caveats, or a compact sequence. Prefer fragments and short bullets to a dense essay.
- **Bottom subtitle**: `script` is spoken narration and also the synchronized subtitle. Each sentence is a separate audio/video segment. Keep sentences concise enough to fit the reduced width to the right of the avatar; split long thoughts into natural sentences rather than shrinking all text.
- **Lower-left avatar**: reserve x=35..285 and y=795..1045 at the default 1920x1080 canvas. The subtitle begins at x=330. Pick gestures that support the spoken intent; idle is appropriate for neutral explanation.
- **Hierarchy**: let the main visual dominate, make note labels scannable, and make the current spoken line legible. Keep labels, notes, and narration semantically distinct rather than repeating the same paragraph in all three places.
- **Adaptation**: Biim layouts vary. Preserve the functional separation of primary content, supporting explanation, spoken text, and character while tuning sizes to the story and supplied background. Do not force a classic RTA/gameplay look when a modern or calmer design better fits the user's topic.

## Research sources

- Interview with biim, section “biimシステムの誕生”: [Denfaminicogamer](https://denfaminicogamer.jp/interview/190514c). The interview directly describes the layout and its viewing rationale.
- [Biim system overview](https://w.atwiki.jp/cookie_kaisetu/pages/552.html) records the common upper-left play area, lower explanation/subtitles and character, and right supplemental panel.
- AivisSpeech Engine's [official README](https://github.com/Aivis-Project/AivisSpeech-Engine) documents the `/audio_query` and `/synthesis` flow and the style ID passed as `speaker`.

The community overview is secondary and can contain fandom-specific language. Use it only to corroborate the layout; the creator interview is the primary source for the design intent.
