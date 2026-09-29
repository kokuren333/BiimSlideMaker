---
name: aivis-pronunciation
description: Prepare Japanese narration for AivisSpeech by adding selective katakana readings for terms likely to be misread while preserving the original displayed subtitle.
---

# Selective AivisSpeech reading overrides

Use this skill when authoring narrated projects that synthesize through AivisSpeech. Read [reading-guidelines.md](references/reading-guidelines.md) for term selection and markup rules.

- Keep the user's exact visible text in `slides[].script`.
- Split the subtitle text by the CLI's sentence punctuation rules (`。！？!?`) and align any pronunciation overrides by sentence index in `slides[].tts_texts`.
- Inspect for uncommon kanji, ambiguous names, acronyms, English product names, and foreign terms that are likely to be pronounced incorrectly in context.
- Create an override only for sentences containing a term that needs help. In that TTS-only sentence, replace only the uncertain token(s) with a reliable katakana reading; leave surrounding wording, normal kanji, punctuation, and prosody intact.
- Leave other entries blank or omit trailing entries. The CLI sends the original sentence to AivisSpeech when no override is present. The subtitle always remains unchanged.
- Preserve meaning, names, numbers, sentence boundaries, and emphasis. Do not convert a whole sentence to kana, introduce alternate phrasing, or guess uncertain readings without checking context or a source.
- Add `tts_texts` only when at least one override is useful. State pronunciation uncertainty in notes instead of silently making a weak guess.
