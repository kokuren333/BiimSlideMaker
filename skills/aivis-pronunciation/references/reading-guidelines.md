# Selective katakana readings

## When to override

Add a reading only where default Japanese TTS is plausibly ambiguous or likely to be wrong:

- Rare, specialist, or context-dependent kanji readings (especially names and places).
- Acronyms that should be spoken letter by letter, such as API or UUID.
- English brands, product names, abbreviations, and code terms whose expected Japanese reading is not obvious from their spelling.
- Mixed Japanese/English compounds where a specific term is liable to be read as ordinary English rather than its expected Japanese pronunciation.

Usually leave ordinary kanji, common loanwords, numbers with obvious context, and familiar technical words unchanged. The purpose is to correct likely mistakes, not to phoneticize all text.

## Choosing a reading

- Check user-provided context first; for names and branded terms, verify the preferred pronunciation from an authoritative or first-party source when possible.
- Write the corrected token in katakana within an otherwise natural Japanese sentence. Examples: `AivisSpeech API` → `エイビススピーチ エーピーアイ`; a known surname `東（あずま）` → `アズマ`.
- Keep punctuation and the sentence boundary unchanged. The override must still be one corresponding sentence, so it stays aligned with the sentence's audio segment and motion.
- Avoid inserting spaces between ordinary Japanese words. A small boundary before an English-derived term can help, but keep phrasing natural.
- Do not guess a person's or product's pronunciation from kanji/letters alone. Leave the display text unchanged and report the unresolved reading in project notes if no dependable source is available.

## JSON mapping

The list index corresponds to each sentence in `script` after splitting at `。！？!?`. Empty strings mean “send the original subtitle to TTS.” For example:

```json
{
  "script": "AivisSpeech APIで音声を作ります。UUIDを設定します。",
  "tts_texts": [
    "エイビススピーチ エーピーアイで音声を作ります。",
    "ユーユーアイディーを設定します。"
  ]
}
```

The user-facing captions still show the original English and acronym. Only synthesis receives the katakana forms. Use strings in the list and `""` to fall back to the original.
