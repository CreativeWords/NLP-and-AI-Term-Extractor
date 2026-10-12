# NLP+AI Term Extractor

Build bilingual glossaries from a project's source material. The tool extracts
terminology candidates from **TMX, PDF and DOCX** files with
[spaCy](https://spacy.io), then (optionally) uses **Claude** or **GPT** to clean
the list, translate each term and write a short definition. When a TMX
translation memory is available, the AI reads the human-verified translations
instead of translating from scratch.

Output is a JSON file plus a matching XLSX spreadsheet, ready for review.

## Features

- **Mixed inputs**: drop TMX, PDF and DOCX files into one folder (optionally
  scanned recursively). Legacy `.doc` files are detected and skipped.
- **24 source languages**, including Chinese and Japanese (see
  [Supported languages](#supported-languages)).
- **NLP extraction**: noun chunks and POS-based candidates, term typing,
  frequency and segment counts, up to three in-context examples per term, and a
  junk filter for codes, units and extraction artifacts.
- **Configurable filtering**: n-gram lengths (1–5), minimum counts per n-gram,
  term caps, minimum segment length, custom stopwords.
- **Optional AI steps** (Anthropic or OpenAI):
  - *Cleaning*: merges variants, normalises surface forms, drops generic words.
  - *Translate + define*: target term and definition for each entry. Terms
    found in TMX files use their aligned human translations, and
    inconsistencies in the memory are flagged in a `notes` column.
- **Cost visibility**: per-step token usage and estimated cost in the run log.
- **Editable prompts**: all AI prompts live in `term_extractor_prompts.yaml`.

## Requirements

- Python 3.10+ with Tkinter (included in the standard Windows and macOS
  installers; on Linux install e.g. `python3-tk`)
- An Anthropic and/or OpenAI API key, only if you use the AI steps

## Installation

```bash
git clone https://github.com/CreativeWords/NLP-AI-Term-Extractor.git
cd NLP-AI-Term-Extractor
pip install -r requirements.txt
```

If you only need NLP extraction, the AI packages (`anthropic`, `openai`,
`pyyaml`) can be skipped.

### spaCy models

You do **not** need to install spaCy models manually. On first use of a
language, the tool downloads the matching large (`lg`) model, compatible with
your installed spaCy version, and caches it in:

| OS      | Default location                                        |
|---------|---------------------------------------------------------|
| Windows | `%LOCALAPPDATA%\term-extractor\spacy_models`            |
| macOS   | `~/Library/Caches/term-extractor/spacy_models`          |
| Linux   | `$XDG_CACHE_HOME/term-extractor/spacy_models` (`~/.cache/...`) |

To use a different folder, set the `TERM_EXTRACTOR_MODELS_DIR` environment
variable. Models are several hundred MB each, so the first run for a language
takes a while.

## Usage

```bash
python term_extractor.py
```

1. Choose the folder that contains your source files.
2. Select the source language, n-gram lengths and term cap.
3. *(Optional)* Open the advanced options to tune the junk filter, minimum
   segment length, minimum counts and custom stopwords.
4. *(Optional)* Enable AI cleaning and/or translation, then pick a model and
   enter domain, context, target language and definition language.
5. Click run and follow progress in the log.

Results are written **into the input folder** as
`<folder>_term_extraction_<timestamp>.json` and `.xlsx`.

### API keys

The API key field is pre-filled from the `ANTHROPIC_API_KEY` or
`OPENAI_API_KEY` environment variable, depending on the selected model. You
can also paste a key directly into the field. Keys are never written to disk
by the tool.

If an AI step fails, the AI chain stops and the output reflects the last step
that completed fully.

## Output columns

| Column | Description |
|---|---|
| `source_term` | Extracted (and optionally cleaned) source term |
| `target_term` | Translation (AI step) |
| `term_type` | Part-of-speech based type |
| `ngram` / `characters` | Word count, or character count for Chinese and Japanese |
| `seg_count` | Number of distinct segments containing the term |
| `raw_freq` | Total occurrences |
| `sources` | Source file types the term came from |
| `definition` | Short definition (AI step) |
| `example_1..3` / `target_example_1..3` | Context examples, aligned with TMX targets when available |
| `notes` | AI notes, e.g. inconsistent translations in the TM |

## Supported languages

Catalan, Chinese, Croatian, Danish, Dutch, English, Finnish, French, German,
Greek, Italian, Japanese, Korean, Lithuanian, Macedonian, Norwegian Bokmål,
Polish, Portuguese, Romanian, Russian, Slovenian, Spanish, Swedish, Ukrainian.

Target languages for translation are free text, so any language the selected
AI model handles is accepted.

## Customising prompts

Edit `term_extractor_prompts.yaml` (it must stay next to the script). You can
rewrite the prompt text freely, but keep the top-level keys and the
`{{PLACEHOLDERS}}`; the file header explains which placeholders each prompt
uses.

## Privacy

With the AI steps enabled, candidate terms and example sentences from your
documents are sent to the selected provider (Anthropic or OpenAI). Do not
enable them for material you are not allowed to share with third-party
services. NLP-only runs are fully local, apart from the one-time model
download.

## License

This project is released under the [MIT License](LICENSE).

**spaCy models are licensed separately.** They are not included in this
repository: the tool downloads them from
[explosion/spacy-models](https://github.com/explosion/spacy-models) at runtime.
Licenses vary by language (MIT, CC BY-SA, GPL, and some **non-commercial**
CC BY-NC-SA). Check the license of each model you use on
[spacy.io/models](https://spacy.io/models), especially for commercial use.

---

Developed by [Creative Words](https://www.creative-words.com).
