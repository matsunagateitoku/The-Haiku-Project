# Project Context: Japanese Haiku Archive & Saijiki

## What this project is

A static haiku website (hosted on GitHub, generated via `generate_haiku_pages.py`)
built from two data sources:

1. **AI_Haiku.xlsx** — central poem database. Two sheets:
   - `Edo` sheet (300+ poems, Japanese column header: `Poem`)
   - `Modern` sheet (500+ poems, Japanese column header: `Text`)
   - Both share columns: `Romaji`, `My translation`, `俳人`, `Poet`, `Dates`,
     `Season`, `Kigo`, `Original source ` (NOTE: trailing space in header,
     must match exactly or strip it), `Modern Source`, `Other translation`,
     `Notes`, `Maegaki`

2. **English-language saijiki** (seasonal almanac) — structured entries, one
   per kigo (seasonal word), each with:
   - The kigo term (Japanese, hiragana/romaji, English gloss)
   - A prose description (etymology, usage, cultural context)
   - Exactly 4 curated example poems

## Site structure

- `generate_haiku_pages.py` — run from project root, generates individual poem
  pages, poet bio pages, saijiki entry pages, and a saijiki index
- `bios/{slug}.txt` — poet bio text files
- `essays/{kigo-slug}.txt` — saijiki prose entries (THIS IS WHAT WE'RE EDITING)
- `photos/{slug}.jpg` — associated photos
- `site/` — generated output, gitignored, always regenerated from source, never
  hand-edited
- **Known bug**: "New Year" as a season value creates a directory with a space
  in the path (`new year/`) — worth fixing or working around when touched

## Editing workflow: direct markdown authoring (no Google Docs)

We considered and dropped a Google-Docs-based workflow (author in Docs, export,
run a splitter script to generate per-entry files). That added a translation
layer with no real benefit once Claude Code is doing the editing — it meant an
export step, a parsing step, and a risk of drift between what's in the Doc and
what's published. Decision: **author the saijiki entries directly as markdown/
text files in this repo.** No splitter, no export/import cycle.

### How entries are organized
- 600+ entries across 5 seasons (spring, summer, autumn, winter, new year)
- Each entry ultimately becomes its own `essays/{kigo-slug}.txt` for the site
  generator to pick up
- Still deciding the best *authoring* grouping (one file per entry vs. one
  file per season with clear headings vs. something else) — if one-file-per-
  entry turns out to be annoying to navigate across 600+ files, consider
  season-level files with a consistent `## kigo-term — kanji (romaji)` heading
  convention, and a small script to split those into `essays/` on demand. But
  default to trying direct per-entry files first now that Claude Code can
  navigate/search across many files easily — that was the main pain point
  when this was manual.

### Each entry's content structure
- Kigo term as heading: Japanese, then hiragana/romaji grouped in parentheses
  with forward slash, then English gloss — e.g. `雪仏 (ゆきぼとけ / yukibotoke)
  snow buddha`
- Prose section (etymology, usage, cultural notes)
- Exactly 4 curated example poems

### Spellcheck / editing experience
Editing markdown directly loses Google's native spellcheck. Recommend
installing the "Code Spell Checker" VS Code extension in the Codespace to
cover this gap.

### Status tracking
Still an open decision — options discussed:
- Add a `Status` column to the existing `JW1_kigo_full_list.md` (643-entry
  authoritative index) — keeps one source of truth instead of two systems
  that can drift.
- A separate tracker file/sheet (CSV or new tab in AI_Haiku.xlsx) with one
  row per entry: slug, season, status, last-updated.
Because entries now live in git-tracked files, a lightweight completeness
script (e.g., flag entries under N words, or missing 4 poems) may eventually
replace manual status tracking entirely — worth considering once enough
entries exist to test against.

### Data integrity principles (apply project-wide)
- Never silently correct anomalies in source material — flag them instead.
- If a numbering/count check doesn't add up, look for dropped entries before
  concluding the source data itself is wrong.
- Preserve Japanese text exactly; don't transliterate or normalize unless
  explicitly asked.

## Related, not-yet-started task: featured poems
Each saijiki entry pairs with exactly 4 curated poems, but the site also
displays *all* poems matching that season/kigo (pulled live from
AI_Haiku.xlsx). Plan: show the 4 curated poems first, with a way to reveal the
rest (toggle/button), likely implemented via a new `Featured order` column in
AI_Haiku.xlsx keyed to a stable poem ID. Not yet built — don't start this
without confirming first.

## My working style
I dictate direction conversationally and expect you to synthesize and propose
concrete next steps rather than asking open-ended questions. I catch errors
through close reading — if I flag a mistake, incorporate the fix precisely
without needing it re-explained. I prefer incremental, step-by-step builds I
can check at each stage over one big untested script.
