#!/usr/bin/env python3
"""
generate_haiku_pages.py

Generates a complete static site from AI_Haiku.xlsx:
  poems/         one HTML page per poem
  poets/         one bio page per poet + one full poems-list page per poet
  photos/        place poet photos here as {poet-slug}.jpg
  index.html     master index (future)

Usage:
    python generate_haiku_pages.py
    python generate_haiku_pages.py --xlsx path/to/AI_Haiku.xlsx --out path/to/site

Poet bios:  place plain-text files in a bios/ folder next to this script,
            named {poet-slug}.txt  e.g.  bios/masaoka-shiki.txt
            Separate paragraphs with a blank line.

Poet photos: place image files in a photos/ folder next to this script,
             named {poet-slug}.jpg  e.g.  photos/masaoka-shiki.jpg
             Supported: .jpg .jpeg .png .webp
"""

import argparse
import math
import os
import re
import shutil
import unicodedata
import pandas as pd

# ---------------------------------------------------------------------------
# Shared assets
# ---------------------------------------------------------------------------

FONTS = (
    '<link rel="preconnect" href="https://fonts.googleapis.com">\n'
    '<link href="https://fonts.googleapis.com/css2?family=Noto+Serif+JP:wght@300;400'
    '&family=IM+Fell+English:ital@0;1'
    '&family=Cormorant+Garamond:ital,wght@0,300;0,400;0,600;1,300;1,400;1,600'
    '&display=swap" rel="stylesheet">'
)

BASE_CSS = """
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
html, body { background: #ede8df; width: 100%; min-height: 100vh; }
.page-bg { background: #ede8df; min-height: 100vh; width: 100%; padding: 3rem 1.5rem; }
a { color: inherit; }
.section-divider { display: flex; align-items: center; gap: 1rem; margin: 2.5rem 0 1.25rem; }
.section-label { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 10px; letter-spacing: 0.15em; text-transform: uppercase; color: #888; white-space: nowrap; }
.section-rule { flex: 1; height: 1px; background: #999; }
.ext-links { margin-top: 3rem; padding-top: 1.5rem; border-top: 1px solid #ccc8c0; display: flex; gap: 1.5rem; flex-wrap: wrap; }
.ext-link { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 11px; letter-spacing: 0.1em; text-transform: uppercase; color: #888; text-decoration: none; border-bottom: 1px solid #bbb; padding-bottom: 1px; }
.ext-link:hover { color: #444; }
"""

POEM_CSS = """
.haiku-page { font-family: "Cormorant Garamond", Georgia, serif; color: #1a1a1a; max-width: 560px; margin: 0 auto; }
.season-banner { display: flex; align-items: center; gap: 1rem; margin-bottom: 3rem; }
.season-pill { font-size: 11px; letter-spacing: 0.15em; text-transform: uppercase; font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; color: #666; border: 1px solid #999; padding: 4px 12px; border-radius: 6px; text-decoration: none; }
.season-line { flex: 1; height: 1px; background: #999; }
.season-word-jp { font-family: "Noto Serif JP", serif; font-size: 13px; color: #666; }
.poem-block { margin: 0 0 3.5rem; }
.poem-japanese { font-family: "Noto Serif JP", serif; font-size: 22px; font-weight: 300; letter-spacing: 0.12em; line-height: 2; margin: 0 0 0.5rem; color: #1a1a1a; }
.poem-romaji { font-size: 15px; font-style: italic; color: #666; letter-spacing: 0.05em; line-height: 1.6; font-family: "IM Fell English", Georgia, serif; }
.poem-translation { font-size: 30px; font-weight: 600; line-height: 1.5; color: #1a1a1a; margin: 1.5rem 0; }
.poem-translation span { display: block; padding-left: 96px; }
.prose-section { font-size: 17px; line-height: 1.75; color: #1a1a1a; margin-bottom: 2rem; font-weight: 300; }
.prose-section p { margin: 0 0 1rem; }
.prose-section p:last-child { margin-bottom: 0; }
.translation-notes { font-size: 17px; line-height: 1.75; color: #1a1a1a; margin-bottom: 2rem; font-weight: 300; }
.translation-notes p { margin: 0 0 1rem; }
.translation-notes p:last-child { margin-bottom: 0; }
.metadata-grid { display: grid; grid-template-columns: repeat(3, 1fr); margin-top: 3rem; border: 1px solid #999; border-radius: 8px; overflow: hidden; }
.meta-cell { padding: 12px 16px; border-right: 1px solid #999; border-bottom: 1px solid #999; background: #e6e0d6; }
.meta-cell:nth-child(3n) { border-right: none; }
.meta-cell:nth-child(n+4) { border-bottom: none; }
.meta-label { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 10px; letter-spacing: 0.12em; text-transform: uppercase; color: #888; margin-bottom: 5px; }
.meta-value { font-size: 15px; color: #1a1a1a; font-weight: 400; line-height: 1.3; }
.meta-value-jp { font-family: "Noto Serif JP", serif; font-size: 14px; color: #666; display: block; margin-top: 2px; }
.poet-link { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 11px; letter-spacing: 0.1em; text-transform: uppercase; color: #666; text-decoration: none; border-bottom: 1px solid #bbb; padding-bottom: 1px; }
.poet-link:hover { color: #1a1a1a; }
"""

POET_CSS = """
.poet-page { font-family: "Cormorant Garamond", Georgia, serif; color: #1a1a1a; max-width: 560px; margin: 0 auto; }
.page-label { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 10px; letter-spacing: 0.15em; text-transform: uppercase; color: #888; margin-bottom: 2.5rem; display: flex; align-items: center; gap: 1rem; }
.page-label-line { flex: 1; height: 1px; background: #999; }
.poet-intro { display: flex; gap: 1.75rem; align-items: flex-start; margin-bottom: 3rem; }
.poet-intro-left { flex: 1; min-width: 0; }
.poet-name-en { font-size: 36px; font-weight: 300; line-height: 1.2; color: #1a1a1a; margin-bottom: 0.4rem; }
.poet-name-row { display: flex; align-items: baseline; gap: 1rem; margin-bottom: 1rem; }
.poet-name-jp { font-family: "Noto Serif JP", serif; font-size: 18px; font-weight: 300; color: #666; }
.poet-dates { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 12px; color: #999; letter-spacing: 0.05em; }
.poet-meta-row { display: flex; gap: 0.5rem; flex-wrap: wrap; margin-bottom: 1.5rem; }
.poet-tag { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 10px; letter-spacing: 0.12em; text-transform: uppercase; color: #666; border: 1px solid #999; padding: 3px 10px; border-radius: 4px; }
.bio-text { font-size: 16px; line-height: 1.75; color: #1a1a1a; font-weight: 300; }
.bio-text p { margin: 0 0 1rem; }
.bio-text p:last-child { margin-bottom: 0; }
.poet-photo { flex-shrink: 0; width: 140px; }
.poet-photo img { width: 140px; display: block; border: 1px solid #ccc8c0; filter: grayscale(25%); }
.photo-caption { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 9px; letter-spacing: 0.05em; color: #aaa; margin-top: 6px; text-align: center; line-height: 1.4; }
.poem-entry { padding: 1.25rem 0; border-bottom: 1px solid #ccc8c0; text-decoration: none; display: block; color: inherit; }
.poem-entry:first-child { border-top: 1px solid #ccc8c0; }
.poem-entry:hover .poem-entry-jp { color: #1a1a1a; }
.poem-entry-jp { font-family: "Noto Serif JP", serif; font-size: 16px; font-weight: 300; letter-spacing: 0.1em; color: #444; margin-bottom: 0.3rem; }
.poem-entry-romaji { font-family: "IM Fell English", Georgia, serif; font-size: 13px; font-style: italic; color: #888; margin-bottom: 0.4rem; }
.poem-entry-translation { font-size: 15px; font-weight: 300; color: #555; line-height: 1.4; }
.poem-entry-season { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 10px; letter-spacing: 0.1em; text-transform: uppercase; color: #aaa; margin-top: 0.5rem; }
.see-all { display: inline-block; margin-top: 1.75rem; font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 11px; letter-spacing: 0.12em; text-transform: uppercase; color: #666; text-decoration: none; border-bottom: 1px solid #999; padding-bottom: 2px; }
.see-all:hover { color: #1a1a1a; }
"""

POEM_LIST_CSS = """
.poet-page { font-family: "Cormorant Garamond", Georgia, serif; color: #1a1a1a; max-width: 560px; margin: 0 auto; }
.page-label { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 10px; letter-spacing: 0.15em; text-transform: uppercase; color: #888; margin-bottom: 2.5rem; display: flex; align-items: center; gap: 1rem; }
.page-label-line { flex: 1; height: 1px; background: #999; }
.list-header { margin-bottom: 2.5rem; }
.list-header-name { font-size: 28px; font-weight: 300; color: #1a1a1a; margin-bottom: 0.25rem; }
.list-header-sub { font-family: "Noto Serif JP", serif; font-size: 16px; color: #888; }
.season-group { margin-bottom: 2rem; }
.season-heading { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 10px; letter-spacing: 0.2em; text-transform: uppercase; color: #aaa; margin-bottom: 0.5rem; padding-bottom: 0.5rem; border-bottom: 1px solid #ccc8c0; }
.poem-entry { padding: 1rem 0; border-bottom: 1px solid #e8e3d8; text-decoration: none; display: block; color: inherit; }
.poem-entry:hover .poem-entry-jp { color: #1a1a1a; }
.poem-entry-jp { font-family: "Noto Serif JP", serif; font-size: 15px; font-weight: 300; letter-spacing: 0.1em; color: #444; margin-bottom: 0.25rem; }
.poem-entry-romaji { font-family: "IM Fell English", Georgia, serif; font-size: 13px; font-style: italic; color: #888; margin-bottom: 0.25rem; }
.poem-entry-translation { font-size: 14px; font-weight: 300; color: #666; line-height: 1.4; }
.back-link { display: inline-block; margin-top: 2rem; font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 11px; letter-spacing: 0.12em; text-transform: uppercase; color: #666; text-decoration: none; border-bottom: 1px solid #999; padding-bottom: 2px; }
.back-link:hover { color: #1a1a1a; }
"""


SAIJIKI_CSS = """
.saijiki-page { font-family: "Cormorant Garamond", Georgia, serif; color: #1a1a1a; max-width: 560px; margin: 0 auto; }
.page-label { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 10px; letter-spacing: 0.15em; text-transform: uppercase; color: #888; margin-bottom: 2.5rem; display: flex; align-items: center; gap: 1rem; }
.page-label-line { flex: 1; height: 1px; background: #999; }
.kigo-header { margin-bottom: 3rem; }
.kigo-name-en { font-size: 38px; font-weight: 300; line-height: 1.2; color: #1a1a1a; margin-bottom: 0.4rem; }
.kigo-name-row { display: flex; align-items: baseline; gap: 1rem; margin-bottom: 1rem; }
.kigo-name-jp { font-family: "Noto Serif JP", serif; font-size: 20px; font-weight: 300; color: #666; }
.kigo-romaji { font-family: "IM Fell English", Georgia, serif; font-size: 16px; font-style: italic; color: #888; }
.kigo-meta-row { display: flex; gap: 0.5rem; flex-wrap: wrap; }
.kigo-tag { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 10px; letter-spacing: 0.12em; text-transform: uppercase; color: #666; border: 1px solid #999; padding: 3px 10px; border-radius: 4px; }
.essay-text { font-size: 17px; line-height: 1.85; color: #1a1a1a; font-weight: 300; }
.essay-text p { margin: 0 0 1.25rem; }
.essay-text p:last-child { margin-bottom: 0; }
.saijiki-poem { padding: 1.5rem 0; border-bottom: 1px solid #ccc8c0; }
.saijiki-poem:first-child { border-top: 1px solid #ccc8c0; }
.saijiki-poem-jp { font-family: "Noto Serif JP", serif; font-size: 20px; font-weight: 300; letter-spacing: 0.12em; line-height: 2; margin-bottom: 0.3rem; color: #1a1a1a; }
.saijiki-poem-romaji { font-family: "IM Fell English", Georgia, serif; font-size: 14px; font-style: italic; color: #888; margin-bottom: 0.75rem; }
.saijiki-poem-translation { font-size: 22px; font-weight: 600; line-height: 1.5; color: #1a1a1a; margin-bottom: 0.75rem; }
.saijiki-poem-translation span { display: block; padding-left: 60px; }
.saijiki-poem-poet { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 11px; letter-spacing: 0.1em; text-transform: uppercase; color: #888; display: flex; justify-content: space-between; align-items: center; margin-top: 0.5rem; }
.saijiki-poem-poet-jp { font-family: "Noto Serif JP", serif; font-size: 13px; text-transform: none; letter-spacing: 0.05em; color: #aaa; }
.saijiki-poem-link { font-size: 10px; color: #aaa; text-decoration: none; border-bottom: 1px solid #ccc; padding-bottom: 1px; }
.saijiki-poem-link:hover { color: #666; }
.see-all { display: inline-block; margin-top: 1.75rem; font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 11px; letter-spacing: 0.12em; text-transform: uppercase; color: #666; text-decoration: none; border-bottom: 1px solid #999; padding-bottom: 2px; }
.see-all:hover { color: #1a1a1a; }
"""

SAIJIKI_LIST_CSS = """
.saijiki-page { font-family: "Cormorant Garamond", Georgia, serif; color: #1a1a1a; max-width: 560px; margin: 0 auto; }
.page-label { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 10px; letter-spacing: 0.15em; text-transform: uppercase; color: #888; margin-bottom: 2.5rem; display: flex; align-items: center; gap: 1rem; }
.page-label-line { flex: 1; height: 1px; background: #999; }
.list-header { margin-bottom: 2.5rem; }
.list-header-name { font-size: 28px; font-weight: 300; color: #1a1a1a; margin-bottom: 0.25rem; }
.list-header-sub { font-family: "Noto Serif JP", serif; font-size: 16px; color: #888; }
.poem-entry { padding: 1rem 0; border-bottom: 1px solid #e8e3d8; text-decoration: none; display: block; color: inherit; }
.poem-entry:first-child { border-top: 1px solid #ccc8c0; }
.poem-entry:hover .poem-entry-jp { color: #1a1a1a; }
.poem-entry-jp { font-family: "Noto Serif JP", serif; font-size: 15px; font-weight: 300; letter-spacing: 0.1em; color: #444; margin-bottom: 0.25rem; }
.poem-entry-romaji { font-family: "IM Fell English", Georgia, serif; font-size: 13px; font-style: italic; color: #888; margin-bottom: 0.25rem; }
.poem-entry-translation { font-size: 14px; font-weight: 300; color: #666; line-height: 1.4; }
.poem-entry-poet { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 10px; letter-spacing: 0.08em; text-transform: uppercase; color: #aaa; margin-top: 0.4rem; }
.back-link { display: inline-block; margin-top: 2rem; font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 11px; letter-spacing: 0.12em; text-transform: uppercase; color: #666; text-decoration: none; border-bottom: 1px solid #999; padding-bottom: 2px; }
.back-link:hover { color: #1a1a1a; }
"""

SAIJIKI_INDEX_CSS = """
.saijiki-page { font-family: "Cormorant Garamond", Georgia, serif; color: #1a1a1a; max-width: 560px; margin: 0 auto; }
.page-label { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 10px; letter-spacing: 0.15em; text-transform: uppercase; color: #888; margin-bottom: 2.5rem; display: flex; align-items: center; gap: 1rem; }
.page-label-line { flex: 1; height: 1px; background: #999; }
.index-title { font-size: 36px; font-weight: 300; color: #1a1a1a; margin-bottom: 3rem; }
.season-group { margin-bottom: 3rem; }
.season-heading { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 11px; letter-spacing: 0.2em; text-transform: uppercase; color: #888; margin-bottom: 1rem; padding-bottom: 0.75rem; border-bottom: 1px solid #999; }
.kigo-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 1px; background: #ccc8c0; border: 1px solid #ccc8c0; border-radius: 6px; overflow: hidden; }
.kigo-entry { background: #ede8df; padding: 0.75rem 1rem; text-decoration: none; display: block; color: inherit; }
.kigo-entry:hover { background: #e6e0d6; }
.kigo-entry-en { font-size: 16px; font-weight: 300; color: #1a1a1a; margin-bottom: 0.2rem; }
.kigo-entry-jp { font-family: "Noto Serif JP", serif; font-size: 13px; color: #888; }
.kigo-entry-count { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 10px; color: #aaa; margin-top: 0.2rem; }
"""


SEASON_ORDER = ["Spring", "Summer", "Autumn", "Winter", "New Year"]

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def val(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    return str(v).strip()


def slugify(text, max_len=60):
    text = unicodedata.normalize("NFKD", text)
    text = text.encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^\w\s-]", "", text).strip().lower()
    text = re.sub(r"[\s_-]+", "-", text)
    return text[:max_len].rstrip("-") or "unknown"


def romaji_inline(raw):
    parts = [p.strip() for p in raw.replace("\\n", "\n").splitlines() if p.strip()]
    return " / ".join(parts)


def translation_inline(raw):
    lines = [l.strip() for l in raw.replace("\\n", "\n").splitlines() if l.strip()]
    return " / ".join(lines)


def translation_html(raw):
    lines = [l.strip() for l in raw.replace("\\n", "\n").splitlines() if l.strip()]
    return "\n".join(f'      <span>{l}</span>' for l in lines)


def kigo_parts(kigo_raw):
    raw = val(kigo_raw)
    if not raw:
        return "—", "—", ""
    parts = raw.split(",", 1)
    en = parts[0].strip() if parts else raw
    en_clean = re.sub(r"[^\x00-\x7F【】]+", "", en).strip()
    en_clean = re.sub(r"【[^】]*】", "", en_clean).strip()
    short = re.sub(r".*】\s*", "", en).strip() or en_clean
    jp_match = re.search(r"([\u4e00-\u9fff\u3040-\u30ff]+\s*【[^】]+】)", raw)
    jp = jp_match.group(1) if jp_match else ""
    return short or raw, en_clean or raw, jp


def build_section(label, text, cls="prose-section"):
    if not text:
        return ""
    paras = "\n".join(f"    <p>{p.strip()}</p>" for p in text.split("\n") if p.strip())
    return f"""  <div class="section-divider">
    <span class="section-label">{label}</span>
    <div class="section-rule"></div>
  </div>
  <div class="{cls}">
{paras}
  </div>
"""


def html_page(title, css, body):
    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{title}</title>
{FONTS}
<style>{BASE_CSS}{css}</style>
</head>
<body>
<div class="page-bg">
{body}
</div>
</body>
</html>"""


def load_bio(bios_dir, poet_slug):
    path = os.path.join(bios_dir, f"{poet_slug}.txt")
    if not os.path.exists(path):
        return ""
    with open(path, encoding="utf-8") as f:
        return f.read().strip()


def find_photo(photos_dir, poet_slug):
    """Return filename if a photo exists for this poet, else None."""
    for ext in ("jpg", "jpeg", "png", "webp"):
        fname = f"{poet_slug}.{ext}"
        if os.path.exists(os.path.join(photos_dir, fname)):
            return fname
    return None


def photo_block_html(photo_fname, poet_name, path_prefix=""):
    if not photo_fname:
        return ""
    src = f"{path_prefix}photos/{photo_fname}"
    return f"""    <div class="poet-photo">
      <img src="{src}" alt="{poet_name}">
      <div class="photo-caption">{poet_name}</div>
    </div>"""


def select_poems(poems, n=4):
    """Pick up to n poems with seasonal variety."""
    by_season = {}
    for p in poems:
        s = p.get("season") or ""
        by_season.setdefault(s, []).append(p)
    selected = []
    seasons = [s for s in SEASON_ORDER if s in by_season]
    # also include unclassified
    for s in by_season:
        if s not in seasons:
            seasons.append(s)
    pools = {s: list(v) for s, v in by_season.items()}
    while len(selected) < n:
        added = False
        for s in seasons:
            if pools.get(s) and len(selected) < n:
                selected.append(pools[s].pop(0))
                added = True
        if not added:
            break
    return selected


# ---------------------------------------------------------------------------
# Page builders
# ---------------------------------------------------------------------------

def build_poem_page(row, poem_filename, poet_slug):
    poem_jp   = val(row.get("Poem") or row.get("Text", ""))
    romaji_raw= val(row.get("Romaji", ""))
    trans_raw = val(row.get("My translation", ""))
    poet      = val(row.get("Poet", ""))
    poet_jp   = val(row.get("俳人", ""))
    season    = val(row.get("Season", ""))
    kigo_raw  = val(row.get("Kigo", ""))
    src       = val(row.get("Original source ", "")) or val(row.get("Modern Source", "")) or "—"
    notes     = val(row.get("Notes", ""))
    maegaki   = val(row.get("Maegaki", ""))

    kigo_short, kigo_en, kigo_jp = kigo_parts(kigo_raw)
    romaji = romaji_inline(romaji_raw)
    trans  = translation_html(trans_raw)

    commentary_block = build_section("Commentary", maegaki)
    notes_block      = build_section("Notes", notes, cls="translation-notes")
    poet_link = (f'<a class="poet-link" href="../poets/{poet_slug}.html">← {poet}</a>'
                 if poet_slug and poet else "")

    body = f"""<div class="haiku-page">

  <div class="season-banner">
    <span class="season-pill">{season or "—"}</span>
    <div class="season-line"></div>
    <span class="season-word-jp">{kigo_short}</span>
  </div>

  <div class="poem-block">
    <div class="poem-japanese">{poem_jp}</div>
    <div class="poem-romaji">{romaji}</div>
    <div class="poem-translation">
{trans}
    </div>
  </div>

{commentary_block}{notes_block}
  <div class="metadata-grid">
    <div class="meta-cell">
      <div class="meta-label">Season Word</div>
      <div class="meta-value">{kigo_en}<span class="meta-value-jp">{kigo_jp}</span></div>
    </div>
    <div class="meta-cell">
      <div class="meta-label">Season</div>
      <div class="meta-value">{season or "—"}</div>
    </div>
    <div class="meta-cell">
      <div class="meta-label">Cutting word</div>
      <div class="meta-value">—</div>
    </div>
    <div class="meta-cell">
      <div class="meta-label">Poet</div>
      <div class="meta-value">{poet or "—"}<span class="meta-value-jp">{poet_jp}</span></div>
    </div>
    <div class="meta-cell">
      <div class="meta-label">Date of composition</div>
      <div class="meta-value">—</div>
    </div>
    <div class="meta-cell">
      <div class="meta-label">Source</div>
      <div class="meta-value">{src}</div>
    </div>
  </div>

  <div class="ext-links">{poet_link}</div>

</div>"""

    title = f"{poem_jp[:20]} — {poet}" if poem_jp else poet
    return html_page(title, POEM_CSS, body)


def build_poet_bio_page(poet_slug, data, bio_text, photos_dir):
    poet_name = data["name"]
    poet_jp   = data["jp"]
    dates     = data["dates"]
    period    = data["period"]
    poems     = data["poems"]

    photo_fname = find_photo(photos_dir, poet_slug)
    photo_html  = photo_block_html(photo_fname, poet_name, path_prefix="../")

    tags_html = f'<span class="poet-tag">{period}</span>' if period else ""

    if bio_text:
        paras = [p.strip() for p in bio_text.split("\n\n") if p.strip()]
        bio_inner = "\n".join(f"        <p>{p}</p>" for p in paras)
    else:
        bio_inner = "        <p><em>No biography on file yet.</em></p>"

    selected = select_poems(poems, n=4)
    selected_html = ""
    for p in selected:
        selected_html += f"""    <a class="poem-entry" href="../poems/{p['filename']}">
      <div class="poem-entry-jp">{p['jp']}</div>
      <div class="poem-entry-romaji">{p['romaji']}</div>
      <div class="poem-entry-translation">{p['translation']}</div>
      <div class="poem-entry-season">{p['season']} · {p['kigo_short']}</div>
    </a>
"""

    total = len(poems)
    see_all = (f'  <a class="see-all" href="{poet_slug}-poems.html">'
               f'All {total} poems by {poet_name} →</a>') if total > 4 else ""

    wiki_slug = poet_name.replace(" ", "_")
    wiki_link = f'<a class="ext-link" href="https://en.wikipedia.org/wiki/{wiki_slug}">Wikipedia ↗</a>'

    body = f"""<div class="poet-page">

  <div class="page-label">
    Poet
    <div class="page-label-line"></div>
  </div>

  <div class="poet-intro">
    <div class="poet-intro-left">
      <div class="poet-name-en">{poet_name}</div>
      <div class="poet-name-row">
        <span class="poet-name-jp">{poet_jp}</span>
        <span class="poet-dates">{dates}</span>
      </div>
      <div class="poet-meta-row">
        {tags_html}
      </div>
      <div class="bio-text">
{bio_inner}
      </div>
    </div>
    {photo_html}
  </div>

  <div class="section-divider">
    <span class="section-label">Selected poems</span>
    <div class="section-rule"></div>
  </div>

{selected_html}
{see_all}

  <div class="ext-links">
    {wiki_link}
  </div>

</div>"""

    return html_page(poet_name, POET_CSS, body)


def build_poet_poems_page(poet_slug, data):
    poet_name = data["name"]
    poet_jp   = data["jp"]
    poems     = data["poems"]

    by_season = {}
    for p in poems:
        s = p.get("season") or "Unclassified"
        by_season.setdefault(s, []).append(p)

    groups_html = ""
    ordered_seasons = [s for s in SEASON_ORDER if s in by_season]
    for s in by_season:
        if s not in ordered_seasons:
            ordered_seasons.append(s)

    for season in ordered_seasons:
        group = by_season.get(season, [])
        entries = ""
        for p in group:
            entries += f"""      <a class="poem-entry" href="../poems/{p['filename']}">
        <div class="poem-entry-jp">{p['jp']}</div>
        <div class="poem-entry-romaji">{p['romaji']}</div>
        <div class="poem-entry-translation">{p['translation']}</div>
      </a>
"""
        groups_html += f"""  <div class="season-group">
    <div class="season-heading">{season}</div>
{entries}  </div>
"""

    total = len(poems)
    body = f"""<div class="poet-page">

  <div class="page-label">
    Poet · All poems
    <div class="page-label-line"></div>
  </div>

  <div class="list-header">
    <div class="list-header-name">{poet_name}</div>
    <div class="list-header-sub">{poet_jp} &ensp;·&ensp; {total} poems</div>
  </div>

{groups_html}
  <a class="back-link" href="{poet_slug}.html">← Back to {poet_name}</a>

</div>"""

    return html_page(f"All poems — {poet_name}", POEM_LIST_CSS, body)



# ---------------------------------------------------------------------------
# Saijiki page builders
# ---------------------------------------------------------------------------

def load_essay(essays_dir, kigo_slug):
    path = os.path.join(essays_dir, f"{kigo_slug}.txt")
    if not os.path.exists(path):
        return ""
    with open(path, encoding="utf-8") as f:
        return f.read().strip()


def translation_lines_html(raw, cls="saijiki-poem-translation"):
    lines = [l.strip() for l in raw.replace("\\n", "\n").splitlines() if l.strip()]
    spans = "\n".join(f'        <span>{l}</span>' for l in lines)
    return f'''    <div class="{cls}">
{spans}
    </div>'''


def build_saijiki_entry(kigo_slug, meta, essay_text, exemplars, all_poem_count):
    """
    meta: dict with kigo_en, kigo_jp, kigo_romaji, season, category
    exemplars: list of poem dicts sorted by Saijiki_Order
    all_poem_count: total poems tagged to this kigo
    """
    kigo_en     = meta.get("kigo_en", kigo_slug)
    kigo_jp     = meta.get("kigo_jp", "")
    kigo_romaji = meta.get("kigo_romaji", "")
    season      = meta.get("season", "")
    category    = meta.get("category", "")

    # Tags
    tags = ""
    if season:
        tags += f'<span class="kigo-tag">{season}</span>'
    if category:
        tags += f'  <span class="kigo-tag">{category}</span>'

    # Essay
    essay_html = ""
    if essay_text:
        paras = [p.strip() for p in essay_text.split("\n\n") if p.strip()]
        essay_html = f'''  <div class="section-divider">
    <span class="section-label">Essay</span>
    <div class="section-rule"></div>
  </div>
  <div class="essay-text">
{chr(10).join(f"    <p>{p}</p>" for p in paras)}
  </div>
'''

    # Exemplar poems
    poems_html = ""
    for p in exemplars:
        trans_lines = "\n".join(
            f'        <span>{l.strip()}</span>'
            for l in p["translation"].replace("\\n", "\n").splitlines()
            if l.strip()
        )
        poet_jp_span = f'<span class="saijiki-poem-poet-jp">{p["poet_jp"]}</span>' if p.get("poet_jp") else ""
        poem_link = f'<a class="saijiki-poem-link" href="../../poems/{p["filename"]}">Full poem →</a>' if p.get("filename") else ""
        poems_html += f'''    <div class="saijiki-poem">
      <div class="saijiki-poem-jp">{p["jp"]}</div>
      <div class="saijiki-poem-romaji">{p["romaji"]}</div>
      <div class="saijiki-poem-translation">
{trans_lines}
      </div>
      <div class="saijiki-poem-poet">
        <span>{p["poet"]} {poet_jp_span}</span>
        {poem_link}
      </div>
    </div>
'''

    see_all = ""
    if all_poem_count > len(exemplars):
        see_all = f'  <a class="see-all" href="{kigo_slug}-poems.html">All {all_poem_count} poems with this season word →</a>'

    body = f'''<div class="saijiki-page">

  <div class="page-label">
    Saijiki · {season}
    <div class="page-label-line"></div>
  </div>

  <div class="kigo-header">
    <div class="kigo-name-en">{kigo_en}</div>
    <div class="kigo-name-row">
      <span class="kigo-name-jp">{kigo_jp}</span>
      <span class="kigo-romaji">{kigo_romaji}</span>
    </div>
    <div class="kigo-meta-row">
      {tags}
    </div>
  </div>

{essay_html}
  <div class="section-divider">
    <span class="section-label">Exemplar poems</span>
    <div class="section-rule"></div>
  </div>

{poems_html}
{see_all}

</div>'''

    return html_page(f"{kigo_en} — Saijiki", SAIJIKI_CSS, body)


def build_saijiki_poems_list(kigo_slug, meta, all_poems):
    """All poems tagged to this kigo, flat list."""
    kigo_en = meta.get("kigo_en", kigo_slug)
    kigo_jp = meta.get("kigo_jp", "")
    season  = meta.get("season", "")

    entries = ""
    for p in all_poems:
        entries += f'''    <a class="poem-entry" href="../../poems/{p["filename"]}">
      <div class="poem-entry-jp">{p["jp"]}</div>
      <div class="poem-entry-romaji">{p["romaji"]}</div>
      <div class="poem-entry-translation">{p["translation"]}</div>
      <div class="poem-entry-poet">{p["poet"]}</div>
    </a>
'''

    body = f'''<div class="saijiki-page">

  <div class="page-label">
    Saijiki · {season}
    <div class="page-label-line"></div>
  </div>

  <div class="list-header">
    <div class="list-header-name">{kigo_en}</div>
    <div class="list-header-sub">{kigo_jp} &ensp;·&ensp; {len(all_poems)} poems</div>
  </div>

{entries}
  <a class="back-link" href="{kigo_slug}.html">← Back to {kigo_en}</a>

</div>'''

    return html_page(f"All poems — {kigo_en}", SAIJIKI_LIST_CSS, body)


def build_saijiki_index(saijiki_data):
    """Master index of all kigo entries grouped by season."""
    by_season = {}
    for slug, meta in saijiki_data.items():
        s = meta.get("season", "Unclassified")
        by_season.setdefault(s, []).append((slug, meta))

    groups_html = ""
    ordered = [s for s in SEASON_ORDER if s in by_season]
    for s in by_season:
        if s not in ordered:
            ordered.append(s)

    for season in ordered:
        entries = by_season.get(season, [])
        entries.sort(key=lambda x: x[1].get("kigo_en", ""))
        grid = ""
        for slug, meta in entries:
            count = meta.get("poem_count", 0)
            grid += f'''      <a class="kigo-entry" href="saijiki/{season.lower()}/{slug}.html">
        <div class="kigo-entry-en">{meta.get("kigo_en", slug)}</div>
        <div class="kigo-entry-jp">{meta.get("kigo_jp", "")}</div>
        <div class="kigo-entry-count">{count} poem{"s" if count != 1 else ""}</div>
      </a>
'''
        groups_html += f'''  <div class="season-group">
    <div class="season-heading">{season}</div>
    <div class="kigo-grid">
{grid}    </div>
  </div>
'''

    body = f'''<div class="saijiki-page">

  <div class="page-label">
    Saijiki
    <div class="page-label-line"></div>
  </div>

  <div class="index-title">Season Word Index</div>

{groups_html}
</div>'''

    return html_page("Saijiki — Season Word Index", SAIJIKI_INDEX_CSS, body)

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--xlsx",   default="/mnt/project/AI_Haiku.xlsx")
    parser.add_argument("--out",    default="/mnt/user-data/outputs/haiku_site")
    parser.add_argument("--bios",   default="/home/claude/bios")
    parser.add_argument("--photos", default="/home/claude/photos")
    parser.add_argument("--essays", default="/home/claude/essays")
    parser.add_argument("--sheets", nargs="+", default=["Edo", "Modern"])
    args = parser.parse_args()

    poems_dir  = os.path.join(args.out, "poems")
    poets_dir  = os.path.join(args.out, "poets")
    photos_out = os.path.join(args.out, "photos")
    for d in (poems_dir, poets_dir, photos_out):
        os.makedirs(d, exist_ok=True)

    # Copy photos into site output
    if os.path.isdir(args.photos):
        for f in os.listdir(args.photos):
            shutil.copy2(os.path.join(args.photos, f), os.path.join(photos_out, f))

    # Read all rows
    all_rows = []
    for sheet in args.sheets:
        try:
            df = pd.read_excel(args.xlsx, sheet_name=sheet)
            for _, row in df.iterrows():
                all_rows.append((sheet, row.to_dict()))
        except Exception as e:
            print(f"  Skipping sheet '{sheet}': {e}")

    # Pass 1 — write poem pages, accumulate per-poet data
    slugs_used  = {}
    poets_data  = {}
    poem_count  = 0

    for sheet, row in all_rows:
        poem_jp    = val(row.get("Poem") or row.get("Text", ""))
        romaji_raw = val(row.get("Romaji", ""))
        trans_raw  = val(row.get("My translation", ""))
        poet       = val(row.get("Poet", ""))
        poet_jp    = val(row.get("俳人", ""))
        dates      = val(row.get("Dates", ""))
        season     = val(row.get("Season", ""))
        kigo_raw   = val(row.get("Kigo", ""))

        if not poem_jp and not romaji_raw:
            continue

        # Unique poem slug
        base = slugify(romaji_raw[:40] if romaji_raw else poem_jp)
        if base in slugs_used:
            slugs_used[base] += 1
            slug = f"{base}-{slugs_used[base]}"
        else:
            slugs_used[base] = 0
            slug = base
        poem_filename = f"{sheet.lower()}-{slug}.html"

        poet_slug = slugify(poet) if poet else "unknown"

        # Write poem page
        html = build_poem_page(row, poem_filename, poet_slug)
        with open(os.path.join(poems_dir, poem_filename), "w", encoding="utf-8") as f:
            f.write(html)
        poem_count += 1

        # Accumulate poet
        kigo_short, _, _ = kigo_parts(kigo_raw)
        if poet_slug not in poets_data:
            poets_data[poet_slug] = {
                "name":   poet or "Unknown",
                "jp":     poet_jp,
                "dates":  dates,
                "period": sheet,
                "poems":  [],
            }
        poets_data[poet_slug]["poems"].append({
            "jp":          poem_jp,
            "romaji":      romaji_inline(romaji_raw),
            "translation": translation_inline(trans_raw),
            "season":      season,
            "kigo_short":  kigo_short,
            "filename":    poem_filename,
        })

    # Also accumulate saijiki data per row
    # (done in pass 1 loop — saijiki_poems collected alongside poet data)

    print(f"  Poems: {poem_count} pages written")

    # Pass 2 — poet bio pages and poem-list pages
    poet_count = 0
    for poet_slug, data in poets_data.items():
        bio_text = load_bio(args.bios, poet_slug)

        bio_html = build_poet_bio_page(poet_slug, data, bio_text, args.photos)
        with open(os.path.join(poets_dir, f"{poet_slug}.html"), "w", encoding="utf-8") as f:
            f.write(bio_html)

        list_html = build_poet_poems_page(poet_slug, data)
        with open(os.path.join(poets_dir, f"{poet_slug}-poems.html"), "w", encoding="utf-8") as f:
            f.write(list_html)

        poet_count += 1

    print(f"  Poets: {poet_count} bio pages + {poet_count} poem-list pages")

    # Pass 3 — saijiki pages
    # Load Saijiki metadata sheet if present
    saijiki_meta = {}
    try:
        df_s = pd.read_excel(args.xlsx, sheet_name="Saijiki")
        for _, row in df_s.iterrows():
            slug = val(row.get("Kigo_Slug", ""))
            if not slug:
                continue
            saijiki_meta[slug] = {
                "kigo_en":     val(row.get("Kigo_Slug", slug)).replace("-", " ").title(),
                "kigo_jp":     val(row.get("Kigo_JP", "")),
                "kigo_romaji": val(row.get("Kigo_Romaji", "")),
                "season":      val(row.get("Season", "")),
                "category":    val(row.get("Category", "")),
            }
    except Exception:
        pass  # Sheet doesn't exist yet — metadata comes from poems only

    # Build saijiki entries from tagged poems
    # saijiki_poems[kigo_slug] = list of poem dicts with order
    saijiki_poems = {}
    for sheet, row in all_rows:
        entry_slug = val(row.get("Saijiki_Entry", ""))
        if not entry_slug:
            continue
        try:
            order = int(float(row.get("Saijiki_Order", 999) or 999))
        except (ValueError, TypeError):
            order = 999

        poem_jp   = val(row.get("Poem") or row.get("Text", ""))
        romaji_raw= val(row.get("Romaji", ""))
        trans_raw = val(row.get("My translation", ""))
        poet      = val(row.get("Poet", ""))
        poet_jp   = val(row.get("俳人", ""))
        season    = val(row.get("Season", ""))
        kigo_raw  = val(row.get("Kigo", ""))

        # Build poem filename to link back
        base = slugify(romaji_raw[:40] if romaji_raw else poem_jp)
        poem_filename = f"{sheet.lower()}-{base}.html"

        # Auto-populate meta from poem data if not in Saijiki sheet
        if entry_slug not in saijiki_meta:
            kigo_short, kigo_en_auto, kigo_jp_auto = kigo_parts(kigo_raw)
            saijiki_meta[entry_slug] = {
                "kigo_en":     kigo_short or entry_slug.replace("-", " ").title(),
                "kigo_jp":     kigo_jp_auto,
                "kigo_romaji": entry_slug,
                "season":      season,
                "category":    "",
            }

        saijiki_poems.setdefault(entry_slug, []).append({
            "order":       order,
            "jp":          poem_jp,
            "romaji":      romaji_inline(romaji_raw),
            "translation": trans_raw,
            "poet":        poet,
            "poet_jp":     poet_jp,
            "filename":    poem_filename,
        })

    # Write saijiki pages
    saijiki_count = 0
    for kigo_slug, poems in saijiki_poems.items():
        meta = saijiki_meta.get(kigo_slug, {})
        season = meta.get("season", "unknown").lower()

        season_dir = os.path.join(args.out, "saijiki", season)
        os.makedirs(season_dir, exist_ok=True)

        # Sort all poems by order
        poems_sorted = sorted(poems, key=lambda p: p["order"])

        # Exemplars = poems with explicit order (order < 900)
        exemplars = [p for p in poems_sorted if p["order"] < 900]
        if not exemplars:
            exemplars = poems_sorted[:4]

        essay_text = load_essay(args.essays, kigo_slug)

        # Entry page
        entry_html = build_saijiki_entry(
            kigo_slug, meta, essay_text, exemplars, len(poems_sorted)
        )
        with open(os.path.join(season_dir, f"{kigo_slug}.html"), "w", encoding="utf-8") as f:
            f.write(entry_html)

        # All-poems list page
        list_html = build_saijiki_poems_list(kigo_slug, meta, poems_sorted)
        with open(os.path.join(season_dir, f"{kigo_slug}-poems.html"), "w", encoding="utf-8") as f:
            f.write(list_html)

        # Store poem count for index
        saijiki_meta[kigo_slug]["poem_count"] = len(poems_sorted)
        saijiki_count += 1

    # Saijiki index page
    if saijiki_count:
        index_html = build_saijiki_index(saijiki_meta)
        with open(os.path.join(args.out, "saijiki-index.html"), "w", encoding="utf-8") as f:
            f.write(index_html)
        print(f"  Saijiki: {saijiki_count} entries ({saijiki_count * 2} pages) + index")
    else:
        print(f"  Saijiki: no entries found (add Saijiki_Entry column to spreadsheet)")

    print(f"\nDone — site written to: {args.out}")
    print(f"\n  poems/          {poem_count} poem pages")
    print(f"  poets/          {poet_count * 2} poet pages")
    print(f"  saijiki/        {saijiki_count * 2} saijiki pages + index")
    print(f"\nSource folders (next to this script):")
    print(f"  bios/           poet biographies  ({{slug}}.txt)")
    print(f"  essays/         saijiki essays    ({{kigo-slug}}.txt)")
    print(f"  photos/         poet photos       ({{slug}}.jpg)")


if __name__ == "__main__":
    main()