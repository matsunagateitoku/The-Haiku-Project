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
.meta-cell.span2 { grid-column: span 2; }
.meta-cell.no-right { border-right: none; }
.meta-cell.no-bottom { border-bottom: none; }
.meta-label { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 10px; letter-spacing: 0.12em; text-transform: uppercase; color: #888; margin-bottom: 5px; }
.meta-value { font-size: 15px; color: #1a1a1a; font-weight: 400; line-height: 1.3; }
.meta-value-jp { font-family: "Noto Serif JP", serif; font-size: 14px; color: #666; display: block; margin-top: 2px; }
.poet-link { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 11px; letter-spacing: 0.1em; text-transform: uppercase; color: #666; text-decoration: none; border-bottom: 1px solid #bbb; padding-bottom: 1px; }
.poet-link:hover { color: #1a1a1a; }
.translation-attribution { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 11px; letter-spacing: 0.1em; text-transform: uppercase; color: #aaa; margin-top: 0.75rem; }
.info-box { border: 1px solid #ccc8c0; border-radius: 6px; overflow: hidden; margin: 1.75rem 0; }
.info-box-label { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 10px; letter-spacing: 0.15em; text-transform: uppercase; color: #888; padding: 7px 14px; background: #e6e0d6; border-bottom: 1px solid #ccc8c0; }
.info-box-body { padding: 1rem 1.25rem; font-size: 16px; line-height: 1.75; color: #1a1a1a; font-weight: 300; }
.info-box-body p { margin: 0 0 0.75rem; }
.info-box-body p:last-child { margin-bottom: 0; }
.info-box-translation { padding: 1rem 1.25rem; font-size: 22px; font-weight: 400; line-height: 1.5; color: #1a1a1a; }
.info-box-translation span { display: block; padding-left: 60px; }
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
.poem-entry-translation { font-size: 14px; font-weight: 300; color: #666; line-height: 1.5; }
.poem-entry-translation span { display: block; }
.poem-entry-poet { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 10px; letter-spacing: 0.08em; text-transform: uppercase; color: #aaa; margin-top: 0.4rem; }
.back-link { display: inline-block; margin-top: 2rem; font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 11px; letter-spacing: 0.12em; text-transform: uppercase; color: #666; text-decoration: none; border-bottom: 1px solid #999; padding-bottom: 2px; }
.back-link:hover { color: #1a1a1a; }
"""

SAIJIKI_INDEX_CSS = """
.saijiki-page { font-family: "Cormorant Garamond", Georgia, serif; color: #1a1a1a; max-width: 560px; margin: 0 auto; }
.page-label { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 10px; letter-spacing: 0.15em; text-transform: uppercase; color: #888; margin-bottom: 2.5rem; display: flex; align-items: center; gap: 1rem; }
.page-label-line { flex: 1; height: 1px; background: #999; }
.index-title { font-size: 36px; font-weight: 300; color: #1a1a1a; margin-bottom: 1.5rem; }
.sort-toggle { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 11px; letter-spacing: 0.1em; text-transform: uppercase; color: #888; margin-bottom: 3rem; display: flex; align-items: center; gap: 0.6rem; }
.sort-toggle a { color: #666; text-decoration: none; border-bottom: 1px solid #999; padding-bottom: 1px; }
.sort-toggle a:hover { color: #1a1a1a; }
.sort-toggle .sort-active { color: #1a1a1a; border-bottom: 1px solid #1a1a1a; padding-bottom: 1px; }
.sort-toggle .sort-sep { color: #ccc; }
.season-group { margin-bottom: 3rem; }
.season-heading { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 11px; letter-spacing: 0.2em; text-transform: uppercase; color: #888; margin-bottom: 1rem; padding-bottom: 0.75rem; border-bottom: 1px solid #999; }
.category-group { margin-bottom: 1.75rem; }
.category-group:last-child { margin-bottom: 0; }
.category-heading { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 10px; letter-spacing: 0.15em; text-transform: uppercase; color: #aaa; margin-bottom: 0.6rem; }
.kigo-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 1px; background: #ccc8c0; border: 1px solid #ccc8c0; border-radius: 6px; overflow: hidden; }
.kigo-entry { background: #ede8df; padding: 0.75rem 1rem; text-decoration: none; display: block; color: inherit; }
.kigo-entry:hover { background: #e6e0d6; }
.kigo-entry-en { font-size: 16px; font-weight: 300; color: #1a1a1a; margin-bottom: 0.2rem; }
.kigo-entry-jp { font-family: "Noto Serif JP", serif; font-size: 13px; color: #888; }
.kigo-entry-count { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 10px; color: #aaa; margin-top: 0.2rem; }
"""

INDEX_CSS = """
.index-page { font-family: "Cormorant Garamond", Georgia, serif; color: #1a1a1a; max-width: 560px; margin: 0 auto; }
.index-title { font-size: 42px; font-weight: 300; color: #1a1a1a; margin-bottom: 0.5rem; }
.index-subtitle { font-family: "Noto Serif JP", serif; font-size: 18px; color: #888; margin-bottom: 3rem; }
.index-nav { display: flex; flex-direction: column; gap: 1px; background: #ccc8c0; border: 1px solid #ccc8c0; border-radius: 8px; overflow: hidden; margin-bottom: 3rem; }
.index-nav-item { background: #ede8df; padding: 1.25rem 1.5rem; text-decoration: none; display: flex; justify-content: space-between; align-items: center; color: inherit; }
.index-nav-item:hover { background: #e6e0d6; }
.index-nav-label { font-size: 20px; font-weight: 300; color: #1a1a1a; }
.index-nav-sub { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 11px; letter-spacing: 0.1em; text-transform: uppercase; color: #aaa; margin-top: 0.25rem; }
.index-nav-arrow { font-size: 18px; color: #ccc; }
"""

POETS_INDEX_CSS = """
.poets-page { font-family: "Cormorant Garamond", Georgia, serif; color: #1a1a1a; max-width: 560px; margin: 0 auto; }
.page-label { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 10px; letter-spacing: 0.15em; text-transform: uppercase; color: #888; margin-bottom: 2.5rem; display: flex; align-items: center; gap: 1rem; }
.page-label-line { flex: 1; height: 1px; background: #999; }
.poets-title { font-size: 36px; font-weight: 300; color: #1a1a1a; margin-bottom: 1.5rem; }
.sort-controls { display: flex; align-items: center; gap: 0.5rem; margin-bottom: 2.5rem; flex-wrap: wrap; }
.sort-label { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 10px; letter-spacing: 0.12em; text-transform: uppercase; color: #aaa; margin-right: 0.25rem; }
.sort-btn { background: none; border: 1px solid transparent; border-radius: 4px; cursor: pointer; padding: 3px 10px; font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 10px; letter-spacing: 0.12em; text-transform: uppercase; color: #888; }
.sort-btn:hover { color: #1a1a1a; border-color: #ccc8c0; }
.sort-btn.active { color: #1a1a1a; border-color: #999; }
.alpha-group { margin-bottom: 2.5rem; }
.alpha-heading { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 11px; letter-spacing: 0.2em; text-transform: uppercase; color: #888; margin-bottom: 0.75rem; padding-bottom: 0.75rem; border-bottom: 1px solid #999; }
.poet-list { display: flex; flex-direction: column; gap: 1px; background: #ccc8c0; border: 1px solid #ccc8c0; border-radius: 6px; overflow: hidden; }
.poet-list-item { background: #ede8df; padding: 0.85rem 1rem; text-decoration: none; display: flex; justify-content: space-between; align-items: baseline; color: inherit; }
.poet-list-item:hover { background: #e6e0d6; }
.poet-list-name { font-size: 17px; font-weight: 300; color: #1a1a1a; }
.poet-list-right { display: flex; gap: 1rem; align-items: baseline; }
.poet-list-jp { font-family: "Noto Serif JP", serif; font-size: 13px; color: #888; }
.poet-list-dates { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 11px; color: #aaa; }
"""

NAV_CSS = """
.site-nav { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; font-size: 11px; letter-spacing: 0.12em; text-transform: uppercase; padding: 0.9rem 0 1.5rem; display: flex; align-items: center; gap: 1.5rem; }
.site-nav-home { color: #888; text-decoration: none; border-bottom: 1px solid #ccc8c0; padding-bottom: 1px; }
.site-nav-home:hover { color: #1a1a1a; }
.site-nav-rule { flex: 1; height: 1px; background: #ccc8c0; }
"""

SEASON_ORDER = ["Spring", "Summer", "Autumn", "Winter", "New Year"]
CATEGORY_ORDER = ["Season", "Celestial", "Landscape", "Life", "Animals", "Plants", "Observance"]

def sorted_by_category(entries):
    """Group (slug, meta) entries by Category, in CATEGORY_ORDER, each
    group internally in Saijiki-sheet order (meta['seq']); unrecognized or
    blank categories go last, grouped under 'Other'."""
    by_cat = {}
    for slug, meta in entries:
        cat = meta.get("category") or "Other"
        by_cat.setdefault(cat, []).append((slug, meta))
    order = [c for c in CATEGORY_ORDER if c in by_cat]
    order += [c for c in by_cat if c not in order]
    for cat in order:
        by_cat[cat].sort(key=lambda x: x[1].get("seq", 999999))
    return [(cat, by_cat[cat]) for cat in order]


def sorted_alpha(entries):
    return sorted(entries, key=lambda x: (x[1].get("kigo_en") or "").lower())

ALT_TRANSLATION_COLS = [
    "Ueda Haiku", "Keene World", "Keene Dawn", "Blyth Haiku", "Blyth History",
    "Fay Aoyagi", "English Saijiki", "21 Century", "Other translation",
]

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


def get_alt_translations(row):
    results = []
    for col in ALT_TRANSLATION_COLS:
        v = val(row.get(col, ""))
        if v:
            results.append((col, v))
    return results


JAPANESE_SOURCE_COLS = [
    "NKBZ", "俳句大観", "名歌名句辞典", "俳句の歴史",
    "現代俳句", "自選自解", "蝸牛", "古典俳句",
]


def get_japanese_source(row):
    for col in JAPANESE_SOURCE_COLS:
        v = val(row.get(col, ""))
        if v:
            return col, v
    return "", ""


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


def kigo_leading_kanji(kigo_raw):
    """Extract the season-word term itself (leading CJK run) from a poem's
    free-text Kigo field, e.g. '麦の秋 mugi no aki, wheat's autumn...' -> '麦の秋'.
    Used to match a poem against the Saijiki registry by Kigo_JP."""
    raw = val(kigo_raw)
    if not raw:
        return ""
    m = re.match(r"[一-鿿぀-ヿー]+", raw.strip())
    return m.group(0) if m else ""


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


def html_page(title, css, body, home_href="../../index.html"):
    nav = f'''<nav class="site-nav"><a class="site-nav-home" href="{home_href}">The Haiku Project</a><div class="site-nav-rule"></div></nav>'''
    favicon_href = home_href.replace("index.html", "favicon.svg")
    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{title}</title>
<link rel="icon" type="image/svg+xml" href="{favicon_href}">
{FONTS}
<style>{BASE_CSS}{NAV_CSS}{css}</style>
</head>
<body>
<div class="page-bg">
{nav}
{body}
</div>
</body>
</html>"""


def load_canonical_poets(poets_path):
    """Load canonical poet list and aliases from poets_canonical.xlsx.
    Returns (canonical_poets dict slug→data, aliases dict alias→canonical_slug).
    """
    canonical_poets = {}
    aliases = {}
    if not poets_path or not os.path.exists(poets_path):
        return canonical_poets, aliases
    try:
        xl = pd.ExcelFile(poets_path)
        for _, r in xl.parse('Poets').iterrows():
            slug = val(r.get('Slug', ''))
            if not slug:
                continue
            canonical_poets[slug] = {
                'name':   val(r.get('Display_Name', '')),
                'jp':     val(r.get('Kanji', '')),
                'dates':  val(r.get('Dates', '')),
                'school': val(r.get('School', '')),
                'period': val(r.get('Period', '')),
            }
        for _, r in xl.parse('Aliases').iterrows():
            a = val(r.get('Alias_Slug', ''))
            c = val(r.get('Canonical_Slug', ''))
            if a and c:
                aliases[a] = c
    except Exception as e:
        print(f"  Warning: could not load canonical poets: {e}")
    return canonical_poets, aliases


def resolve_poet_slug(raw_slug, aliases, canonical_poets):
    """Resolve a poet slug to its canonical form via the aliases table."""
    if raw_slug in canonical_poets:
        return raw_slug
    if raw_slug in aliases:
        target = aliases[raw_slug]
        if target in canonical_poets:
            return target
    return raw_slug


def load_bio(bios_dir, poet_slug):
    path = os.path.join(bios_dir, f"{poet_slug}.txt")
    if not os.path.exists(path):
        return ""
    with open(path, encoding="utf-8") as f:
        return f.read().strip()


def load_appreciation(appreciations_dir, poem_slug):
    path = os.path.join(appreciations_dir, f"{poem_slug}.txt")
    if not os.path.exists(path):
        return ""
    with open(path, encoding="utf-8") as f:
        return f.read().strip()


def find_photo(photos_dir, poet_slug):
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
    by_season = {}
    for p in poems:
        s = p.get("season") or ""
        by_season.setdefault(s, []).append(p)
    selected = []
    seasons = [s for s in SEASON_ORDER if s in by_season]
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

def build_poem_page(row, poem_filename, poet_slug, saijiki_slug="", saijiki_season="", appreciation="", canonical_name=""):
    poem_jp   = val(row.get("Poem") or row.get("Text", ""))
    romaji_raw= val(row.get("Romaji", ""))
    poet      = val(row.get("Poet", ""))
    poet_jp   = val(row.get("俳人", ""))
    season    = val(row.get("Season", ""))
    kigo_raw  = val(row.get("Kigo", ""))
    src       = val(row.get("Original source ", "")) or val(row.get("Modern Source", "")) or "—"
    notes     = val(row.get("Notes", ""))
    maegaki   = val(row.get("Maegaki", ""))

    my_trans  = val(row.get("My translation", ""))
    alt_trans = get_alt_translations(row)

    if my_trans:
        trans_raw       = my_trans
        main_translator = ""
        other_pair      = alt_trans[0] if alt_trans else None
    elif alt_trans:
        main_translator, trans_raw = alt_trans[0]
        other_pair      = alt_trans[1] if len(alt_trans) > 1 else None
    else:
        trans_raw       = ""
        main_translator = ""
        other_pair      = None

    kigo_short, kigo_en, kigo_jp = kigo_parts(kigo_raw)
    romaji = romaji_inline(romaji_raw)
    trans  = translation_html(trans_raw)

    attribution_html = ""
    if main_translator:
        attribution_html = f'    <div class="translation-attribution">tr. {main_translator}</div>'

    source_col, source_text = get_japanese_source(row)

    def info_box(label, content, translation=False):
        if not content:
            return ""
        if translation:
            lines = "\n".join(
                f'        <span>{l.strip()}</span>'
                for l in content.replace("\\n", "\n").splitlines()
                if l.strip()
            )
            inner = f'    <div class="info-box-translation">\n{lines}\n    </div>'
        else:
            paras = "\n".join(f'      <p>{p.strip()}</p>' for p in content.split("\n") if p.strip())
            inner = f'    <div class="info-box-body">\n{paras}\n    </div>'
        return f'  <div class="info-box">\n    <div class="info-box-label">{label}</div>\n{inner}\n  </div>\n'

    preface_box     = info_box("Preface", maegaki)
    other_trans_box = info_box(f"Other translation — {other_pair[0]}", other_pair[1], translation=True) if other_pair else ""
    notes_block          = build_section("Notes", notes, cls="translation-notes")
    appreciation_block   = build_section("Appreciation", appreciation, cls="translation-notes")

    poet_link = ""
    if poet_slug and poet:
        display_name = canonical_name or poet
        poet_link = f'<a class="ext-link" href="../poets/{poet_slug}.html">← {display_name}</a>'

    saijiki_link = ""
    if saijiki_slug and saijiki_season:
        season_dir = slugify(saijiki_season)
        saijiki_link = (f'<a class="ext-link" href="../saijiki/{season_dir}/{saijiki_slug}.html">'
                        f'Season word: {saijiki_slug} →</a>')

    body = f"""<div class="haiku-page">

{preface_box}
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
{attribution_html}
  </div>

{other_trans_box}{notes_block}{appreciation_block}
  <div class="metadata-grid">
    <div class="meta-cell">
      <div class="meta-label">Season Word</div>
      <div class="meta-value">{kigo_en}<span class="meta-value-jp">{kigo_jp}</span></div>
    </div>
    <div class="meta-cell">
      <div class="meta-label">Season</div>
      <div class="meta-value">{season or "—"}</div>
    </div>
    <div class="meta-cell no-right">
      <div class="meta-label">Cutting word</div>
      <div class="meta-value">—</div>
    </div>
    <div class="meta-cell">
      <div class="meta-label">Poet</div>
      <div class="meta-value">{canonical_name or poet or "—"}<span class="meta-value-jp">{poet_jp}</span></div>
    </div>
    <div class="meta-cell span2 no-right">
      <div class="meta-label">Date of composition</div>
      <div class="meta-value">—</div>
    </div>
    <div class="meta-cell no-bottom">
      <div class="meta-label">Original Source</div>
      <div class="meta-value">{src}</div>
    </div>
    <div class="meta-cell span2 no-right no-bottom">
      <div class="meta-label">Modern Source</div>
      <div class="meta-value">{source_text.replace(chr(10), "<br>") if source_text else "—"}</div>
    </div>
  </div>

  <div class="ext-links">{poet_link}{saijiki_link}</div>

</div>"""

    title = f"{poem_jp[:20]} — {canonical_name or poet}" if poem_jp else (canonical_name or poet)
    return html_page(title, POEM_CSS, body, home_href="../../index.html")


def build_poet_bio_page(poet_slug, data, bio_text, photos_dir):
    poet_name = data["name"]
    poet_jp   = data["jp"]
    dates     = data["dates"]
    period    = data["period"]
    school    = data.get("school", "")
    poems     = data["poems"]

    photo_fname = find_photo(photos_dir, poet_slug)
    photo_html  = photo_block_html(photo_fname, poet_name, path_prefix="../")

    # Strip leading "N. " prefix from school strings like "4. Shōmon"
    school_label = re.sub(r'^\d+\.\s*', '', school).strip()
    tag_text  = school_label or period
    tags_html = f'<span class="poet-tag">{tag_text}</span>' if tag_text else ""

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

    return html_page(poet_name, POET_CSS, body, home_href="../../index.html")


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

    return html_page(f"All poems — {poet_name}", POEM_LIST_CSS, body, home_href="../../index.html")


# ---------------------------------------------------------------------------
# Saijiki page builders
# ---------------------------------------------------------------------------

def load_essay(essays_dir, kigo_slug):
    path = os.path.join(essays_dir, f"{kigo_slug}.txt")
    if not os.path.exists(path):
        return ""
    with open(path, encoding="utf-8") as f:
        return f.read().strip()


def build_saijiki_entry(kigo_slug, meta, essay_text, exemplars, all_poem_count):
    kigo_en     = meta.get("kigo_en", kigo_slug)
    kigo_jp     = meta.get("kigo_jp", "")
    kigo_romaji = meta.get("kigo_romaji", "")
    season      = meta.get("season", "")
    category    = meta.get("category", "")

    tags = ""
    if season:
        tags += f'<span class="kigo-tag">{season}</span>'
    if category:
        tags += f'  <span class="kigo-tag">{category}</span>'

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

    exemplars_html = ""
    if exemplars:
        exemplars_html = f'''  <div class="section-divider">
    <span class="section-label">Exemplar poems</span>
    <div class="section-rule"></div>
  </div>

{poems_html}'''

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
{exemplars_html}
{see_all}

</div>'''

    return html_page(f"{kigo_en} — Saijiki", SAIJIKI_CSS, body, home_href="../../../index.html")


def build_saijiki_poems_list(kigo_slug, meta, all_poems):
    kigo_en = meta.get("kigo_en", kigo_slug)
    kigo_jp = meta.get("kigo_jp", "")
    season  = meta.get("season", "")

    entries = ""
    for p in all_poems:
        trans_lines = "\n".join(
            f'        <span>{l.strip()}</span>'
            for l in p["translation"].replace("\\n", "\n").splitlines()
            if l.strip()
        )
        entries += f'''    <a class="poem-entry" href="../../poems/{p["filename"]}">
      <div class="poem-entry-jp">{p["jp"]}</div>
      <div class="poem-entry-romaji">{p["romaji"]}</div>
      <div class="poem-entry-translation">
{trans_lines}
      </div>
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

    return html_page(f"All poems — {kigo_en}", SAIJIKI_LIST_CSS, body, home_href="../../../index.html")


def _kigo_grid_html(entries):
    grid = ""
    for slug, meta in entries:
        count = meta.get("poem_count", 0)
        grid += f'''      <a class="kigo-entry" href="{slug}.html">
        <div class="kigo-entry-en">{meta.get("kigo_en", slug.replace("-", " "))}</div>
        <div class="kigo-entry-jp">{meta.get("kigo_jp", "")}</div>
        <div class="kigo-entry-count">{count} poem{"s" if count != 1 else ""}</div>
      </a>
'''
    return grid


def build_saijiki_season_index(season, entries, sort="category"):
    if sort == "alpha":
        body_html = f'  <div class="kigo-grid">\n{_kigo_grid_html(sorted_alpha(entries))}  </div>\n'
        toggle = ('<a href="index.html">By category</a>'
                  '<span class="sort-sep">·</span><span class="sort-active">Alphabetical</span>')
    else:
        groups_html = ""
        for cat, cat_entries in sorted_by_category(entries):
            groups_html += f'''  <div class="category-group">
    <div class="category-heading">{cat}</div>
    <div class="kigo-grid">
{_kigo_grid_html(cat_entries)}    </div>
  </div>
'''
        body_html = groups_html
        toggle = ('<span class="sort-active">By category</span>'
                  '<span class="sort-sep">·</span><a href="index-alpha.html">Alphabetical</a>')

    body = f'''<div class="saijiki-page">

  <div class="page-label">
    Saijiki · {season}
    <div class="page-label-line"></div>
  </div>

  <div class="index-title">{season} Season Words</div>
  <div class="sort-toggle">Sort: {toggle}</div>

{body_html}
  <div class="ext-links">
    <a class="ext-link" href="../../saijiki-index.html">← All season words</a>
  </div>

</div>'''
    return html_page(f"{season} — Saijiki", SAIJIKI_INDEX_CSS, body, home_href="../../../index.html")


def _master_kigo_grid_html(entries, season):
    grid = ""
    for slug, meta in entries:
        count = meta.get("poem_count", 0)
        grid += f'''      <a class="kigo-entry" href="saijiki/{slugify(season)}/{slug}.html">
        <div class="kigo-entry-en">{meta.get("kigo_en", slug)}</div>
        <div class="kigo-entry-jp">{meta.get("kigo_jp", "")}</div>
        <div class="kigo-entry-count">{count} poem{"s" if count != 1 else ""}</div>
      </a>
'''
    return grid


def build_saijiki_index(saijiki_data, sort="category"):
    by_season = {}
    for slug, meta in saijiki_data.items():
        s = meta.get("season", "Unclassified")
        by_season.setdefault(s, []).append((slug, meta))

    ordered = [s for s in SEASON_ORDER if s in by_season]
    for s in by_season:
        if s not in ordered:
            ordered.append(s)

    groups_html = ""
    for season in ordered:
        entries = by_season.get(season, [])
        if sort == "alpha":
            inner_html = f'    <div class="kigo-grid">\n{_master_kigo_grid_html(sorted_alpha(entries), season)}    </div>\n'
        else:
            inner_html = ""
            for cat, cat_entries in sorted_by_category(entries):
                inner_html += f'''    <div class="category-group">
      <div class="category-heading">{cat}</div>
      <div class="kigo-grid">
{_master_kigo_grid_html(cat_entries, season)}      </div>
    </div>
'''
        groups_html += f'''  <div class="season-group">
    <div class="season-heading">{season}</div>
{inner_html}  </div>
'''

    if sort == "alpha":
        toggle = ('<a href="saijiki-index.html">By category</a>'
                  '<span class="sort-sep">·</span><span class="sort-active">Alphabetical</span>')
    else:
        toggle = ('<span class="sort-active">By category</span>'
                  '<span class="sort-sep">·</span><a href="saijiki-index-alpha.html">Alphabetical</a>')

    body = f'''<div class="saijiki-page">

  <div class="page-label">
    Saijiki
    <div class="page-label-line"></div>
  </div>

  <div class="index-title">Season Word Index</div>
  <div class="sort-toggle">Sort: {toggle}</div>

{groups_html}
</div>'''

    return html_page("Saijiki — Season Word Index", SAIJIKI_INDEX_CSS, body, home_href="../index.html")


def update_home_poems(home_path, poems_by_season):
    import json
    if not os.path.exists(home_path):
        return
    with open(home_path, encoding="utf-8") as f:
        content = f.read()
    data_json = json.dumps(poems_by_season, ensure_ascii=False)
    start_tag = '<script id="haiku-poems-data" type="application/json">'
    end_tag = "</script>"
    i = content.find(start_tag)
    if i == -1:
        return
    i += len(start_tag)
    j = content.find(end_tag, i)
    content = content[:i] + data_json + content[j:]
    with open(home_path, "w", encoding="utf-8") as f:
        f.write(content)


def update_home_stats(home_path, poem_count, poet_count, season_counts, kigo_count=0):
    if not os.path.exists(home_path):
        return
    with open(home_path, encoding="utf-8") as f:
        content = f.read()

    content = re.sub(
        r'(?<=class="stat-number">)\d+(?=</span>\s*<span class="stat-label">Poems)',
        str(poem_count), content
    )
    content = re.sub(
        r'(?<=class="stat-number">)\d+(?=</span>\s*<span class="stat-label">Poets)',
        str(poet_count), content
    )
    if kigo_count:
        content = re.sub(
            r'(?<=class="stat-number">)\d+(?=</span>\s*<span class="stat-label">Season Words)',
            str(kigo_count), content
        )
    known = {"Spring": 0, "Summer": 0, "Autumn": 0, "Winter": 0}
    for raw, count in season_counts.items():
        for k in known:
            if raw.lower() == k.lower():
                known[k] += count
    for season_name, count in known.items():
        content = re.sub(
            rf'(class="season-card {season_name.lower()}"[^>]*>.*?class="season-count">)\d+ poems',
            rf'\g<1>{count} poems',
            content, flags=re.DOTALL
        )
    content = re.sub(r'All \d+ poets', f'All {poet_count} poets', content)

    with open(home_path, "w", encoding="utf-8") as f:
        f.write(content)


def build_poets_index(poets_data):
    import json

    poet_records = []
    for slug, data in sorted(poets_data.items(), key=lambda x: x[1]["name"].split()[-1]):
        poet_records.append({
            "href":  f"{slug}.html",
            "name":  data["name"],
            "jp":    data["jp"],
            "dates": data["dates"],
        })

    poets_json = json.dumps(poet_records, ensure_ascii=False, separators=(',', ':'))
    total = len(poets_data)

    body = f'''<div class="poets-page">

  <div class="page-label">
    Poets
    <div class="page-label-line"></div>
  </div>

  <div class="poets-title">All {total} Poets</div>

  <div class="sort-controls">
    <span class="sort-label">Sort</span>
    <button class="sort-btn active" data-sort="given">Given name</button>
    <button class="sort-btn" data-sort="family">Family name</button>
    <button class="sort-btn" data-sort="date">Date</button>
  </div>

  <div id="poets-list"></div>

</div>
<script>
(function(){{
  var POETS={poets_json};
  function birthYear(d){{
    if(!d)return 9999;
    var m=d.match(/\\d{{3,4}}/);
    return m?parseInt(m[0]):9999;
  }}
  function givenKey(p){{return p.name.split(/\\s+/).pop().toLowerCase();}}
  function familyKey(p){{return p.name.split(/\\s+/)[0].toLowerCase();}}
  function givenGroup(p){{return p.name.split(/\\s+/).pop()[0].toUpperCase();}}
  function familyGroup(p){{return p.name.split(/\\s+/)[0][0].toUpperCase();}}
  function dateGroup(y){{
    if(y===9999)return'Dates unknown';
    if(y<1600)return'Before 1600';
    if(y<1700)return'1600s';
    if(y<1800)return'1700s';
    if(y<1868)return'Early 1800s';
    if(y<1900)return'1868–1899';
    if(y<1945)return'1900–1944';
    return'1945 and after';
  }}
  function render(sort){{
    var sorted=[].concat(POETS).sort(function(a,b){{
      var ka,kb;
      if(sort==='given'){{ka=givenKey(a);kb=givenKey(b);}}
      else if(sort==='family'){{ka=familyKey(a);kb=familyKey(b);}}
      else{{ka=birthYear(a.dates);kb=birthYear(b.dates);}}
      return ka<kb?-1:ka>kb?1:0;
    }});
    var groups={{}},order=[];
    sorted.forEach(function(p){{
      var g;
      if(sort==='given')g=givenGroup(p);
      else if(sort==='family')g=familyGroup(p);
      else g=dateGroup(birthYear(p.dates));
      if(!groups[g]){{groups[g]=[];order.push(g);}}
      groups[g].push(p);
    }});
    var html='';
    order.forEach(function(g){{
      html+='<div class="alpha-group"><div class="alpha-heading">'+g+'</div><div class="poet-list">';
      groups[g].forEach(function(p){{
        html+='<a class="poet-list-item" href="'+p.href+'"><span class="poet-list-name">'+p.name+'</span><div class="poet-list-right">';
        if(p.jp)html+='<span class="poet-list-jp">'+p.jp+'</span>';
        if(p.dates)html+='<span class="poet-list-dates">'+p.dates+'</span>';
        html+='</div></a>';
      }});
      html+='</div></div>';
    }});
    document.getElementById('poets-list').innerHTML=html;
    document.querySelectorAll('.sort-btn').forEach(function(btn){{
      btn.classList.toggle('active',btn.dataset.sort===sort);
    }});
  }}
  document.querySelectorAll('.sort-btn').forEach(function(btn){{
    btn.addEventListener('click',function(){{render(btn.dataset.sort);}});
  }});
  render('given');
}})();
</script>'''

    return html_page("Poets — The Haiku Project", POETS_INDEX_CSS, body, home_href="../../index.html")


def build_index(poem_count, poet_count, saijiki_count):
    """Master site index linking to poems, poets, and saijiki."""
    body = f'''<div class="index-page">

  <div class="page-label">
    歳時記
    <div class="page-label-line"></div>
  </div>

  <div class="index-title">The Haiku Project</div>
  <div class="index-subtitle">歳時記 · Saijiki</div>

  <nav class="index-nav">
    <a class="index-nav-item" href="saijiki-index.html">
      <div>
        <div class="index-nav-label">Season Words</div>
        <div class="index-nav-sub">{saijiki_count} kigo entries</div>
      </div>
      <span class="index-nav-arrow">→</span>
    </a>
    <a class="index-nav-item" href="poets/">
      <div>
        <div class="index-nav-label">Poets</div>
        <div class="index-nav-sub">{poet_count} haijin</div>
      </div>
      <span class="index-nav-arrow">→</span>
    </a>
    <a class="index-nav-item" href="poems/">
      <div>
        <div class="index-nav-label">Poems</div>
        <div class="index-nav-sub">{poem_count} haiku</div>
      </div>
      <span class="index-nav-arrow">→</span>
    </a>
  </nav>

</div>'''

    return html_page("The Haiku Project · 歳時記", INDEX_CSS, body, home_href="../index.html")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--xlsx",   default="/mnt/project/AI_Haiku.xlsx")
    parser.add_argument("--out",    default="/mnt/user-data/outputs/haiku_site")
    parser.add_argument("--bios",   default="/home/claude/bios")
    parser.add_argument("--photos", default="/home/claude/photos")
    parser.add_argument("--essays",        default="/home/claude/essays")
    parser.add_argument("--appreciations", default="appreciations")
    parser.add_argument("--poets",  default="")
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

    # Load canonical poets registry
    canonical_poets, poet_aliases = load_canonical_poets(args.poets)
    if canonical_poets:
        print(f"  Canonical poets: {len(canonical_poets)} entries, {len(poet_aliases)} aliases")
    else:
        print("  No canonical poets file — using spreadsheet data directly")

    # Read all rows
    all_rows = []
    for sheet in args.sheets:
        try:
            df = pd.read_excel(args.xlsx, sheet_name=sheet)
            for _, row in df.iterrows():
                all_rows.append((sheet, row.to_dict()))
        except Exception as e:
            print(f"  Skipping sheet '{sheet}': {e}")

    # Saijiki registry — loaded up front so poem pages (Pass 1) can link to
    # their season-word page. See the fuller comment above Pass 3 below for
    # how a poem is matched to a term (by its Kigo field's leading kanji).
    saijiki_slug_by_jp = {}
    try:
        df_s = pd.read_excel(args.xlsx, sheet_name="Saijiki")
        for _, row in df_s.iterrows():
            slug    = val(row.get("Kigo_Slug", ""))
            kigo_jp = val(row.get("Kigo_JP", ""))
            if slug and kigo_jp:
                saijiki_slug_by_jp[kigo_jp] = slug
    except Exception:
        pass

    # Pass 1 — write poem pages, accumulate per-poet data
    slugs_used      = {}
    poets_data      = {}
    poem_count      = 0
    season_counts   = {}
    kigo_seen       = set()
    home_poems      = {}   # season → list of poem dicts for home page cycling

    for sheet, row in all_rows:
        poem_jp    = val(row.get("Poem") or row.get("Text", ""))
        romaji_raw = val(row.get("Romaji", ""))
        trans_raw  = val(row.get("My translation", ""))
        poet       = val(row.get("Poet", ""))
        poet_jp    = val(row.get("俳人", ""))
        dates      = val(row.get("Dates", ""))
        season     = val(row.get("Season", ""))
        kigo_raw   = val(row.get("Kigo", ""))
        saijiki_entry = saijiki_slug_by_jp.get(kigo_leading_kanji(kigo_raw), "")

        if not poem_jp and not romaji_raw:
            continue

        # Unique poem slug
        base = slugify(romaji_raw if romaji_raw else poem_jp)
        if base in slugs_used:
            slugs_used[base] += 1
            slug = f"{base}-{slugs_used[base]}"
        else:
            slugs_used[base] = 0
            slug = base
        poem_filename = f"{sheet.lower()}-{slug}.html"

        raw_slug  = slugify(poet) if poet else "unknown"
        poet_slug = resolve_poet_slug(raw_slug, poet_aliases, canonical_poets)

        # Write poem page — pass saijiki_entry and season for link generation
        poem_slug    = poem_filename[:-5]  # strip .html
        appreciation = load_appreciation(args.appreciations, poem_slug)
        canon_name = canonical_poets.get(poet_slug, {}).get("name", "")
        html = build_poem_page(row, poem_filename, poet_slug,
                               saijiki_slug=saijiki_entry,
                               saijiki_season=season,
                               appreciation=appreciation,
                               canonical_name=canon_name)
        with open(os.path.join(poems_dir, poem_filename), "w", encoding="utf-8") as f:
            f.write(html)
        poem_count += 1
        if kigo_raw:
            kigo_seen.add(kigo_raw)
        if season:
            key = season.strip().title()
            season_counts[key] = season_counts.get(key, 0) + 1
            if key in ("Spring", "Summer", "Autumn", "Winter") and poem_jp and trans_raw and poet:
                kigo_short_h, _, _ = kigo_parts(kigo_raw)
                home_poems.setdefault(key, []).append({
                    "jp":          poem_jp,
                    "romaji":      romaji_inline(romaji_raw),
                    "translation": trans_raw,
                    "poet":        poet,
                    "kigo":        kigo_short_h,
                    "url":         f"site/poems/{poem_filename}",
                })

        # Accumulate poet — prefer canonical registry data over spreadsheet data
        kigo_short, _, _ = kigo_parts(kigo_raw)
        if poet_slug not in poets_data:
            canon = canonical_poets.get(poet_slug, {})
            poets_data[poet_slug] = {
                "name":   canon.get("name") or poet or "Unknown",
                "jp":     canon.get("jp")   or poet_jp,
                "dates":  canon.get("dates") or dates,
                "period": canon.get("period") or sheet,
                "school": canon.get("school", ""),
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

    print(f"  Poems: {poem_count} pages written")

    # Add canonical poets who have a bio file but no haiku in the spreadsheet
    for poet_slug, canon in canonical_poets.items():
        if poet_slug in poets_data:
            continue
        bio_text = load_bio(args.bios, poet_slug)
        if not bio_text:
            continue
        poets_data[poet_slug] = {
            "name":   canon.get("name", ""),
            "jp":     canon.get("jp", ""),
            "dates":  canon.get("dates", ""),
            "period": canon.get("period", ""),
            "school": canon.get("school", ""),
            "poems":  [],
        }

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

    poets_index_html = build_poets_index(poets_data)
    with open(os.path.join(poets_dir, "index.html"), "w", encoding="utf-8") as f:
        f.write(poets_index_html)

    # Pass 3 — saijiki pages
    #
    # Pipeline, in order:
    #   1. Check the Saijiki sheet — this is the master list of kigo terms
    #      (one row per term: slug, kanji, romaji, English gloss, season,
    #      category). It drives which pages exist at all.
    #   2. For each term, get the essay text from essays/{slug}.txt.
    #   3. Get the example (exemplar) poems for that term from AI_Haiku —
    #      poems whose Kigo field names this term AND carry a "Saijiki
    #      priority" value, ranked by that value.
    #   4. Get all the other matching poems from AI_Haiku — every poem
    #      whose Kigo field names this term, priority or not — for the
    #      "all N poems" list page.
    #
    # A poem is matched to a term by comparing the leading kanji run of its
    # free-text Kigo column against the term's Kigo_JP — there is no manual
    # per-poem tagging step.

    # 1. Saijiki sheet
    saijiki_meta = {}
    slug_by_jp = {}
    try:
        df_s = pd.read_excel(args.xlsx, sheet_name="Saijiki")
        for _, row in df_s.iterrows():
            slug    = val(row.get("Kigo_Slug", ""))
            kigo_jp = val(row.get("Kigo_JP", ""))
            if not slug or not kigo_jp:
                continue
            try:
                seq = int(float(row.get("Seq", "") or ""))
            except (ValueError, TypeError):
                seq = 999999
            saijiki_meta[slug] = {
                "kigo_en":     val(row.get("Kigo_EN", "")) or slug.replace("-", " "),
                "kigo_jp":     kigo_jp,
                "kigo_romaji": val(row.get("Kigo_Romaji", "")),
                "season":      val(row.get("Season", "")),
                "category":    val(row.get("Category", "")),
                "seq":         seq,
            }
            slug_by_jp[kigo_jp] = slug
        print(f"  Saijiki sheet: {len(saijiki_meta)} kigo terms")
    except Exception as e:
        print(f"  Skipping Saijiki sheet: {e}")

    # 3 & 4. Match every poem row to a term by its Kigo field's leading kanji
    saijiki_poems = {}
    for sheet, row in all_rows:
        entry_slug = slug_by_jp.get(kigo_leading_kanji(row.get("Kigo", "")))
        if not entry_slug:
            continue

        try:
            priority = float(row.get("Saijiki priority", "") or "")
            if priority != priority:  # NaN
                priority = None
        except (ValueError, TypeError):
            priority = None

        poem_jp    = val(row.get("Poem") or row.get("Text", ""))
        romaji_raw = val(row.get("Romaji", ""))
        trans_raw  = val(row.get("My translation", ""))
        poet       = val(row.get("Poet", ""))
        poet_jp    = val(row.get("俳人", ""))

        base = slugify(romaji_raw if romaji_raw else poem_jp)
        poem_filename = f"{sheet.lower()}-{base}.html"

        saijiki_poems.setdefault(entry_slug, []).append({
            "priority":    priority,
            "jp":          poem_jp,
            "romaji":      romaji_inline(romaji_raw),
            "translation": trans_raw,
            "poet":        poet,
            "poet_jp":     poet_jp,
            "filename":    poem_filename,
        })

    saijiki_count = 0
    saijiki_built = {}  # only terms that actually got a page — index/season pages link only to these
    for kigo_slug, meta in saijiki_meta.items():
        poems = saijiki_poems.get(kigo_slug, [])

        # 2. essay text
        essay_text = load_essay(args.essays, kigo_slug)

        if not essay_text and not poems:
            continue  # nothing to show for this term yet

        season = slugify(meta.get("season") or "unknown")
        season_dir = os.path.join(args.out, "saijiki", season)
        os.makedirs(season_dir, exist_ok=True)

        # 3. exemplars = prioritized matches, ranked; falls back to the
        #    first few matches if no priority has been set yet
        prioritized = sorted((p for p in poems if p["priority"] is not None),
                              key=lambda p: p["priority"])
        exemplars = prioritized[:4] if prioritized else poems[:4]

        # 4. all matches, prioritized ones first, for the full list page
        poems_sorted = prioritized + [p for p in poems if p["priority"] is None]

        entry_html = build_saijiki_entry(
            kigo_slug, meta, essay_text, exemplars, len(poems)
        )
        with open(os.path.join(season_dir, f"{kigo_slug}.html"), "w", encoding="utf-8") as f:
            f.write(entry_html)

        list_html = build_saijiki_poems_list(kigo_slug, meta, poems_sorted)
        with open(os.path.join(season_dir, f"{kigo_slug}-poems.html"), "w", encoding="utf-8") as f:
            f.write(list_html)

        meta["poem_count"] = len(poems)
        saijiki_built[kigo_slug] = meta
        saijiki_count += 1

    if saijiki_count:
        with open(os.path.join(args.out, "saijiki-index.html"), "w", encoding="utf-8") as f:
            f.write(build_saijiki_index(saijiki_built, sort="category"))
        with open(os.path.join(args.out, "saijiki-index-alpha.html"), "w", encoding="utf-8") as f:
            f.write(build_saijiki_index(saijiki_built, sort="alpha"))

        by_season = {}
        for slug, meta in saijiki_built.items():
            s = meta.get("season", "")
            if s:
                by_season.setdefault(s, []).append((slug, meta))
        for season, entries in by_season.items():
            season_dir = os.path.join(args.out, "saijiki", slugify(season))
            os.makedirs(season_dir, exist_ok=True)
            with open(os.path.join(season_dir, "index.html"), "w", encoding="utf-8") as f:
                f.write(build_saijiki_season_index(season, entries, sort="category"))
            with open(os.path.join(season_dir, "index-alpha.html"), "w", encoding="utf-8") as f:
                f.write(build_saijiki_season_index(season, entries, sort="alpha"))

        print(f"  Saijiki: {saijiki_count} entries ({saijiki_count * 2} pages) + category/alpha indexes")
    else:
        print(f"  Saijiki: no entries found (need a 'Saijiki' sheet with a Kigo_JP that matches poems' Kigo field, and/or essays/{{slug}}.txt files)")

    # Write master index
    index_html = build_index(poem_count, poet_count, saijiki_count)
    with open(os.path.join(args.out, "index.html"), "w", encoding="utf-8") as f:
        f.write(index_html)
    print(f"  Index: index.html written")

    home_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "index.html")
    update_home_stats(home_path, poem_count, poet_count, season_counts, kigo_count=len(kigo_seen))
    update_home_poems(home_path, home_poems)
    print(f"  Home:  index.html stats + poems updated")

    print(f"\nDone — site written to: {args.out}")
    print(f"\n  poems/          {poem_count} poem pages")
    print(f"  poets/          {poet_count * 2} poet pages")
    print(f"  saijiki/        {saijiki_count * 2} saijiki pages + index")
    print(f"  index.html      master index")
    print(f"\nSource folders (next to this script):")
    print(f"  bios/           poet biographies  ({{slug}}.txt)")
    print(f"  essays/         saijiki essays    ({{kigo-slug}}.txt)")
    print(f"  photos/         poet photos       ({{slug}}.jpg)")


if __name__ == "__main__":
    main()