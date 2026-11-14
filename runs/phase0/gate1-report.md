# Phase 0 — Anchor Spec (frozen from your real inputs)

Generated offline from: guidelines PDF (Stream A) + Sample CoverPage.docx & Sample Chapter-1.docx (Stream C).
Precedence applied: Stream C (sample) > Stream A (guidelines). Review at the **spec-approval** gate.

## Conflicts (2) — sample wins by default; override below if the guidelines should win

| Anchor | Guidelines (A) | Sample (C) — current spec value | Recommendation |
|---|---|---|---|
| `FONT.FAMILY` | Times New Roman | Calibri (theme font in the sample) | **Override to Times New Roman** — the guidelines mandate it; the sample only carries it via a Word theme default |
| `FONT.MAIN_HEADING_ALIGNMENT` | LEFT | BOTH (justified) | **Override to LEFT** — guidelines say main headings are left-aligned |

## Missing anchors (4) — not present in either stream

- `REFERENCES.SEPARATE_FILE` — suggest `false` (references chapter inside the same file)
- `REFERENCES.INDENT_STYLE` — suggest `HANGING` (standard IEEE-style)
- `REFERENCES.SIZE_PT` — suggest `12`
- `REFERENCES.ALIGNMENT` — suggest `LEFT`

## Resolved spec (all other anchors)

| Group | Values |
|---|---|
| PAGE | A4 portrait, margins L=1.5" R=1.0" T=1.0" B=1.0" (11906x16838 twips) |
| FONT | body 12pt regular justified 1.5-spacing; main heading 14pt bold; subheading 12pt bold+italic; cover title 16pt bold right |
| COVER_PAGE | 19 empty paragraphs offset, no header/footer, not counted |
| FIRST_CONTENT_PAGE | repeats cover title, no header/footer, page counted |
| SUBSEQUENT_PAGES | header "Chapter -{{CHAPTER_NUM}}" 10pt right; footer centered PAGE field; numbering starts at 2 (displayed as 1 on first content page via titlePg) |

## Suggested anchor_overrides for Gate 1

```json
{
  "FONT.FAMILY": "Times New Roman",
  "FONT.MAIN_HEADING_ALIGNMENT": "LEFT",
  "REFERENCES.SEPARATE_FILE": false,
  "REFERENCES.INDENT_STYLE": "HANGING",
  "REFERENCES.SIZE_PT": 12,
  "REFERENCES.ALIGNMENT": "LEFT"
}
```
