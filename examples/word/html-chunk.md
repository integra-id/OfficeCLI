# Word HTML Chunk Showcase

This demo consists of these files:

- **html-chunk.sh** — CLI script that calls `officecli` to build the document.
  Each chunk is preceded by a `# Features:` comment listing what it exercises.
- **html-chunk.py** — Python SDK twin (`officecli-sdk`). One resident, every item
  shipped in a single `doc.batch(...)`. Produces identical chunk payloads.
- **html-chunk.docx** — the generated document: 11 chunks (9 HTML, 1 RTF,
  1 plain text) between native Word headings, plus one chunk in a table cell.
- **html-chunk-report.html** — a complete HTML document with `<style>` CSS,
  embedded from a file with `src=`.
- **html-chunk-memo.rtf** / **html-chunk-notes.txt** — RTF and plain-text sources.
- **html-chunk.md** — this file.

## How HTML chunks work

```bash
officecli add file.docx /body --type htmlchunk --prop html='<h2>Title</h2><table>…</table>'
officecli add file.docx /body --type htmlchunk --prop src=page.html
```

The payload is stored **verbatim** in its own package part
(`word/afchunkN.htm`) and referenced from the body by
`<w:altChunk r:id="…"/>` — Word's native *alternative format import*. When the
document is opened, **Word converts each chunk into native paragraphs, tables,
lists and images**, so the HTML keeps its CSS fidelity (colspan/rowspan, colors,
fonts, borders) without OfficeCLI translating the markup.

Until Word has opened and re-saved the file, each chunk is a single opaque node:

```bash
officecli get html-chunk.docx /body/altChunk[8]
# /body/altChunk[8] (altChunk) contentType=text/html format=html partUri=/word/afchunk7.htm size=2806 matchSrc=true
```

- The HTML preview (`view html`, `watch`) shows a chunk's **source as escaped
  text**, never as live HTML.
- A chunk's text is not queryable or editable through paragraph paths. Use
  `--type markdown` or native adds when you need to edit or check the content
  afterwards.

## Properties

| Property | Aliases | Meaning |
|---|---|---|
| `html` | `content`, `text` | Inline markup: a fragment (auto-wrapped in a UTF-8 HTML document) or a complete document |
| `src` | `path` | File to embed. Complete documents are stored byte-for-byte |
| `format` | | `html` (default), `xhtml`, `mht`, `rtf`, `text`. Inferred from the `src` extension |
| `matchSrc` | | `true` → `<w:altChunkPr><w:matchSrc/>`: keep the source formatting instead of the document styles |
| `title` | | `<title>` for the wrapper generated around a fragment |

Parents: `/body` (with `--index` / `--after` / `--before`) and table cells
(`/body/tbl[N]/tr[M]/tc[K]`). Headers, footers and footnotes are not supported.

## Sections of html-chunk.docx

| # | Section | Demonstrates |
|---|---|---|
| — | Note box at the top | positioned insert with `--after p[2]` (added last) |
| 1 | Inline text formatting | `b` `i` `u` `s` `sub` `sup` `mark` `small` `big`, inline `color` / `background` / `font-family` / `font-size` / `letter-spacing` / `text-transform`, Unicode (Latin accents, CJK, Arabic, Greek, currency) |
| 2 | Headings & paragraph layout | `h1`–`h4`, `text-align` left/center/right/justify, `text-indent`, `margin-left`, `line-height`, bordered and shaded block |
| 3 | Lists | `ul` disc/square, 3 nesting levels, `ol start=3`, `lower-alpha`, `upper-roman`, mixed nesting |
| 4 | Tables | `rowspan` + `colspan` two-row header, header fills, zebra row, per-cell color and alignment, `tfoot` total row |
| 5 | Quote, code, rule, links | `blockquote`, inline `code`, `pre` with preserved whitespace, `hr`, `https:` and `mailto:` links |
| 6 | Image | `<img>` with a base64 **data URI** (relative paths are not resolved inside a chunk) |
| 7 | Complete HTML file | `src=html-chunk-report.html`: `<style>` CSS classes, rowspan/colspan, callout boxes, nested lists, `matchSrc=true` |
| 8 | RTF & plain text | `src=….rtf` (format inferred) and `src=….txt` with `format=text` |
| 9 | Chunk in a table cell | parent `/body/tbl[1]/tr[1]/tc[2]`, next to a native text cell |

## Viewer support

| Application | HTML chunk | RTF / text chunk |
|---|---|---|
| Microsoft Word (Windows / macOS) | ✓ converted on open | ✓ |
| LibreOffice Writer | ✓ (HTML import; CSS support is narrower than Word's) | partial |
| Google Docs, Pages, most previewers | ✗ chunk is skipped | ✗ |

Open the file once in Word and **save it** to replace every chunk with native
content that all viewers understand.

## Headless conversion (`materialize`)

`officecli materialize file.docx` replaces HTML, XHTML and plain-text altChunks
with native paragraphs, lists, tables and links without opening Word. RTF and
MHT stay as altChunks. `--strict` refuses to change the file when any whole
chunk cannot be converted.

Data-URI images whose type the picture pipeline already accepts (png, jpeg,
gif, bmp, tiff, emf, wmf) are embedded as inline `w:drawing` pictures. The
`alt` attribute becomes the picture description. webp and svg are not
embedded. `http(s)` URLs are not downloaded and relative `src` values are not
resolved; they stay as italic `[image: …]` alt text, and the rest of the chunk
is still converted. A data URI that fails to decode is the same kind of
warning and does not trip `--strict`.

### CSS borders

Materialize maps a small border subset onto real OOXML. It does not run a browser.

| CSS | Word |
|---|---|
| `border`, `border-top` / `right` / `bottom` / `left` | all four edges, or one edge |
| `border-width`, `border-style`, `border-color` | the matching longhands (1–4 values, the usual box order) |
| `border-*-width`, `border-*-style`, `border-*-color` | that edge |
| paragraph, heading, `pre`, or a callout wrapper (`div`, `blockquote`) | `w:pBdr` |
| `td` / `th` (a `tr` border fills a side the cell did not set) | `w:tcBorders` |
| every cell has the same four-edge border | `w:tblBorders` too, including inside lines |
| padding on a bordered paragraph | `w:space` on that edge (points) |
| padding on a cell | `w:tcMar` (twips) |

Styles Word can draw: `solid` → `single`, `dashed`, `dotted`, `double`, `inset`, `outset`. `none` and `hidden` draw no line (`nil` on a cell, so they cover the table grid). `groove` is drawn as `inset` and `ridge` as `outset`; materialize warns. Any other line style, or a width that is not a length or `thin` / `medium` / `thick`, is dropped and warned. `transparent` draws no line. `currentcolor` and `inherit` use `w:color="auto"`.

A border on a wrapper is copied onto each paragraph inside it. It is not one rectangle around the group. Borders on `span`, `code`, and other inline elements are not turned into run borders. `border-radius`, `border-image`, `border-spacing`, `outline`, `box-shadow`, and logical properties (`border-inline-*`, `border-block-*`) are ignored. `border-collapse` is not modeled separately. The HTML `border` attribute is not read; a table with no CSS border still gets the built-in single grid.

These warnings do not trip `--strict`.

### Nested lists

A nested list (`ul` / `ol` inside an `li`, including one wrapped in a `div` or `blockquote`) is one Word numbering instance. Every item in that tree shares a single `w:numId`. `w:ilvl` is the depth: 0 on the outermost list, then 1, 2, and so on, clamped at 8.

A list that is not inside that tree gets its own `numId`. That includes the next `ul` / `ol` in the body and any list in a table cell, even when the table itself sits inside an `li`. Two adjacent lists are not merged. A paragraph between a marker and a nested list does not split the instance.

Markers still come from the existing list-style helper. An `ol` level cycles decimal, then lower-letter, then lower-roman. A `ul` level cycles •, ◦, ▪. The first item at a depth picks that level's marker; a later sublist at the same depth keeps it. The `type` attribute on `ol` is not read.

Sharing one `numId` is what makes Word treat the tree as a single multilevel list: a later parent item stays on that instance instead of starting a new list at 1. Before this, each nested list minted its own `numId` at the right `ilvl`. The parent usually continued, but a nested list always restarted, and a paragraph or a change of list type between items could mint a new `numId` for the parent too.

The definition is that helper's hybrid multilevel list (`w:multiLevelType` = `hybridMultilevel`) and does not write `w:lvlRestart`. Two restart edges remain:

- ECMA-376 says a hybrid level with no `lvlRestart` does not restart. Word may keep a nested ordered counter going across parent items (a, b under item 1, then c under item 2). `officecli view` restarts a level whose `lvlRestart` is omitted whenever a shallower level advances, and it does not special-case hybrid lists, so the preview can show that counter starting again where Word continues.
- `start` on the outermost list is that level's `w:start`. `start` on a nested list is written once, as `w:startOverride` on that `ilvl` of the shared instance. It applies to the whole level, not to one sublist. A second nested list at the same depth does not get its own start. `start="1"` adds no override.

See `materialize-altchunk.sh`.

Round-trip: `officecli dump` re-emits body-level HTML/RTF/text chunks as
`add htmlchunk` items. Chunks inside table cells are reported as a dump warning.

## Regenerate

```bash
bash html-chunk.sh
# or
python3 html-chunk.py
```
