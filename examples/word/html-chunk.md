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

Round-trip: `officecli dump` re-emits body-level HTML/RTF/text chunks as
`add htmlchunk` items. Chunks inside table cells are reported as a dump warning.

## Regenerate

```bash
bash html-chunk.sh
# or
python3 html-chunk.py
```
