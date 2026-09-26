# OfficeCLI

> **OfficeCLI is the world's first and the best Office suite designed for AI agents.**

**Give any AI agent full control over Word, Excel, and PowerPoint — in one line of code.**

Open-source. Single binary. No Office installation. No dependencies. Works everywhere.

**OfficeCLI's built-in HTML rendering engine reproduces documents with high fidelity — and that's what gives AI eyes.** It renders `.docx` / `.xlsx` / `.pptx` to HTML or PNG, closing the *render → look → fix* loop.

[![GitHub Release](https://img.shields.io/github/v/release/iOfficeAI/OfficeCLI)](https://github.com/iOfficeAI/OfficeCLI/releases)
[![License](https://img.shields.io/badge/license-Apache%202.0-blue.svg)](LICENSE)

**English** | [中文](README_zh.md) | [日本語](README_ja.md) | [한국어](README_ko.md)

> **This repository is the [integra-id/OfficeCLI](https://github.com/integra-id/OfficeCLI) fork** of [iOfficeAI/OfficeCLI](https://github.com/iOfficeAI/OfficeCLI). The working line is `feature/html-chunk` (not `main`). It keeps the upstream CLI and adds a headless Word pipeline for technical documents. Read [This fork](#this-fork-headless-word-technical-documents) first if you are contributing or writing a skill.

## This fork: headless Word technical documents

Upstream OfficeCLI already creates and edits `.docx` / `.xlsx` / `.pptx` with no Office install. This fork’s branch adds the pieces a technical-document skill needs: embed HTML as `w:altChunk`, turn supported chunks into native OOXML without Word, rebuild TOC entries, apply an A4 + Moderate page preset, and fit ID columns so short codes stay on one line.

Assembly version on this branch: `1.0.152-htmlchunk.1` (`src/officecli/officecli.csproj`). The guide below is the upstream README. Fork-only behavior is this section, [examples/word/html-chunk.md](examples/word/html-chunk.md), and the skills under `skills/`.

### What this fork adds

| Piece | Role |
| --- | --- |
| `add --type htmlchunk` | Store HTML, XHTML, MHT, RTF, or plain text as a Word `w:altChunk` (aliases `html`, `altchunk`). |
| `materialize` | Replace HTML / XHTML / plain-text chunks with native paragraphs, lists, tables, links, and accepted data-URI pictures. RTF and MHT stay. |
| `refresh --toc` | Rebuild TOC entry titles and in-document hyperlinks. `PAGEREF` page numbers are the placeholder `0`. |
| `pageSetup` / `pageSize` / `margins` | Named presets. `a4-moderate` (aliases `tech-doc`, `techdoc`) is A4 plus Word’s Moderate margins. |
| ID columns | `idColumn` / `idColumns` / `noWrap` / `width=fit` on native tables. Htmlchunk cells with class `col-id` or `white-space: nowrap` get the same treatment when materialized. |
| `finalize` | One headless pass: materialize, then TOC (if a TOC field exists), then `a4-moderate` only on sections with no `w:pgSz`, then schema validate. |
| Skills | `dokumen-teknis-officecli` and `rule-gaya-dokumen-teknis`, registered in `SkillMap`. |

Excel, PowerPoint, and the rest of the Word DOM are the upstream commands documented further down.

### Build and run

Compilation needs the [.NET 10 SDK](https://dotnet.microsoft.com/download). The published binary is self-contained; the SDK is not required at runtime.

Framework-dependent build (the smoke scripts look for this binary; running it needs the .NET 10 runtime):

```bash
dotnet build -c Release src/officecli/officecli.csproj
# src/officecli/bin/Release/net10.0/officecli
# or src/officecli/bin/Release/net10.0/<rid>/officecli
```

Self-contained publish for the machine you are on (`./build.sh` detects the RID and writes `bin/release/`):

```bash
./build.sh            # Release, current platform (default)
./build.sh debug
./build.sh all        # every RID listed in build.sh (includes macOS)
./dev-install.sh      # publish Release for this machine and copy onto PATH
```

`./build.sh` usage is only `release`, `debug`, or `all`. Fork CI (`.github/workflows/build.yml`) publishes Linux and Windows RIDs (`linux-x64`, `linux-arm64`, `linux-musl-x64`, `linux-musl-arm64`, `win-x64`, `win-arm64`). It does not publish macOS binaries; a local `./build.sh` on macOS still can.

Check the binary you built:

```bash
officecli --help
officecli materialize --help
officecli finalize --help
officecli help docx htmlchunk
officecli help docx section
officecli help docx table-column
```

### `htmlchunk` (altChunk)

```bash
officecli create spec.docx
officecli add spec.docx /body --type htmlchunk \
  --prop html='<h2>Report</h2><p>Sales grew <b>12%</b>.</p>'
officecli add spec.docx /body --type htmlchunk --prop src=page.html --prop matchSrc=true
officecli add spec.docx "/body/tbl[1]/tr[1]/tc[2]" --type htmlchunk --prop html='<p>In a cell</p>'
```

| Property | Meaning |
| --- | --- |
| `html` | Inline markup. Aliases: `content`, `text`. Wins over `src`. |
| `src` | File to embed. Alias: `path`. Format is inferred from the extension. |
| `format` | `html` (default), `xhtml`, `mht`, `rtf`, `text`. |
| `matchSrc` | `true` writes `<w:matchSrc/>` so Word keeps the chunk’s own formatting. |
| `title` | `<title>` of the wrapper around an HTML fragment. |

Parents are `/body` (with `--index`, `--after`, or `--before`) and a body table cell. Headers, footers, footnotes, and endnotes are rejected. A fragment without `<html>` or `<!DOCTYPE>` is wrapped as UTF-8 HTML; a complete document from `src` is stored as-is. `add` returns `/body/altChunk[N]`. Until Word re-saves the file, or you run `materialize`, the chunk is one opaque node: `get` shows format and size, and `view html` shows the source as escaped text.

Walkthrough and the generated sample: [examples/word/html-chunk.md](examples/word/html-chunk.md), [examples/word/html-chunk.sh](examples/word/html-chunk.sh).

### `materialize`

```bash
officecli materialize spec.docx
officecli materialize spec.docx --strict
officecli materialize spec.docx --json
```

Headless. Does not run Microsoft Word. Converts HTML, XHTML, and plain text into native paragraphs, lists, tables (including colspan / rowspan), hyperlinks, and inline pictures. Formatting is written on the runs; `h1`–`h6` also reference Heading styles. RTF and MHT chunks are left in place. Scripts are dropped. Float, flex, grid, and media queries are not reproduced.

Pictures: a data-URI whose type the picture pipeline already accepts (`png`, `jpeg`, `gif`, `bmp`, `tiff`, `emf`, `wmf`) becomes an inline `w:drawing`. The `alt` attribute is the description. `webp`, `svg`, video, and iframe are not embedded. `http(s)` URLs are not downloaded and relative `src` values are not resolved; those stay as italic `[image: …]` text and the rest of the chunk is still converted. A data URI that fails to decode is the same kind of warning. Picture warnings do not trip `--strict`. `officecli materialize --help` is the converter contract. The `htmlchunk` element note (`officecli help docx htmlchunk`) still says images are not embedded; that sentence is behind the converter.

`--strict` leaves the file unchanged and exits 1 when any whole chunk cannot be converted (RTF, MHT, a missing part, or a chunk outside the body). Without `--strict`, convertible chunks are still written and the rest stay.

CSS borders are a subset mapped to `w:pBdr` / `w:tcBorders` (and `w:tblBorders` when every cell shares the same four-edge border). Word line styles kept: `solid`, `dashed`, `dotted`, `double`, `inset`, `outset`. `groove` is drawn as `inset` and `ridge` as `outset` (a warning). The mapping, and what is ignored (`border-radius`, inline borders, and the rest), is tabulated in [examples/word/html-chunk.md](examples/word/html-chunk.md). Those warnings do not trip `--strict`.

Nested lists: a `ul` / `ol` inside an `li` (including one wrapped in a `div` or `blockquote`) is one numbering instance. Every item in that tree shares a single `w:numId`. `w:ilvl` is the depth, clamped at 8. A list outside that tree — the next list in the body, or any list in a table cell — gets its own `numId`.

### `refresh --toc`

```bash
officecli add spec.docx /body --type toc \
  --prop title="Contents" --prop levels=1-3 \
  --prop hyperlinks=true --prop pageNumbers=true
officecli refresh spec.docx --toc
```

`--toc` rebuilds TOC entries from the current headings (titles and in-document hyperlinks) and writes every `PAGEREF` result as `0`. It does not run Word or a browser, and it does not keep page numbers that were resolved earlier. Property names for the field itself: `officecli help docx toc`.

Omit `--toc` to try Microsoft Word on Windows first, then HTML pagination when a headless browser exists. If neither can number pages, entries are still rebuilt and page numbers stay `0`. `.docx` / `.docm` only.

### Page setup

```bash
officecli set spec.docx /section[1] --prop pageSetup=a4-moderate
# alias: --prop pageSetup=tech-doc
officecli set spec.docx /section[1] --prop pageSize=a4 --prop margins=moderate
```

`a4-moderate` / `tech-doc` / `techdoc` sets A4 (210×297 mm, 11906×16838 twips) and Word Moderate body margins: top/bottom 2.54 cm (1 in / 25.4 mm), left/right 1.905 cm (0.75 in / 19.05 mm). Header and footer distance is not part of the preset. The same `pageSetup`, `pageSize`, and `margins` keys exist on the final section via the document properties (`officecli help docx document`).

`pageSize`: `a4`, `letter`, or `none` (removes `w:pgSz`). `margins` (alias key `marginPreset`): `normal`, `narrow`, `moderate`, `wide`. An explicit `pageWidth` / `pageHeight` / `margin*` in the same call overrides the preset. `orientation=landscape` swaps the edges after a size preset.

`officecli create` already writes an A4 `w:pgSz` and its own margins (not Moderate). Setting `pageSetup=a4-moderate` on the section is what applies Moderate. `finalize` will not overwrite a section that already has `w:pgSz`.

### Table helpers (ID columns)

An ID column is a short code such as `T-01`, `NFR-01`, or `FR-001`. Marking it writes `w:noWrap` on each single-span cell that starts in the column and sizes `w:gridCol` / `w:tcW` to the longest line: `max(18mm, longest × 2.0mm + 8mm)`. A Latin character counts as 1; a full-width character counts as about 1.8. Six Latin characters (`NFR-01`) → 1134 twips. Other columns are not resized. A table whose width is in twips (`dxa`) shifts by the same delta.

```bash
officecli set spec.docx "/body/tbl[1]/col[1]" --prop idColumn=true
officecli set spec.docx "/body/tbl[1]/col[2]" --prop noWrap=true
officecli set spec.docx "/body/tbl[1]/col[2]" --prop width=fit
officecli set spec.docx "/body/tbl[1]" --prop idColumns=1,3
officecli add spec.docx /body --type table --prop rows=3 --prop cols=3 --prop idColumn=1
```

| Form | Effect |
| --- | --- |
| `set …/col[C] --prop idColumn=true` | `w:noWrap` plus the fit width. `false` removes `w:noWrap` and leaves the width. An explicit `width=` on the same call is used instead of the fit. |
| `set …/col[C] --prop noWrap=true` | Wrap only. Alias `nowrap`. Does not change the width. |
| `set …/col[C] --prop width=fit` | Fit width only (`auto` is the same). Does not change wrap. |
| `set …/tbl[N] --prop idColumn=1` or `idColumns=1,3` | 1-based indexes on the table. A boolean is not accepted on the table path. |
| `add --type table --prop idColumn=1` | Same index form, applied after the grid and cells exist. |

Htmlchunk, consumed by `materialize`: a `td` / `th` with class `col-id`, or with `white-space: nowrap` (inline or from a class/id/element rule the converter understands), marks that grid column when the cell’s colspan is 1. A colspan cell keeps `w:noWrap` on itself and does not resize the column.

```html
<td class="col-id">NFR-01</td>
```

Schema: `officecli help docx table-column` and `officecli help docx table`.

### `finalize`

```bash
officecli finalize spec.docx
officecli finalize spec.docx --strict --json
officecli finalize spec.docx --no-materialize --no-page-setup
```

One headless pass for a `.docx` or `.docm`, in this order:

1. **materialize** — HTML, XHTML, and plain text, same rules as `materialize` (RTF and MHT stay).
2. **`refresh --toc`** — only when the file has a TOC field. Titles and in-document hyperlinks are rebuilt. `PAGEREF` stays `0`. No Word and no browser. Skipped when there is no TOC field. Materialize runs first, so a heading that lived inside an HTML chunk is collected.
3. **page setup** — `pageSetup=a4-moderate` only on sections that have no `w:pgSz`. A section that already has a page size is left alone, including the A4 size `officecli create` writes, so its margins stay. A section that has margins but no `w:pgSz` receives the preset’s Moderate margins as well.
4. **validate** — OpenXML schema. Errors are reported and the command exits 1. Earlier steps are kept.

Skip a step with `--no-materialize`, `--no-toc`, `--no-page-setup`, or `--no-validate`. `--strict` is materialize’s `--strict`: the file is left unchanged, later steps do not run, and the command exits 1. If the TOC rebuild fails, later steps are not run either. `--json` prints one envelope; `data.steps[]` uses `ran`, `skipped`, `failed`, or `not-run`.

This command does not calculate real page numbers. Apply `pageSetup=a4-moderate` yourself while authoring if the section already has `w:pgSz`.

### Skills

Both skills ship in the repo and are registered in `SkillMap` (`src/officecli/Core/SkillInstaller.cs`). Names stay generic. `dokumen-teknis-officecli` is the build guide for a technical Word document (cover, TOC, htmlchunk, page setup, ID columns, `finalize`). `rule-gaya-dokumen-teknis` is the short style checklist that points at that guide.

```bash
officecli skills list
officecli skills install dokumen-teknis-officecli
officecli skills install rule-gaya-dokumen-teknis
officecli load_skill dokumen-teknis-officecli
officecli load_skill rule-gaya-dokumen-teknis
```

`skills install <name>` copies the skill into every detected agent skill directory. `load_skill <name>` prints the embedded `SKILL.md` without installing it. `load_skill` with no name lists the catalog. Sources: [skills/dokumen-teknis-officecli/SKILL.md](skills/dokumen-teknis-officecli/SKILL.md), [skills/rule-gaya-dokumen-teknis/SKILL.md](skills/rule-gaya-dokumen-teknis/SKILL.md).

### Smoke examples

These scripts live under [examples/word/](examples/word/). They are not invented wrappers. The four smoke scripts expect a Release build and default `OFFICECLI` to `src/officecli/bin/Release/net10.0/officecli` or `…/linux-x64/officecli`. Override with `OFFICECLI=/path/to/officecli`. Run them from the repo root.

| Script | What it exercises |
| --- | --- |
| [examples/word/html-chunk.sh](examples/word/html-chunk.sh) | Showcase. Builds `html-chunk.docx` (HTML, RTF, text, in-cell chunk). Run from `examples/word/`. Python twin: `html-chunk.py`. |
| [examples/word/materialize-altchunk.sh](examples/word/materialize-altchunk.sh) | `materialize`, including data-URI pictures, CSS borders, shared `numId` nested lists, and `--strict` leaving RTF in place. |
| [examples/word/toc-refresh.sh](examples/word/toc-refresh.sh) | `refresh --toc`. Asserts titles and hyperlinks, and that `PAGEREF` stays `0`. |
| [examples/word/table-id-column.sh](examples/word/table-id-column.sh) | `idColumn`, `idColumns`, `noWrap`, `width=fit`. |
| [examples/word/finalize.sh](examples/word/finalize.sh) | Default `finalize` order, skip flags, `--strict`, `--json`, and the “do not overwrite an existing `w:pgSz`” rule. |

```bash
dotnet build -c Release src/officecli/officecli.csproj
bash examples/word/finalize.sh
bash examples/word/materialize-altchunk.sh
bash examples/word/toc-refresh.sh
bash examples/word/table-id-column.sh
( cd examples/word && bash html-chunk.sh )
```

### Known limits

- **TOC page numbers.** `refresh --toc` and the TOC step inside `finalize` write `PAGEREF` as `0`. Real page numbers need Microsoft Word (Update Field), or `officecli refresh` without `--toc` where Word on Windows or a headless browser can paginate. Do not treat `0` as a finished page number. `officecli view` pagination can also disagree with Word.
- **Nested-list restart.** The shared list is `hybridMultilevel` and does not write `w:lvlRestart`. ECMA-376 says a hybrid level with no `lvlRestart` does not restart, so Word may keep a nested ordered counter going across parent items (a, b under item 1, then c under item 2). `officecli view` restarts a level whose `lvlRestart` is omitted whenever a shallower level advances, and it does not special-case hybrid lists, so the preview can restart where Word continues. `start` on the outermost list is that level’s `w:start`. `start` on a nested list is one `w:startOverride` for that `ilvl` of the shared instance: it applies to the whole level, not to one sublist. `start="1"` adds no override. The `type` attribute on `ol` is not read.
- **`finalize` page setup does not restyle a document `create` already sized.** Sections with `w:pgSz` keep their margins. Set `pageSetup=a4-moderate` while building the file.
- **Materialize is not Word’s HTML importer.** Unsupported pictures, remote images, RTF/MHT, and the CSS gaps above stay warnings or leftover altChunks. `--strict` only refuses when a whole chunk cannot be converted.

<p align="center">
  <strong>🌐 Website:</strong> <a href="https://officecli.ai" target="_blank">officecli.ai</a> &nbsp;|&nbsp; <strong>💬 Community:</strong> <a href="https://discord.gg/2QAwJn7Egx" target="_blank">Discord</a>
</p>

<p align="center">
  <img src="assets/ppt-process.webp" alt="OfficeCLI creating a PowerPoint presentation on AionUi" width="100%">
</p>

<p align="center"><em>PPT creation process using OfficeCLI on <a href="https://github.com/iOfficeAI/AionUi">AionUi</a></em></p>

<p align="center"><strong>PowerPoint Presentations</strong></p>

<table>
<tr>
<td width="33%"><img src="assets/designwhatmovesyou.gif" alt="OfficeCLI design presentation (PowerPoint)"></td>
<td width="33%"><img src="assets/horizon.gif" alt="OfficeCLI business presentation (PowerPoint)"></td>
<td width="33%"><img src="assets/efforless.gif" alt="OfficeCLI tech presentation (PowerPoint)"></td>
</tr>
<tr>
<td width="33%"><img src="assets/blackhole.gif" alt="OfficeCLI space presentation (PowerPoint)"></td>
<td width="33%"><img src="assets/first-ppt-aionui.gif" alt="OfficeCLI gaming presentation (PowerPoint)"></td>
<td width="33%"><img src="assets/shiba.gif" alt="OfficeCLI creative presentation (PowerPoint)"></td>
</tr>
</table>

<p align="center">—</p>
<p align="center"><strong>Word Documents</strong></p>

<table>
<tr>
<td width="33%"><img src="assets/showcase/word1.gif" alt="OfficeCLI academic paper (Word)"></td>
<td width="33%"><img src="assets/showcase/word2.gif" alt="OfficeCLI project proposal (Word)"></td>
<td width="33%"><img src="assets/showcase/word3.gif" alt="OfficeCLI annual report (Word)"></td>
</tr>
</table>

<p align="center">—</p>
<p align="center"><strong>Excel Spreadsheets</strong></p>

<table>
<tr>
<td width="33%"><img src="assets/showcase/excel1.gif" alt="OfficeCLI budget tracker (Excel)"></td>
<td width="33%"><img src="assets/showcase/excel2.gif" alt="OfficeCLI gradebook (Excel)"></td>
<td width="33%"><img src="assets/showcase/excel3.gif" alt="OfficeCLI sales dashboard (Excel)"></td>
</tr>
</table>

<p align="center"><em>All documents above were created entirely by AI agents using OfficeCLI — no templates, no manual editing.</em></p>

## For AI Agents — Get Started in One Line

Paste this into your AI agent's chat — it will read the skill file and install everything automatically:

```
curl -fsSL https://officecli.ai/SKILL.md
```

That's it. The skill file teaches the agent how to install the binary and use all commands.

## For Humans

**Option A — GUI:** Install [**AionUi**](https://github.com/iOfficeAI/AionUi) — a desktop app that lets you create and edit Office documents through natural language, powered by OfficeCLI under the hood. Just describe what you want, and AionUi handles the rest.

**Option B — CLI:** Download the binary for your platform from [GitHub Releases](https://github.com/iOfficeAI/OfficeCLI/releases), then run:

```bash
officecli install
```

This copies the binary to your PATH and installs the **officecli skill** into every AI coding agent it detects — Claude Code, Cursor, Windsurf, GitHub Copilot, and more. Your agent can immediately create, read, and edit Office documents on your behalf, no extra configuration needed.

## For Developers — See It Live in 30 Seconds

```bash
# 1. Install (macOS / Linux) — or: brew install officecli / npm install -g @officecli/officecli
curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash
# Windows (PowerShell): irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex

# 2. Create a blank PowerPoint
officecli create deck.pptx

# 3. Start live preview — opens http://localhost:26315 in your browser
officecli watch deck.pptx

# 4. Open another terminal, add a slide — watch the browser update instantly
officecli add deck.pptx / --type slide --prop title="Hello, World!"
```

That's it. Every `add`, `set`, or `remove` command you run will refresh the preview in real time. Keep experimenting — the browser is your live feedback loop.

## Quick Start

```bash
# Create a presentation and add content
officecli create deck.pptx
officecli add deck.pptx / --type slide --prop title="Q4 Report" --prop background=1A1A2E
officecli add deck.pptx '/slide[1]' --type shape \
  --prop text="Revenue grew 25%" --prop x=2cm --prop y=5cm \
  --prop font=Arial --prop size=24 --prop color=FFFFFF

# View as outline
officecli view deck.pptx outline
# → Slide 1: Q4 Report
# →   Shape 1 [TextBox]: Revenue grew 25%

# View as HTML — opens a rendered preview in your browser, no server needed
officecli view deck.pptx html

# Get structured JSON for any element
officecli get deck.pptx '/slide[1]/shape[1]' --json

# Save and close — flushes the resident session to disk
officecli close deck.pptx
```

```json
{
  "tag": "shape",
  "path": "/slide[1]/shape[1]",
  "attributes": {
    "name": "TextBox 1",
    "text": "Revenue grew 25%",
    "x": "720000",
    "y": "1800000"
  }
}
```

## Why OfficeCLI?

What used to take 50 lines of Python and 3 separate libraries:

```python
from pptx import Presentation
from pptx.util import Inches, Pt
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[0])
title = slide.shapes.title
title.text = "Q4 Report"
# ... 45 more lines ...
prs.save('deck.pptx')
```

Now takes one command:

```bash
officecli add deck.pptx / --type slide --prop title="Q4 Report"
```

**What OfficeCLI can do:**

- **Create** documents from scratch -- blank or with content
- **Read** text, structure, styles, formulas -- in plain text or structured JSON
- **Analyze** formatting issues, style inconsistencies, and structural problems
- **Modify** any element -- text, fonts, colors, layout, formulas, charts, images
- **Reorganize** content -- add, remove, move, copy elements across documents

| Format | Read | Modify | Create |
|--------|------|--------|--------|
| Word (.docx) | ✅ | ✅ | ✅ |
| Excel (.xlsx) | ✅ | ✅ | ✅ |
| PowerPoint (.pptx) | ✅ | ✅ | ✅ |

**Word** — full [i18n & RTL support](https://github.com/iOfficeAI/OfficeCLI/wiki/i18n) (per-script font slots, per-script BCP-47 lang tags `lang.latin/ea/cs`, complex-script bold/italic/size, `direction=rtl` cascading through paragraph/run/section/table/style/header/footer/docDefaults, `rtlGutter` + `pgBorders` shorthand, locale-aware page numbering for Hindi/Arabic/Thai/CJK; `create --locale ar-SA` auto-enables RTL), [paragraphs](https://github.com/iOfficeAI/OfficeCLI/wiki/word-paragraph) (framePr, tabs shorthand, char-based indents), [runs](https://github.com/iOfficeAI/OfficeCLI/wiki/word-run) (underline.color, position half-pts), [tables](https://github.com/iOfficeAI/OfficeCLI/wiki/word-table) (virtual column ops add/remove/move/copyfrom, hMerge), [styles](https://github.com/iOfficeAI/OfficeCLI/wiki/word-style), [textbox](https://github.com/iOfficeAI/OfficeCLI/wiki/word-textbox) / [shape](https://github.com/iOfficeAI/OfficeCLI/wiki/word-shape) (textbox: rotation, `textDirection` `eaVert`/`vert270`, gradient, shadow, opacity), [headers/footers](https://github.com/iOfficeAI/OfficeCLI/wiki/word-header-footer), [images](https://github.com/iOfficeAI/OfficeCLI/wiki/word-picture) (PNG/JPG/GIF/SVG), [equations](https://github.com/iOfficeAI/OfficeCLI/wiki/word-equation) (LaTeX input), [diagrams](https://github.com/iOfficeAI/OfficeCLI/wiki/diagram) (mermaid → native editable shapes, or any mermaid type as a full-fidelity PNG), [comments](https://github.com/iOfficeAI/OfficeCLI/wiki/word-comment), [footnotes](https://github.com/iOfficeAI/OfficeCLI/wiki/word-footnote), [watermarks](https://github.com/iOfficeAI/OfficeCLI/wiki/word-watermark), [bookmarks](https://github.com/iOfficeAI/OfficeCLI/wiki/word-bookmark), [TOC](https://github.com/iOfficeAI/OfficeCLI/wiki/word-toc), [charts](https://github.com/iOfficeAI/OfficeCLI/wiki/word-chart), [hyperlinks](https://github.com/iOfficeAI/OfficeCLI/wiki/word-hyperlink), [sections](https://github.com/iOfficeAI/OfficeCLI/wiki/word-section), [form fields](https://github.com/iOfficeAI/OfficeCLI/wiki/word-formfield), [content controls (SDT)](https://github.com/iOfficeAI/OfficeCLI/wiki/word-sdt), [fields](https://github.com/iOfficeAI/OfficeCLI/wiki/word-field) (22 zero-param types + MERGEFIELD / REF / PAGEREF / SEQ / STYLEREF / DOCPROPERTY / IF), [OLE objects](https://github.com/iOfficeAI/OfficeCLI/wiki/word-ole), [revisions / tracked changes](https://github.com/iOfficeAI/OfficeCLI/wiki/word-revision) (`revision.type=ins\|del\|format\|moveFrom\|moveTo` + `revision.action=accept\|reject`, per-target `/revision[@author=Alice]` selector, tracked Find&Replace), page background color, [document properties](https://github.com/iOfficeAI/OfficeCLI/wiki/word-document)

**Excel** — [cells](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-cell) (phonetic guide / furigana on add, Excel-UI `--shift left\|up` on remove / `shift=right\|down` on add), formulas (350+ built-in functions with auto-evaluation, spilling dynamic arrays with `_xlfn.` auto-prefix, financial / bond and statistical families, OFFSET/INDIRECT, defined-name formula bodies inlined at parse, formula-ref rewrite on row/col insert), [sheets](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-sheet) (visible/hidden/veryHidden, print margins, printTitleRows/Cols, RTL `sheetView`, cascade-aware sheet rename, empty-cell bloat filter on open), boolean `and`/`or` selectors (`row[Salary>5000 and Region=EMEA]`), [tables](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-table), [sort](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-sort) (sheet / range, multi-key, sidecar-aware), [conditional formatting](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-conditionalformatting), [charts](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-chart) (including box-whisker, [pareto](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-chart-add) with auto-sort + cumulative-%, log axis), [pivot tables](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-pivottable) (multi-field, date grouping, showDataAs, sort, grandTotals, subtotals, compact/outline/tabular layout, repeat item labels, blank rows, calculated fields, persistent `labelFilter` / `topN` filters, cache CoW + cross-pivot sharing), [slicers](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-slicer), [named ranges](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-namedrange), [data validation](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-validation), [images](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-picture) (PNG/JPG/GIF/SVG with dual-representation fallback), [sparklines](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-sparkline), [comments](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-comment) (RTL), [autofilter](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-autofilter), [shapes](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-shape), [OLE objects](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-ole), CSV/TSV import, `$Sheet:A1` cell addressing

**PowerPoint** — [slides](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-slide) (header/footer/date/slidenum toggles, hidden), [shapes](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-shape) (pattern fill, blur effect, hyperlink tooltip + slide-jump links, **highlight color** on runs, `slideMaster`/`slideLayout` typed add/set/remove, arrow alias, effective.X + effective.X.src), [images](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-picture) (PNG/JPG/GIF/SVG, fill modes: stretch/contain/cover/tile, brightness/contrast/glow/shadow, rotation, link + tooltip), [tables](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-table) (built-in PowerPoint style catalogue, virtual `/col[C]` get + swap/copyFrom, row/col Move/CopyFrom, fill/background alias), [charts](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-chart) (pieOfPie, barOfPie, per-attr axisLine/gridline setters, series add/remove with theme palette, `anchor=x,y,w,h` shorthand), [animations](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-shape-set) (15 emphasis + 16 exit template-backed presets, multi-effect chains, motion-path presets, repeat/restart/autoReverse, chart animations + `chartBuild`), [transitions](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-morph-check) (morph + p14 + 12 p15 PowerPoint 2013+ presets), [3D models (.glb)](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-3dmodel) (combined `rotation=ax,ay,az`), [slide zoom](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-zoom), [equations](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-equation) (LaTeX input), [diagrams](https://github.com/iOfficeAI/OfficeCLI/wiki/diagram) (mermaid flowchart / sequence → native editable shapes, or any mermaid type as a full-fidelity PNG), [themes](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-theme), [connectors](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-connector) (`from`/`to` accept a full `/slide[N]/shape[@name=Foo]` path), [video/audio](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-video) (loop, autoStart), [groups](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-group) (link + tooltip; Get/Query/Add/Remove all descend into groups), [notes](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-notes) (RTL, lang), [comments](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-comment) (RTL, legacy + modern p188 threaded round-trip), SmartArt (round-trip via add-part + raw-set), [OLE objects](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-ole), [placeholders](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-placeholder) (add/set by phType)

## Use Cases

**For Developers:**
- Automate report generation from databases or APIs
- Batch-process documents (bulk find/replace, style updates)
- Build document pipelines in CI/CD environments (generate docs from test results)
- Headless Office automation in Docker/containerized environments

**For AI Agents:**
- Generate presentations from user prompts (see examples above)
- Extract structured data from documents to JSON
- Validate and check document quality before delivery

**For Teams:**
- Clone document templates and populate with data
- Automated document validation in CI/CD pipelines

## Installation

Ships as a single self-contained binary. The .NET runtime is embedded -- nothing to install, no runtime to manage.

**One-line install:**

```bash
# macOS / Linux
curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash

# Windows (PowerShell)
irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex
```

**Or via a package manager:**

```bash
# Homebrew (macOS / Linux)
brew install officecli

# Scoop (Windows)
scoop install officecli

# npm (all platforms — fetches the native binary for your platform)
npm install -g @officecli/officecli
```

**Or download manually** from [GitHub Releases](https://github.com/iOfficeAI/OfficeCLI/releases):

| Platform | Binary |
|----------|--------|
| macOS Apple Silicon | `officecli-mac-arm64` |
| macOS Intel | `officecli-mac-x64` |
| Linux x64 | `officecli-linux-x64` |
| Linux ARM64 | `officecli-linux-arm64` |
| Windows x64 | `officecli-win-x64.exe` |
| Windows ARM64 | `officecli-win-arm64.exe` |

Verify installation: `officecli --version`

**Or self-install from a downloaded binary (or run bare `officecli` to auto-install):**

```bash
officecli install    # explicit
officecli            # bare invocation also triggers install
```

Updates are checked automatically in the background. Disable with `officecli config autoUpdate false` or skip per-invocation with `OFFICECLI_SKIP_UPDATE=1`. Configuration lives under `~/.officecli/config.json`.

## Key Features

### Built-in Engines & Generation Primitives

OfficeCLI is self-contained. The capabilities below ship inside the binary — **no Office required**.

#### Rendering engine — high-fidelity, built-in

OfficeCLI's keystone: a from-scratch, high-fidelity HTML rendering engine that lets an AI agent *see* the rendered document instead of guessing from the DOM. It covers shapes, charts (trendlines, error bars, waterfall, candlestick, sparklines), equations (OMML → LaTeX, rendered with KaTeX), 3D `.glb` models via Three.js, morph transitions, slide zoom, and shape effects. Per-page PNG screenshots are produced by piping the rendered HTML through a headless browser. Three modes:

- **`view html`** — standalone HTML file, assets inlined. Open in any browser.
- **`view screenshot`** — per-page PNG, ready for multimodal agents to read.
- **`watch`** — local HTTP server with auto-refreshing preview; every `add` / `set` / `remove` updates the browser instantly. Excel watch supports inline cell editing and drag-to-reposition charts.

```bash
officecli view deck.pptx html -o /tmp/deck.html
officecli view deck.pptx screenshot -o /tmp/deck.png # add --page 1-N for more slides
officecli watch deck.pptx                            # http://localhost:26315
```

> Without visualization, an agent generating slides is flying blind — it can read the DOM but can't tell if the title overflows or two shapes overlap. Because rendering is built into the binary, the *render → look → fix* loop works in CI, in Docker, on a server with no display — anywhere the binary runs.

#### Formula & pivot engine

350+ built-in Excel functions evaluated automatically on write — write `=SUM(A1:A2)`, `get` the cell, the value is already there. No round-trip through Office to recalc. Covers spilling dynamic arrays (`FILTER` / `SORT` / `UNIQUE` / `SEQUENCE` / `LET` / `LAMBDA` / `MAP`), `VLOOKUP` / `XLOOKUP` / `INDEX` / `MATCH`, financial & bond math (`XIRR` / `PRICE` / `YIELD` / `DURATION` / `COUPNUM`), statistical distributions, tests & regression (`NORM.DIST` / `T.TEST` / `LINEST`), and date & text functions.

Plus native OOXML pivot tables from a source range with one command — multi-field rows/cols/filters, 10 aggregations, `showDataAs` modes, date grouping, calculated fields, top-N, layouts. Pivot cache + definition are written to OOXML, so Excel opens the file with the aggregation already populated:

```bash
officecli add sales.xlsx '/Sheet1' --type pivottable \
  --prop source='Data!A1:E10000' --prop rows='Region,Category' \
  --prop cols=Quarter --prop values='Revenue:sum,Units:avg' \
  --prop showDataAs=percentOfTotal
```

#### Template merge — generate once, fill many

`merge` replaces `{{key}}` placeholders in any `.docx` / `.xlsx` / `.pptx` with JSON data — across paragraphs, table cells, shapes, headers, footers, and chart titles. Agent designs the layout once (expensive); production code fills it N times (cheap, deterministic, zero token cost). Avoids the failure mode where an agent regenerates each report from scratch and produces N inconsistent layouts.

```bash
officecli merge invoice-template.docx out-001.docx --data '{"client":"Acme","total":"$5,200"}'
officecli merge q4-template.pptx q4-acme.pptx --data data.json
```

#### Round-trip dump — learn from existing docs

`dump` serializes any `.docx`, `.pptx`, or `.xlsx` — whole document **or any subtree** (a single paragraph, table, slide, worksheet, the styles part, numbering, theme, or settings) — into a replayable batch JSON; `batch` replays it. Given a sample the user wants to imitate, an agent reads the structured spec instead of raw OOXML XML, mutates, and replays. Bridges "I have an existing template" and "generate me 100 variations."

```bash
officecli dump existing.docx -o blueprint.json                  # whole document
officecli dump existing.docx /body/tbl[1] -o table.json         # any subtree
officecli dump existing.xlsx /Sheet1 -o sheet.json              # a single worksheet
officecli batch new.docx --input blueprint.json
```

### Resident Mode & Batch

For multi-step workflows, resident mode keeps the document in memory. Batch mode applies multiple operations in a single pass.

```bash
# Resident mode — near-zero latency via named pipes
officecli open report.docx
officecli set report.docx /body/p[1]/r[1] --prop bold=true
officecli set report.docx /body/p[2]/r[1] --prop color=FF0000
officecli close report.docx

# Batch mode — multi-command execution (atomic by default: any failed item rolls back the whole batch)
echo '[{"command":"set","path":"/slide[1]/shape[1]","props":{"text":"Hello"}},
      {"command":"set","path":"/slide[1]/shape[2]","props":{"fill":"FF0000"}}]' \
  | officecli batch deck.pptx --json

# Inline batch with --commands (no stdin needed)
officecli batch deck.pptx --commands '[{"op":"set","path":"/slide[1]/shape[1]","props":{"text":"Hi"}}]'

# Keep whatever succeeds even if some items fail (pre-1.0.137 behavior)
officecli batch deck.pptx --input updates.json --best-effort --json

# Stop at the first failing command instead of running the rest (still rolls back everything unless combined with --best-effort)
officecli batch deck.pptx --input updates.json --stop-on-error --json
```

> **Reading the file with another tool? Flush to disk first.**
> officecli's own reads (`get`/`query`/`view`) always see your latest edits, so within officecli you never need to save. But a live resident defers the disk write, so **before a non‑officecli program reads the file** — python‑docx/openpyxl, Microsoft Word, a renderer, delivery/upload — flush it:
> ```bash
> officecli set report.docx /body/p[1] --prop bold=true
> officecli save report.docx           # flush, keep the resident warm (or `close` to flush + release)
> python my_reader.py report.docx      # now sees the edit
> ```
> A live resident also auto‑flushes shortly after going idle (adaptive 2–10s, scaled to the document's measured save cost). For a pipeline where another program reads after every command, set `OFFICECLI_RESIDENT_FLUSH=each` — every mutation is on disk before the command returns, while the resident stays warm. Full flush model (`each`/`auto`/fixed/`off`, save / close, env tuning): [wiki → open / close](https://github.com/iOfficeAI/OfficeCLI/wiki/command-open#when-the-file-on-disk-is-refreshed).

### Three-Layer Architecture

Start simple, go deep only when needed.

| Layer | Purpose | Commands |
|-------|---------|----------|
| **L1: Read** | Semantic views of content | `view` (text, annotated, outline, stats, issues, html, svg, screenshot) |
| **L2: DOM** | Structured element operations | `get`, `query`, `set`, `add`, `remove`, `move`, `swap` |
| **L3: Raw XML** | Direct XPath access — universal fallback | `raw`, `raw-set`, `add-part`, `validate` |

```bash
# L1 — high-level views
officecli view report.docx annotated
officecli view budget.xlsx text --cols A,B,C --max-lines 50

# L2 — element-level operations
officecli query report.docx "run:contains(TODO)"
officecli add budget.xlsx / --type sheet --prop name="Q2 Report"
officecli move report.docx /body/p[5] --to /body --index 1

# L3 — raw XML when L2 isn't enough
officecli raw deck.pptx '/slide[1]'
officecli raw-set report.docx document \
  --xpath "//w:p[1]" --action append \
  --xml '<w:r><w:t>Injected text</w:t></w:r>'
```

## AI Integration

### MCP Server

Built-in [MCP](https://modelcontextprotocol.io) server — register with one command:

```bash
officecli mcp claude       # Claude Code
officecli mcp cursor       # Cursor
officecli mcp vscode       # VS Code / Copilot
officecli mcp lmstudio     # LM Studio
officecli mcp list         # Check registration status
```

Exposes all document operations as tools over JSON-RPC — no shell access needed.

### Direct CLI Integration

Get OfficeCLI working with your AI agent in two steps:

1. **Install the binary** -- one command (see [Installation](#installation))
2. **Done.** OfficeCLI automatically detects your AI tools (Claude Code, GitHub Copilot, Codex) by checking known config directories and installs its skill file. Your agent can immediately create, read, and modify any Office document.

<details>
<summary><strong>Manual setup (optional)</strong></summary>

If auto-install doesn't cover your setup, you can install the skill file manually:

**Feed SKILL.md to your agent directly:**

```bash
curl -fsSL https://officecli.ai/SKILL.md
```

**Install as a local skill for Claude Code:**

```bash
curl -fsSL https://officecli.ai/SKILL.md -o ~/.claude/skills/officecli.md
```

**Other agents:** Include the contents of `SKILL.md` in your agent's system prompt or tool description.

</details>

### Why your agent will thrive on OfficeCLI

- **Deterministic JSON output** — every command supports `--json` with consistent schemas. No regex parsing, no scraping stdout.
- **Path-based addressing** — every element has a stable path (`/slide[1]/shape[2]`). Agents navigate documents without understanding XML namespaces. (OfficeCLI syntax: 1-based indexing, element local names — not XPath.)
- **Progressive complexity (L1 → L2 → L3)** — agents start with read-only views, escalate to DOM ops, fall back to raw XML only when needed. Minimizes token usage.
- **Self-healing workflow** — `validate`, `view issues`, and the structured error codes (`not_found`, `invalid_value`, `unsupported_property`) return suggestions and valid ranges. Agents self-correct without human intervention.
- **Built-in agent-friendly rendering engine** — `view html` / `view screenshot` / `watch` emit HTML and PNG natively. No Office required. Agents can *see* their output and fix layout issues, even inside CI / Docker / headless environments.
- **Built-in formula & pivot engine** — 350+ Excel functions auto-evaluated on write (incl. spilling dynamic arrays, financial / bond and statistical families); native OOXML pivot tables from a source range with one command. Agents read computed values and shipped aggregations immediately, without round-tripping through Office.
- **Template merge** — agent designs the layout once, downstream code fills `{{key}}` placeholders N times. Avoids burning tokens regenerating every report from scratch.
- **Round-trip dump** — `dump` turns any `.docx`, `.pptx`, or `.xlsx` into replayable batch JSON. Agents learn from human-authored samples by reading a structured spec, not raw OOXML XML.
- **Built-in help** — when unsure about property names or value formats, the agent runs `officecli <format> set <element>` instead of guessing.
- **Auto-install** — OfficeCLI detects your AI tooling (Claude Code, Cursor, VS Code, …) and configures itself. No manual skill-file setup.

### Built-in Help

Don't guess property names — drill into the help:

```bash
officecli help pptx set              # All settable elements and properties
officecli help pptx set shape        # Detail for one element type
officecli help docx query            # Selector reference: attributes, :contains, :has(), etc.
```

Run `officecli --help` for the full overview.

### JSON Output Schemas

All commands support `--json`. The general response shapes:

**Single element** (`get --json`):

```json
{"tag": "shape", "path": "/slide[1]/shape[1]", "attributes": {"name": "TextBox 1", "text": "Hello"}}
```

**List of elements** (`query --json`):

```json
[
  {"tag": "paragraph", "path": "/body/p[1]", "attributes": {"style": "Heading1", "text": "Title"}},
  {"tag": "paragraph", "path": "/body/p[5]", "attributes": {"style": "Heading1", "text": "Summary"}}
]
```

**Errors** return a non-zero exit code with a structured error object including error code, suggestion, and valid values when available:

```json
{
  "success": false,
  "error": {
    "error": "Slide 50 not found (total: 8)",
    "code": "not_found",
    "suggestion": "Valid Slide index range: 1-8"
  }
}
```

Error codes: `not_found`, `invalid_value`, `unsupported_property`, `invalid_path`, `unsupported_type`, `missing_property`, `file_not_found`, `file_locked`, `invalid_selector`. Property names are auto-corrected -- misspelling a property returns a suggestion with the closest match.

**Error Recovery** -- Agents self-correct by inspecting available elements:

```bash
# Agent tries an invalid path
officecli get report.docx /body/p[99] --json
# Returns: {"success": false, "error": {"error": "...", "code": "not_found", "suggestion": "..."}}

# Agent self-corrects by checking available elements
officecli get report.docx /body --depth 1 --json
# Returns the list of available children, agent picks the right path
```

**Mutation confirmations** (`set`, `add`, `remove`, `move`, `create` with `--json`):

```json
{"success": true, "path": "/slide[1]/shape[1]"}
```

**Exit codes.** `0` — the operation landed. `1` — it did not (process or business failure; the envelope carries `error`). `2` — it ran with caveats: an `unsupported_property` was dropped, or a validator raised a warning. Exit 2 covers both "landed with a warning" (`success: true`) and "nothing applied" (`success: false`, e.g. every property refused), so scripts must read `success` in the envelope, not the exit code alone; `warnings[]` lists what was dropped. The MCP server derives `isError` from the envelope the same way.

## Comparison

| | OfficeCLI | Microsoft Office | LibreOffice | python-docx / openpyxl |
|---|---|---|---|---|
| Open source & free | ✓ (Apache 2.0) | ✗ (paid license) | ✓ | ✓ |
| AI-native CLI + JSON | ✓ | ✗ | ✗ | ✗ |
| Zero install (single binary) | ✓ | ✗ | ✗ | ✗ (Python + pip) |
| Call from any language | ✓ (CLI) | ✗ (COM/Add-in) | ✗ (UNO API) | Python only |
| Path-based element access | ✓ | ✗ | ✗ | ✗ |
| Raw XML fallback | ✓ | ✗ | ✗ | Partial |
| Built-in agent-friendly rendering engine | ✓ | ✗ | ✗ | ✗ |
| Headless HTML/PNG output | ✓ | ✗ | Partial | ✗ |
| Template merge (`{{key}}`) across formats | ✓ | ✗ | ✗ | ✗ |
| Round-trip dump → batch JSON | ✓ | ✗ | ✗ | ✗ |
| Live preview (auto-refresh on edit) | ✓ | ✗ | ✗ | ✗ |
| Headless / CI | ✓ | ✗ | Partial | ✓ |
| Cross-platform | ✓ | Windows/Mac | ✓ | ✓ |
| Word + Excel + PowerPoint | ✓ | ✓ | ✓ | Separate libs |

## Command Reference

| Command | Description |
|---------|-------------|
| [`create`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-create) | Create a blank .docx, .xlsx, or .pptx (type from extension) |
| [`view`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-view) | View content (modes: `outline`, `text`, `annotated`, `stats` (`--page-count`), `issues`, `html`, `svg`, `screenshot`, `pdf` (via exporter plugin), `forms` (via format-handler plugin)). docx supports `--render auto\|native\|html`. |
| [`load_skill`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-skills) | Print embedded SKILL.md content for a specialized skill (no install) |
| [`get`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-get) | Get element and children (`--depth N`, `--json`) |
| [`query`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-query) | CSS-like query with boolean `and`/`or`, row-by-column-name (`row[Salary>5000]`), `--find` flag |
| [`set`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-set) | Modify element properties; accepts selectors and Excel-native paths (parity with `get`/`query`), `--find`/`--replace` flags |
| [`add`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-add) | Add element (or clone with `--from <path>`) |
| [`remove`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-remove) | Remove an element |
| [`move`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-move) | Move element (`--to <parent>`, `--index N`, `--after <path>`, `--before <path>`) |
| [`swap`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-swap) | Swap two elements |
| [`validate`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-validate) | Validate against OpenXML schema |
| `view <file> issues` | Enumerate document issues (text overflow, missing alt text, formula errors, ...) |
| [`batch`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-batch) | Multiple operations applied in a single pass (stdin, `--input`, or `--commands`; atomic by default — any failed item rolls back the whole batch — `--best-effort` to keep partial progress, `--stop-on-error` to abort early) |
| [`dump`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-dump) | Serialize a `.docx`, `.pptx`, or `.xlsx` into a replayable batch JSON (round-trip via `batch`); accepts a subtree path |
| [`refresh`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-refresh) | Recalculate derived fields (`.docx` / `.docm`). `--toc` rebuilds TOC entries and writes `PAGEREF` as `0` (no Word, no browser). Without `--toc`: Word on Windows, then HTML pagination; entries are still rebuilt if neither can number pages. See [This fork](#refresh---toc). |
| `materialize` | Replace HTML / XHTML / plain-text `w:altChunk` parts with native OOXML. RTF and MHT stay. `--strict` refuses the whole file when a chunk cannot be converted. `--json`. See [This fork](#materialize). |
| `finalize` | Headless technical-doc pass: materialize → `refresh --toc` if a TOC field exists → `pageSetup=a4-moderate` only on sections with no `w:pgSz` → validate. Skip with `--no-materialize`, `--no-toc`, `--no-page-setup`, `--no-validate`. `--strict`, `--json`. See [This fork](#finalize). |
| [`plugins`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-plugins) | List / inspect / lint installed plugins (extend to `.doc`, `.hwpx`, `.pdf` export via dump-reader / exporter / format-handler kinds) |
| [`merge`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-merge) | Template merge — replace `{{key}}` placeholders with JSON data |
| [`watch`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-watch) | Live HTML preview in browser with auto-refresh |
| [`mcp`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-mcp) | Start MCP server for AI tool integration |
| [`raw`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-raw) | View raw XML of a document part |
| [`raw-set`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-raw) | Modify raw XML via XPath |
| `add-part` | Add a new document part (header, chart, etc.) |
| [`open`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-open) | Start resident mode (keep document in memory) |
| `close` | Save and close resident mode |
| [`install`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-install) | Install binary + skills + MCP (`all`, `claude`, `cursor`, etc.) |
| `config` | Get or set configuration |
| `help <format> <command>` | [Built-in help](https://github.com/iOfficeAI/OfficeCLI/wiki/command-reference) (e.g. `officecli help pptx set shape`) |

## End-to-End Workflow Example

A typical self-healing agent workflow: create a presentation, populate it, verify, and fix issues -- all without human intervention.

```bash
# 1. Create
officecli create report.pptx

# 2. Add content
officecli add report.pptx / --type slide --prop title="Q4 Results"
officecli add report.pptx '/slide[1]' --type shape \
  --prop text="Revenue: $4.2M" --prop x=2cm --prop y=5cm --prop size=28
officecli add report.pptx / --type slide --prop title="Details"
officecli add report.pptx '/slide[2]' --type shape \
  --prop text="Growth driven by new markets" --prop x=2cm --prop y=5cm

# 3. Verify
officecli view report.pptx outline
officecli validate report.pptx

# 4. Fix any issues found
officecli view report.pptx issues --json
# Address issues based on output, e.g.:
officecli set report.pptx '/slide[1]/shape[1]' --prop font=Arial
```

### Units & Colors

All dimension and color properties accept flexible input formats:

| Type | Accepted formats | Examples |
|------|-----------------|----------|
| **Dimensions** | cm, in, pt, px, or raw EMU | `2cm`, `1in`, `72pt`, `96px`, `914400` |
| **Colors** | Hex, named, RGB, theme | `#FF0000`, `FF0000`, `red`, `rgb(255,0,0)`, `accent1` |
| **Font sizes** | Bare number or pt-suffixed | `14`, `14pt`, `10.5pt` |
| **Spacing** | pt, cm, in, or multiplier | `12pt`, `0.5cm`, `1.5x`, `150%` |

## Common Patterns

```bash
# Replace all Heading1 text in a Word doc
officecli query report.docx "paragraph[style=Heading1]" --json | ...
officecli set report.docx /body/p[1]/r[1] --prop text="New Title"

# Export all slide content as JSON
officecli get deck.pptx / --depth 2 --json

# Bulk-update Excel cells
officecli batch budget.xlsx --input updates.json --json

# Import CSV data into an Excel sheet
officecli add budget.xlsx / --type sheet --prop name="Q1 Data"
officecli import budget.xlsx "/Q1 Data" sales.csv --header

# Template merge for batch reports
officecli merge invoice-template.docx invoice-001.docx --data '{"client":"Acme","total":"$5,200"}'

# Check document quality before delivery
officecli validate report.docx && officecli view report.docx issues --json
```

**From Python or Node.js** — install one of the thin resident-pipe SDKs (no per-call process spawn):

```python
# Python — `pip install officecli-sdk`
import officecli
with officecli.create("deck.pptx") as doc:          # or officecli.open("deck.pptx")
    doc.send({"command": "add", "parent": "/", "type": "slide"})
    print(doc.send({"command": "get", "path": "/slide[1]"}))
```

```javascript
// Node.js — `npm install @officecli/sdk`
const oc = require("@officecli/sdk");
const doc = await oc.create("deck.pptx");            // or oc.open("deck.pptx")
await doc.send({ command: "add", parent: "/", type: "slide" });
console.log(await doc.send({ command: "get", path: "/slide[1]" }));
await doc.close();
```

Both SDKs auto-provision the native CLI when missing (mirror-first, Windows-capable) and announce the install rather than doing it silently.

Or wrap subprocess directly, one-shot:

```python
import json, subprocess
def cli(*args):
    return json.loads(subprocess.check_output(["officecli", *args, "--json"], text=True))
cli("create", "deck.pptx")
```

## Documentation

The [Wiki](https://github.com/iOfficeAI/OfficeCLI/wiki) has detailed guides for every command, element type, and property:

- **By format:** [Word](https://github.com/iOfficeAI/OfficeCLI/wiki/word-reference) | [Excel](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-reference) | [PowerPoint](https://github.com/iOfficeAI/OfficeCLI/wiki/powerpoint-reference)
- **Workflows:** [End-to-end examples](https://github.com/iOfficeAI/OfficeCLI/wiki/workflows) -- Word reports, Excel dashboards, PowerPoint decks, batch modifications, resident mode
- **This fork:** [headless Word technical documents](#this-fork-headless-word-technical-documents) — `htmlchunk`, `materialize`, `refresh --toc`, page presets, ID columns, `finalize`, and the two in-repo skills
- **Runnable examples:** [examples/](examples/) -- copy-paste scripts (.sh/.py) for Word, Excel, and PowerPoint, with output files included. Fork smoke scripts: [examples/word/](examples/word/) (`finalize.sh`, `materialize-altchunk.sh`, `toc-refresh.sh`, `table-id-column.sh`)
- **Troubleshooting:** [Common errors and solutions](https://github.com/iOfficeAI/OfficeCLI/wiki/troubleshooting)
- **AI agent guide:** [Decision tree for navigating the wiki](https://github.com/iOfficeAI/OfficeCLI/wiki/agent-guide)

## Build from Source

Requires the [.NET 10 SDK](https://dotnet.microsoft.com/download) to compile. `./build.sh` publishes a self-contained binary (the runtime is embedded). `dotnet build` without `-r` is framework-dependent and needs the .NET 10 runtime to run. Fork-specific commands, smoke scripts, and the CI RID list are in [Build and run](#build-and-run).

```bash
dotnet build -c Release src/officecli/officecli.csproj
./build.sh          # self-contained Release for this machine → bin/release/
./dev-install.sh    # publish and copy onto PATH
```

## License

[Apache License 2.0](LICENSE)

Bug reports and contributions are welcome on [GitHub Issues](https://github.com/iOfficeAI/OfficeCLI/issues).

---

If you find OfficeCLI useful, please [give it a star on GitHub](https://github.com/iOfficeAI/OfficeCLI) — it helps others discover the project.

[OfficeCLI.AI](https://OfficeCLI.AI) | [GitHub](https://github.com/iOfficeAI/OfficeCLI)

<!--
yaml-frontmatter:
  name: officecli
  version: see GitHub Releases
  description: The standard CLI for AI agents to create, read, and modify Office documents
  formats: [docx, xlsx, pptx]
  platforms: [macos, linux, windows]
  license: Apache-2.0
  keywords: [office, cli, ai-agent, automation, docx, xlsx, pptx, openxml, document-generation, mcp-server, word, excel, powerpoint]
  ai-agent-compatible: true
  mcp-server: true
  skill-file: SKILL.md

  install-command-unix: curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash
  install-command-windows: irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex
-->

<!-- LLM/agent discovery metadata
tool: officecli
type: cli
canonical: the standard CLI for AI-driven Office document manipulation
formats: docx, xlsx, pptx
capabilities: create, read, modify, validate, batch, resident-mode, mcp-server, live-preview, auto-install
output: json, text, annotated
platforms: macos, linux, windows
license: Apache-2.0
keywords: office, cli, ai-agent, automation, docx, xlsx, pptx, openxml, document-generation, mcp-server, word, excel, powerpoint, ai-tools, command-line, structured-output
ai-agent-compatible: true
mcp-server: true
skill-file: SKILL.md
skill-file-lines: 403
alternatives: python-docx, openpyxl, python-pptx, libreoffice --headless
install-command-unix: curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash
install-command-windows: irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex
-->
