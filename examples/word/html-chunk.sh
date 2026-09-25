#!/bin/bash
# Word HTML Chunk Showcase — embed HTML (and RTF / plain text) into a document
# with `add --type htmlchunk` (aliases: html, altchunk).
# CLI twin of html-chunk.py (officecli Python SDK). Both produce an equivalent
# html-chunk.docx.
#
# An HTML chunk is stored VERBATIM in its own package part (word/afchunkN.htm)
# and referenced from the body by <w:altChunk r:id="…"/>. Word converts it to
# native paragraphs / tables / lists the first time the file is opened, so the
# HTML keeps its full CSS fidelity without OfficeCLI translating the markup.
# Until then each chunk is a single /body/altChunk[N] node (get shows its
# format / size; `view html` shows its source as escaped text).
#
# Covered: inline text formatting, headings, alignment, colors & fonts, lists
# (nested / start / types), tables (colspan / rowspan / zebra / tfoot),
# blockquote, pre/code, hr, links, a data-URI image, a complete HTML document
# from a file (src=) with <style> CSS + matchSrc, RTF and plain-text chunks
# (format=), a chunk inside a table cell, and positioned inserts (--after).
#
# Usage:
#   ./html-chunk.sh
set -e
DIR="$(cd "$(dirname "$0")" && pwd)"
FILE="$DIR/html-chunk.docx"
rm -f "$FILE"

officecli create "$FILE"
officecli open "$FILE"

h2() { officecli add "$FILE" /body --type paragraph --prop text="$1" --prop style=Heading2; }

# Native title + intro (ordinary Word paragraphs around the chunks).
officecli add "$FILE" /body --type paragraph --prop text="HTML Chunk Showcase" --prop style=Heading1 --prop align=center
officecli add "$FILE" /body --type paragraph --prop text="Every section below is an HTML (or RTF / plain-text) chunk embedded with 'add --type htmlchunk'. Word converts each chunk into native content when this file is opened."

# ==========================================================================
# 1. Inline text formatting
# ==========================================================================
h2 "1. Inline text formatting"
# Features: html= fragment (auto-wrapped in a UTF-8 document), b/i/u/s,
#   sup/sub, mark, small/big, inline style color/font/size/background, Unicode
officecli add "$FILE" /body --type htmlchunk --prop html='
<p>Teks <b>tebal</b>, <i>miring</i>, <u>garis bawah</u>, <s>coret</s>,
   <b><i><u>gabungan</u></i></b>, H<sub>2</sub>O, E = mc<sup>2</sup>,
   <mark>disorot</mark>, <small>kecil</small> dan <big>besar</big>.</p>
<p><span style="color:#C00000">merah</span>,
   <span style="color:#2E75B6;font-weight:bold">biru tebal</span>,
   <span style="background:#FFFF00">latar kuning</span>,
   <span style="font-family:Georgia,serif;font-size:16pt">Georgia 16pt</span>,
   <span style="font-family:Consolas,monospace">Consolas</span>,
   <span style="letter-spacing:3pt">renggang</span>,
   <span style="text-transform:uppercase">huruf besar</span>.</p>
<p>Unicode: café — naïve — ✓ ✗ ★ — 日本語 — العربية — Ελληνικά — € £ ¥ Rp</p>'

# ==========================================================================
# 2. Headings & paragraph layout
# ==========================================================================
h2 "2. Headings & paragraph layout"
# Features: h1–h4, text-align (left/center/right/justify), text-indent,
#   margin-left, line-height, border + padding + background on a block
officecli add "$FILE" /body --type htmlchunk --prop html='
<h1>Heading 1 dari HTML</h1>
<h2>Heading 2 dari HTML</h2>
<h3>Heading 3 dari HTML</h3>
<h4>Heading 4 dari HTML</h4>
<p style="text-align:center">Paragraf rata tengah.</p>
<p style="text-align:right">Paragraf rata kanan.</p>
<p style="text-align:justify">Paragraf rata kiri-kanan (justify): Lorem ipsum dolor sit amet,
   consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.
   Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris.</p>
<p style="text-indent:1cm">Paragraf dengan indentasi baris pertama 1 cm.</p>
<p style="margin-left:2cm">Paragraf dengan indentasi kiri 2 cm.</p>
<p style="line-height:200%">Paragraf dengan spasi baris ganda (line-height 200%).
   Baris ini sengaja cukup panjang supaya jarak antarbaris terlihat jelas.</p>
<p style="border:1px solid #2E75B6;background:#DDEBF7;padding:6px">Kotak berbingkai dengan latar biru muda.</p>'

# ==========================================================================
# 3. Lists
# ==========================================================================
h2 "3. Lists"
# Features: ul (disc/square), ol (decimal/lower-alpha/upper-roman), start=,
#   three nesting levels, mixed ol/ul nesting, inline formatting inside items
officecli add "$FILE" /body --type htmlchunk --prop html='
<ul>
  <li>Butir pertama</li>
  <li>Butir kedua dengan <b>teks tebal</b>
    <ul style="list-style-type:square">
      <li>Sub-butir (square)</li>
      <li>Sub-butir lain
        <ul><li>Level ketiga</li></ul>
      </li>
    </ul>
  </li>
</ul>
<ol start="3">
  <li>Mulai dari nomor 3</li>
  <li>Nomor 4
    <ol style="list-style-type:lower-alpha">
      <li>sub a</li>
      <li>sub b</li>
    </ol>
  </li>
</ol>
<ol style="list-style-type:upper-roman">
  <li>Romawi I</li>
  <li>Romawi II</li>
</ol>'

# ==========================================================================
# 4. Tables
# ==========================================================================
h2 "4. Tables"
# Features: colspan, rowspan, thead/tbody/tfoot, header fill, zebra rows,
#   per-cell alignment, cell colors, column widths, border-collapse
officecli add "$FILE" /body --type htmlchunk --prop html='
<table style="border-collapse:collapse;width:100%" border="1" cellpadding="4">
  <thead>
    <tr style="background:#1F4E79;color:#FFFFFF">
      <th rowspan="2" style="width:30%">Produk</th>
      <th colspan="2">Semester 1</th>
      <th colspan="2">Semester 2</th>
    </tr>
    <tr style="background:#2E75B6;color:#FFFFFF">
      <th>Unit</th><th>Omzet</th><th>Unit</th><th>Omzet</th>
    </tr>
  </thead>
  <tbody>
    <tr><td>Laptop</td><td align="right">120</td><td align="right">1.800</td><td align="right">140</td><td align="right">2.100</td></tr>
    <tr style="background:#DDEBF7"><td>Monitor</td><td align="right">300</td><td align="right">900</td><td align="right">280</td><td align="right">840</td></tr>
    <tr><td>Printer</td><td align="right">75</td><td align="right">300</td><td align="right" style="color:#C00000">60</td><td align="right" style="color:#C00000">240</td></tr>
  </tbody>
  <tfoot>
    <tr style="background:#FFF2CC;font-weight:bold">
      <td>Total</td><td align="right">495</td><td align="right">3.000</td><td align="right">480</td><td align="right">3.180</td>
    </tr>
  </tfoot>
</table>'

# ==========================================================================
# 5. Blocks: quote, code, rule, links
# ==========================================================================
h2 "5. Quote, code, horizontal rule & links"
# Features: blockquote, pre (whitespace preserved), inline code, hr,
#   absolute http link, mailto link
officecli add "$FILE" /body --type htmlchunk --prop html='
<blockquote style="border-left:4px solid #A5A5A5;padding-left:10px;color:#555">
  <p><i>"Kesederhanaan adalah kecanggihan tertinggi."</i><br>— Leonardo da Vinci</p>
</blockquote>
<p>Perintah: <code>officecli add doc.docx /body --type htmlchunk --prop src=page.html</code></p>
<pre style="background:#F2F2F2;border:1px solid #D9D9D9;padding:6px">
def halo(nama):
    return f"Halo, {nama}!"

print(halo("OfficeCLI"))
</pre>
<hr>
<p>Tautan: <a href="https://github.com/iOfficeAI/OfficeCLI">OfficeCLI di GitHub</a> ·
   <a href="mailto:hello@example.com">kirim email</a></p>'

# ==========================================================================
# 6. Image (data: URI)
# ==========================================================================
h2 "6. Image from a data: URI"
# Features: <img> with an embedded base64 data URI (relative paths are NOT
#   resolved inside a chunk — use data: or absolute http(s) URLs), width/height
LOGO=$(base64 < "$DIR/pictures-logo.png" | tr -d '\n')
officecli add "$FILE" /body --type htmlchunk --prop html="
<p><img src=\"data:image/png;base64,$LOGO\" width=\"96\" height=\"96\" alt=\"logo\">
   Gambar di samping disematkan sebagai data URI.</p>"

# ==========================================================================
# 7. Complete HTML document from a file
# ==========================================================================
h2 "7. Complete HTML document from a file (src=)"
# Features: src= file (stored byte-for-byte), <style> CSS classes, rowspan +
#   colspan header, tfoot, styled callout boxes, nested lists; matchSrc=true
#   asks Word to keep the source formatting instead of the document styles
officecli add "$FILE" /body --type htmlchunk --prop src="$DIR/html-chunk-report.html" --prop matchSrc=true

# ==========================================================================
# 8. Other alternate formats: RTF and plain text
# ==========================================================================
h2 "8. RTF and plain-text chunks (format=)"
# Features: format inferred from the .rtf / .txt extension (format= forces it)
officecli add "$FILE" /body --type htmlchunk --prop src="$DIR/html-chunk-memo.rtf"
officecli add "$FILE" /body --type htmlchunk --prop src="$DIR/html-chunk-notes.txt" --prop format=text

# ==========================================================================
# 9. HTML chunk inside a table cell
# ==========================================================================
h2 "9. HTML chunk inside a native table cell"
# Features: parent /body/tbl[N]/tr[M]/tc[K]; a trailing empty paragraph is
#   kept after the chunk (Word requires a cell to end with a paragraph)
officecli add "$FILE" /body --type table --prop rows=1 --prop cols=2
officecli set "$FILE" "/body/tbl[1]/tr[1]/tc[1]" --prop text="Sel native (teks biasa)"
officecli add "$FILE" "/body/tbl[1]/tr[1]/tc[2]" --type htmlchunk --prop html='
<p><b>Sel berisi HTML</b></p>
<ul><li>butir satu</li><li><span style="color:#2E7D32">butir dua (hijau)</span></li></ul>'

# ==========================================================================
# 10. Positioned insert
# ==========================================================================
# Features: --after anchors the chunk next to an existing paragraph (here the
#   intro paragraph p[2]); --index / --before work the same way
officecli add "$FILE" /body --type html --after "p[2]" --prop html='
<p style="background:#FFF2CC;border:1px dashed #BF9000;padding:6px">
  <b>Catatan:</b> kotak ini disisipkan belakangan dengan <code>--after p[2]</code>.</p>'

officecli close "$FILE"

# Read back: every chunk is a positional /body/altChunk[N] node.
officecli get "$FILE" /body --depth 1
officecli validate "$FILE"
echo "Created: $FILE"
