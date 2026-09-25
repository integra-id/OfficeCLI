#!/usr/bin/env python3
"""
HTML Chunk Showcase — generates html-chunk.docx exercising the docx
`htmlchunk` element (aliases: html, altchunk): HTML, RTF and plain-text chunks
embedded via <w:altChunk>, which Word converts into native content on open.

SDK twin of html-chunk.sh (officecli CLI). Both produce an equivalent
html-chunk.docx. This one drives the **officecli Python SDK**
(`pip install officecli-sdk`): one resident is started and every item is
shipped over the named pipe in a single `doc.batch(...)` round-trip. Each item
is the same `{"command","parent","type","props"}` dict you'd put in an
`officecli batch` list.

Usage:
  pip install officecli-sdk          # plus the `officecli` binary on PATH
  python3 html-chunk.py
"""

import base64
import os
import sys

# --- locate the SDK: prefer an installed `officecli-sdk`, else the in-repo copy
try:
    import officecli  # pip install officecli-sdk
except ImportError:
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                    "..", "..", "sdk", "python"))
    import officecli

DIR = os.path.dirname(os.path.abspath(__file__))
FILE = os.path.join(DIR, "html-chunk.docx")


def para(text, **props):
    """One `add paragraph` item in batch-shape."""
    return {"command": "add", "parent": "/body", "type": "paragraph",
            "props": {"text": text, **props}}


def h2(text):
    return para(text, style="Heading2")


def chunk(parent="/body", **props):
    """One `add htmlchunk` item in batch-shape."""
    return {"command": "add", "parent": parent, "type": "htmlchunk", "props": props}


INLINE = """
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
<p>Unicode: café — naïve — ✓ ✗ ★ — 日本語 — العربية — Ελληνικά — € £ ¥ Rp</p>"""

LAYOUT = """
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
<p style="border:1px solid #2E75B6;background:#DDEBF7;padding:6px">Kotak berbingkai dengan latar biru muda.</p>"""

LISTS = """
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
</ol>"""

TABLE = """
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
</table>"""

BLOCKS = """
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
   <a href="mailto:hello@example.com">kirim email</a></p>"""

with open(os.path.join(DIR, "pictures-logo.png"), "rb") as f:
    LOGO = base64.b64encode(f.read()).decode("ascii")
IMAGE = f"""
<p><img src="data:image/png;base64,{LOGO}" width="96" height="96" alt="logo">
   Gambar di samping disematkan sebagai data URI.</p>"""

CELL = """
<p><b>Sel berisi HTML</b></p>
<ul><li>butir satu</li><li><span style="color:#2E7D32">butir dua (hijau)</span></li></ul>"""

NOTE = """
<p style="background:#FFF2CC;border:1px dashed #BF9000;padding:6px">
  <b>Catatan:</b> kotak ini disisipkan belakangan dengan <code>--after p[2]</code>.</p>"""

print(f"Building {FILE} ...")

with officecli.create(FILE, "--force") as doc:
    items = [
        para("HTML Chunk Showcase", style="Heading1", align="center"),
        para("Every section below is an HTML (or RTF / plain-text) chunk embedded with "
             "'add --type htmlchunk'. Word converts each chunk into native content when "
             "this file is opened."),

        h2("1. Inline text formatting"), chunk(html=INLINE),
        h2("2. Headings & paragraph layout"), chunk(html=LAYOUT),
        h2("3. Lists"), chunk(html=LISTS),
        h2("4. Tables"), chunk(html=TABLE),
        h2("5. Quote, code, horizontal rule & links"), chunk(html=BLOCKS),
        h2("6. Image from a data: URI"), chunk(html=IMAGE),

        h2("7. Complete HTML document from a file (src=)"),
        chunk(src=os.path.join(DIR, "html-chunk-report.html"), matchSrc="true"),

        h2("8. RTF and plain-text chunks (format=)"),
        chunk(src=os.path.join(DIR, "html-chunk-memo.rtf")),
        chunk(src=os.path.join(DIR, "html-chunk-notes.txt"), format="text"),

        h2("9. HTML chunk inside a native table cell"),
        {"command": "add", "parent": "/body", "type": "table",
         "props": {"rows": "1", "cols": "2"}},
        {"command": "set", "path": "/body/tbl[1]/tr[1]/tc[1]",
         "props": {"text": "Sel native (teks biasa)"}},
        chunk(parent="/body/tbl[1]/tr[1]/tc[2]", html=CELL),

        # Positioned insert: lands right after the intro paragraph p[2].
        {"command": "add", "parent": "/body", "type": "html", "after": "p[2]",
         "props": {"html": NOTE}},
    ]

    doc.batch(items)
    print(f"  added {len(items)} items")

print(f"Generated: {FILE}")
