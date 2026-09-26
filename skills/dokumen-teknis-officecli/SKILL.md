---
name: Dokumen teknis (OfficeCLI)
description: >-
  Pakai saat membuat atau mengedit dokumen Word teknis project (panduan, spek,
  runbook, BA) agar layout tipografi warna tabel callout kode header-footer
  spasi numbering TOC mengikuti desain dokumen teknis memakai OfficeCLI htmlchunk.
---
# Dokumen teknis (OfficeCLI)

Buat dokumen Word teknis project (panduan integrasi, spesifikasi, runbook, BA, desain) yang **layout, tipografi, warna, spasi, numbering, TOC, tabel, callout, kode, header/footer-nya mengikuti desain dokumen teknis** — memakai **OfficeCLI** (`officecli`), terutama `htmlchunk`.

## Kapan dipakai

- User minta buat/update dokumen teknis `.docx` bergaya desain dokumen teknis / klien / project sejenis
- Perlu cover + meta dokumen + riwayat versi + TOC terhubung + bab bernomor + tabel + callout + blok kode
- Repo/tooling: `integra-id/OfficeCLI` (fitur `htmlchunk`) atau binary `officecli` yang mendukung `--type htmlchunk`

Jangan pakai skill ini untuk surat singkat, CV, atau dokumen yang sengaja template lain.

## Prasyarat

```bash
command -v officecli || { echo "Install officecli dulu (binary fork integra-id/OfficeCLI atau upstream)"; exit 1; }
officecli help docx htmlchunk   # pastikan elemen htmlchunk ada
officecli help docx toc         # field TOC dengan hyperlink
```

Bekerja di direktori kerja project; simpan aset logo/diagram di `assets/`.

## Design tokens (wajib)

Salin token ini ke setiap dokumen baru. Jangan mengganti palet kecuali user minta rebrand.

| Token | Nilai |
|---|---|
| Page | A4 |
| Margin | **Moderate** (Word): atas/bawah **25,4 mm (1")**, kiri/kanan **19,05 mm (0,75")** |
| Body font | Calibri 10.5 pt, warna `#1F2937` |
| Mono / kode | Consolas (atau Cascadia Mono), fill blok `#F3F4F6` |
| H1 | Calibri Bold ~15 pt, `#0F766E` |
| H2 | Calibri Bold ~12.5 pt, `#111827` |
| H3 | Calibri Bold ~11 pt, `#111827` |
| Muted | `#6B7280` |
| Teal gelap (teks th) | `#0B4F49` |
| Table header fill | `#E7F2F0` |
| Table border | `#D1D5DB` |
| Callout warn | `#FEF3C7` |
| Callout success/inti | `#ECFDF5` |
| Callout info | `#EFF6FF` |
| Callout danger | `#FEE2E2` + teks `#B91C1C` |
| Bahasa UI dokumen | id-ID (kecuali user minta lain) |

### Spasi paragraf & heading (wajib)

Atur lewat **style** (bukan hanya CSS htmlchunk), supaya native paragraph ikut konsisten:

| Style | spaceBefore | spaceAfter |
|---|---|---|
| Normal | **6 pt** | **6 pt** |
| Heading1 | **18 pt** | **8 pt** |
| Heading2 | **14 pt** | **6 pt** |
| Heading3 | **10 pt** | **4 pt** |

```bash
officecli set "$FILE" /styles/Normal --prop spaceBefore=6pt --prop spaceAfter=6pt --prop align=justify \
  --prop font=Calibri --prop size=10.5pt --prop color=#1F2937
officecli set "$FILE" /styles/Heading1 --prop spaceBefore=18pt --prop spaceAfter=8pt \
  --prop font=Calibri --prop size=15pt --prop bold=true --prop color=#0F766E
officecli set "$FILE" /styles/Heading2 --prop spaceBefore=14pt --prop spaceAfter=6pt \
  --prop font=Calibri --prop size=12.5pt --prop bold=true --prop color=#111827
officecli set "$FILE" /styles/Heading3 --prop spaceBefore=10pt --prop spaceAfter=4pt \
  --prop font=Calibri --prop size=11pt --prop bold=true --prop color=#111827
```

Untuk paragraf yang di-set properti run secara eksplisit, tambahkan juga `--prop spaceBefore=6pt` bila style Normal belum terwariskan. Di htmlchunk, CSS `p` / `h1`–`h3` harus memakai margin atas yang selaras (lihat CSS bersama).

## Anatomi dokumen (skeleton)

Urutan baku — sesuaikan judul/isi project, **jangan hilangkan blok kecuali user bilang**:

1. **Cover** (section `titlePg`): logo kiri (klien) + logo kanan (vendor/implementor), eyebrow organisasi, judul (baris netral + baris aksen teal), subtitle audiens, paragraf ringkas tujuan, catatan versi singkat, tabel meta (Nomor dokumen, Versi, Tanggal, Klasifikasi, Disusun oleh, Ditujukan kepada)
2. **Riwayat Versi** — tabel: Versi | Tanggal | Ringkasan perubahan | Penyusun
3. **Ruang Lingkup** — apa yang masuk / tidak masuk
4. **Dokumen Terkait** — tabel referensi
5. **Daftar Isi** — **field TOC native** dengan hyperlink ke Heading 1–3 (bukan daftar manual)
6. **Bab isi** — Heading 1 bernomor (`1. …`), Heading 2 (`1.1 …`), Heading 3 bila perlu
7. **Lampiran** — contoh stack, checklist go-live, referensi

Header (halaman isi, bukan cover): `{Judul singkat} · v{versi}` warna muted.
Footer: `{DOC-ID} | Halaman {PAGE} dari {NUMPAGES}` muted, PAGE/NUMPAGES harus **field hidup**, bukan teks statis.

## Strategi build OfficeCLI

**Prinsip:** native untuk struktur yang perlu diedit/di-query (heading, TOC, list numbering, header/footer, page setup); **`htmlchunk` + `matchSrc=true`** untuk blok visual kaya (cover band, callout, tabel styled, kode, meta cards). HTML disimpan verbatim sebagai `altChunk`; Word mengonversi saat dibuka — preview CLI menampilkan source escaped, itu normal.

```bash
FILE="out/NAMA-DOKUMEN.docx"
rm -f "$FILE"
officecli create "$FILE"
officecli open "$FILE"

# Page setup A4 + margin Moderate (Word: atas/bawah 25,4 mm / 1", kiri/kanan 19,05 mm / 0,75")
officecli set "$FILE" /section[1] --prop pageSetup=a4-moderate
# setara: --prop pageSetup=tech-doc
# atau terpisah: --prop pageSize=a4 --prop margins=moderate
officecli help docx section
```

### CSS bersama untuk htmlchunk

```css
body { font-family: Calibri, Arial, sans-serif; font-size: 10.5pt; color: #1F2937; }
h1 { font-size: 15pt; color: #0F766E; margin: 18pt 0 8pt; }
h2 { font-size: 12.5pt; color: #111827; margin: 14pt 0 6pt; }
h3 { font-size: 11pt; color: #111827; margin: 10pt 0 4pt; }
p { margin: 6pt 0; text-align: justify; }
.muted { color: #6B7280; }
.accent { color: #0F766E; }
table.meta, table.data { border-collapse: collapse; width: 100%; font-size: 10pt; }
table.meta td, table.data td, table.data th {
  border: 1px solid #D1D5DB; padding: 6px 8px; vertical-align: top;
}
table.data th { background: #E7F2F0; color: #0B4F49; font-weight: bold; text-align: left; }
.callout { padding: 10px 12px; margin: 10pt 0; border-radius: 2px; }
.callout.warn { background: #FEF3C7; }
.callout.ok { background: #ECFDF5; }
.callout.info { background: #EFF6FF; }
.callout.danger { background: #FEE2E2; color: #B91C1C; }
.callout .title { font-weight: bold; text-transform: uppercase; letter-spacing: .02em; margin-bottom: 4px; }
p { margin: 6pt 0; text-align: justify; }
pre, code, .mono { font-family: Consolas, "Courier New", monospace; font-size: 9.5pt; }
/* Inline / table mono — chip ala GitHub markdown (bukan full code section) */
code, span.mono {
  background: #F6F8FA;
  border: 1px solid #D0D7DE;
  border-radius: 3px;
  padding: 1px 5px;
}
pre.block {
  background: #F6F8FA; padding: 10px 12px; white-space: pre-wrap;
  border: 1px solid #D0D7DE; border-radius: 3px;
}
ol, ul { margin: 6pt 0; padding-left: 24pt; }
li { margin: 3pt 0; }
```

### Cover via htmlchunk

```bash
officecli add "$FILE" /body --type htmlchunk --prop matchSrc=true --prop src=build/cover.html
officecli add "$FILE" /body --type paragraph --prop text="" --prop pageBreakBefore=true
```

`cover.html` memuat logo sebagai `data:` URI (path relatif di dalam chunk **tidak** di-resolve).

### Daftar isi (TOC) terhubung — wajib

Pakai **field TOC native**, bukan daftar manual / htmlchunk palsu:

```bash
officecli add "$FILE" /body --type toc \
  --prop title="Daftar Isi" \
  --prop levels=1-3 \
  --prop hyperlinks=true \
  --prop pageNumbers=true
# Cek nama prop persis: officecli help docx toc

# Setelah heading final ada: isi entri (judul + hyperlink) tanpa Word.
officecli refresh "$FILE" --toc
# Atau biarkan `officecli finalize` di akhir (lihat bawah) — langkah itu
# sudah mencakup refresh --toc bila field TOC ada.
```

`title` memakai style TOCHeading dan **tidak** masuk daftar. Jangan memakai Heading1 untuk judul "Daftar Isi" — paragraf itu ikut terkumpul sebagai entri.

Syarat agar hyperlink benar:

- Semua bab/subbab memakai style **Heading1 / Heading2 / Heading3** atau `outlineLvl` (bukan bold di Normal).
- `refresh --toc` menulis judul dan tautan internal (`_Toc`). **Nomor halaman di entri adalah placeholder `0`**, bukan nomor halaman sungguhan.
- Nomor halaman sungguhan butuh Word (Update Field) atau `officecli refresh` tanpa `--toc` bila ada browser headless. Jangan melaporkan `0` sebagai nomor halaman jadi.

### Heading native + isi

```bash
officecli add "$FILE" /body --type paragraph --prop text="1. Pendahuluan" --prop style=Heading1
officecli add "$FILE" /body --type paragraph --prop text="1.1 Untuk siapa dokumen ini" --prop style=Heading2
officecli add "$FILE" /body --type paragraph --prop text="Paragraf tubuh…" --prop style=Normal
```

Untuk callout / tabel / kode setelah heading: htmlchunk + `matchSrc=true` (bungkus HTML lengkap ber-`<style>`).

### Numbering & indentasi daftar (wajib)

**Jangan** menulis nomor manual di teks (`"1. Internal …"`, `"2. Eksternal …"`). Pakai numbering Word agar indent hanging / left indent otomatis benar:

```bash
# Ordered list — teks TANPA prefix "1." / "2."
officecli add "$FILE" /body --type paragraph \
  --prop listStyle=ordered \
  --prop text="Internal [perusahaan/klien] — Admin Pusat, …" \
  --prop style=Normal
officecli add "$FILE" /body --type paragraph \
  --prop listStyle=ordered \
  --prop text="Eksternal mitra/operator — Admin Operasional, …"

# Bullet
officecli add "$FILE" /body --type paragraph \
  --prop listStyle=bullet \
  --prop text="Butir tanpa nomor"
```

Untuk kontrol indent eksplisit (opsional):

```bash
officecli help docx abstractNum
officecli help docx level
# set level indent (twips): --prop indent=720 --prop hanging=360
```

Aturan:

- Satu rangkaian `listStyle=ordered` berurutan = satu list; paragraf Normal di tengah **mereset** counter.
- Nested list: `numLevel=1` (atau `ilvl=1`) pada item anak.
- Di htmlchunk, pakai `<ol>` / `<ul>` / `<li>` (bukan `1.` di plain text).

Contoh yang benar untuk § seperti "2.2 Pengguna utama": heading native + tiga paragraf `listStyle=ordered` tanpa angka di string.

### Tabel data

Prefer htmlchunk `table.data` (header styled). Native `table` hanya jika harus di-query belakangan.

### Header & footer

```bash
officecli add "$FILE" / --type footer --prop type=default --prop size=9pt --prop color=6B7280 \
  --prop text="DOC-ID-01 | Halaman "
officecli add "$FILE" "/footer[1]/p[1]" --type field --prop fieldType=page
officecli add "$FILE" "/footer[1]/p[1]" --type run --prop text=" dari " --prop color=6B7280 --prop size=9pt
officecli add "$FILE" "/footer[1]/p[1]" --type field --prop fieldType=numpages
officecli add "$FILE" / --type footer --prop type=first --prop text=""
```

### Diagram

Ekspor PNG; sisipkan native image atau `<img src="data:image/png;base64,…">` di htmlchunk. Selalu isi `alt`. `officecli materialize` menyematkan data-URI png/jpeg/gif/bmp/tiff/emf/wmf sebagai `w:drawing` (alt jadi deskripsi gambar). webp, svg, serta URL http(s) atau path relatif tetap teks alt dan tidak menggagalkan konversi. Border CSS pada paragraf, callout, atau sel tabel ikut menjadi `w:pBdr` / `w:tcBorders` (`solid`, `dashed`, `dotted`, `double`, `inset`, `outset`). Border pada `span`/`code`, `border-radius`, dan `border-image` tidak dipetakan. Daftar bersarang (`ul`/`ol` di dalam `li`) memakai satu `numId`; `ilvl` adalah kedalaman. Daftar terpisah, termasuk daftar di sel tabel, mendapat `numId` sendiri.

### Langkah akhir (`finalize`)

Setelah isi, heading, field TOC, dan htmlchunk selesai, satu perintah headless menutup build. Tidak menjalankan Word dan tidak menghitung nomor halaman sungguhan.

```bash
officecli finalize "$FILE"
# Default, berurutan:
#   1. materialize     HTML/XHTML/teks → native (RTF/MHT tetap, sama seperti materialize)
#   2. refresh --toc   hanya jika ada field TOC; PAGEREF tetap placeholder 0
#   3. page setup      pageSetup=a4-moderate hanya pada section tanpa w:pgSz
#   4. validate        skema OpenXML; error → exit 1, langkah sebelumnya tetap
```

Lewati langkah: `--no-materialize`, `--no-toc`, `--no-page-setup`, `--no-validate`. `--strict` sama seperti `materialize --strict`: file tidak diubah, langkah berikutnya tidak jalan, exit 1. `--json` satu envelope; `data.steps[]` memakai status `ran`, `skipped`, `failed`, atau `not-run`.

**Page setup tidak menimpa ukuran halaman yang sudah ada.** `officecli create` sudah menulis A4 (`w:pgSz`), jadi `finalize` tidak mengganti margin dokumen itu menjadi Moderate. Tetap panggil `set /section[1] --prop pageSetup=a4-moderate` saat membuat dokumen. `finalize` hanya mengisi preset itu pada section yang tidak punya `w:pgSz` (margin preset ikut tertulis di section itu).

Materialize jalan sebelum TOC, jadi heading di dalam htmlchunk ikut terkumpul. Nomor halaman entri tetap `0` sampai di-update di Word.

## Alur kerja agen (checklist)

1. Kumpulkan meta + outline + aset logo.
2. Outline Heading 1–3 bernomor konsisten.
3. `officecli create` + `set /section[1] --prop pageSetup=a4-moderate` (A4 + margin Moderate) + **set spaceBefore/spaceAfter** pada Normal & Heading1–3.
4. Cover htmlchunk → page break → front matter → **TOC field (hyperlinks)** → bab.
5. Semua daftar bertingkat memakai `listStyle=ordered|bullet` (bukan angka di string).
6. Header/footer field PAGE.
7. Tutup build: `officecli finalize "$FILE"` (materialize + `refresh --toc` bila ada field TOC + validate). Page setup A4+Moderate tetap di langkah 3 — `finalize` tidak menimpa `w:pgSz` yang sudah ada.
8. QA:
   - `officecli view "$FILE" outline`
   - Cek entri TOC memuat judul heading (nomor halaman boleh `0`)
   - Spot-check satu list (mis. § pengguna): `listStyle`/`numId` ada; teks tanpa prefix `1.`
   - Heading2 spaceBefore ≈ 14pt
   - Update Field di Word hanya bila nomor halaman sungguhan diperlukan
9. Serahkan `.docx`. Sebutkan jika `finalize` melaporkan altChunk yang sengaja dibiarkan (RTF/MHT) atau validate gagal.


## Mono inline (chip GitHub) — wajib

Chip abu + border hanya pada **keyword/teks monospace** (bukan seluruh sel tabel, bukan paragraf penuh). Berlaku di paragraf dan di dalam sel tabel, sama seperti inline code GitHub Markdown.

| Token | Nilai |
|---|---|
| Background | `#F6F8FA` |
| Border | `1px solid #D0D7DE` |
| Radius | ~3px |
| Padding inline | ~1px 5px |
| Font | Consolas / mono 9.5pt |

HTML: bungkus keyword dengan `<code>…</code>` atau `<span class="mono">…</span>` — termasuk di dalam `<td>`. Contoh: `<td><span class="mono">docs/testing/README.md</span></td>`.

**Dilarang** `class="mono"` / fill abu pada `<td>` / `<th>` (itu mewarnai seluruh cell).

Blok kode penuh (`pre.block`) tetap full-width abu + border — terpisah dari chip keyword.

Jangan pakai Consolas polos tanpa chip untuk path/command/identifier.

## Alignment paragraf body — wajib

Paragraf body default **rata kiri-kanan** (`align=justify` / CSS `text-align: justify`):

- Style Normal: `--prop align=justify`
- Helper `p` / daftar body: `--prop align=justify`
- CSS htmlchunk: `p { text-align: justify; }`
- Heading, cover title, meta table, header/footer: biarkan left/center sesuai konteks (jangan force justify pada judul)


## Kolom ID pada tabel — wajib

Jika tabel punya kolom ID/kode bernilai pendek seperti `T-01`, `T-02`, `NFR-01`, `FR-001`, `REQ-12`:

1. Tandai header & sel: `th.col-id` / `td.col-id` (isi tetap chip mono: `<span class="mono">NFR-01</span>`).
2. `white-space: nowrap` pada `.col-id` — **ID tidak boleh turun baris**.
3. Set lebar kolom dari **nilai ID terpanjang** di tabel itu (bukan lebar merata):
   - Hitung `max_len` = panjang string ID terpanjang (atau header `ID` jika lebih panjang).
   - Lebar ≈ `max(18mm, max_len × 2.0mm + 8mm)` (mengakomodasi chip mono + padding sel).
   - Terapkan via `<colgroup><col class="col-id" style="width:…mm" />…</colgroup>`.
4. Kolom lain memakai sisa lebar tabel (`width: 100%` pada `table.data`).
5. Jangan mewarnai seluruh sel sebagai chip — chip hanya pada keyword mono di dalam sel.

CSS wajib:

```css
table.data th.col-id, table.data td.col-id {
  white-space: nowrap;
  vertical-align: top;
  width: 1%;
}
```

### Tabel native (bukan htmlchunk)

Untuk tabel yang sudah ada, jangan set `noWrap` sel per sel. Tandai kolomnya:

```bash
officecli set doc.docx "/body/tbl[1]/col[1]" --prop idColumn=true
```

`idColumn=true` menulis `w:noWrap` pada setiap sel satu-kolom di kolom itu dan mengatur lebar (`w:gridCol` + `w:tcW`) dari baris terpanjang: `max(18mm, panjang × 2.0mm + 8mm)`. Beberapa kolom sekaligus, atau saat tabel dibuat:

```bash
officecli set doc.docx "/body/tbl[1]" --prop idColumns=1,3
officecli add doc.docx /body --type table --prop idColumn=1 \
  --prop data="ID,Requirement;T-01,The operator exports the register;NFR-01,Response stays under one second"
```

Htmlchunk tetap memakai `th.col-id` / `td.col-id` dan `white-space: nowrap`. `officecli materialize` menerapkan perlakuan kolom ID yang sama. Lihat `officecli help docx table-column`.

## RULES (wajib dipatuhi)

1. **Satu sistem desain** — hanya token di atas.
2. **Cover dulu** — dokumen teknis >3 halaman: cover + meta + riwayat versi.
3. **Nomor dokumen + versi** di cover, footer, dan riwayat.
4. **Heading untuk struktur** — jangan palsukan H1 dengan bold di Normal (kecuali cover htmlchunk).
5. **Spasi wajib** — Normal & setiap level heading punya spaceBefore/spaceAfter sesuai tabel token; jangan rapat menempel.
6. **TOC field + hyperlink** — Daftar Isi = `toc` native levels 1–3 dengan hyperlinks; bukan bullet manual.
7. **Numbering native** — daftar bernomor/bullet pakai `listStyle` / `numId` + indent Word; dilarang `"1. …"` sebagai plain text.
8. **Callout = satu pesan** — warna sesuai jenis (warn/ok/info/danger).
9. **Kode & secret** — placeholder saja; jangan tempel secret nyata.
10. **Tabel** — header row styled; kolom teknis mono.
11. **Bahasa** — Indonesia baku teknis; istilah OIDC/OAuth boleh Inggris.
12. **htmlchunk** — `matchSrc=true`; gambar data URI atau native.
13. **Tanpa placeholder final** — `$xxx$`, `TODO`, `lorem` dilarang di output.
14. **Field halaman hidup** di footer.
15. **Reusable** — ganti konten + logo + DOC-ID, bukan gaya.
16. **Mono keyword = chip** — hanya teks keyword (`code`/`span.mono`) yang ber-chip; jangan warnai seluruh sel tabel; full `pre.block` terpisah.
17. **Body justify** — paragraf Normal rata kiri-kanan; heading/cover tidak ikut force-justify.
18. **Kolom ID = nowrap + fit max** — `T-01`/`NFR-01`/`FR-001` dll. satu baris; lebar kolom dari ID terpanjang; dilarang wrap. Tabel native: `set …/col[N] --prop idColumn=true`.

## Adaptasi project lain

| Yang diganti | Yang tetap |
|---|---|
| Judul, DOC-ID, versi, klien, logo | Palet teal, tipografi, margin, spasi, pola TOC/list/cover |
| Outline bab | Header `Judul · vX` + footer `ID \| Halaman X dari Y` |
| Stack contoh di lampiran | Aturan callout, mono, numbering |

## Referensi cepat OfficeCLI

```bash
officecli help docx htmlchunk
officecli help docx toc
officecli help docx section      # pageSetup=a4-moderate, margins=moderate, pageSize=a4
officecli help docx paragraph   # listStyle, spaceBefore
officecli help docx style       # spaceBefore pada Heading/Normal
officecli add doc.docx /body --type toc --prop levels=1-3 --prop hyperlinks=true
officecli add doc.docx /body --type paragraph --prop listStyle=ordered --prop text='…'
officecli finalize doc.docx     # materialize + refresh --toc + validate
officecli finalize --help       # --no-materialize, --no-toc, --no-page-setup, --no-validate, --strict
```

Contoh fork: `examples/word/html-chunk.sh` di `integra-id/OfficeCLI` branch fitur htmlchunk.

## Anti-pola

- TOC manual / htmlchunk berisi daftar heading palsu tanpa field TOC
- `"1. Item"` di paragraf Normal alih-alih `listStyle=ordered`
- Heading tanpa spaceBefore (menempel ke paragraf sebelumnya)
- Seluruh dokumen hanya Normal + warna manual tanpa Heading
- Mengandalkan `view html` untuk menilai htmlchunk (escaped = expected)
- Path gambar relatif di dalam HTML chunk
- Consolas/mono polos tanpa background+border di tengah paragraf/tabel
- Margin selain Moderate (atau yang diminta user) / page size non-A4 tanpa permintaan
- Kolom ID yang wrap (T-01 / NFR-01 turun baris) karena lebar kolom tidak di-fit
- Paragraf body left-only padahal desain dokumen teknis default justify
