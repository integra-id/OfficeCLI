---
name: Rule gaya dokumen teknis Integra
description: >-
  Aturan wajib desain dokumen Word teknis Integra (A4, teal, cover, TOC
  hyperlink, spasi heading/paragraf, numbering native, mono chip, justify, margin Moderate, kolom ID nowrap, htmlchunk). Pakai
  bersama skill Dokumen teknis Integra (OfficeCLI).
---
# Rule: Gaya dokumen teknis Integra

Berlaku setiap kali membuat atau mengedit dokumen Word teknis project dengan OfficeCLI.

1. Ikuti skill Dokumen teknis Integra (OfficeCLI) end-to-end.
2. A4, margin **Moderate** (atas/bawah 25,4 mm, kiri/kanan 19,05 mm), Calibri body `#1F2937`, aksen teal `#0F766E` / `#0B4F49`.
3. Wajib: cover + meta + riwayat versi + heading bernomor + header/footer field PAGE.
4. **Spasi:** Normal spaceBefore/After 6 pt; Heading1 18/8; Heading2 14/6; Heading3 10/4 (pt).
5. **TOC:** field native `toc` levels 1–3 dengan hyperlinks (+ page numbers); bukan daftar manual.
6. **Numbering:** `listStyle=ordered|bullet` (atau numId) dengan indent Word; dilarang prefix `"1."` di plain text (contoh § pengguna utama).
7. Callout & tabel & kode: htmlchunk + matchSrc, CSS token Integra (termasuk margin `p`/`h*`).
8. Tanpa secret nyata; tanpa lorem/TODO di output final.
9. **Mono keyword = chip:** hanya teks keyword (`code`/`span.mono`) di paragraf atau di dalam sel — chip `#F6F8FA` + border `#D0D7DE`; dilarang fill seluruh `td`; `pre.block` = full section.
10. **Body justify:** Normal + paragraf htmlchunk `text-align: justify` / `align=justify`.
11. **Kolom ID:** `col-id` + `nowrap` + lebar dari ID terpanjang (`T-01`, `NFR-01`, `FR-001`, …) agar tidak wrap.
12. Validasi outline + TOC + list numPr + sample mono chip + kolom ID satu baris + footer field sebelum menyerahkan file.

Detail lengkap dan perintah OfficeCLI ada di skill Dokumen teknis Integra (OfficeCLI).
