#!/usr/bin/env bash
# Smoke test for headless altChunk materialization.
#
# Creates a docx with htmlchunk payloads, runs `officecli materialize`, and
# checks the package: w:altChunk and word/afchunk*.htm are gone, and the body
# contains real paragraphs, a table (with gridSpan), a hyperlink, and list
# numbering. A separate file keeps an RTF chunk in place and checks that
# --strict refuses to modify the file.
#
# Usage (from the repo root, after `dotnet build -c Release`):
#   OFFICECLI=path/to/officecli bash examples/word/materialize-altchunk.sh
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/../.." && pwd)"
if [[ -z "${OFFICECLI:-}" ]]; then
  if [[ -x "$ROOT/src/officecli/bin/Release/net10.0/officecli" ]]; then
    OFFICECLI="$ROOT/src/officecli/bin/Release/net10.0/officecli"
  else
    OFFICECLI="$ROOT/src/officecli/bin/Release/net10.0/linux-x64/officecli"
  fi
fi
if [[ ! -x "$OFFICECLI" ]]; then
  echo "officecli binary not found at $OFFICECLI" >&2
  echo "Build first: dotnet build -c Release src/officecli/officecli.csproj" >&2
  exit 1
fi

# Direct open/save. A resident would hold the package in memory and the
# unzip assertions below would race the idle flush.
export OFFICECLI_NO_AUTO_RESIDENT=1

WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT

fail() { echo "FAIL: $*" >&2; exit 1; }

doc_xml() { unzip -p "$1" word/document.xml; }

echo "== simple HTML chunk =="
SIMPLE="$WORK/simple.docx"
"$OFFICECLI" create "$SIMPLE"
"$OFFICECLI" add "$SIMPLE" /body --type htmlchunk --prop html='<h1>Title</h1><p>Hello <b>bold</b> and <a href="https://example.com/docs">link</a>.</p><ul><li>one</li><li>two</li></ul><table><tr><th>A</th><th colspan="2">B</th></tr><tr><td>1</td><td>2</td><td>3</td></tr></table>'
doc_xml "$SIMPLE" | grep -q 'w:altChunk' || fail "expected w:altChunk before materialize"
unzip -l "$SIMPLE" | grep -q 'afchunk' || fail "expected afchunk part before materialize"

"$OFFICECLI" materialize "$SIMPLE"
if doc_xml "$SIMPLE" | grep -q 'w:altChunk'; then
  fail "w:altChunk still present after materialize"
fi
if unzip -l "$SIMPLE" | grep -q 'afchunk'; then
  fail "afchunk part still in the package"
fi
XML="$(doc_xml "$SIMPLE")"
echo "$XML" | grep -q 'Heading1' || fail "heading style missing"
unzip -p "$SIMPLE" word/styles.xml | grep -q 'w:styleId="Heading1"' || fail "Heading1 style was not defined in styles.xml"
echo "$XML" | grep -q '>Title<' || fail "heading text missing"
echo "$XML" | grep -q '>bold<' || fail "bold text missing"
echo "$XML" | grep -q '<w:b' || fail "bold mark missing"
echo "$XML" | grep -q 'w:tbl' || fail "table missing"
echo "$XML" | grep -q 'w:gridSpan' || fail "colspan was not emitted as gridSpan"
echo "$XML" | grep -q 'w:hyperlink' || fail "hyperlink missing"
echo "$XML" | grep -q 'w:numPr' || fail "list numbering missing"
echo "$XML" | grep -Eq '<w:ilvl w:val="0"[[:space:]]*/>' || fail "top-level list should be ilvl 0"
unzip -p "$SIMPLE" word/_rels/document.xml.rels | grep -q 'example.com/docs' || fail "hyperlink relationship missing"
python3 - "$SIMPLE" <<'PY' || fail "afchunk content type still registered"
import sys, zipfile
z = zipfile.ZipFile(sys.argv[1])
ct = z.read("[Content_Types].xml").decode("utf-8", "replace")
if "afchunk" in ct or any("afchunk" in n for n in z.namelist()):
    sys.exit(1)
PY
"$OFFICECLI" validate "$SIMPLE" >/dev/null

echo "== second pass is a no-op =="
"$OFFICECLI" materialize "$SIMPLE" | grep -q 'No altChunks' || fail "second materialize should report nothing to do"

echo "== styled HTML file (CSS, rowspan, nested list) =="
STYLED="$WORK/styled.docx"
"$OFFICECLI" create "$STYLED"
"$OFFICECLI" add "$STYLED" /body --type htmlchunk --prop matchSrc=true --prop src="$ROOT/examples/word/html-chunk-report.html"
"$OFFICECLI" materialize "$STYLED"
SXML="$(doc_xml "$STYLED")"
echo "$SXML" | grep -q 'w:altChunk' && fail "styled chunk was not materialized"
echo "$SXML" | grep -q '1F4E79' || fail "h1/th CSS color was not applied"
echo "$SXML" | grep -q 'w:gridSpan' || fail "report colspan missing"
echo "$SXML" | grep -q 'w:vMerge' || fail "report rowspan missing"
echo "$SXML" | grep -q 'Jakarta' || fail "table text missing"
echo "$SXML" | grep -Eq '<w:ilvl w:val="0"[[:space:]]*/>' || fail "top-level ordered list should be ilvl 0"
echo "$SXML" | grep -Eq '<w:ilvl w:val="1"[[:space:]]*/>' || fail "nested list should be ilvl 1"
"$OFFICECLI" validate "$STYLED" >/dev/null

echo "== chunk inside a table cell =="
CELL="$WORK/cell.docx"
"$OFFICECLI" create "$CELL"
"$OFFICECLI" add "$CELL" /body --type table --prop rows=1 --prop cols=1
"$OFFICECLI" add "$CELL" "/body/tbl[1]/tr[1]/tc[1]" --type htmlchunk --prop html='<p>In cell <b>yes</b></p>'
"$OFFICECLI" materialize "$CELL"
CXML="$(doc_xml "$CELL")"
echo "$CXML" | grep -q 'w:altChunk' && fail "cell chunk was not materialized"
echo "$CXML" | grep -q 'In cell' || fail "cell text missing"
"$OFFICECLI" validate "$CELL" >/dev/null

echo "== plain text chunk =="
TEXT="$WORK/text.docx"
"$OFFICECLI" create "$TEXT"
"$OFFICECLI" add "$TEXT" /body --type htmlchunk --prop format=text --prop html=$'alpha\nbeta'
"$OFFICECLI" materialize "$TEXT"
doc_xml "$TEXT" | grep -q 'w:altChunk' && fail "text chunk was not materialized"
doc_xml "$TEXT" | grep -q '>alpha<' || fail "plain-text line missing"
"$OFFICECLI" validate "$TEXT" >/dev/null

echo "== RTF-only chunk is left in place (not an error) =="
RTF="$WORK/rtf.docx"
"$OFFICECLI" create "$RTF"
"$OFFICECLI" add "$RTF" /body --type htmlchunk --prop format=rtf --prop html='{\rtf1 only rtf}'
"$OFFICECLI" materialize "$RTF" | grep -q 'nothing was materialized' || fail "RTF-only file should stay unchanged without failing"
doc_xml "$RTF" | grep -q 'w:altChunk' || fail "RTF-only chunk should remain"

echo "== RTF stays; --strict does not modify =="
MIXED="$WORK/mixed.docx"
"$OFFICECLI" create "$MIXED"
"$OFFICECLI" add "$MIXED" /body --type htmlchunk --prop html='<p>Keep me</p>'
"$OFFICECLI" add "$MIXED" /body --type htmlchunk --prop format=rtf --prop html='{\rtf1 hello}'
BEFORE="$(doc_xml "$MIXED")"
set +e
"$OFFICECLI" materialize "$MIXED" --strict
STRICT_RC=$?
set -e
[[ "$STRICT_RC" -ne 0 ]] || fail "--strict should fail when an RTF chunk is present"
AFTER="$(doc_xml "$MIXED")"
[[ "$BEFORE" == "$AFTER" ]] || fail "--strict modified the file"
echo "$AFTER" | grep -q 'w:altChunk' || fail "RTF chunk should still be an altChunk"

"$OFFICECLI" materialize "$MIXED"
MX="$(doc_xml "$MIXED")"
echo "$MX" | grep -q '>Keep me<' || fail "HTML chunk in the mixed file was not materialized"
echo "$MX" | grep -q 'w:altChunk' || fail "RTF chunk should remain after a non-strict materialize"
# The HTML part is gone; one afchunk (the RTF) remains.
AFCOUNT="$(unzip -l "$MIXED" | grep -c 'afchunk' || true)"
[[ "$AFCOUNT" -eq 1 ]] || fail "expected exactly one remaining afchunk part, found $AFCOUNT"
"$OFFICECLI" validate "$MIXED" >/dev/null

echo "== data-URI images become drawings; remote/svg/webp stay alt text =="
# 1x1 PNG and 1x1 GIF. jpeg/bmp/tiff/emf/wmf use the same picture helper
# (ImageSource + AddPicture). webp is not in that pipeline.
PNG='iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
GIF='R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7'
# Newline inside the cell image's base64 — materialize must still embed it.
PNG_WRAPPED=$(printf '%s\n%s' "${PNG:0:20}" "${PNG:20}")
IMG="$WORK/images.docx"
"$OFFICECLI" create "$IMG"
"$OFFICECLI" add "$IMG" /body --type htmlchunk --prop html="<p>Before <img src=\"data:image/png;base64,$PNG\" width=\"32\" height=\"32\" alt=\"tiny logo\"> mid <img src=\"data:image/gif;base64,$GIF\" width=\"32\" height=\"32\" alt=\"pixel gif\"> after</p><p><img src=\"https://example.com/x.png\" alt=\"remote\"></p><p><img src=\"logo.png\" alt=\"local file\"></p><p><img src=\"data:image/svg+xml;base64,PHN2Zy8+\" alt=\"icon\"></p><p><img src=\"data:image/webp;base64,UklGRg==\" alt=\"wpic\"></p><p>Still here <img src=\"data:image/png;base64,!!!!\" alt=\"broken\"></p><table><tr><td><img src=\"data:image/png;base64,$PNG_WRAPPED\" width=\"16\" height=\"16\" alt=\"cell pic\"></td></tr></table>"
# A bad data URI is a warning, not a --strict failure. The rest of the chunk
# (including the good pictures) is still converted.
"$OFFICECLI" materialize "$IMG" --strict 2>"$WORK/images.err"
grep -q 'not downloaded' "$WORK/images.err" || fail "expected a warning that remote images are not downloaded"
grep -q 'not resolved' "$WORK/images.err" || fail "expected a warning that relative images are not resolved"
grep -q 'unsupported type' "$WORK/images.err" || fail "expected a warning for svg/webp"
grep -q 'could not be embedded' "$WORK/images.err" || fail "expected a warning for the corrupt data URI"
python3 - "$IMG" <<'PY' || fail "data-URI pictures were not embedded as drawings"
import sys, zipfile
z = zipfile.ZipFile(sys.argv[1])
xml = z.read("word/document.xml").decode("utf-8")
names = z.namelist()
if "w:altChunk" in xml or any("afchunk" in n for n in names):
    sys.exit("altChunk still present")
import re
drawings = xml.count("<w:drawing")
if drawings < 3:
    sys.exit(f"expected at least 3 drawings (png, gif, cell), found {drawings}")
doc_ids = re.findall(r"<wp:docPr[^>]*\sid=\"(\d+)\"", xml)
if len(doc_ids) < 3 or len(doc_ids) != len(set(doc_ids)):
    sys.exit(f"docPr ids are missing or not unique: {doc_ids}")
for needle in ("Before", "mid", "after", "Still here", "tiny logo", "pixel gif", "cell pic"):
    if needle not in xml:
        sys.exit(f"missing {needle!r}")
for placeholder in ("[image: remote]", "[image: local file]", "[image: icon]", "[image: wpic]", "[image: broken]"):
    if placeholder not in xml:
        sys.exit(f"missing alt-text placeholder {placeholder!r}")
# 32px at 96dpi = 304800 EMU; 16px = 152400 EMU.
if xml.count('cx="304800"') < 2:
    sys.exit("png/gif extent was not 32px")
if 'cx="152400"' not in xml:
    sys.exit("cell picture extent was not 16px")
# add picture stores bytes at package-root media/ (Target="/media/…"),
# not only under word/media/. Accept either layout.
media = [n for n in names if n.startswith("media/") or n.startswith("word/media/")]
if len(media) < 3:
    sys.exit(f"expected at least 3 media parts, found {media}")
blobs = [z.read(n)[:4] for n in media]
if not any(b.startswith(b"\x89PNG") for b in blobs):
    sys.exit("no PNG media part")
if not any(b.startswith(b"GIF8") for b in blobs):
    sys.exit("no GIF media part")
rels = z.read("word/_rels/document.xml.rels").decode("utf-8")
if "relationships/image" not in rels:
    sys.exit("document rels have no image relationship")
PY
"$OFFICECLI" validate "$IMG" >/dev/null

echo "SMOKE OK"
