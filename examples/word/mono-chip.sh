#!/usr/bin/env bash
# Smoke test for the native mono keyword chip (w:rPr, not cell shading).
#
# mono=true (aliases chip, monoChip) writes Consolas 9.5pt, run shading
# #F6F8FA, and a character border 1px solid #D0D7DE. htmlchunk materialize
# paints the font, size, and shading for code/span.mono but does not emit
# a run border; this prop adds w:bdr. The chip stays on the run — a table
# cell that contains one chipped keyword must not gain w:tcPr/w:shd.
#
# Usage (from the repo root, after `dotnet build -c Release`):
#   OFFICECLI=path/to/officecli bash examples/word/mono-chip.sh
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

export OFFICECLI_NO_AUTO_RESIDENT=1

WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT
fail() { echo "FAIL: $*" >&2; exit 1; }

echo "== help documents mono =="
"$OFFICECLI" help docx run | grep -q 'mono' || fail "help docx run missing mono"
"$OFFICECLI" help docx paragraph | grep -q 'F6F8FA' || fail "help docx paragraph missing chip fill"
"$OFFICECLI" help docx table-cell | grep -q 'mono' || fail "help docx table-cell missing mono"

DOC="$WORK/mono.docx"
"$OFFICECLI" create "$DOC"

echo "== add paragraph mono=true =="
"$OFFICECLI" add "$DOC" /body --type paragraph --prop text="Intro sentence"
"$OFFICECLI" add "$DOC" /body --type paragraph --prop text="NFR-01" --prop mono=true

echo "== range chip inside a sentence =="
"$OFFICECLI" add "$DOC" /body --type paragraph --prop text="see docs/testing/README.md now"
"$OFFICECLI" set "$DOC" "/body/p[3]" --prop range=4:26 --prop mono=true

echo "== chip one run inside a table cell; keyword cell via set =="
"$OFFICECLI" add "$DOC" /body --type table --prop rows=2 --prop cols=2
"$OFFICECLI" set "$DOC" "/body/tbl[1]/tr[1]/tc[1]" --prop text="see docs/testing/README.md now"
"$OFFICECLI" set "$DOC" "/body/tbl[1]/tr[1]/tc[1]/p[1]" --prop range=4:26 --prop chip=true
"$OFFICECLI" set "$DOC" "/body/tbl[1]/tr[1]/tc[2]" --prop text="plain"
"$OFFICECLI" set "$DOC" "/body/tbl[1]/tr[2]/tc[1]" --prop text="NFR-01"
"$OFFICECLI" set "$DOC" "/body/tbl[1]/tr[2]/tc[1]" --prop mono=true
"$OFFICECLI" add "$DOC" "/body/tbl[1]/tr[2]/tc[1]/p[1]" --type run --prop text=" extra" --prop monochip=true

"$OFFICECLI" validate "$DOC" >/dev/null

python3 - "$DOC" <<'PY' || fail "chip OOXML mismatch"
import sys, zipfile
import xml.etree.ElementTree as ET
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
root = ET.fromstring(zipfile.ZipFile(sys.argv[1]).read("word/document.xml"))
body = root.find(f"{W}body")

def text_of(el):
    return "".join(t.text or "" for t in el.iter(f"{W}t"))

def is_chip(r):
    rPr = r.find(f"{W}rPr")
    if rPr is None:
        return False
    fonts = rPr.find(f"{W}rFonts")
    sz = rPr.find(f"{W}sz")
    shd = rPr.find(f"{W}shd")
    bdr = rPr.find(f"{W}bdr")
    if fonts is None or sz is None or shd is None or bdr is None:
        return False
    ok = (
        fonts.get(f"{W}ascii") == "Consolas"
        and fonts.get(f"{W}hAnsi") == "Consolas"
        and fonts.get(f"{W}eastAsia") == "Consolas"
        and sz.get(f"{W}val") == "19"
        and (shd.get(f"{W}fill") or "").upper() == "F6F8FA"
        and (shd.get(f"{W}val") or "") == "clear"
        and (bdr.get(f"{W}val") or "") == "single"
        and bdr.get(f"{W}sz") == "6"
        and (bdr.get(f"{W}color") or "").upper() == "D0D7DE"
        and bdr.get(f"{W}space") == "1"
    )
    return ok

paras = [c for c in list(body) if c.tag == f"{W}p"]
assert len(paras) >= 3, len(paras)
# p2 is the whole-keyword paragraph.
p2_runs = paras[1].findall(f"{W}r")
assert len(p2_runs) == 1 and is_chip(p2_runs[0]) and text_of(p2_runs[0]) == "NFR-01", text_of(paras[1])
assert paras[1].find(f"{W}pPr/{W}shd") is None

# p3 range: only the path run is a chip.
p3_runs = paras[2].findall(f"{W}r")
chipped = [r for r in p3_runs if is_chip(r)]
plain = [r for r in p3_runs if not is_chip(r)]
assert len(chipped) == 1 and text_of(chipped[0]) == "docs/testing/README.md", [text_of(r) for r in p3_runs]
assert any(text_of(r) == "see " for r in plain) and any("now" in text_of(r) for r in plain)

tbl = body.find(f"{W}tbl")
rows = tbl.findall(f"{W}tr")
def cell(r, c):
    return rows[r].findall(f"{W}tc")[c]

def cell_shd(tc):
    tcPr = tc.find(f"{W}tcPr")
    if tcPr is None:
        return None
    return tcPr.find(f"{W}shd")

c00 = cell(0, 0)
assert cell_shd(c00) is None, "range chip must not shade the cell"
c00_runs = c00.findall(f".//{W}r")
c00_chip = [r for r in c00_runs if is_chip(r)]
assert len(c00_chip) == 1 and text_of(c00_chip[0]) == "docs/testing/README.md"
assert cell_shd(cell(0, 1)) is None
assert all(not is_chip(r) for r in cell(0, 1).findall(f".//{W}r"))

c10 = cell(1, 0)
assert cell_shd(c10) is None, "mono on the cell must not write w:tcPr/w:shd"
c10_chips = [r for r in c10.findall(f".//{W}r") if is_chip(r)]
assert {text_of(r) for r in c10_chips} >= {"NFR-01", " extra"}, [text_of(r) for r in c10.findall(f".//{W}r")]
print("chip xml ok")
PY

echo "== mono=false clears a matching chip =="
"$OFFICECLI" set "$DOC" "/body/p[2]/r[1]" --prop mono=false
python3 - "$DOC" <<'PY' || fail "mono=false left chip markup"
import sys, zipfile
import xml.etree.ElementTree as ET
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
root = ET.fromstring(zipfile.ZipFile(sys.argv[1]).read("word/document.xml"))
body = root.find(f"{W}body")
paras = [c for c in list(body) if c.tag == f"{W}p"]
rPr = paras[1].find(f"{W}r/{W}rPr")
assert rPr is not None
assert rPr.find(f"{W}rFonts") is None
assert rPr.find(f"{W}sz") is None
assert rPr.find(f"{W}shd") is None
assert rPr.find(f"{W}bdr") is None
print("cleared")
PY

echo "== invalid mono value is rejected =="
if "$OFFICECLI" set "$DOC" "/body/p[1]/r[1]" --prop mono=banana >/dev/null 2>"$WORK/err"; then
  fail "mono=banana should fail"
fi
grep -q -i "boolean" "$WORK/err" || fail "expected a boolean error, got: $(cat "$WORK/err")"

"$OFFICECLI" validate "$DOC" >/dev/null
echo "OK mono-chip"
