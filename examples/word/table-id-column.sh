#!/usr/bin/env bash
# Smoke test for Word ID columns (w:noWrap + width fitted to the longest line).
#
# An ID column holds short codes such as T-01, NFR-01, FR-001, REQ-12. Marking
# it writes w:noWrap on each single-span cell and sizes w:gridCol / w:tcW to
# max(18mm, longest × 2.0mm + 8mm). Six Latin characters → 1134 twips.
#
# Usage (from the repo root, after `dotnet build -c Release`):
#   OFFICECLI=path/to/officecli bash examples/word/table-id-column.sh
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

# Six character units (NFR-01, FR-001, REQ-12) at the ID-column formula.
FIT6=1134
# 2.2cm via ParseTwips (22mm is not a width suffix; cm/in/pt/dxa are).
EXPLICIT_2_2CM=1247

echo "== help documents idColumn =="
"$OFFICECLI" help docx table-column | grep -q 'idColumn' || fail "help docx table-column missing idColumn"
"$OFFICECLI" help docx column | grep -q 'width=fit' || fail "help docx column missing width=fit"
"$OFFICECLI" help docx table | grep -q 'idColumns' || fail "help docx table missing idColumns"

echo "== set /tbl/col idColumn=true =="
DOC="$WORK/ids.docx"
"$OFFICECLI" create "$DOC"
"$OFFICECLI" add "$DOC" /body --type table --prop rows=4 --prop cols=3
"$OFFICECLI" set "$DOC" "/body/tbl[1]/tr[1]/tc[1]" --prop text="ID"
"$OFFICECLI" set "$DOC" "/body/tbl[1]/tr[1]/tc[2]" --prop text="Requirement"
"$OFFICECLI" set "$DOC" "/body/tbl[1]/tr[1]/tc[3]" --prop text="Code"
"$OFFICECLI" set "$DOC" "/body/tbl[1]/tr[2]/tc[1]" --prop text="T-01"
"$OFFICECLI" set "$DOC" "/body/tbl[1]/tr[2]/tc[2]" --prop text="The operator exports the register"
"$OFFICECLI" set "$DOC" "/body/tbl[1]/tr[2]/tc[3]" --prop text="FR-001"
"$OFFICECLI" set "$DOC" "/body/tbl[1]/tr[3]/tc[1]" --prop text="NFR-01"
"$OFFICECLI" set "$DOC" "/body/tbl[1]/tr[3]/tc[2]" --prop text="Response stays under one second"
"$OFFICECLI" set "$DOC" "/body/tbl[1]/tr[3]/tc[3]" --prop text="REQ-12"
"$OFFICECLI" set "$DOC" "/body/tbl[1]/tr[4]/tc[1]" --prop text="T-02"
"$OFFICECLI" set "$DOC" "/body/tbl[1]/tr[4]/tc[2]" --prop text="A longer requirement statement that must keep wrapping"
"$OFFICECLI" set "$DOC" "/body/tbl[1]/tr[4]/tc[3]" --prop text="T-02"

python3 - "$DOC" "$WORK/before.txt" <<'PY'
import sys, zipfile
import xml.etree.ElementTree as ET
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
root = ET.fromstring(zipfile.ZipFile(sys.argv[1]).read("word/document.xml"))
tbl = root.find(f"{W}body").find(f"{W}tbl")
widths = [c.get(W + "w") for c in tbl.find(f"{W}tblGrid").findall(f"{W}gridCol")]
open(sys.argv[2], "w").write(",".join(widths))
print("before", widths)
PY

"$OFFICECLI" set "$DOC" "/body/tbl[1]/col[1]" --prop idColumn=true
"$OFFICECLI" validate "$DOC" >/dev/null

python3 - "$DOC" "$WORK/before.txt" "$FIT6" <<'PY' || fail "column idColumn did not fit + nowrap"
import sys, zipfile
import xml.etree.ElementTree as ET
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
fit = int(sys.argv[3])
before = [int(x) for x in open(sys.argv[2]).read().split(",")]
root = ET.fromstring(zipfile.ZipFile(sys.argv[1]).read("word/document.xml"))
tbl = root.find(f"{W}body").find(f"{W}tbl")
grid = [int(c.get(W + "w")) for c in tbl.find(f"{W}tblGrid").findall(f"{W}gridCol")]
if grid[0] != fit:
    sys.exit(f"grid col0 {grid[0]} != {fit}")
if grid[1] != before[1] or grid[2] != before[2]:
    sys.exit(f"other columns changed: {before} -> {grid}")
if grid[0] >= grid[1]:
    sys.exit(f"ID column is not narrower than the requirement column: {grid}")

def cell_at(row, slot):
    acc = 0
    for tc in row.findall(f"{W}tc"):
        tcpr = tc.find(f"{W}tcPr")
        span = 1
        if tcpr is not None:
            gs = tcpr.find(f"{W}gridSpan")
            if gs is not None:
                span = int(gs.get(W + "val") or "1")
        if slot >= acc and slot < acc + span:
            return tc, acc, span
        acc += span
    return None, 0, 0

nowrap_ids = 0
for tr in tbl.findall(f"{W}tr"):
    tc, start, span = cell_at(tr, 0)
    text = "".join(t.text or "" for t in tc.iter(W + "t"))
    tcpr = tc.find(f"{W}tcPr")
    if tcpr is None or tcpr.find(f"{W}noWrap") is None:
        sys.exit(f"missing noWrap on {text!r}")
    tcw = tcpr.find(f"{W}tcW")
    if tcw is None or tcw.get(W + "w") != str(fit) or tcw.get(W + "type") != "dxa":
        sys.exit(f"tcW on {text!r}: {tcw.attrib if tcw is not None else None}")
    nowrap_ids += 1
    other, _, _ = cell_at(tr, 1)
    other_pr = other.find(f"{W}tcPr")
    if other_pr is not None and other_pr.find(f"{W}noWrap") is not None:
        sys.exit("requirement column must still wrap")
if nowrap_ids != 4:
    sys.exit(f"expected 4 ID cells, got {nowrap_ids}")
print("ok set col", grid)
PY

echo "== get reports idColumn, nowrap, and width =="
GET="$("$OFFICECLI" get "$DOC" "/body/tbl[1]/col[1]" --json)"
echo "$GET" | grep -Eq '"idColumn"[[:space:]]*:[[:space:]]*true' || fail "get missing idColumn true: $GET"
echo "$GET" | grep -Eq '"nowrap"[[:space:]]*:[[:space:]]*true' || fail "get missing nowrap: $GET"
echo "$GET" | grep -Eq "\"width\"[[:space:]]*:[[:space:]]*\"${FIT6}dxa\"" || fail "get width is not ${FIT6}dxa: $GET"

echo "== idColumn=false clears noWrap and keeps the width =="
"$OFFICECLI" set "$DOC" "/body/tbl[1]/col[1]" --prop idColumn=false
python3 - "$DOC" "$FIT6" <<'PY' || fail "idColumn=false did not clear noWrap"
import sys, zipfile
import xml.etree.ElementTree as ET
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
fit = sys.argv[2]
root = ET.fromstring(zipfile.ZipFile(sys.argv[1]).read("word/document.xml"))
tbl = root.find(f"{W}body").find(f"{W}tbl")
grid0 = tbl.find(f"{W}tblGrid").find(f"{W}gridCol").get(W + "w")
if grid0 != fit:
    sys.exit(f"width changed to {grid0}")
if next(tbl.iter(W + "noWrap"), None) is not None:
    sys.exit("noWrap still present")
print("ok cleared")
PY

echo "== restore, then width=fit without nowrap =="
"$OFFICECLI" set "$DOC" "/body/tbl[1]/col[1]" --prop idColumn=true
"$OFFICECLI" set "$DOC" "/body/tbl[1]/col[2]" --prop width=fit
python3 - "$DOC" "$FIT6" <<'PY' || fail "width=fit on the requirement column"
import sys, zipfile
import xml.etree.ElementTree as ET
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
root = ET.fromstring(zipfile.ZipFile(sys.argv[1]).read("word/document.xml"))
tbl = root.find(f"{W}body").find(f"{W}tbl")
grid = [int(c.get(W + "w")) for c in tbl.find(f"{W}tblGrid").findall(f"{W}gridCol")]
# Longest requirement line is longer than NFR-01, so column 2 is wider than 1134.
if grid[1] <= int(sys.argv[2]):
    sys.exit(f"fitted requirement column {grid[1]} is not wider than the ID column")
# Column 2 must not gain noWrap. Column 1 still has it.
def starts(row, slot):
    acc = 0
    for tc in row.findall(f"{W}tc"):
        tcpr = tc.find(f"{W}tcPr")
        span = 1
        if tcpr is not None and (gs := tcpr.find(f"{W}gridSpan")) is not None:
            span = int(gs.get(W + "val") or "1")
        if acc == slot:
            return tc
        acc += span
    return None
for tr in tbl.findall(f"{W}tr"):
    tc = starts(tr, 1)
    tcpr = tc.find(f"{W}tcPr")
    if tcpr is not None and tcpr.find(f"{W}noWrap") is not None:
        sys.exit("width=fit must not set noWrap")
print("ok width=fit", grid)
PY

echo "== add table --prop idColumn=1 =="
"$OFFICECLI" add "$DOC" /body --type table --prop idColumn=1 \
  --prop data="ID,Requirement;T-01,The operator exports the register;NFR-01,Response stays under one second"
python3 - "$DOC" "$FIT6" <<'PY' || fail "add table idColumn=1"
import sys, zipfile
import xml.etree.ElementTree as ET
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
fit = int(sys.argv[2])
root = ET.fromstring(zipfile.ZipFile(sys.argv[1]).read("word/document.xml"))
tables = root.find(f"{W}body").findall(f"{W}tbl")
tbl = tables[1]
grid = [int(c.get(W + "w")) for c in tbl.find(f"{W}tblGrid").findall(f"{W}gridCol")]
if grid[0] != fit or grid[0] >= grid[1]:
    sys.exit(f"unexpected grid {grid}")
nowrap = list(tbl.iter(W + "noWrap"))
if len(nowrap) != 3:
    sys.exit(f"expected 3 noWrap, got {len(nowrap)}")
print("ok add", grid)
PY

echo "== idColumns=1,3 keeps both ID columns fitted =="
"$OFFICECLI" add "$DOC" /body --type table --prop idColumns=1,3 \
  --prop data="ID,Name,Code;T-01,Export register,FR-001;NFR-01,Response time,REQ-12"
python3 - "$DOC" "$FIT6" <<'PY' || fail "idColumns=1,3"
import sys, zipfile
import xml.etree.ElementTree as ET
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
fit = int(sys.argv[2])
root = ET.fromstring(zipfile.ZipFile(sys.argv[1]).read("word/document.xml"))
tbl = root.find(f"{W}body").findall(f"{W}tbl")[2]
grid = [int(c.get(W + "w")) for c in tbl.find(f"{W}tblGrid").findall(f"{W}gridCol")]
if grid[0] != fit or grid[2] != fit:
    sys.exit(f"ID columns {grid[0]}, {grid[2]} != {fit}")
if grid[1] <= fit:
    sys.exit(f"middle column was squeezed to {grid[1]}")
# noWrap only on columns 0 and 2 (6 cells).
count = 0
for tr in tbl.findall(f"{W}tr"):
    tcs = tr.findall(f"{W}tc")
    for i, tc in enumerate(tcs):
        has = tc.find(f"{W}tcPr/{W}noWrap") is not None
        if i in (0, 2):
            if not has:
                sys.exit(f"missing noWrap on col {i}")
            count += 1
        elif has:
            sys.exit("middle column has noWrap")
if count != 6:
    sys.exit(f"nowrap count {count}")
print("ok two columns", grid)
PY

echo "== add column --prop idColumn=true =="
"$OFFICECLI" add "$DOC" "/body/tbl[3]" --type column --prop idColumn=true --prop text="REQ-12"
python3 - "$DOC" "$FIT6" <<'PY' || fail "add column idColumn"
import sys, zipfile
import xml.etree.ElementTree as ET
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
fit = int(sys.argv[2])
root = ET.fromstring(zipfile.ZipFile(sys.argv[1]).read("word/document.xml"))
tbl = root.find(f"{W}body").findall(f"{W}tbl")[2]
grid = [int(c.get(W + "w")) for c in tbl.find(f"{W}tblGrid").findall(f"{W}gridCol")]
if len(grid) != 4 or grid[3] != fit:
    sys.exit(f"new column grid {grid}")
for tr in tbl.findall(f"{W}tr"):
    tc = tr.findall(f"{W}tc")[-1]
    text = "".join(t.text or "" for t in tc.iter(W + "t"))
    if text != "REQ-12":
        sys.exit(f"seed text {text!r}")
    if tc.find(f"{W}tcPr/{W}noWrap") is None:
        sys.exit("new column missing noWrap")
print("ok added column", grid)
PY

echo "== explicit width wins over the fit =="
"$OFFICECLI" add "$DOC" /body --type table --prop rows=2 --prop cols=2
"$OFFICECLI" set "$DOC" "/body/tbl[4]/tr[1]/tc[1]" --prop text="ID"
"$OFFICECLI" set "$DOC" "/body/tbl[4]/tr[2]/tc[1]" --prop text="NFR-01"
"$OFFICECLI" set "$DOC" "/body/tbl[4]/col[1]" --prop idColumn=true --prop width=2.2cm
python3 - "$DOC" "$EXPLICIT_2_2CM" <<'PY' || fail "explicit width was ignored"
import sys, zipfile
import xml.etree.ElementTree as ET
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
want = sys.argv[2]
root = ET.fromstring(zipfile.ZipFile(sys.argv[1]).read("word/document.xml"))
tbl = root.find(f"{W}body").findall(f"{W}tbl")[3]
got = tbl.find(f"{W}tblGrid").find(f"{W}gridCol").get(W + "w")
if got != want:
    sys.exit(f"grid {got} != {want}")
if next(tbl.iter(W + "noWrap"), None) is None:
    sys.exit("missing noWrap")
print("ok explicit", got)
PY

echo "== boolean idColumn on a table is rejected =="
if "$OFFICECLI" set "$DOC" "/body/tbl[4]" --prop idColumn=true >/dev/null 2>"$WORK/bool.err"; then
  fail "table idColumn=true should fail"
fi
grep -q '1-based' "$WORK/bool.err" || fail "expected index hint, got: $(cat "$WORK/bool.err")"

echo "== a fully merged column is rejected =="
"$OFFICECLI" add "$DOC" /body --type table --prop rows=1 --prop cols=2
"$OFFICECLI" set "$DOC" "/body/tbl[5]/tr[1]/tc[1]" --prop gridspan=2
if "$OFFICECLI" set "$DOC" "/body/tbl[5]/col[1]" --prop idColumn=true >/dev/null 2>"$WORK/merge.err"; then
  fail "merged column should be rejected"
fi
grep -q 'Unmerge' "$WORK/merge.err" || fail "expected unmerge hint, got: $(cat "$WORK/merge.err")"

echo "== htmlchunk col-id and white-space:nowrap =="
cat > "$WORK/ids.html" <<'EOF'
<style>
  table.data th.col-id, table.data td.col-id { white-space: nowrap; vertical-align: top; }
</style>
<table class="data">
  <tr><th class="col-id">ID</th><th>Requirement</th></tr>
  <tr><td class="col-id">NFR-01</td><td>Response stays under one second</td></tr>
  <tr><td class="col-id">T-01</td><td>The operator exports the register</td></tr>
</table>
<table>
  <tr><td style="white-space:nowrap">FR-001</td><td>A longer requirement statement</td></tr>
  <tr><td style="white-space:nowrap">REQ-12</td><td>Another requirement</td></tr>
</table>
EOF
HTML="$WORK/html.docx"
"$OFFICECLI" create "$HTML"
"$OFFICECLI" add "$HTML" /body --type htmlchunk --prop matchSrc=true --prop src="$WORK/ids.html"
"$OFFICECLI" materialize "$HTML"
"$OFFICECLI" validate "$HTML" >/dev/null
python3 - "$HTML" "$FIT6" <<'PY' || fail "materialize did not fit ID columns"
import sys, zipfile
import xml.etree.ElementTree as ET
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
fit = int(sys.argv[2])
root = ET.fromstring(zipfile.ZipFile(sys.argv[1]).read("word/document.xml"))
if root.find(f".//{W}altChunk") is not None:
    sys.exit("altChunk still present")
tables = root.find(f"{W}body").findall(f"{W}tbl")
if len(tables) != 2:
    sys.exit(f"expected 2 tables, got {len(tables)}")
for i, tbl in enumerate(tables):
    grid = [int(c.get(W + "w")) for c in tbl.find(f"{W}tblGrid").findall(f"{W}gridCol")]
    if grid[0] != fit or grid[0] >= grid[1]:
        sys.exit(f"table {i} grid {grid}")
    texts = []
    for tc in tbl.iter(W + "tc"):
        text = "".join(t.text or "" for t in tc.iter(W + "t"))
        has = tc.find(f"{W}tcPr/{W}noWrap") is not None
        texts.append((text, has))
        if text in ("ID", "NFR-01", "T-01", "FR-001", "REQ-12") and not has:
            sys.exit(f"missing noWrap on {text!r}")
        if has and text not in ("ID", "NFR-01", "T-01", "FR-001", "REQ-12"):
            sys.exit(f"nowrap on description {text!r}")
    print("ok materialize", i, grid)
PY

"$OFFICECLI" validate "$DOC" >/dev/null
echo "OK table-id-column"
