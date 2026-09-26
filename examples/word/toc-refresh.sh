#!/usr/bin/env bash
# Smoke test for headless TOC entry rebuild (`officecli refresh --toc`).
#
# Builds a docx with Heading styles, a direct outline level, a custom style
# mapped with \t, and one native TOC field. Asserts that refresh --toc:
#   - writes entry titles and internal hyperlinks to _Toc bookmarks
#   - drops headings outside the requested level range
#   - keeps PAGEREF page numbers at the placeholder 0
#   - follows a later heading edit
#   - reuses the same _Toc bookmark names on a second run
#   - omits PAGEREF when pageNumbers=false
#   - honors \b bookmark scope
#   - succeeds with a clear message when the file has no TOC field
#
# Page numbers are NOT asserted to be real. This path does not paginate.
#
# Usage (from the repo root, after `dotnet build -c Release`):
#   OFFICECLI=path/to/officecli bash examples/word/toc-refresh.sh
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
styles_xml() { unzip -p "$1" word/styles.xml; }

echo "== refresh --help mentions --toc =="
"$OFFICECLI" refresh --help | grep -q -- '--toc' || fail "refresh --help does not document --toc"

echo "== build headings + toc =="
DOC="$WORK/toc.docx"
"$OFFICECLI" create "$DOC"
"$OFFICECLI" add "$DOC" /styles --type style \
  --prop id=Heading1 --prop name="heading 1" --prop type=paragraph --prop outlineLvl=0
"$OFFICECLI" add "$DOC" /styles --type style \
  --prop id=Heading2 --prop name="heading 2" --prop type=paragraph --prop outlineLvl=1
# No outlineLvl: the HeadingN style-id fallback must still collect it.
"$OFFICECLI" add "$DOC" /styles --type style \
  --prop id=Heading3 --prop name="heading 3" --prop type=paragraph
"$OFFICECLI" add "$DOC" /styles --type style \
  --prop id=Heading4 --prop name="heading 4" --prop type=paragraph --prop outlineLvl=3
"$OFFICECLI" add "$DOC" /styles --type style \
  --prop id=Procedure --prop name=Procedure --prop type=paragraph

"$OFFICECLI" add "$DOC" /body --type paragraph --prop text="Before Toc" --prop style=Heading1
"$OFFICECLI" add "$DOC" /body --type toc \
  --prop title="Contents" --prop levels=1-3 \
  --prop hyperlinks=true --prop pageNumbers=true \
  --prop customStyles="Procedure,2"
"$OFFICECLI" add "$DOC" /body --type paragraph --prop text="After Toc" --prop style=Heading1
"$OFFICECLI" add "$DOC" /body --type paragraph --prop text="Nested Section" --prop style=Heading2
"$OFFICECLI" add "$DOC" /body --type paragraph --prop text="Regex Heading" --prop style=Heading3
"$OFFICECLI" add "$DOC" /body --type paragraph --prop text="Direct Outline" --prop outlineLvl=1
"$OFFICECLI" add "$DOC" /body --type paragraph --prop text="How to run" --prop style=Procedure
"$OFFICECLI" add "$DOC" /body --type paragraph --prop text="Too Deep" --prop style=Heading4
"$OFFICECLI" add "$DOC" /body --type paragraph --prop text="Just body" --prop style=Normal

# Placeholder only, before refresh.
doc_xml "$DOC" | grep -q 'Update field to see table of contents' \
  || fail "expected the TOC placeholder before refresh"

echo "== refresh --toc =="
MSG="$("$OFFICECLI" refresh "$DOC" --toc)"
echo "$MSG" | grep -q 'backend: toc-entries' || fail "text result missing toc-entries: $MSG"
echo "$MSG" | grep -q 'placeholder 0' || fail "text result does not admit the page-number placeholder: $MSG"

JSON="$("$OFFICECLI" refresh "$DOC" --toc --json)"
echo "$JSON" | grep -q '"backend": *"toc-entries"' || fail "json missing backend: $JSON"
echo "$JSON" | grep -q '"pageNumbers": *"placeholder"' || fail "json pageNumbers should be placeholder: $JSON"
echo "$JSON" | grep -q '"entries": *6' || fail "expected 6 entries: $JSON"

XML="$(doc_xml "$DOC")"
echo "$XML" | grep -q 'Update field to see table of contents' \
  && fail "placeholder survived refresh"

for title in "Before Toc" "After Toc" "Nested Section" "Regex Heading" "Direct Outline" "How to run"; do
  echo "$XML" | grep -q ">$title<" || fail "missing TOC/heading text: $title"
done
echo "$XML" | grep -q '>Too Deep<' || fail "Heading4 paragraph missing from the body"
# Too Deep is a body heading but must not be a TOC hyperlink target text inside w:hyperlink.
# Count hyperlink anchors: one per included heading, not for Too Deep / Just body / Contents.
anchors="$(echo "$XML" | grep -o 'w:anchor="_Toc[^"]*"' | wc -l | tr -d ' ')"
[[ "$anchors" == "6" ]] || fail "expected 6 TOC hyperlinks, got $anchors"

echo "$XML" | grep -q 'PAGEREF _Toc' || fail "missing PAGEREF"
echo "$XML" | grep -q '>0<' || fail "missing placeholder page number 0"
echo "$XML" | grep -q 'w:name="_Toc' || fail "heading is missing a _Toc bookmark"
echo "$XML" | grep -q '>Contents<' || fail "TOC title was dropped"
# Title uses TOCHeading, not a TOC entry style, and is not a 7th hyperlink (checked above).

styles_xml "$DOC" | grep -q 'w:styleId="TOC1"' || fail "TOC1 style was not created"
styles_xml "$DOC" | grep -q 'w:styleId="TOC2"' || fail "TOC2 style was not created"
echo "$XML" | grep -q 'w:val="TOC1"' || fail "entry missing TOC1 style"
echo "$XML" | grep -q 'w:val="TOC2"' || fail "entry missing TOC2 style"
begins="$(echo "$XML" | grep -o 'w:fldCharType="begin"' | wc -l | tr -d ' ')"
ends="$(echo "$XML" | grep -o 'w:fldCharType="end"' | wc -l | tr -d ' ')"
[[ "$begins" == "$ends" ]] || fail "unbalanced field chars: begin=$begins end=$ends"
while read -r anchor; do
  [[ -z "$anchor" ]] && continue
  echo "$XML" | grep -q "w:name=\"${anchor}\"" || fail "hyperlink anchor ${anchor} has no heading bookmark"
done < <(echo "$XML" | grep -o 'w:anchor="_Toc[^"]*"' | sed -E 's/w:anchor="([^"]+)"/\1/')
"$OFFICECLI" validate "$DOC" | grep -q 'no errors found' || fail "validate failed after TOC rebuild"

names="$(echo "$XML" | grep -o 'w:name="_Toc[^"]*"' | sort)"

echo "== heading edit is picked up; bookmarks stay =="
# Body order: title para, empty? Let's find the After Toc paragraph via set by text.
# Paragraphs: [1] Before Toc, then TOC title, TOC field paras, After Toc...
# Use query/set by matching text through a known path. Before Toc is /body/p[1].
"$OFFICECLI" set "$DOC" '/body/p[1]' --prop text="Renamed Chapter"
"$OFFICECLI" refresh "$DOC" --toc >/dev/null
XML="$(doc_xml "$DOC")"
echo "$XML" | grep -q '>Renamed Chapter<' || fail "renamed heading missing after second refresh"
echo "$XML" | grep -q '>Before Toc<' && fail "old heading title still present"
names2="$(echo "$XML" | grep -o 'w:name="_Toc[^"]*"' | sort)"
[[ "$names" == "$names2" ]] || fail " _Toc bookmark names changed on the second refresh"

echo "== pageNumbers=false omits PAGEREF =="
NOPAGE="$WORK/nopage.docx"
"$OFFICECLI" create "$NOPAGE"
"$OFFICECLI" add "$NOPAGE" /styles --type style \
  --prop id=Heading1 --prop name="heading 1" --prop type=paragraph --prop outlineLvl=0
"$OFFICECLI" add "$NOPAGE" /body --type paragraph --prop text="Only Title" --prop style=Heading1
"$OFFICECLI" add "$NOPAGE" /body --type toc \
  --prop levels=1-3 --prop hyperlinks=true --prop pageNumbers=false
"$OFFICECLI" refresh "$NOPAGE" --toc >/dev/null
NXML="$(doc_xml "$NOPAGE")"
echo "$NXML" | grep -q 'w:anchor="_Toc' || fail "hyperlink missing when page numbers are off"
echo "$NXML" | grep -q 'PAGEREF' && fail "PAGEREF present despite pageNumbers=false"
echo "$NXML" | grep -q '>Only Title<' || fail "title entry missing"

echo "== no headings =="
EMPTY="$WORK/empty.docx"
"$OFFICECLI" create "$EMPTY"
"$OFFICECLI" add "$EMPTY" /body --type toc --prop levels=1-3 --prop hyperlinks=true
"$OFFICECLI" refresh "$EMPTY" --toc >/dev/null
doc_xml "$EMPTY" | grep -q 'No table of contents entries found.' \
  || fail "empty TOC did not report that there are no entries"

echo "== no TOC field is a successful no-op =="
NONE="$WORK/none.docx"
"$OFFICECLI" create "$NONE"
"$OFFICECLI" add "$NONE" /body --type paragraph --prop text="Hello"
NOMSG="$("$OFFICECLI" refresh "$NONE" --toc)"
echo "$NOMSG" | grep -q 'No TOC field found' || fail "missing no-TOC message: $NOMSG"

echo "== bookmark scope (\\b) =="
SCOPE="$WORK/scope.docx"
"$OFFICECLI" create "$SCOPE"
"$OFFICECLI" add "$SCOPE" /styles --type style \
  --prop id=Heading1 --prop name="heading 1" --prop type=paragraph --prop outlineLvl=0
"$OFFICECLI" add "$SCOPE" /body --type paragraph --prop text="Inside" --prop style=Heading1
"$OFFICECLI" add "$SCOPE" /body --type paragraph --prop text="Outside" --prop style=Heading1
# Bookmark the first heading only. /body/p[1] is "Inside".
"$OFFICECLI" add "$SCOPE" '/body/p[1]' --type bookmark --prop name=OnlyHere --prop open=true --index 0
"$OFFICECLI" add "$SCOPE" '/body/p[1]' --type bookmark --prop name=OnlyHere --prop end=true
"$OFFICECLI" add "$SCOPE" /body --type toc \
  --prop levels=1-3 --prop hyperlinks=true --prop pageNumbers=false \
  --prop bookmark=OnlyHere
"$OFFICECLI" refresh "$SCOPE" --toc >/dev/null
SXML="$(doc_xml "$SCOPE")"
# Both headings remain in the body. Only Inside is hyperlinked from the TOC.
echo "$SXML" | grep -q '>Inside<' || fail "scoped heading missing"
echo "$SXML" | grep -q '>Outside<' || fail "out-of-scope heading was deleted from the body"
sanchors="$(echo "$SXML" | grep -o 'w:anchor="_Toc[^"]*"' | wc -l | tr -d ' ')"
[[ "$sanchors" == "1" ]] || fail "bookmark scope should produce 1 TOC hyperlink, got $sanchors"

echo "== default refresh still rebuilds entries without --toc =="
DEF="$WORK/default.docx"
"$OFFICECLI" create "$DEF"
"$OFFICECLI" add "$DEF" /styles --type style \
  --prop id=Heading1 --prop name="heading 1" --prop type=paragraph --prop outlineLvl=0
"$OFFICECLI" add "$DEF" /body --type paragraph --prop text="Fallback Chapter" --prop style=Heading1
"$OFFICECLI" add "$DEF" /body --type toc --prop levels=1-3 --prop hyperlinks=true --prop pageNumbers=true
set +e
DEFMSG="$(timeout 90 "$OFFICECLI" refresh "$DEF" 2>&1)"
defrc=$?
set -e
[[ "$defrc" == "0" ]] || fail "default refresh should succeed when a TOC can be rebuilt (exit $defrc): $DEFMSG"
echo "$DEFMSG" | grep -Eq 'backend: (toc-entries|html|word)' \
  || fail "default refresh did not name a backend: $DEFMSG"
doc_xml "$DEF" | grep -q 'w:anchor="_Toc' || fail "default refresh dropped TOC hyperlinks"
doc_xml "$DEF" | grep -q '>Fallback Chapter<' || fail "default refresh dropped the heading title"

echo "OK"
