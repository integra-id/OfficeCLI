#!/usr/bin/env bash
# Smoke test for `officecli finalize`.
#
# Default pipeline, in order:
#   1. materialize HTML/XHTML/plain-text altChunks (RTF/MHT stay)
#   2. refresh --toc when a TOC field exists (page numbers stay 0)
#   3. pageSetup=a4-moderate only on sections with no w:pgSz
#   4. validate
#
# Also checks skip flags, --strict (file unchanged, later steps not run),
# --json, and that an existing page size — including the A4 stamp from
# `officecli create` — is not overwritten.
#
# Usage (from the repo root, after `dotnet build -c Release`):
#   OFFICECLI=path/to/officecli bash examples/word/finalize.sh
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
sha() { sha256sum "$1" | awk '{print $1}'; }

echo "== finalize --help documents the default steps and skip flags =="
HELP="$("$OFFICECLI" finalize --help)"
echo "$HELP" | grep -q -- '--no-materialize' || fail "help missing --no-materialize"
echo "$HELP" | grep -q -- '--no-toc' || fail "help missing --no-toc"
echo "$HELP" | grep -q -- '--no-page-setup' || fail "help missing --no-page-setup"
echo "$HELP" | grep -q -- '--no-validate' || fail "help missing --no-validate"
echo "$HELP" | grep -q -- '--strict' || fail "help missing --strict"
echo "$HELP" | grep -q 'a4-moderate' || fail "help does not document a4-moderate"
echo "$HELP" | grep -q 'w:pgSz' || fail "help does not say page setup is limited to sections without w:pgSz"
"$OFFICECLI" --help | grep -q 'finalize' || fail "root help does not list finalize"

echo "== unsupported type =="
"$OFFICECLI" create "$WORK/deck.pptx"
set +e
"$OFFICECLI" finalize "$WORK/deck.pptx" >/dev/null 2>"$WORK/pptx.err"
RC=$?
set -e
[[ "$RC" -ne 0 ]] || fail "finalize should reject .pptx"
grep -q 'docx' "$WORK/pptx.err" || fail "expected an unsupported-type error: $(cat "$WORK/pptx.err")"

echo "== default pipeline: materialize then TOC, page size left alone =="
DOC="$WORK/tech.docx"
"$OFFICECLI" create "$DOC"
"$OFFICECLI" add "$DOC" /body --type toc \
  --prop title="Contents" --prop levels=1-3 \
  --prop hyperlinks=true --prop pageNumbers=true
"$OFFICECLI" add "$DOC" /body --type paragraph --prop text="Native Chapter" --prop style=Heading1
"$OFFICECLI" add "$DOC" /body --type htmlchunk \
  --prop html='<h1>Chunk Title</h1><p>From the chunk.</p>'

doc_xml "$DOC" | grep -q 'w:altChunk' || fail "expected an altChunk before finalize"
doc_xml "$DOC" | grep -q 'Update field to see table of contents' \
  || fail "expected the TOC placeholder before finalize"
# officecli create stamps A4 and its own margins (right 1800 twips, not Moderate 1080).
doc_xml "$DOC" | grep -q 'w:w="11906"' || fail "create should stamp A4 width"
doc_xml "$DOC" | grep -q 'w:right="1800"' || fail "create margin right should still be 1800 before finalize"

MSG="$("$OFFICECLI" finalize "$DOC")"
echo "$MSG" | grep -q '^materialize: Materialized ' || fail "materialize step did not convert: $MSG"
echo "$MSG" | grep -q '^toc: Rebuilt TOC entries:' || fail "toc step did not rebuild: $MSG"
echo "$MSG" | grep -q 'placeholder 0' || fail "toc step did not admit placeholder page numbers: $MSG"
echo "$MSG" | grep -q '^pageSetup: Skipped' || fail "page setup should skip when w:pgSz exists: $MSG"
echo "$MSG" | grep -q 'a4-moderate was not applied' || fail "page setup skip message missing: $MSG"
echo "$MSG" | grep -q '^validate: Validation passed' || fail "validate step failed: $MSG"
echo "$MSG" | grep -q '^finalize: ok$' || fail "finalize did not report ok: $MSG"

XML="$(doc_xml "$DOC")"
echo "$XML" | grep -q 'w:altChunk' && fail "altChunk still present after finalize"
echo "$XML" | grep -q 'Update field to see table of contents' && fail "TOC placeholder survived finalize"
echo "$XML" | grep -q '>Chunk Title<' || fail "materialized heading text missing"
chunk_hits="$(echo "$XML" | grep -o '>Chunk Title<' | wc -l | tr -d ' ')"
[[ "$chunk_hits" -ge 2 ]] || fail "Chunk Title should appear in the body and the TOC (got $chunk_hits)"
echo "$XML" | grep -q '>Native Chapter<' || fail "native heading missing"
echo "$XML" | grep -q 'w:anchor="_Toc' || fail "TOC hyperlink missing — materialize should run before the TOC rebuild"
echo "$XML" | grep -q 'PAGEREF _Toc' || fail "PAGEREF missing"
echo "$XML" | grep -q '>0<' || fail "placeholder page number 0 missing"
# Existing page size and the create margins must survive.
echo "$XML" | grep -q 'w:w="11906"' || fail "A4 width was dropped"
echo "$XML" | grep -q 'w:right="1800"' || fail "existing right margin was overwritten"
echo "$XML" | grep -q 'w:right="1080"' && fail "Moderate margin was applied over an existing page size"

echo "== second finalize is a no-op on chunks and still refreshes the TOC =="
MSG2="$("$OFFICECLI" finalize "$DOC")"
echo "$MSG2" | grep -q 'No altChunks to materialize' || fail "second finalize should find no chunks: $MSG2"
echo "$MSG2" | grep -q '^toc: Rebuilt TOC entries:' || fail "second finalize should still rebuild the TOC: $MSG2"
echo "$MSG2" | grep -q '^pageSetup: Skipped' || fail "second finalize should still skip page setup: $MSG2"
echo "$MSG2" | grep -q '^finalize: ok$' || fail "second finalize failed: $MSG2"

echo "== --json reports steps =="
JSON="$("$OFFICECLI" finalize "$DOC" --json)"
echo "$JSON" | grep -q '"success": true' || fail "json success should be true: $JSON"
echo "$JSON" | grep -q '"name": "materialize"' || fail "json missing materialize step: $JSON"
echo "$JSON" | grep -q '"name": "toc"' || fail "json missing toc step: $JSON"
echo "$JSON" | grep -q '"name": "pageSetup"' || fail "json missing pageSetup step: $JSON"
echo "$JSON" | grep -q '"name": "validate"' || fail "json missing validate step: $JSON"
echo "$JSON" | grep -q '"status": "skipped"' || fail "json should skip page setup: $JSON"
echo "$JSON" | grep -q '"pageNumbers": "placeholder"' || fail "json toc detail should say placeholder: $JSON"

echo "== no TOC field skips the rebuild =="
NOTOC="$WORK/notoc.docx"
"$OFFICECLI" create "$NOTOC"
"$OFFICECLI" add "$NOTOC" /body --type paragraph --prop text="Just a paragraph"
MSGN="$("$OFFICECLI" finalize "$NOTOC")"
echo "$MSGN" | grep -q '^toc: Skipped - no TOC field' || fail "missing no-TOC skip: $MSGN"
echo "$MSGN" | grep -q '^finalize: ok$' || fail "no-TOC finalize failed: $MSGN"

echo "== --no-toc leaves the placeholder; --no-materialize leaves the chunk =="
SKIP="$WORK/skip.docx"
"$OFFICECLI" create "$SKIP"
"$OFFICECLI" add "$SKIP" /body --type toc --prop levels=1-3 --prop hyperlinks=true --prop pageNumbers=true
"$OFFICECLI" add "$SKIP" /body --type paragraph --prop text="Kept Heading" --prop style=Heading1
"$OFFICECLI" add "$SKIP" /body --type htmlchunk --prop html='<p>Still a chunk</p>'
"$OFFICECLI" finalize "$SKIP" --no-materialize --no-toc >/dev/null
SXML="$(doc_xml "$SKIP")"
echo "$SXML" | grep -q 'w:altChunk' || fail "--no-materialize removed the altChunk"
echo "$SXML" | grep -q 'Update field to see table of contents' || fail "--no-toc rebuilt the TOC"
echo "$SXML" | grep -q 'w:right="1800"' || fail "skip flags should not change create margins"

echo "== all skip flags leave the package bytes alone =="
IDLE="$WORK/idle.docx"
"$OFFICECLI" create "$IDLE"
"$OFFICECLI" add "$IDLE" /body --type paragraph --prop text="Untouched"
BEFORE="$(sha "$IDLE")"
"$OFFICECLI" finalize "$IDLE" --no-materialize --no-toc --no-page-setup --no-validate >/dev/null
AFTER="$(sha "$IDLE")"
[[ "$BEFORE" == "$AFTER" ]] || fail "skip-all finalize rewrote the package"

echo "== missing w:pgSz receives a4-moderate; a later run does not touch it =="
BARE="$WORK/bare.docx"
"$OFFICECLI" create "$BARE"
"$OFFICECLI" set "$BARE" /section[1] --prop pageSize=none
doc_xml "$BARE" | grep -q 'w:pgSz' && fail "pageSize=none should remove w:pgSz"
doc_xml "$BARE" | grep -q 'w:right="1800"' || fail "margins should remain after pageSize=none"
MSGB="$("$OFFICECLI" finalize "$BARE")"
echo "$MSGB" | grep -q '^pageSetup: Applied a4-moderate to /section\[1\]' \
  || fail "expected page setup on the bare section: $MSGB"
echo "$MSGB" | grep -q '^toc: Skipped - no TOC field' || fail "bare doc should skip TOC: $MSGB"
echo "$MSGB" | grep -q '^finalize: ok$' || fail "bare-section finalize failed: $MSGB"
BXML="$(doc_xml "$BARE")"
echo "$BXML" | grep -q 'w:w="11906"' || fail "a4 width was not written"
echo "$BXML" | grep -q 'w:h="16838"' || fail "a4 height was not written"
echo "$BXML" | grep -q 'w:left="1080"' || fail "moderate left margin missing"
echo "$BXML" | grep -q 'w:right="1080"' || fail "moderate right margin missing"
MSG3="$("$OFFICECLI" finalize "$BARE")"
echo "$MSG3" | grep -q '^pageSetup: Skipped' || fail "second pass should not reapply page setup: $MSG3"
doc_xml "$BARE" | grep -q 'w:right="1080"' || fail "second pass changed the moderate margin"

echo "== --no-page-setup does not invent w:pgSz =="
NOPG="$WORK/nopg.docx"
"$OFFICECLI" create "$NOPG"
"$OFFICECLI" set "$NOPG" /section[1] --prop pageSize=none
"$OFFICECLI" finalize "$NOPG" --no-page-setup >/dev/null
doc_xml "$NOPG" | grep -q 'w:pgSz' && fail "--no-page-setup wrote a page size"
doc_xml "$NOPG" | grep -q 'w:right="1800"' || fail "--no-page-setup changed margins"

echo "== RTF stays without --strict; HTML in the same file is converted =="
MIX="$WORK/mix.docx"
"$OFFICECLI" create "$MIX"
"$OFFICECLI" add "$MIX" /body --type htmlchunk --prop html='<p>Keep me</p>'
"$OFFICECLI" add "$MIX" /body --type htmlchunk --prop format=rtf --prop html='{\rtf1 hello}'
"$OFFICECLI" finalize "$MIX" --no-toc >/dev/null
MX="$(doc_xml "$MIX")"
echo "$MX" | grep -q '>Keep me<' || fail "HTML chunk was not materialized"
echo "$MX" | grep -q 'w:altChunk' || fail "RTF chunk should remain without --strict"

echo "== --strict refuses, leaves the file unchanged, and skips later steps =="
STRICT="$WORK/strict.docx"
"$OFFICECLI" create "$STRICT"
"$OFFICECLI" add "$STRICT" /body --type paragraph --prop text="Body Heading" --prop style=Heading1
"$OFFICECLI" add "$STRICT" /body --type toc --prop levels=1-3 --prop hyperlinks=true --prop pageNumbers=true
"$OFFICECLI" add "$STRICT" /body --type htmlchunk --prop html='<p>Will not convert</p>'
"$OFFICECLI" add "$STRICT" /body --type htmlchunk --prop format=rtf --prop html='{\rtf1 hello}'
BEFORE_S="$(sha "$STRICT")"
set +e
SMSG="$("$OFFICECLI" finalize "$STRICT" --strict 2>"$WORK/strict.err")"
SRC=$?
set -e
[[ "$SRC" -ne 0 ]] || fail "--strict should exit 1"
echo "$SMSG" | grep -q '^materialize: .*unchanged' || fail "strict materialize message missing: $SMSG"
echo "$SMSG" | grep -q '^toc: Not run' || fail "toc should not run after --strict: $SMSG"
echo "$SMSG" | grep -q '^pageSetup: Not run' || fail "page setup should not run after --strict: $SMSG"
echo "$SMSG" | grep -q '^validate: Not run' || fail "validate should not run after --strict: $SMSG"
echo "$SMSG" | grep -q '^finalize: failed$' || fail "strict finalize should report failed: $SMSG"
AFTER_S="$(sha "$STRICT")"
[[ "$BEFORE_S" == "$AFTER_S" ]] || fail "--strict modified the file"
doc_xml "$STRICT" | grep -q 'Update field to see table of contents' \
  || fail "TOC was rebuilt even though --strict should have stopped the pipeline"

set +e
SJSON="$("$OFFICECLI" finalize "$STRICT" --strict --json 2>/dev/null)"
SJRC=$?
set -e
[[ "$SJRC" -ne 0 ]] || fail "--strict --json should exit 1"
echo "$SJSON" | grep -q '"success": false' || fail "strict json should be success false: $SJSON"
echo "$SJSON" | grep -q '"status": "failed"' || fail "strict json missing failed step: $SJSON"
echo "$SJSON" | grep -q '"status": "not-run"' || fail "strict json missing not-run steps: $SJSON"
# The refusal must still have left the package alone.
[[ "$(sha "$STRICT")" == "$BEFORE_S" ]] || fail "strict --json modified the file"

echo "OK"
