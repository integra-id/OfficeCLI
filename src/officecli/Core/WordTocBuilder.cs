// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeCli.Core;

/// <summary>
/// Rebuild TOC field results from the document's current headings.
///
/// Headless: titles and in-document hyperlinks are written into the TOC
/// field's cached result. PAGEREF page numbers are the placeholder "0".
/// Real pagination still needs Microsoft Word or the HTML layout fallback
/// in <see cref="WordHtmlRefresh"/> — this builder does not invent page numbers.
///
/// <c>\z</c> suppresses page numbers. That matches <c>add --type toc</c>,
/// which emits <c>\z</c> when <c>pageNumbers=false</c>. It is not Word's
/// "hide in web layout" meaning of <c>\z</c>.
/// </summary>
internal static class WordTocBuilder
{
    internal readonly record struct TocRebuildResult(
        int TocFields, int Entries, int Skipped, bool WrotePagePlaceholders);

    /// <summary>
    /// Rebuild every TOC field in the body. The package is not saved; the
    /// caller saves when <see cref="TocRebuildResult.TocFields"/> exceeds
    /// <see cref="TocRebuildResult.Skipped"/>.
    /// </summary>
    public static TocRebuildResult RegenerateAllTocs(WordprocessingDocument doc)
    {
        var body = doc.MainDocumentPart?.Document?.Body;
        if (body == null) return default;

        var fields = FindTocFields(body);
        if (fields.Count == 0) return default;

        var rebuildable = new List<TocField>();
        int skipped = 0;
        foreach (var field in fields)
        {
            if (field.Separate == null || field.EndParagraph == null || !field.SafeToReplace)
            {
                skipped++;
                continue;
            }
            rebuildable.Add(field);
        }

        var insideResult = new HashSet<Paragraph>();
        var openers = new HashSet<Paragraph>();
        foreach (var field in rebuildable)
        {
            openers.Add(field.Opener);
            MarkResultParagraphs(field, insideResult);
        }

        var headings = EnumerateHeadings(doc, body, insideResult, openers);
        var included = new List<(HeadingInfo Heading, int Level, TocField Field)>();
        foreach (var field in rebuildable)
        {
            foreach (var h in headings)
            {
                var level = LevelFor(h, field.Spec);
                if (level != null)
                    included.Add((h, level.Value, field));
            }
        }

        EnsureHeadingBookmarks(doc, included.Select(x => x.Heading).Distinct());

        int entries = 0;
        bool wrotePages = false;
        int maxLevel = 0;
        foreach (var field in rebuildable)
        {
            var paras = new List<Paragraph>();
            foreach (var h in headings)
            {
                var level = LevelFor(h, field.Spec);
                if (level == null) continue;
                bool page = PageFor(field.Spec, level.Value);
                if (page) wrotePages = true;
                if (level.Value > maxLevel) maxLevel = level.Value;
                paras.Add(BuildEntryParagraph(h, level.Value, field.Spec.Hyperlinks, page));
                entries++;
            }
            if (paras.Count == 0)
                paras.Add(EmptyResultParagraph());
            ReplaceTocFieldContent(field, paras);
        }

        if (rebuildable.Count > 0)
            EnsureTocStyles(doc, maxLevel);

        return new TocRebuildResult(fields.Count, entries, skipped, wrotePages);
    }

    // ==================== TOC field location ====================

    sealed class TocField
    {
        public Paragraph Opener { get; init; } = null!;
        public TocSpec Spec { get; init; } = null!;
        public Run? Separate { get; init; }
        public Run? EndRun { get; init; }
        public Paragraph? EndParagraph { get; init; }
        public bool SafeToReplace { get; init; }
    }

    sealed record TocSpec(
        int MinLevel,
        int MaxLevel,
        bool Hyperlinks,
        bool SuppressAllPages,
        int NoPageMin,
        int NoPageMax,
        string? Bookmark,
        Dictionary<string, int> CustomStyles);

    static List<TocField> FindTocFields(Body body)
    {
        var list = new List<TocField>();
        foreach (var p in body.Descendants<Paragraph>())
        {
            if (p.Ancestors<TextBoxContent>().Any()) continue;
            var begin = FindTocBegin(p, out var instr);
            if (begin == null || instr == null) continue;
            var spec = ParseTocSwitches(instr);
            FindFieldEnd(body, begin, out var separate, out var endRun, out var endPara);
            bool safe = false;
            if (separate != null && endRun != null && endPara != null)
            {
                if (ReferenceEquals(endPara, p))
                {
                    var sepRoot = RootChild(separate, p);
                    var endRoot = RootChild(endRun, p);
                    safe = IsFollowingSibling(sepRoot, endRoot);
                }
                else
                    safe = endPara.Parent == p.Parent && IsFollowingSibling(p, endPara);
            }
            list.Add(new TocField
            {
                Opener = p,
                Spec = spec,
                Separate = separate,
                EndRun = endRun,
                EndParagraph = endPara,
                SafeToReplace = safe,
            });
        }
        return list;
    }

    static Run? FindTocBegin(Paragraph opener, out string? instruction)
    {
        instruction = null;
        Run? begin = null;
        int depth = 0;
        var sb = new StringBuilder();
        foreach (var run in opener.Descendants<Run>())
        {
            var fc = run.GetFirstChild<FieldChar>();
            var kind = fc?.FieldCharType?.Value;
            if (kind == FieldCharValues.Begin)
            {
                if (depth == 0)
                {
                    begin = run;
                    sb.Clear();
                }
                depth++;
            }
            else if (kind == FieldCharValues.End)
            {
                depth = Math.Max(0, depth - 1);
                if (depth == 0) begin = null;
            }
            else if (kind == FieldCharValues.Separate && depth == 1)
            {
                if (IsTocInstruction(sb.ToString()))
                {
                    instruction = sb.ToString();
                    return begin;
                }
            }
            else if (depth >= 1)
            {
                var code = run.GetFirstChild<FieldCode>();
                if (code?.Text != null) sb.Append(code.Text);
            }
        }
        return null;
    }

    static bool IsTocInstruction(string raw)
    {
        var instr = raw.Trim();
        if (instr.Length < 3) return false;
        if (!instr.StartsWith("TOC", StringComparison.OrdinalIgnoreCase)) return false;
        return instr.Length == 3 || char.IsWhiteSpace(instr[3]);
    }

    static void FindFieldEnd(Body body, Run begin, out Run? separate, out Run? endRun, out Paragraph? endPara)
    {
        separate = null;
        endRun = null;
        endPara = null;
        bool started = false;
        int depth = 0;
        foreach (var run in body.Descendants<Run>())
        {
            if (!started)
            {
                if (!ReferenceEquals(run, begin)) continue;
                started = true;
                depth = 1;
                continue;
            }
            var fc = run.GetFirstChild<FieldChar>();
            var kind = fc?.FieldCharType?.Value;
            if (kind == FieldCharValues.Begin) depth++;
            else if (kind == FieldCharValues.Separate && depth == 1 && separate == null)
                separate = run;
            else if (kind == FieldCharValues.End)
            {
                depth--;
                if (depth == 0)
                {
                    endRun = run;
                    endPara = run.Ancestors<Paragraph>().FirstOrDefault();
                    return;
                }
            }
        }
    }

    static bool IsFollowingSibling(OpenXmlElement start, OpenXmlElement candidate)
    {
        for (var n = start.NextSibling(); n != null; n = n.NextSibling())
            if (ReferenceEquals(n, candidate)) return true;
        return false;
    }

    static TocSpec ParseTocSwitches(string instr)
    {
        var min = 1;
        var max = 3;
        var m = Regex.Match(instr, @"\\o\s+""\s*(\d+)\s*-\s*(\d+)\s*""");
        if (!m.Success)
            m = Regex.Match(instr, @"\\o\s+(\d+)\s*-\s*(\d+)");
        if (m.Success)
        {
            min = ClampLevel(int.Parse(m.Groups[1].Value));
            max = ClampLevel(int.Parse(m.Groups[2].Value));
            if (max < min) (min, max) = (max, min);
        }

        var hyperlinks = Regex.IsMatch(instr, @"\\h\b");
        // \z is pageNumbers=false in this fork (see AddToc). A bare \n is
        // Word's "omit page numbers" switch. \n "2-4" omits only that band.
        var rangedPages = Regex.Match(instr, @"\\n\s*""\s*(\d+)\s*-\s*(\d+)\s*""");
        var suppressAll = Regex.IsMatch(instr, @"\\z\b")
            || (Regex.IsMatch(instr, @"\\n\b") && !rangedPages.Success);
        var noMin = 0;
        var noMax = 0;
        if (!suppressAll && rangedPages.Success)
        {
            noMin = ClampLevel(int.Parse(rangedPages.Groups[1].Value));
            noMax = ClampLevel(int.Parse(rangedPages.Groups[2].Value));
            if (noMax < noMin) (noMin, noMax) = (noMax, noMin);
        }

        string? bookmark = null;
        var cb = Regex.Match(instr, @"\\b\s+""([^""]+)""");
        if (cb.Success) bookmark = cb.Groups[1].Value;

        var custom = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var ct = Regex.Match(instr, @"\\t\s+""([^""]+)""");
        if (ct.Success) ParseCustomStyles(ct.Groups[1].Value, custom);

        return new TocSpec(min, max, hyperlinks, suppressAll, noMin, noMax, bookmark, custom);
    }

    static void ParseCustomStyles(string raw, Dictionary<string, int> into)
    {
        var parts = raw.Split(',');
        var i = 0;
        while (i + 1 < parts.Length)
        {
            var name = parts[i].Trim();
            if (name.Length > 0
                && int.TryParse(parts[i + 1].Trim(), out var lvl)
                && lvl >= 1 && lvl <= 9)
            {
                into[name] = lvl;
                i += 2;
            }
            else i++;
        }
    }

    static int ClampLevel(int level) => Math.Clamp(level, 1, 9);

    static bool PageFor(TocSpec spec, int level)
    {
        if (spec.SuppressAllPages) return false;
        if (spec.NoPageMin > 0 && level >= spec.NoPageMin && level <= spec.NoPageMax) return false;
        return true;
    }

    static void MarkResultParagraphs(TocField field, HashSet<Paragraph> sink)
    {
        var end = field.EndParagraph;
        if (end == null) return;
        sink.Add(end);
        if (ReferenceEquals(field.Opener, end)) return;
        for (var n = field.Opener.NextSibling(); n != null && !ReferenceEquals(n, end); n = n.NextSibling())
        {
            if (n is Paragraph p) sink.Add(p);
            foreach (var d in n.Descendants<Paragraph>()) sink.Add(d);
        }
    }

    // ==================== Phase 1: Heading enumeration ====================

    internal sealed class HeadingInfo
    {
        public Paragraph Para { get; }
        public int OutlineLevel { get; }
        public string StyleId { get; }
        public string StyleName { get; }
        public string Text { get; }
        public string BookmarkName { get; set; } = "";
        public HeadingInfo(Paragraph p, int outlineLevel, string styleId, string styleName, string text)
        {
            Para = p;
            OutlineLevel = outlineLevel;
            StyleId = styleId;
            StyleName = styleName;
            Text = text;
        }
    }

    static List<HeadingInfo> EnumerateHeadings(
        WordprocessingDocument doc, Body body,
        HashSet<Paragraph> insideResult, HashSet<Paragraph> openers)
    {
        var styleLevels = ResolveHeadingStyleLevels(doc);
        var styleNames = ResolveStyleNames(doc);
        var list = new List<HeadingInfo>();
        foreach (var p in body.Descendants<Paragraph>())
        {
            if (p.Ancestors<TextBoxContent>().Any()) continue;
            if (insideResult.Contains(p) || openers.Contains(p)) continue;

            var styleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "";
            styleNames.TryGetValue(styleId, out var styleName);
            styleName ??= "";
            var level = ResolveOutlineLevel(p, styleId, styleName, styleLevels);
            var text = ExtractHeadingText(p);
            if (string.IsNullOrWhiteSpace(text)) continue;
            list.Add(new HeadingInfo(p, level, styleId, styleName, text));
        }
        return list;
    }

    static int? LevelFor(HeadingInfo h, TocSpec spec)
    {
        if (!string.IsNullOrEmpty(spec.Bookmark) && !ParagraphInBookmark(h.Para, spec.Bookmark))
            return null;

        if (h.OutlineLevel >= spec.MinLevel && h.OutlineLevel <= spec.MaxLevel)
            return h.OutlineLevel;

        if (spec.CustomStyles.Count > 0)
        {
            if (spec.CustomStyles.TryGetValue(h.StyleId, out var byId)
                && byId >= spec.MinLevel && byId <= spec.MaxLevel)
                return byId;
            if (h.StyleName.Length > 0
                && spec.CustomStyles.TryGetValue(h.StyleName, out var byName)
                && byName >= spec.MinLevel && byName <= spec.MaxLevel)
                return byName;
        }
        return null;
    }

    static bool ParagraphInBookmark(Paragraph para, string name)
    {
        var body = para.Ancestors<Body>().FirstOrDefault();
        if (body == null) return false;
        BookmarkStart? start = null;
        foreach (var bs in body.Descendants<BookmarkStart>())
        {
            if (string.Equals(bs.Name?.Value, name, StringComparison.Ordinal))
            {
                start = bs;
                break;
            }
        }
        if (start?.Id?.Value == null) return false;
        var id = start.Id.Value;
        var startPara = start.Ancestors<Paragraph>().FirstOrDefault();
        if (startPara == null) return false;
        BookmarkEnd? end = null;
        foreach (var be in body.Descendants<BookmarkEnd>())
        {
            if (be.Id?.Value == id) { end = be; break; }
        }
        var endPara = end?.Ancestors<Paragraph>().FirstOrDefault() ?? startPara;

        var paras = body.Descendants<Paragraph>().ToList();
        int i0 = IndexOfRef(paras, startPara);
        int i1 = IndexOfRef(paras, endPara);
        int ip = IndexOfRef(paras, para);
        if (i0 < 0 || i1 < 0 || ip < 0) return false;
        if (i1 < i0) (i0, i1) = (i1, i0);
        return ip >= i0 && ip <= i1;
    }

    static int IndexOfRef<T>(List<T> list, T item) where T : class
    {
        for (int i = 0; i < list.Count; i++)
            if (ReferenceEquals(list[i], item)) return i;
        return -1;
    }

    static Dictionary<string, int> ResolveHeadingStyleLevels(WordprocessingDocument doc)
    {
        var styles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        var direct = new Dictionary<string, int>(StringComparer.Ordinal);
        var basedOn = new Dictionary<string, string>(StringComparer.Ordinal);
        if (styles != null)
        {
            foreach (var s in styles.Elements<Style>())
            {
                var id = s.StyleId?.Value;
                if (string.IsNullOrEmpty(id)) continue;
                var lvl = s.StyleParagraphProperties?.OutlineLevel?.Val?.Value;
                if (lvl is >= 0 and <= 8)
                    direct[id] = lvl.Value + 1;
                var parent = s.BasedOn?.Val?.Value;
                if (!string.IsNullOrEmpty(parent))
                    basedOn[id] = parent;
            }
        }

        var resolved = new Dictionary<string, int>(direct, StringComparer.Ordinal);
        foreach (var id in basedOn.Keys)
        {
            if (resolved.ContainsKey(id)) continue;
            var seen = new HashSet<string>(StringComparer.Ordinal) { id };
            var cur = basedOn[id];
            while (!string.IsNullOrEmpty(cur) && seen.Add(cur))
            {
                if (direct.TryGetValue(cur, out var lv))
                {
                    resolved[id] = lv;
                    break;
                }
                if (!basedOn.TryGetValue(cur, out cur!)) break;
            }
        }
        return resolved;
    }

    static Dictionary<string, string> ResolveStyleNames(WordprocessingDocument doc)
    {
        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        var styles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (styles == null) return map;
        foreach (var s in styles.Elements<Style>())
        {
            var id = s.StyleId?.Value;
            var name = s.StyleName?.Val?.Value;
            if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(name))
                map[id] = name;
        }
        return map;
    }

    static int ResolveOutlineLevel(Paragraph p, string styleId, string styleName, Dictionary<string, int> styleLevels)
    {
        var direct = p.ParagraphProperties?.OutlineLevel?.Val?.Value;
        if (direct is >= 0 and <= 8) return direct.Value + 1;
        if (styleId.Length > 0 && styleLevels.TryGetValue(styleId, out var sl)) return sl;
        var idMatch = Regex.Match(styleId, @"^Heading([1-9])$", RegexOptions.IgnoreCase);
        if (idMatch.Success) return int.Parse(idMatch.Groups[1].Value);
        var name = styleName.Length > 0 ? styleName : styleId;
        if (name.Contains("Heading", StringComparison.OrdinalIgnoreCase)
            || name.Contains("标题", StringComparison.Ordinal)
            || name.Contains("標題", StringComparison.Ordinal))
        {
            foreach (var ch in name)
                if (char.IsDigit(ch) && ch >= '1' && ch <= '9')
                    return ch - '0';
        }
        return -1;
    }

    static string ExtractHeadingText(Paragraph p)
    {
        var sb = new StringBuilder();
        foreach (var t in p.Descendants<Text>())
            sb.Append(t.Text);
        return sb.ToString().Trim();
    }

    // ==================== Phase 2: Bookmark management ====================

    static void EnsureHeadingBookmarks(WordprocessingDocument doc, IEnumerable<HeadingInfo> headings)
    {
        int maxId = MaxBookmarkId(doc);
        foreach (var h in headings)
        {
            var existing = h.Para.Descendants<BookmarkStart>()
                .FirstOrDefault(b => b.Name?.Value?.StartsWith("_Toc", StringComparison.Ordinal) == true);
            if (existing?.Name?.Value != null)
            {
                h.BookmarkName = existing.Name.Value;
                continue;
            }
            var name = $"_Toc{Guid.NewGuid().ToString("N")[..8]}";
            var bookmarkId = (++maxId).ToString();
            var pPr = h.Para.GetFirstChild<ParagraphProperties>();
            var bs = new BookmarkStart { Id = bookmarkId, Name = name };
            var be = new BookmarkEnd { Id = bookmarkId };
            if (pPr != null) pPr.InsertAfterSelf(bs);
            else h.Para.PrependChild(bs);
            h.Para.AppendChild(be);
            h.BookmarkName = name;
        }
    }

    static int MaxBookmarkId(WordprocessingDocument doc)
    {
        int max = 0;
        void Scan(OpenXmlElement? root)
        {
            if (root == null) return;
            foreach (var b in root.Descendants<BookmarkStart>())
                if (int.TryParse(b.Id?.Value, out var n) && n > max) max = n;
        }
        var main = doc.MainDocumentPart;
        if (main == null) return 0;
        Scan(main.Document);
        foreach (var h in main.HeaderParts) Scan(h.Header);
        foreach (var f in main.FooterParts) Scan(f.Footer);
        if (main.FootnotesPart != null) Scan(main.FootnotesPart.Footnotes);
        if (main.EndnotesPart != null) Scan(main.EndnotesPart.Endnotes);
        return max;
    }

    // ==================== Phase 3: Entry generation ====================

    static Paragraph BuildEntryParagraph(HeadingInfo h, int level, bool hyperlinks, bool pageNumber)
    {
        var p = new Paragraph(new ParagraphProperties(
            new ParagraphStyleId { Val = "TOC" + level }));

        OpenXmlElement host = p;
        if (hyperlinks && h.BookmarkName.Length > 0)
        {
            var hyper = new Hyperlink { Anchor = h.BookmarkName, History = OnOffValue.FromBoolean(true) };
            p.AppendChild(hyper);
            host = hyper;
        }
        host.AppendChild(new Run(new Text(h.Text) { Space = SpaceProcessingModeValues.Preserve }));

        if (pageNumber)
        {
            host.AppendChild(new Run(new TabChar()));
            // Placeholder. WordHtmlRefresh overwrites this when a layout
            // engine can map the _Toc bookmark to a page. Do not invent one.
            host.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
            host.AppendChild(new Run(new FieldCode($" PAGEREF {h.BookmarkName} \\h ")
            { Space = SpaceProcessingModeValues.Preserve }));
            host.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
            host.AppendChild(new Run(new Text("0") { Space = SpaceProcessingModeValues.Preserve }));
            host.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        }
        return p;
    }

    static Paragraph EmptyResultParagraph()
        => new(new Run(new Text("No table of contents entries found.") { Space = SpaceProcessingModeValues.Preserve }));

    static void ReplaceTocFieldContent(TocField field, List<Paragraph> entries)
    {
        var sepRun = field.Separate!;
        var endRun = field.EndRun!;
        var opener = field.Opener;
        var endPara = field.EndParagraph!;

        var sepRoot = RootChild(sepRun, opener);
        if (ReferenceEquals(opener, endPara))
        {
            var endRoot = RootChild(endRun, opener);
            RemoveAfterThrough(sepRoot, endRoot);
        }
        else
        {
            RemoveAfterThrough(sepRoot, null);
            RemoveSiblingsThrough(opener, endPara);
        }

        OpenXmlElement cursor = opener;
        foreach (var entry in entries)
            cursor = cursor.InsertAfterSelf(entry);
        cursor.InsertAfterSelf(new Paragraph(new Run(new FieldChar { FieldCharType = FieldCharValues.End })));
    }

    static OpenXmlElement RootChild(OpenXmlElement node, OpenXmlElement parent)
    {
        var cur = node;
        while (cur.Parent != null && !ReferenceEquals(cur.Parent, parent))
            cur = cur.Parent;
        return cur;
    }

    /// <summary>Remove siblings after <paramref name="startExclusive"/> through
    /// <paramref name="endInclusive"/> (or the rest of the parent when it is null).</summary>
    static void RemoveAfterThrough(OpenXmlElement startExclusive, OpenXmlElement? endInclusive)
    {
        var node = startExclusive.NextSibling();
        while (node != null)
        {
            var next = node.NextSibling();
            bool done = endInclusive != null && ReferenceEquals(node, endInclusive);
            node.Remove();
            if (done) break;
            node = next;
        }
    }

    static void RemoveSiblingsThrough(OpenXmlElement start, OpenXmlElement endInclusive)
    {
        var node = start.NextSibling();
        while (node != null)
        {
            var next = node.NextSibling();
            bool done = ReferenceEquals(node, endInclusive);
            node.Remove();
            if (done) break;
            node = next;
        }
    }

    static void EnsureTocStyles(WordprocessingDocument doc, int maxLevel)
    {
        if (maxLevel < 1) return;
        var main = doc.MainDocumentPart;
        if (main == null) return;
        var stylesPart = main.StyleDefinitionsPart ?? main.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles ??= new Styles();
        var styles = stylesPart.Styles;
        bool hasNormal = styles.Elements<Style>()
            .Any(s => string.Equals(s.StyleId?.Value, "Normal", StringComparison.OrdinalIgnoreCase));

        for (int level = 1; level <= maxLevel && level <= 9; level++)
        {
            var id = "TOC" + level;
            if (styles.Elements<Style>().Any(s => string.Equals(s.StyleId?.Value, id, StringComparison.OrdinalIgnoreCase)))
                continue;
            var style = new Style { Type = StyleValues.Paragraph, StyleId = id };
            style.AppendChild(new StyleName { Val = "toc " + level });
            if (hasNormal)
                style.AppendChild(new BasedOn { Val = "Normal" });
            style.AppendChild(new UIPriority { Val = 39 });
            style.AppendChild(new UnhideWhenUsed());
            // CT_PPr sequence: tabs precedes ind. The reverse is rejected
            // as an unexpected child.
            var pPr = new StyleParagraphProperties();
            pPr.AppendChild(new Tabs(new TabStop
            {
                Val = TabStopValues.Right,
                Leader = TabStopLeaderCharValues.Dot,
                Position = 9350,
            }));
            if (level > 1)
                pPr.AppendChild(new Indentation { Left = ((level - 1) * 220).ToString() });
            style.AppendChild(pPr);
            styles.AppendChild(style);
        }
    }
}
