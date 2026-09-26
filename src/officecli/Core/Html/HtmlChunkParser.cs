// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeCli.Core.Html;

/// <summary>
/// Neutral block model produced by <see cref="HtmlChunkParser"/>. The Word
/// handler turns this into OOXML; the parser itself does not reference the
/// Open XML SDK.
/// </summary>
internal abstract class HtmlFlowBlock;

internal sealed class HtmlFlowParagraph : HtmlFlowBlock
{
    public List<HtmlFlowRun> Runs { get; } = new();
    public string? StyleId { get; set; }
    public string? Align { get; set; }
    public string? Fill { get; set; }
    public int IndentLeftTwips { get; set; }
    public int? FirstLineTwips { get; set; }
    public int? SpaceBeforeTwips { get; set; }
    public int? SpaceAfterTwips { get; set; }
    /// <summary><c>bullet</c> or <c>ordered</c>. Null for a normal paragraph.</summary>
    public string? ListKind { get; set; }
    public int ListLevel { get; set; }
    /// <summary>Numbering start. Set only on the first item of a list so the
    /// following items continue the same definition.</summary>
    public int? ListStart { get; set; }
    /// <summary>
    /// One nested list tree. The outermost <c>ul</c>/<c>ol</c> and every list
    /// inside its items share an id so materialize can bind them to one Word
    /// <c>numId</c>. A list that is not inside those items gets a new id.
    /// Zero when this paragraph is not a list item.
    /// </summary>
    public int ListTree { get; set; }
    public bool HorizontalRule { get; set; }
    /// <summary>
    /// CSS border on this block, or inherited from a wrapper (div, blockquote)
    /// when this paragraph is the one that draws it. Null when no side was set.
    /// A nil side is explicit <c>none</c>/<c>hidden</c> and blocks inheritance.
    /// </summary>
    public HtmlFlowBorder? Border { get; set; }
}

internal sealed class HtmlFlowRun
{
    public string Text { get; set; } = "";
    public bool Break { get; set; }
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public bool Underline { get; set; }
    public bool Strike { get; set; }
    public string? VertAlign { get; set; }
    public string? Color { get; set; }
    public string? Fill { get; set; }
    public double? SizePt { get; set; }
    public string? Font { get; set; }
    public string? Href { get; set; }
    /// <summary>
    /// Data URI to embed as an inline picture. Null on a text run.
    /// Set only for types <c>ImageSource</c> already accepts, except SVG
    /// (png, jpeg, gif, bmp, tiff, emf, wmf). webp stays text.
    /// </summary>
    public string? ImageSrc { get; set; }
    /// <summary>Picture description from the <c>alt</c> attribute. Empty when absent.</summary>
    public string? ImageAlt { get; set; }
    /// <summary>CSS <c>width</c> or the HTML width attribute, in source units.</summary>
    public string? ImageWidth { get; set; }
    /// <summary>CSS <c>height</c> or the HTML height attribute, in source units.</summary>
    public string? ImageHeight { get; set; }
}

internal sealed class HtmlFlowTable : HtmlFlowBlock
{
    public List<HtmlFlowRow> Rows { get; } = new();
    /// <summary>CSS border on the <c>table</c> element itself (outer edges).</summary>
    public HtmlFlowBorder? Border { get; set; }
}

internal sealed class HtmlFlowRow
{
    public bool Header { get; set; }
    public List<HtmlFlowCell> Cells { get; } = new();
}

internal sealed class HtmlFlowCell
{
    public List<HtmlFlowBlock> Blocks { get; } = new();
    public int ColSpan { get; set; } = 1;
    public int RowSpan { get; set; } = 1;
    public bool Header { get; set; }
    public string? Fill { get; set; }
    /// <summary>CSS border on the cell, filled in from the row when a side is absent.</summary>
    public HtmlFlowBorder? Border { get; set; }
    /// <summary>CSS padding in twips. Emitted as <c>w:tcMar</c>, not as border space.</summary>
    public HtmlFlowPadding? Padding { get; set; }
}

/// <summary>
/// One CSS edge after the border shorthand has been resolved.
/// <see cref="Style"/> is a Word line token (<c>single</c>, <c>dashed</c>,
/// <c>dotted</c>, <c>double</c>, <c>inset</c>, <c>outset</c>, <c>nil</c>).
/// <see cref="SizeEighths"/> is OOXML <c>w:sz</c> (eighths of a point).
/// <see cref="Color"/> is <c>RRGGBB</c>, or null for <c>w:color="auto"</c>.
/// <see cref="SpacePoints"/> is <c>w:space</c> from CSS padding, in points.
/// A null edge means the property was absent. A <c>nil</c> edge was set to none.
/// </summary>
internal readonly record struct HtmlFlowBorderSide(string Style, int SizeEighths, string? Color, int? SpacePoints);

/// <summary>The four physical CSS edges. Missing edges stay null.</summary>
internal sealed class HtmlFlowBorder
{
    public HtmlFlowBorderSide? Top { get; set; }
    public HtmlFlowBorderSide? Right { get; set; }
    public HtmlFlowBorderSide? Bottom { get; set; }
    public HtmlFlowBorderSide? Left { get; set; }
    public bool Any => Top != null || Right != null || Bottom != null || Left != null;
}

/// <summary>CSS padding in twips. Null on an edge means that longhand was absent.</summary>
internal readonly record struct HtmlFlowPadding(int? Top, int? Right, int? Bottom, int? Left)
{
    public bool Any => Top != null || Right != null || Bottom != null || Left != null;
}

internal sealed class HtmlChunkParseResult
{
    public List<HtmlFlowBlock> Blocks { get; } = new();
    public List<string> Warnings { get; } = new();
}

/// <summary>
/// HTML (and plain text) → block model for headless altChunk materialization.
///
/// Supported: paragraphs, headings, bold/italic/underline/strike, color,
/// font size and family, sub/sup, links, lists, simple tables (colspan /
/// rowspan), inline pictures from a <c>data:</c> URI (png, jpeg, gif, bmp,
/// tiff, emf, wmf — the picture pipeline's raster types, not svg or webp),
/// and a small CSS subset (element, class, id, descendant and child
/// combinators; the properties listed above plus text-align,
/// background-color, margin, and borders). Remote and relative images, svg,
/// webp, scripts, and layout CSS (float, flex, grid, media queries) are
/// reported and not embedded. Border coverage is documented on
/// <see cref="HtmlChunkParser"/>'s border helpers: physical shorthands and
/// the Word line styles, not radius, image, outline, or inline borders.
/// </summary>
internal static partial class HtmlChunkParser
{
    public static HtmlChunkParseResult ParsePlainText(string text)
    {
        var result = new HtmlChunkParseResult();
        var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        if (normalized.Length == 0)
        {
            result.Blocks.Add(new HtmlFlowParagraph());
            return result;
        }
        foreach (var line in normalized.Split('\n'))
        {
            var para = new HtmlFlowParagraph();
            if (line.Length > 0)
                para.Runs.Add(new HtmlFlowRun { Text = SanitizeXml(line) });
            result.Blocks.Add(para);
        }
        return result;
    }

    public static HtmlChunkParseResult ParseHtml(string html)
    {
        var result = new HtmlChunkParseResult();
        var stats = new CssStats();
        var root = BuildDom(html, stats);
        var sheet = new Stylesheet();
        int ignoredSelectors = 0;
        foreach (var style in Enumerate(root, n => n.Tag == "style"))
        {
            var css = string.Concat(style.Children.Where(c => c.IsText).Select(c => c.Text));
            ignoredSelectors += sheet.AddCss(css, stats);
        }
        if (ignoredSelectors > 0)
            result.Warnings.Add($"{ignoredSelectors} CSS selector(s) ignored (supported: element, class, id, descendant and child combinators)");
        if (stats.BordersApproximated > 0)
            result.Warnings.Add($"{stats.BordersApproximated} CSS border style(s) approximated (groove as inset, ridge as outset)");
        if (stats.BordersIgnored > 0)
            result.Warnings.Add($"{stats.BordersIgnored} CSS border declaration(s) ignored (unsupported style or width; supported styles: solid, dashed, dotted, double, inset, outset)");

        var conv = new Converter(sheet);
        var body = Enumerate(root, n => n.Tag == "body").FirstOrDefault();
        var start = body ?? root;
        conv.EmitElement(start, FlowFormat.Empty, result.Blocks, ListContext.None, inPre: false);
        if (conv.ImagesRemote > 0)
            result.Warnings.Add($"{conv.ImagesRemote} image(s) kept as alt text (http(s) URL is not downloaded)");
        if (conv.ImagesLocal > 0)
            result.Warnings.Add($"{conv.ImagesLocal} image(s) kept as alt text (relative or local URL is not resolved)");
        if (conv.ImagesUnsupported > 0)
            result.Warnings.Add($"{conv.ImagesUnsupported} image(s) kept as alt text (unsupported type such as svg or webp, or no src)");
        if (conv.ImagesInvalid > 0)
            result.Warnings.Add($"{conv.ImagesInvalid} image(s) kept as alt text (data URI could not be read)");
        if (conv.Objects > 0)
            result.Warnings.Add($"{conv.Objects} embedded object(s) dropped (svg, video, iframe, object)");
        if (conv.Scripts > 0)
            result.Warnings.Add($"{conv.Scripts} script element(s) dropped");
        return result;
    }

    // ---------------------------------------------------------------- DOM

    internal sealed class HtmlNode
    {
        public string? Tag { get; init; }
        public string Text { get; set; } = "";
        public bool IsText => Tag == null;
        public Dictionary<string, string> Attrs { get; } = new(StringComparer.OrdinalIgnoreCase);
        public List<HtmlNode> Children { get; } = new();
        public HtmlNode? Parent { get; set; }
        public HashSet<string> Classes { get; } = new(StringComparer.Ordinal);
        public string? Id { get; set; }
        public Dictionary<string, string> Declarations { get; } = new(StringComparer.OrdinalIgnoreCase);
    }

    static HtmlNode BuildDom(string html, CssStats stats)
    {
        var root = new HtmlNode { Tag = "#root" };
        var stack = new Stack<HtmlNode>();
        stack.Push(root);
        foreach (var tok in Tokenize(html))
        {
            if (tok.Kind == TokKind.Text)
            {
                if (tok.Text.Length == 0) continue;
                var text = new HtmlNode { Text = tok.Text, Parent = stack.Peek() };
                stack.Peek().Children.Add(text);
                continue;
            }
            if (tok.Kind == TokKind.End)
            {
                PopUntil(stack, tok.Name);
                continue;
            }
            PrepareFor(stack, tok.Name);
            var node = new HtmlNode { Tag = tok.Name, Parent = stack.Peek() };
            foreach (var kv in tok.Attrs) node.Attrs[kv.Key] = kv.Value;
            if (node.Attrs.TryGetValue("class", out var cls))
            {
                foreach (var part in cls.Split((char[])[' ', '\t', '\n', '\r'], StringSplitOptions.RemoveEmptyEntries))
                    node.Classes.Add(part);
            }
            if (node.Attrs.TryGetValue("id", out var id) && id.Length > 0)
                node.Id = id;
            AbsorbPresentationalAttrs(node);
            if (node.Attrs.TryGetValue("style", out var style))
                ParseDeclarations(style, node.Declarations, stats);
            stack.Peek().Children.Add(node);
            if (!tok.SelfClosing && !IsVoid(tok.Name))
                stack.Push(node);
        }
        return root;
    }

    static void AbsorbPresentationalAttrs(HtmlNode node)
    {
        var style = node.Attrs.TryGetValue("style", out var existing) ? existing : "";
        void Append(string decl)
        {
            if (style.Length > 0 && !style.TrimEnd().EndsWith(';')) style += ";";
            style += decl;
        }
        if (node.Attrs.TryGetValue("align", out var align)
            && align.Length > 0
            && style.IndexOf("text-align", StringComparison.OrdinalIgnoreCase) < 0)
            Append("text-align:" + align);
        if (node.Attrs.TryGetValue("bgcolor", out var bg) && bg.Length > 0
            && style.IndexOf("background", StringComparison.OrdinalIgnoreCase) < 0)
            Append("background-color:" + bg);
        if (node.Tag == "font")
        {
            if (node.Attrs.TryGetValue("color", out var color) && color.Length > 0)
                Append("color:" + color);
            if (node.Attrs.TryGetValue("face", out var face) && face.Length > 0)
                Append("font-family:" + face);
            if (node.Attrs.TryGetValue("size", out var size)
                && int.TryParse(size, out var n))
            {
                var pt = n switch { 1 => 8, 2 => 10, 3 => 12, 4 => 14, 5 => 18, 6 => 24, >= 7 => 36, _ => 11 };
                Append("font-size:" + pt + "pt");
            }
        }
        if (node.Tag == "center" && style.IndexOf("text-align", StringComparison.OrdinalIgnoreCase) < 0)
            Append("text-align:center");
        if (style.Length > 0) node.Attrs["style"] = style;
    }

    static void PrepareFor(Stack<HtmlNode> stack, string tag)
    {
        if (ClosesP.Contains(tag))
            PopWhile(stack, t => t == "p");
        switch (tag)
        {
            case "li":
                PopWhile(stack, t => t is "p" or "li");
                break;
            case "td" or "th":
                PopWhile(stack, t => t is "p" or "td" or "th");
                break;
            case "tr":
                PopWhile(stack, t => t is "p" or "td" or "th" or "tr");
                break;
            case "thead" or "tbody" or "tfoot":
                PopWhile(stack, t => t is "p" or "td" or "th" or "tr" or "thead" or "tbody" or "tfoot");
                break;
        }
    }

    static void PopWhile(Stack<HtmlNode> stack, Func<string, bool> pred)
    {
        while (stack.Count > 1 && stack.Peek().Tag is { } t && pred(t))
            stack.Pop();
    }

    static void PopUntil(Stack<HtmlNode> stack, string tag)
    {
        while (stack.Count > 1 && !string.Equals(stack.Peek().Tag, tag, StringComparison.Ordinal))
            stack.Pop();
        if (stack.Count > 1) stack.Pop();
    }

    static readonly HashSet<string> ClosesP = new(StringComparer.Ordinal)
    {
        "address", "article", "aside", "blockquote", "div", "dl", "fieldset",
        "figcaption", "figure", "footer", "form", "h1", "h2", "h3", "h4", "h5",
        "h6", "header", "hr", "main", "nav", "ol", "p", "pre", "section",
        "table", "ul",
    };

    static bool IsVoid(string tag) => tag is "area" or "base" or "br" or "col" or "embed"
        or "hr" or "img" or "input" or "link" or "meta" or "source" or "track" or "wbr";

    static IEnumerable<HtmlNode> Enumerate(HtmlNode node, Func<HtmlNode, bool> pred)
    {
        if (pred(node)) yield return node;
        foreach (var child in node.Children)
        {
            foreach (var hit in Enumerate(child, pred))
                yield return hit;
        }
    }

    // ---------------------------------------------------------------- tokenize

    enum TokKind { Text, Start, End }

    readonly struct Tok
    {
        public TokKind Kind { get; init; }
        public string Name { get; init; }
        public string Text { get; init; }
        public bool SelfClosing { get; init; }
        public Dictionary<string, string> Attrs { get; init; }
    }

    static IEnumerable<Tok> Tokenize(string html)
    {
        int i = 0;
        int n = html.Length;
        while (i < n)
        {
            if (html[i] != '<')
            {
                int start = i;
                while (i < n && html[i] != '<') i++;
                yield return new Tok { Kind = TokKind.Text, Name = "", Text = WebUtility.HtmlDecode(html[start..i]), Attrs = EmptyAttrs };
                continue;
            }
            if (i + 3 < n && html[i + 1] == '!' && html[i + 2] == '-' && html[i + 3] == '-')
            {
                i += 4;
                int commentEnd = html.IndexOf("-->", i, StringComparison.Ordinal);
                i = commentEnd < 0 ? n : commentEnd + 3;
                continue;
            }
            if (i + 1 < n && (html[i + 1] == '!' || html[i + 1] == '?'))
            {
                i += 2;
                while (i < n && html[i] != '>') i++;
                if (i < n) i++;
                continue;
            }
            bool isEnd = i + 1 < n && html[i + 1] == '/';
            int j = i + (isEnd ? 2 : 1);
            while (j < n && char.IsWhiteSpace(html[j])) j++;
            int nameStart = j;
            while (j < n && (char.IsLetterOrDigit(html[j]) || html[j] is '-' or ':')) j++;
            if (j == nameStart)
            {
                yield return new Tok { Kind = TokKind.Text, Name = "", Text = "<", Attrs = EmptyAttrs };
                i++;
                continue;
            }
            var name = html[nameStart..j].ToLowerInvariant();
            if (name.Contains(':')) name = name[(name.LastIndexOf(':') + 1)..];
            var attrs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            bool self = false;
            while (j < n && html[j] != '>')
            {
                while (j < n && char.IsWhiteSpace(html[j])) j++;
                if (j >= n) break;
                if (html[j] == '>') break;
                if (html[j] == '/' && j + 1 < n && html[j + 1] == '>') { self = true; j++; break; }
                int an = j;
                while (j < n && !char.IsWhiteSpace(html[j]) && html[j] is not '=' and not '>' and not '/') j++;
                var aname = html[an..j];
                string aval = "";
                while (j < n && char.IsWhiteSpace(html[j])) j++;
                if (j < n && html[j] == '=')
                {
                    j++;
                    while (j < n && char.IsWhiteSpace(html[j])) j++;
                    if (j < n && html[j] is '"' or '\'')
                    {
                        char q = html[j++];
                        int vs = j;
                        while (j < n && html[j] != q) j++;
                        aval = html[vs..j];
                        if (j < n) j++;
                    }
                    else
                    {
                        int vs = j;
                        while (j < n && !char.IsWhiteSpace(html[j]) && html[j] != '>') j++;
                        aval = html[vs..j];
                    }
                    aval = WebUtility.HtmlDecode(aval);
                }
                if (aname.Length > 0) attrs[aname] = aval;
            }
            if (j < n && html[j] == '>') j++;
            i = j;
            if (isEnd)
            {
                yield return new Tok { Kind = TokKind.End, Name = name, Text = "", Attrs = EmptyAttrs };
                continue;
            }
            self = self || IsVoid(name);
            yield return new Tok { Kind = TokKind.Start, Name = name, Text = "", SelfClosing = self, Attrs = attrs };
            if (!self && name is "script" or "style" or "title" or "textarea")
            {
                var close = "</" + name;
                int endIdx = IndexOfIgnoreCase(html, close, i);
                var raw = endIdx < 0 ? html[i..] : html[i..endIdx];
                if (raw.Length > 0)
                    yield return new Tok { Kind = TokKind.Text, Name = "", Text = raw, Attrs = EmptyAttrs };
                i = endIdx < 0 ? n : endIdx;
            }
        }
    }

    static int IndexOfIgnoreCase(string hay, string needle, int start)
    {
        return hay.IndexOf(needle, start, StringComparison.OrdinalIgnoreCase);
    }

    static readonly Dictionary<string, string> EmptyAttrs = new(StringComparer.OrdinalIgnoreCase);

    // ---------------------------------------------------------------- CSS

    sealed class Stylesheet
    {
        public List<CssRule> Rules { get; } = new();

        public int AddCss(string css, CssStats? stats = null)
        {
            css = StripComments(css);
            int ignored = 0;
            int i = 0;
            while (i < css.Length)
            {
                while (i < css.Length && char.IsWhiteSpace(css[i])) i++;
                if (i >= css.Length) break;
                if (css[i] == '@')
                {
                    i = SkipAtRule(css, i);
                    continue;
                }
                int brace = css.IndexOf('{', i);
                if (brace < 0) break;
                var selectorText = css[i..brace];
                int depth = 1;
                int k = brace + 1;
                while (k < css.Length && depth > 0)
                {
                    if (css[k] == '{') depth++;
                    else if (css[k] == '}') depth--;
                    k++;
                }
                var body = css[(brace + 1)..(k - 1)];
                var decls = new Dictionary<string, (string Value, bool Important)>(StringComparer.OrdinalIgnoreCase);
                ParseDeclarations(body, decls, stats);
                foreach (var sel in SplitSelectors(selectorText))
                {
                    if (!TryParseSelector(sel, out var compounds, out var childJoins))
                    {
                        ignored++;
                        continue;
                    }
                    Rules.Add(new CssRule
                    {
                        Compounds = compounds,
                        ChildJoins = childJoins,
                        Decls = decls,
                        Order = Rules.Count,
                    });
                }
                i = k;
            }
            return ignored;
        }
    }

    sealed class CssRule
    {
        public List<CssCompound> Compounds { get; init; } = new();
        public List<bool> ChildJoins { get; init; } = new();
        public Dictionary<string, (string Value, bool Important)> Decls { get; init; } = new(StringComparer.OrdinalIgnoreCase);
        public int Order { get; init; }
        public int Specificity => Compounds.Sum(c => c.Specificity);
    }

    readonly struct CssCompound
    {
        public string? Tag { get; init; }
        public string? Id { get; init; }
        public List<string> Classes { get; init; }
        public int Specificity =>
            (Id != null ? 100 : 0)
            + Classes.Count * 10
            + (Tag != null && Tag != "*" ? 1 : 0);
    }

    static string StripComments(string css)
    {
        var sb = new StringBuilder(css.Length);
        for (int i = 0; i < css.Length; i++)
        {
            if (i + 1 < css.Length && css[i] == '/' && css[i + 1] == '*')
            {
                int end = css.IndexOf("*/", i + 2, StringComparison.Ordinal);
                i = end < 0 ? css.Length : end + 1;
                continue;
            }
            sb.Append(css[i]);
        }
        return sb.ToString();
    }

    static int SkipAtRule(string css, int i)
    {
        int semi = css.IndexOf(';', i);
        int brace = css.IndexOf('{', i);
        if (brace < 0 || (semi >= 0 && semi < brace))
            return semi < 0 ? css.Length : semi + 1;
        int depth = 0;
        for (int k = brace; k < css.Length; k++)
        {
            if (css[k] == '{') depth++;
            else if (css[k] == '}')
            {
                depth--;
                if (depth == 0) return k + 1;
            }
        }
        return css.Length;
    }

    static IEnumerable<string> SplitSelectors(string text)
    {
        var sb = new StringBuilder();
        foreach (var ch in text)
        {
            if (ch == ',')
            {
                var s = sb.ToString().Trim();
                if (s.Length > 0) yield return s;
                sb.Clear();
            }
            else sb.Append(ch);
        }
        var last = sb.ToString().Trim();
        if (last.Length > 0) yield return last;
    }

    static bool TryParseSelector(string text, out List<CssCompound> compounds, out List<bool> childJoins)
    {
        compounds = new List<CssCompound>();
        childJoins = new List<bool>();
        if (text.IndexOfAny(['[', ':', '+', '~']) >= 0) return false;
        int i = 0;
        bool? pendingChild = null;
        while (i < text.Length)
        {
            while (i < text.Length && char.IsWhiteSpace(text[i])) i++;
            if (i >= text.Length) break;
            if (text[i] == '>')
            {
                if (compounds.Count == 0 || pendingChild == true) return false;
                pendingChild = true;
                i++;
                continue;
            }
            if (!TryReadCompound(text, ref i, out var compound)) return false;
            if (compounds.Count > 0)
                childJoins.Add(pendingChild == true);
            pendingChild = false;
            compounds.Add(compound);
        }
        return compounds.Count > 0 && pendingChild != true;
    }

    static bool TryReadCompound(string text, ref int i, out CssCompound compound)
    {
        string? tag = null;
        string? id = null;
        var classes = new List<string>();
        if (i < text.Length && (char.IsLetter(text[i]) || text[i] == '*'))
        {
            int s = i;
            if (text[i] == '*') i++;
            else while (i < text.Length && (char.IsLetterOrDigit(text[i]) || text[i] == '-')) i++;
            tag = text[s..i].ToLowerInvariant();
        }
        while (i < text.Length && text[i] is '.' or '#')
        {
            char sigil = text[i++];
            int s = i;
            while (i < text.Length && (char.IsLetterOrDigit(text[i]) || text[i] is '-' or '_')) i++;
            if (i == s) { compound = default; return false; }
            var ident = text[s..i];
            if (sigil == '#') id = ident;
            else classes.Add(ident);
        }
        if (tag == null && id == null && classes.Count == 0)
        {
            compound = default;
            return false;
        }
        compound = new CssCompound { Tag = tag, Id = id, Classes = classes };
        return true;
    }

    static void ParseDeclarations(string body, Dictionary<string, string> dest, CssStats? stats = null)
    {
        var rich = new Dictionary<string, (string Value, bool Important)>(StringComparer.OrdinalIgnoreCase);
        ParseDeclarations(body, rich, stats);
        foreach (var kv in rich) dest[kv.Key] = kv.Value.Value;
    }

    static void ParseDeclarations(string body, Dictionary<string, (string Value, bool Important)> dest, CssStats? stats = null)
    {
        foreach (var piece in body.Split(';'))
        {
            int colon = piece.IndexOf(':');
            if (colon <= 0) continue;
            var prop = piece[..colon].Trim().ToLowerInvariant();
            var val = piece[(colon + 1)..].Trim();
            if (prop.Length == 0 || val.Length == 0) continue;
            bool important = false;
            int bang = val.IndexOf("!important", StringComparison.OrdinalIgnoreCase);
            if (bang >= 0)
            {
                important = true;
                val = val[..bang].Trim();
            }
            if (val.Length == 0) continue;
            if (TryExpandBorderOrPadding(prop, val, important, dest, stats))
                continue;
            if (IsBorderStyleLonghand(prop))
            {
                var style = val.Trim().ToLowerInvariant();
                NoteBorderStyle(style, stats);
                dest[prop] = (style, important);
                continue;
            }
            if (IsBorderWidthLonghand(prop) && !IsBorderWidth(val))
            {
                if (stats != null) stats.BordersIgnored++;
            }
            else if (IsBorderColorLonghand(prop) && !IsBorderColor(val))
            {
                if (stats != null) stats.BordersIgnored++;
            }
            dest[prop] = (val, important);
        }
    }

    /// <summary>
    /// Where a list is being emitted. <see cref="InItem"/> means the list sits
    /// inside an <c>li</c> (directly, or under a wrapper such as <c>div</c>).
    /// <see cref="Tree"/> is the ancestor list's id; zero means there is none.
    /// </summary>
    readonly record struct ListContext(int Level, bool InItem, int Tree)
    {
        public static ListContext None => default;
    }

    sealed partial class Converter
    {
        readonly Stylesheet _sheet;
        public int ImagesRemote, ImagesLocal, ImagesUnsupported, ImagesInvalid, Objects, Scripts;
        int _nextListTree;

        public Converter(Stylesheet sheet)
        {
            _sheet = sheet;
        }

        public void EmitElement(HtmlNode node, FlowFormat parent, List<HtmlFlowBlock> blocks, ListContext ctx, bool inPre)
        {
            var tag = node.Tag;
            if (tag is null or "#root") return;
            if (tag is "script" or "noscript") { Scripts++; return; }
            if (tag is "style" or "head" or "title" or "meta" or "link" or "col" or "colgroup") return;
            var fmt = FormatOf(node, parent);
            if (tag is "ul" or "ol")
            {
                ConvertList(node, fmt, blocks, ctx);
                return;
            }
            if (tag == "table")
            {
                // A cell is its own block container. Lists there do not join
                // a list that happens to contain the table.
                EmitCaption(node, fmt, blocks);
                var table = ConvertTable(node, fmt);
                if (table.Rows.Count > 0) blocks.Add(table);
                return;
            }
            if (tag == "hr")
            {
                blocks.Add(new HtmlFlowParagraph { HorizontalRule = true });
                return;
            }
            if (tag is "html" or "body")
            {
                EmitContainer(node, fmt, blocks, ctx, inPre: false);
                return;
            }
            EmitContainer(node, fmt, blocks, ctx, inPre || tag == "pre");
        }

        void EmitCaption(HtmlNode table, FlowFormat fmt, List<HtmlFlowBlock> blocks)
        {
            var cap = table.Children.FirstOrDefault(c => c.Tag == "caption");
            if (cap == null) return;
            EmitContainer(cap, FormatOf(cap, fmt), blocks, ListContext.None, inPre: false);
        }

        void ConvertList(HtmlNode list, FlowFormat parent, List<HtmlFlowBlock> blocks, ListContext ctx)
        {
            var kind = list.Tag == "ol" ? "ordered" : "bullet";
            int start = 1;
            if (list.Attrs.TryGetValue("start", out var raw)
                && int.TryParse(raw, out var n) && n > 0)
                start = n;
            // Inside an <li>, this list continues the ancestor tree one level
            // deeper. Anywhere else (body, div, table cell) it is a new tree.
            int level = Math.Clamp(ctx.InItem ? ctx.Level + 1 : ctx.Level, 0, 8);
            int tree = ctx.InItem && ctx.Tree > 0 ? ctx.Tree : ++_nextListTree;
            var itemCtx = new ListContext(level, true, tree);
            bool first = true;
            foreach (var child in list.Children)
            {
                if (child.Tag != "li") continue;
                var item = new List<HtmlFlowBlock>();
                EmitContainer(child, FormatOf(child, parent), item, itemCtx, inPre: false);
                var marker = item.OfType<HtmlFlowParagraph>().FirstOrDefault(p => p.ListKind == null && !p.HorizontalRule);
                if (marker == null)
                {
                    marker = new HtmlFlowParagraph();
                    item.Insert(0, marker);
                }
                marker.ListKind = kind;
                marker.ListLevel = level;
                marker.ListTree = tree;
                if (first) marker.ListStart = start;
                first = false;
                blocks.AddRange(item);
            }
        }

        HtmlFlowTable ConvertTable(HtmlNode table, FlowFormat parent)
        {
            var flow = new HtmlFlowTable
            {
                // Outer edges only. Cell borders are resolved per td/th.
                Border = ResolveBorder(table, parent.SizePt ?? 11, includeSpace: false),
            };
            foreach (var tr in CollectRows(table))
            {
                var row = new HtmlFlowRow { Header = tr.Parent?.Tag == "thead" };
                foreach (var cellNode in tr.Children)
                {
                    if (cellNode.Tag is not ("td" or "th")) continue;
                    var cellFmt = FormatOf(cellNode, FormatOf(tr, parent));
                    var cell = new HtmlFlowCell
                    {
                        Header = cellNode.Tag == "th" || row.Header,
                        ColSpan = ClampSpan(cellNode, "colspan"),
                        RowSpan = ClampSpan(cellNode, "rowspan"),
                        Fill = OwnFill(cellNode),
                        Border = CellBorder(cellNode, tr, cellFmt.SizePt ?? 11),
                        Padding = CellPadding(cellNode, tr, cellFmt.SizePt ?? 11),
                    };
                    EmitContainer(cellNode, cellFmt, cell.Blocks, ListContext.None, inPre: false);
                    if (cell.Blocks.Count == 0)
                        cell.Blocks.Add(MakeParagraph(cellNode, cellFmt, new List<HtmlFlowRun>()));
                    row.Cells.Add(cell);
                }
                if (row.Cells.Count > 0) flow.Rows.Add(row);
            }
            return flow;
        }

        static int ClampSpan(HtmlNode node, string attr)
        {
            if (!node.Attrs.TryGetValue(attr, out var raw)) return 1;
            return int.TryParse(raw, out var n) && n > 1 ? Math.Min(n, 63) : 1;
        }

        static List<HtmlNode> CollectRows(HtmlNode table)
        {
            var rows = new List<HtmlNode>();
            void Walk(HtmlNode node)
            {
                foreach (var child in node.Children)
                {
                    if (child.Tag == "tr") rows.Add(child);
                    else if (child.Tag is "thead" or "tbody" or "tfoot") Walk(child);
                }
            }
            Walk(table);
            return rows;
        }

        void EmitContainer(HtmlNode node, FlowFormat fmt, List<HtmlFlowBlock> blocks, ListContext ctx, bool inPre)
        {
            var runs = new List<HtmlFlowRun>();
            bool pending = false;
            bool preTrim = !inPre;
            bool sawBlock = false;
            foreach (var child in node.Children)
            {
                if (child.IsText && !inPre && IsOnlyWhitespace(child.Text) && IgnoresLooseWhitespace(node.Tag))
                    continue;
                if (child.IsText || (child.Tag != null && IsInline(child.Tag)))
                {
                    AppendInline(child, fmt, runs, ref pending, inPre, ref preTrim);
                    continue;
                }
                // A block child ends the current phrasing run. Don't invent
                // an empty paragraph just because <blockquote> or <p> is
                // allowed to survive when it is itself empty.
                Flush(node, fmt, runs, blocks, ref pending, inPre, keepEmpty: false);
                sawBlock = true;
                if (child.Tag is "ul" or "ol")
                    ConvertList(child, fmt, blocks, ctx);
                else
                    EmitElement(child, fmt, blocks, ctx, inPre);
            }
            bool emptyElement = !sawBlock && node.Tag is "p" or "h1" or "h2" or "h3" or "h4" or "h5" or "h6" or "pre" or "blockquote";
            Flush(node, fmt, runs, blocks, ref pending, inPre, keepEmpty: emptyElement);
        }

        void Flush(HtmlNode node, FlowFormat fmt, List<HtmlFlowRun> runs, List<HtmlFlowBlock> blocks, ref bool pending, bool inPre, bool keepEmpty)
        {
            pending = false;
            if (!inPre) TrimTrailingSpace(runs);
            if (runs.Count == 0 && !keepEmpty) return;
            blocks.Add(MakeParagraph(node, fmt, runs));
            runs.Clear();
        }

        HtmlFlowParagraph MakeParagraph(HtmlNode node, FlowFormat fmt, List<HtmlFlowRun> runs)
        {
            var para = new HtmlFlowParagraph { Align = fmt.Align };
            para.Runs.AddRange(runs);
            if (node.Tag is "h1" or "h2" or "h3" or "h4" or "h5" or "h6")
                para.StyleId = "Heading" + node.Tag[1];
            // Cell/body backgrounds shade the cell (or are page chrome), not
            // every run. Inline backgrounds (span, mark) stay on the run.
            para.Fill = node.Tag is "td" or "th" or "tr" or "body" or "html" or "#root"
                ? null
                : OwnFill(node);
            int indent = fmt.ExtraIndent;
            if (LengthTwips(Own(node, "margin-left"), fmt.SizePt ?? 11) is int ml) indent += ml;
            para.IndentLeftTwips = indent;
            para.FirstLineTwips = LengthTwips(Own(node, "text-indent"), fmt.SizePt ?? 11);
            para.SpaceBeforeTwips = MarginEdge(node, fmt, top: true);
            para.SpaceAfterTwips = MarginEdge(node, fmt, top: false);
            // A cell's own border is w:tcBorders. Direct text in a td is still
            // flushed as a paragraph of that td, so don't also paint pBdr.
            para.Border = ParagraphBorder(node, fmt.SizePt ?? 11);
            return para;
        }

        int? MarginEdge(HtmlNode node, FlowFormat fmt, bool top)
        {
            var side = top ? "margin-top" : "margin-bottom";
            var direct = LengthTwips(Own(node, side), fmt.SizePt ?? 11);
            if (direct != null) return direct;
            if (Own(node, "margin") is not string shorthand) return null;
            var box = ParseBox(shorthand, fmt.SizePt ?? 11);
            return top ? box.Top : box.Bottom;
        }

        void AppendInline(HtmlNode node, FlowFormat fmt, List<HtmlFlowRun> runs, ref bool pending, bool inPre, ref bool preTrim)
        {
            if (node.IsText)
            {
                AppendText(node.Text, fmt, runs, ref pending, inPre, ref preTrim);
                return;
            }
            var tag = node.Tag ?? "";
            if (tag is "script" or "noscript") { Scripts++; return; }
            if (tag is "style") return;
            if (tag == "br")
            {
                runs.Add(new HtmlFlowRun { Break = true });
                pending = false;
                return;
            }
            if (tag == "wbr") return;
            if (tag == "img")
            {
                var alt = node.Attrs.TryGetValue("alt", out var a) ? a : "";
                var label = alt.Length > 0 ? alt : "image";
                node.Attrs.TryGetValue("src", out var src);
                var kind = ClassifyImageSrc(src ?? "", out var normalized);
                if (kind == ImageSrcKind.Embed)
                {
                    runs.Add(new HtmlFlowRun
                    {
                        ImageSrc = normalized,
                        ImageAlt = SanitizeXml(alt),
                        ImageWidth = ImageDimension(node, "width"),
                        ImageHeight = ImageDimension(node, "height"),
                        Href = string.IsNullOrWhiteSpace(fmt.Href) ? null : fmt.Href,
                    });
                }
                else
                {
                    switch (kind)
                    {
                        case ImageSrcKind.Remote: ImagesRemote++; break;
                        case ImageSrcKind.Local: ImagesLocal++; break;
                        case ImageSrcKind.Invalid: ImagesInvalid++; break;
                        default: ImagesUnsupported++; break;
                    }
                    runs.Add(RunFrom(fmt with { Italic = true }, "[image: " + SanitizeXml(label) + "]"));
                }
                pending = false;
                return;
            }
            if (tag is "svg" or "video" or "iframe" or "object" or "embed" or "canvas" or "audio")
            {
                Objects++;
                return;
            }
            if (!IsInline(tag) && tag is not "a" and not "font" and not "span")
            {
                // A block nested inside phrasing content: flush is the caller's
                // job. Treat the element's text so the words are not lost.
            }
            var child = FormatOf(node, fmt);
            if (tag == "a" && node.Attrs.TryGetValue("href", out var href) && !string.IsNullOrWhiteSpace(href))
                child = child with { Href = href.Trim() };
            foreach (var c in node.Children)
                AppendInline(c, child, runs, ref pending, inPre || tag == "pre", ref preTrim);
        }

        void AppendText(string text, FlowFormat fmt, List<HtmlFlowRun> runs, ref bool pending, bool inPre, ref bool preTrim)
        {
            if (inPre)
            {
                var raw = text.Replace("\r\n", "\n").Replace('\r', '\n');
                if (!preTrim && raw.StartsWith('\n')) raw = raw[1..];
                preTrim = true;
                var lines = raw.Split('\n');
                for (int i = 0; i < lines.Length; i++)
                {
                    if (i > 0) runs.Add(new HtmlFlowRun { Break = true });
                    if (lines[i].Length > 0)
                        runs.Add(RunFrom(fmt, SanitizeXml(lines[i])));
                }
                return;
            }
            var sb = new StringBuilder();
            foreach (var c in text)
            {
                if (c is ' ' or '\t' or '\n' or '\r' or '\f')
                {
                    pending = true;
                    continue;
                }
                if (pending && (sb.Length > 0 || HasContent(runs)))
                    sb.Append(' ');
                pending = false;
                if (c >= ' ' || c == '\t') sb.Append(c);
            }
            if (pending && sb.Length > 0)
            {
                sb.Append(' ');
                pending = false;
            }
            if (sb.Length > 0)
                runs.Add(RunFrom(fmt, SanitizeXml(sb.ToString())));
        }

        FlowFormat FormatOf(HtmlNode node, FlowFormat parent)
        {
            var fmt = parent with { VertAlign = null, Fill = null };
            if (node.Tag == "blockquote")
                fmt = fmt with { ExtraIndent = parent.ExtraIndent + 720 };

            if (Own(node, "font-weight") is string weight) fmt = fmt with { Bold = IsBold(weight) };
            else if (node.Tag is "b" or "strong" or "th" or "h1" or "h2" or "h3" or "h4" or "h5" or "h6")
                fmt = fmt with { Bold = true };

            if (Own(node, "font-style") is string fs) fmt = fmt with { Italic = fs is "italic" or "oblique" };
            else if (node.Tag is "i" or "em" or "cite" or "dfn" or "var") fmt = fmt with { Italic = true };

            if (Own(node, "text-decoration") is string dec)
            {
                bool none = dec.Contains("none", StringComparison.OrdinalIgnoreCase);
                fmt = fmt with
                {
                    Underline = dec.Contains("underline", StringComparison.OrdinalIgnoreCase) || (!none && fmt.Underline),
                    Strike = dec.Contains("line-through", StringComparison.OrdinalIgnoreCase) || (!none && fmt.Strike),
                };
                if (none && !dec.Contains("underline", StringComparison.OrdinalIgnoreCase))
                    fmt = fmt with { Underline = false };
                if (none && !dec.Contains("line-through", StringComparison.OrdinalIgnoreCase))
                    fmt = fmt with { Strike = false };
            }
            else if (node.Tag is "u" or "a" or "ins") fmt = fmt with { Underline = true };
            else if (node.Tag is "s" or "strike" or "del") fmt = fmt with { Strike = true };

            if (Own(node, "color") is string color && ParseColor(color) is string hex)
                fmt = fmt with { Color = hex };
            else if (node.Tag == "a" && Own(node, "color") == null)
                fmt = fmt with { Color = "0563C1" };

            if (Own(node, "font-size") is string size)
                fmt = fmt with { SizePt = ComputeSize(size, parent.SizePt ?? 11) };
            else if (DefaultSize(node.Tag) is double d)
                fmt = fmt with { SizePt = d };

            if (Own(node, "font-family") is string family && ParseFont(family) is string font)
                fmt = fmt with { Font = font };
            else if (node.Tag is "code" or "pre" or "kbd" or "samp" or "tt")
                fmt = fmt with { Font = "Consolas" };

            if (Own(node, "text-align") is string align && MapAlign(align) is string mapped)
                fmt = fmt with { Align = mapped };

            // Block backgrounds become paragraph/cell shading. Only phrasing
            // elements paint a run shading (otherwise a <td> fill would shade
            // every run on top of the cell).
            if (node.Tag != null && IsInline(node.Tag))
            {
                if (OwnFill(node) is string fill)
                    fmt = fmt with { Fill = fill };
                else if (node.Tag == "mark")
                    fmt = fmt with { Fill = "FFFF00" };
            }

            if (node.Tag == "sub") fmt = fmt with { VertAlign = "subscript" };
            else if (node.Tag == "sup") fmt = fmt with { VertAlign = "superscript" };
            return fmt;
        }

        string? ImageDimension(HtmlNode node, string name)
        {
            // CSS wins over the presentational attribute, same as a browser.
            if (Own(node, name) is string css && css.Length > 0) return css;
            if (node.Attrs.TryGetValue(name, out var raw) && raw.Length > 0) return raw;
            return null;
        }

        string? Own(HtmlNode node, string prop)
        {
            if (node.Declarations.TryGetValue(prop, out var inline)) return inline;
            CssRule? best = null;
            (string Value, bool Important) chosen = default;
            foreach (var rule in _sheet.Rules)
            {
                if (!rule.Decls.TryGetValue(prop, out var decl)) continue;
                if (!Matches(node, rule)) continue;
                if (best == null
                    || decl.Important && !chosen.Important
                    || decl.Important == chosen.Important && rule.Specificity > best.Specificity
                    || decl.Important == chosen.Important && rule.Specificity == best.Specificity && rule.Order >= best.Order)
                {
                    best = rule;
                    chosen = decl;
                }
            }
            return best == null ? null : chosen.Value;
        }

        string? OwnFill(HtmlNode node)
        {
            var raw = Own(node, "background-color") ?? Own(node, "background");
            return raw == null ? null : ParseFill(raw);
        }

        static bool Matches(HtmlNode node, CssRule rule)
        {
            var compounds = rule.Compounds;
            if (!MatchCompound(node, compounds[^1])) return false;
            var cur = node;
            for (int i = compounds.Count - 2; i >= 0; i--)
            {
                bool child = rule.ChildJoins.Count > i && rule.ChildJoins[i];
                cur = cur.Parent;
                if (cur == null) return false;
                if (child)
                {
                    if (!MatchCompound(cur, compounds[i])) return false;
                }
                else
                {
                    while (cur != null && !MatchCompound(cur, compounds[i]))
                        cur = cur.Parent;
                    if (cur == null) return false;
                }
            }
            return true;
        }

        static bool MatchCompound(HtmlNode node, CssCompound compound)
        {
            if (node.Tag == null || node.Tag == "#root") return false;
            if (compound.Tag != null && compound.Tag != "*"
                && !string.Equals(node.Tag, compound.Tag, StringComparison.Ordinal))
                return false;
            if (compound.Id != null && !string.Equals(node.Id, compound.Id, StringComparison.Ordinal))
                return false;
            foreach (var cls in compound.Classes)
                if (!node.Classes.Contains(cls)) return false;
            return true;
        }
    }

    readonly record struct FlowFormat(
        bool Bold, bool Italic, bool Underline, bool Strike,
        string? VertAlign, string? Color, string? Fill,
        double? SizePt, string? Font, string? Align, string? Href,
        int ExtraIndent)
    {
        public static FlowFormat Empty => default;
    }

    static HtmlFlowRun RunFrom(FlowFormat fmt, string text) => new()
    {
        Text = text,
        Bold = fmt.Bold,
        Italic = fmt.Italic,
        Underline = fmt.Underline,
        Strike = fmt.Strike,
        VertAlign = fmt.VertAlign,
        Color = fmt.Color,
        Fill = fmt.Fill,
        SizePt = fmt.SizePt,
        Font = fmt.Font,
        Href = fmt.Href,
    };

    static bool IsInline(string tag) => tag is "a" or "abbr" or "b" or "bdo" or "big" or "br"
        or "cite" or "code" or "del" or "dfn" or "em" or "font" or "i" or "img" or "ins"
        or "kbd" or "mark" or "q" or "s" or "samp" or "small" or "span" or "strike"
        or "strong" or "sub" or "sup" or "tt" or "u" or "var" or "wbr";

    static bool IgnoresLooseWhitespace(string? tag) => tag is "html" or "body" or "#root"
        or "div" or "li" or "td" or "th" or "blockquote" or "section" or "article"
        or "ul" or "ol" or "table" or "tr" or "thead" or "tbody" or "tfoot"
        or "header" or "footer" or "main" or "figure" or "center" or "h1"
        or "h2" or "h3" or "h4" or "h5" or "h6";

    static bool IsOnlyWhitespace(string text)
    {
        foreach (var c in text)
            if (c is not (' ' or '\t' or '\n' or '\r' or '\f')) return false;
        return true;
    }

        static bool HasContent(List<HtmlFlowRun> runs)
        {
            foreach (var r in runs)
                if (r.ImageSrc != null || r.Text.Length > 0) return true;
            return false;
        }

        static void TrimTrailingSpace(List<HtmlFlowRun> runs)
        {
            for (int i = runs.Count - 1; i >= 0; i--)
            {
                // A picture run has no text. It is still content, and it must
                // not be dropped the way a collapsed whitespace run is.
                if (runs[i].ImageSrc != null || runs[i].Break) return;
                if (runs[i].Text.Length == 0) { runs.RemoveAt(i); continue; }
                runs[i].Text = runs[i].Text.TrimEnd(' ');
                if (runs[i].Text.Length == 0) runs.RemoveAt(i);
                return;
            }
        }

    static bool IsBold(string weight)
    {
        weight = weight.Trim().ToLowerInvariant();
        if (weight is "bold" or "bolder") return true;
        return int.TryParse(weight, out var n) && n >= 600;
    }

    static double? DefaultSize(string? tag) => tag switch
    {
        "h1" => 18,
        "h2" => 16,
        "h3" => 14,
        "h4" => 13,
        "h5" => 12,
        "h6" => 11,
        "small" => 9,
        "big" => 14,
        _ => null,
    };

    static string? MapAlign(string value) => value.Trim().ToLowerInvariant() switch
    {
        "left" or "start" => "left",
        "center" => "center",
        "right" or "end" => "right",
        "justify" => "both",
        _ => null,
    };

    static double ComputeSize(string value, double parentPt)
    {
        value = value.Trim().ToLowerInvariant();
        if (value.EndsWith("pt", StringComparison.Ordinal) && double.TryParse(value[..^2], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var pt))
            return pt;
        if (value.EndsWith("px", StringComparison.Ordinal) && double.TryParse(value[..^2], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var px))
            return px * 0.75;
        if (value.EndsWith("em", StringComparison.Ordinal) && double.TryParse(value[..^2], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var em))
            return em * parentPt;
        if (value.EndsWith('%') && double.TryParse(value[..^1], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var pct))
            return pct / 100.0 * parentPt;
        return value switch
        {
            "small" or "smaller" => parentPt * 0.85,
            "large" or "larger" => parentPt * 1.2,
            "x-small" => 9,
            "medium" => 11,
            "x-large" => 18,
            _ => parentPt,
        };
    }

    static int? LengthTwips(string? value, double basePt)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        value = value.Trim().ToLowerInvariant();
        if (value is "0" or "0px" or "0pt" or "0em") return 0;
        if (value is "auto" or "inherit" or "initial") return null;
        double pt;
        if (value.EndsWith("pt", StringComparison.Ordinal) && double.TryParse(value[..^2], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var p))
            pt = p;
        else if (value.EndsWith("px", StringComparison.Ordinal) && double.TryParse(value[..^2], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var px))
            pt = px * 0.75;
        else if (value.EndsWith("em", StringComparison.Ordinal) && double.TryParse(value[..^2], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var em))
            pt = em * basePt;
        else if (value.EndsWith("in", StringComparison.Ordinal) && double.TryParse(value[..^2], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var inches))
            pt = inches * 72;
        else if (value.EndsWith("cm", StringComparison.Ordinal) && double.TryParse(value[..^2], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var cm))
            pt = cm / 2.54 * 72;
        else if (value.EndsWith("mm", StringComparison.Ordinal) && value.Length > 2
            && double.TryParse(value[..^2], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var mm))
            pt = mm / 25.4 * 72;
        else return null;
        return (int)Math.Round(pt * 20, MidpointRounding.AwayFromZero);
    }

    readonly struct Box(int? top, int? bottom)
    {
        public int? Top { get; } = top;
        public int? Bottom { get; } = bottom;
    }

    static Box ParseBox(string shorthand, double basePt)
    {
        var parts = shorthand.Split((char[])[' ', '\t', '\n', '\r'], StringSplitOptions.RemoveEmptyEntries);
        int? L(string s) => LengthTwips(s, basePt);
        return parts.Length switch
        {
            1 => new Box(L(parts[0]), L(parts[0])),
            2 => new Box(L(parts[0]), L(parts[0])),
            3 => new Box(L(parts[0]), L(parts[2])),
            >= 4 => new Box(L(parts[0]), L(parts[2])),
            _ => new Box(null, null),
        };
    }

    static string? ParseFont(string value)
    {
        foreach (var raw in SplitFontList(value))
        {
            var name = raw.Trim().Trim('"', '\'');
            if (name.Length == 0) continue;
            switch (name.ToLowerInvariant())
            {
                case "monospace": return "Consolas";
                case "serif": return "Times New Roman";
                case "sans-serif": return "Calibri";
                case "cursive": return "Segoe Script";
                case "fantasy": return "Impact";
                default: return name;
            }
        }
        return null;
    }

    static IEnumerable<string> SplitFontList(string value)
    {
        var sb = new StringBuilder();
        char? quote = null;
        foreach (var c in value)
        {
            if (quote != null)
            {
                if (c == quote) quote = null;
                else sb.Append(c);
                continue;
            }
            if (c is '"' or '\'') { quote = c; continue; }
            if (c == ',')
            {
                yield return sb.ToString();
                sb.Clear();
                continue;
            }
            sb.Append(c);
        }
        if (sb.Length > 0) yield return sb.ToString();
    }

    static string? ParseFill(string value)
    {
        foreach (var token in value.Split((char[])[' ', '\t'], StringSplitOptions.RemoveEmptyEntries))
        {
            if (ParseColor(token) is string hex) return hex;
        }
        return null;
    }

    static string? ParseColor(string value)
    {
        value = value.Trim().ToLowerInvariant();
        if (value is "transparent" or "inherit" or "currentcolor" or "auto") return null;
        if (value.StartsWith('#'))
        {
            var h = value[1..];
            if (h.Length == 3)
                h = string.Concat(h[0], h[0], h[1], h[1], h[2], h[2]);
            if (h.Length >= 6 && IsHex(h[..6])) return h[..6].ToUpperInvariant();
            return null;
        }
        var rgb = Regex.Match(value, @"^rgba?\(\s*(\d{1,3})\s*[, ]\s*(\d{1,3})\s*[, ]\s*(\d{1,3})");
        if (rgb.Success)
        {
            int R = ClampByte(rgb.Groups[1].Value);
            int G = ClampByte(rgb.Groups[2].Value);
            int B = ClampByte(rgb.Groups[3].Value);
            return $"{R:X2}{G:X2}{B:X2}";
        }
        return NamedColors.TryGetValue(value, out var named) ? named : null;
    }

    static int ClampByte(string s) => Math.Clamp(int.Parse(s, System.Globalization.CultureInfo.InvariantCulture), 0, 255);

    static bool IsHex(string s)
    {
        foreach (var c in s)
            if (!Uri.IsHexDigit(c)) return false;
        return true;
    }

    static readonly Dictionary<string, string> NamedColors = new(StringComparer.Ordinal)
    {
        ["black"] = "000000", ["white"] = "FFFFFF", ["red"] = "FF0000", ["green"] = "008000",
        ["blue"] = "0000FF", ["yellow"] = "FFFF00", ["gray"] = "808080", ["grey"] = "808080",
        ["navy"] = "000080", ["orange"] = "FFA500", ["purple"] = "800080", ["silver"] = "C0C0C0",
        ["maroon"] = "800000", ["teal"] = "008080", ["aqua"] = "00FFFF", ["fuchsia"] = "FF00FF",
        ["lime"] = "00FF00", ["olive"] = "808000",
    };

    enum ImageSrcKind { Embed, Remote, Local, Unsupported, Invalid }

    /// <summary>
    /// Decide whether an <c>img</c> src can be handed to the existing picture
    /// pipeline. Only a base64 <c>data:</c> URI of a type
    /// <c>ImageSource</c> already embeds is <see cref="ImageSrcKind.Embed"/>.
    /// SVG is refused here even though the picture pipeline can wrap it.
    /// http(s) and every other URL are not fetched and not read from disk.
    /// </summary>
    static ImageSrcKind ClassifyImageSrc(string src, out string normalized)
    {
        normalized = src.Trim();
        if (normalized.Length == 0) return ImageSrcKind.Unsupported;
        if (normalized.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
        {
            normalized = NormalizeDataUri(normalized);
            return ClassifyDataUri(normalized);
        }
        if (normalized.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
            || normalized.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
            || normalized.StartsWith("//", StringComparison.Ordinal))
            return ImageSrcKind.Remote;
        return ImageSrcKind.Local;
    }

    static ImageSrcKind ClassifyDataUri(string src)
    {
        int comma = src.IndexOf(',');
        if (comma < 5) return ImageSrcKind.Invalid;
        var header = src[..comma];
        if (header.IndexOf("base64", StringComparison.OrdinalIgnoreCase) < 0)
            return ImageSrcKind.Invalid;
        int mimeStart = header.IndexOf(':') + 1;
        int mimeEnd = header.IndexOf(';');
        if (mimeEnd < mimeStart) mimeEnd = header.Length;
        var mime = header[mimeStart..mimeEnd].Trim().ToLowerInvariant();
        // Keep this list aligned with ImageSource.MimeToContentType, minus SVG.
        if (mime is "image/png"
            or "image/jpeg" or "image/jpg"
            or "image/gif"
            or "image/bmp"
            or "image/tiff" or "image/tif"
            or "image/emf" or "image/x-emf"
            or "image/wmf" or "image/x-wmf")
            return ImageSrcKind.Embed;
        return ImageSrcKind.Unsupported;
    }

    /// <summary>
    /// Drop whitespace inside the base64 payload. HTML pretty-printers wrap
    /// long data URIs; <c>Convert.FromBase64String</c> rejects that whitespace.
    /// </summary>
    static string NormalizeDataUri(string src)
    {
        int comma = src.IndexOf(',');
        if (comma < 0) return src;
        var payload = src[(comma + 1)..];
        bool dirty = false;
        foreach (var c in payload)
        {
            if (c is ' ' or '\t' or '\n' or '\r') { dirty = true; break; }
        }
        if (!dirty) return src;
        var sb = new StringBuilder(src.Length);
        sb.Append(src[..(comma + 1)]);
        foreach (var c in payload)
            if (c is not (' ' or '\t' or '\n' or '\r')) sb.Append(c);
        return sb.ToString();
    }

    internal static string SanitizeXml(string text)
    {
        var sb = new StringBuilder(text.Length);
        foreach (var c in text)
        {
            if (c == '\t' || c == '\n' || c == '\r' || c >= ' ')
                sb.Append(c);
        }
        return sb.ToString();
    }
}
