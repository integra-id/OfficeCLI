// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core.Html;

internal static partial class HtmlChunkParser
{
    /// <summary>
    /// Counts border declarations the applicator cannot draw as specified.
    /// Warnings are advisory: they do not fail <c>materialize --strict</c>.
    /// </summary>
    sealed class CssStats
    {
        public int BordersApproximated;
        public int BordersIgnored;
    }

    // CSS borders honored by materialize
    // ----------------------------------
    // Properties: border, border-top|right|bottom|left, border-width,
    // border-style, border-color, and the physical longhands
    // (border-top-width, border-left-style, border-bottom-color, …).
    // Padding on a bordered paragraph becomes w:space. Cell padding becomes
    // w:tcMar. A wrapper (div, blockquote, …) copies its border onto each
    // paragraph inside it; it is not one rectangle around the group.
    // Table-cell borders stay on the cell. A row border fills sides the cell
    // did not set. A uniform cell grid also updates w:tblBorders.
    //
    // Styles kept: solid→single, dashed, dotted, double, inset, outset,
    // none/hidden→nil. Approximated (warning): groove→inset, ridge→outset.
    // Ignored (warning): any other line style, and a width that is not a
    // length or thin/medium/thick. transparent color draws no line.
    // currentcolor and inherit use w:color="auto" (the paragraph's text).
    // Alpha in rgba() is dropped; the RGB channels are kept.
    //
    // Not applied: border-radius, border-image, border-spacing, outline,
    // box-shadow, logical properties (border-inline-*, border-block-*),
    // and borders on inline elements (span, code, a). Those do not become
    // run borders. border-collapse is not a separate model — cell borders
    // are written directly. The HTML border attribute is not read.

    static readonly string[] BoxSides = ["top", "right", "bottom", "left"];

    static bool TryExpandBorderOrPadding(
        string prop, string val, bool important,
        Dictionary<string, (string Value, bool Important)> dest, CssStats? stats)
    {
        if (prop == "padding")
        {
            ExpandPadding(val, important, dest);
            return true;
        }
        if (prop is "border-width" or "border-style" or "border-color")
        {
            ExpandBorderBox(prop, val, important, dest, stats);
            return true;
        }
        if (prop == "border" || prop is "border-top" or "border-right" or "border-bottom" or "border-left")
        {
            ExpandBorderEdge(prop, val, important, dest, stats);
            return true;
        }
        return false;
    }

    static void ExpandPadding(string val, bool important, Dictionary<string, (string Value, bool Important)> dest)
    {
        var tokens = TokenizeCssValue(val);
        if (!TryAssignBox(tokens, out var assigned)) return;
        foreach (var token in assigned)
            if (!IsPaddingLength(token)) return;
        for (int i = 0; i < 4; i++)
            dest["padding-" + BoxSides[i]] = (assigned[i], important);
    }

    static void ExpandBorderBox(
        string prop, string val, bool important,
        Dictionary<string, (string Value, bool Important)> dest, CssStats? stats)
    {
        var component = prop["border-".Length..]; // width | style | color
        var tokens = TokenizeCssValue(val);
        if (!TryAssignBox(tokens, out var assigned))
        {
            if (stats != null) stats.BordersIgnored++;
            return;
        }
        var noted = new HashSet<string>(StringComparer.Ordinal);
        foreach (var token in assigned)
        {
            bool ok = component switch
            {
                "width" => IsBorderWidth(token),
                "style" => true,
                "color" => IsBorderColor(token),
                _ => false,
            };
            if (!ok)
            {
                if (stats != null) stats.BordersIgnored++;
                return;
            }
            if (component == "style" && noted.Add(token.Trim().ToLowerInvariant()))
                NoteBorderStyle(token, stats);
        }
        for (int i = 0; i < 4; i++)
        {
            var stored = component == "style" ? assigned[i].Trim().ToLowerInvariant() : assigned[i];
            dest[$"border-{BoxSides[i]}-{component}"] = (stored, important);
        }
    }

    static void ExpandBorderEdge(
        string prop, string val, bool important,
        Dictionary<string, (string Value, bool Important)> dest, CssStats? stats)
    {
        if (!TryParseEdge(val, out var width, out var style, out var color))
        {
            if (stats != null) stats.BordersIgnored++;
            return;
        }
        // Omitted components reset to the CSS initial value, so a later
        // shorthand replaces an earlier longhand in the same declaration block.
        width ??= "medium";
        style ??= "none";
        color ??= "currentcolor";
        NoteBorderStyle(style, stats);
        var sides = prop == "border" ? BoxSides : new[] { prop["border-".Length..] };
        foreach (var side in sides)
        {
            dest[$"border-{side}-width"] = (width, important);
            dest[$"border-{side}-style"] = (style.Trim().ToLowerInvariant(), important);
            dest[$"border-{side}-color"] = (color, important);
        }
    }

    static bool TryParseEdge(string val, out string? width, out string? style, out string? color)
    {
        width = style = color = null;
        var tokens = TokenizeCssValue(val);
        if (tokens.Count == 0) return false;
        foreach (var token in tokens)
        {
            if (IsBorderStyle(token))
            {
                if (style != null) return false;
                style = token;
                continue;
            }
            if (IsBorderWidth(token))
            {
                if (width != null) return false;
                width = token;
                continue;
            }
            if (IsBorderColor(token))
            {
                if (color != null) return false;
                color = token;
                continue;
            }
            return false;
        }
        return style != null || width != null || color != null;
    }

    static bool TryAssignBox(List<string> tokens, out string[] assigned)
    {
        assigned = new string[4];
        switch (tokens.Count)
        {
            case 1:
                assigned[0] = assigned[1] = assigned[2] = assigned[3] = tokens[0];
                return true;
            case 2:
                assigned[0] = assigned[2] = tokens[0];
                assigned[1] = assigned[3] = tokens[1];
                return true;
            case 3:
                assigned[0] = tokens[0];
                assigned[1] = assigned[3] = tokens[1];
                assigned[2] = tokens[2];
                return true;
            case 4:
                for (int i = 0; i < 4; i++) assigned[i] = tokens[i];
                return true;
            default:
                return false;
        }
    }

    static List<string> TokenizeCssValue(string value)
    {
        var tokens = new List<string>();
        var sb = new System.Text.StringBuilder();
        int depth = 0;
        foreach (var c in value)
        {
            if (c == '(') { depth++; sb.Append(c); continue; }
            if (c == ')') { if (depth > 0) depth--; sb.Append(c); continue; }
            if (depth == 0 && char.IsWhiteSpace(c))
            {
                if (sb.Length > 0) { tokens.Add(sb.ToString()); sb.Clear(); }
                continue;
            }
            sb.Append(c);
        }
        if (sb.Length > 0) tokens.Add(sb.ToString());
        return tokens;
    }

    static bool IsBorderStyleLonghand(string prop) => prop is
        "border-top-style" or "border-right-style" or "border-bottom-style" or "border-left-style";

    static bool IsBorderWidthLonghand(string prop) => prop is
        "border-top-width" or "border-right-width" or "border-bottom-width" or "border-left-width";

    static bool IsBorderColorLonghand(string prop) => prop is
        "border-top-color" or "border-right-color" or "border-bottom-color" or "border-left-color";

    static bool IsBorderStyle(string token) => token.Trim().ToLowerInvariant() is
        "none" or "hidden" or "solid" or "dashed" or "dotted" or "double"
        or "groove" or "ridge" or "inset" or "outset";

    static void NoteBorderStyle(string style, CssStats? stats)
    {
        if (stats == null) return;
        switch (style.Trim().ToLowerInvariant())
        {
            case "groove":
            case "ridge":
                stats.BordersApproximated++;
                break;
            case "none":
            case "hidden":
            case "solid":
            case "dashed":
            case "dotted":
            case "double":
            case "inset":
            case "outset":
                break;
            default:
                stats.BordersIgnored++;
                break;
        }
    }

    static bool IsBorderWidth(string token)
    {
        var v = token.Trim().ToLowerInvariant();
        if (v is "thin" or "medium" or "thick") return true;
        var twips = LengthTwips(v, 11);
        return twips != null && twips >= 0;
    }

    static bool IsPaddingLength(string token)
    {
        var v = token.Trim().ToLowerInvariant();
        if (v is "thin" or "medium" or "thick") return false;
        return IsBorderWidth(v);
    }

    static bool IsBorderColor(string token)
    {
        var v = token.Trim().ToLowerInvariant();
        if (v is "transparent" or "currentcolor" or "inherit") return true;
        return ParseColor(token) != null;
    }

    /// <summary>
    /// CSS px/pt/em/cm/mm/in → OOXML eighths of a point. 1px = 0.75pt = 6.
    /// thin/medium/thick follow the CSS guide (1px/3px/5px).
    /// </summary>
    static int? BorderEighths(string value, double basePt)
    {
        var v = value.Trim().ToLowerInvariant();
        if (v is "thin") v = "1px";
        else if (v is "medium") v = "3px";
        else if (v is "thick") v = "5px";
        var twips = LengthTwips(v, basePt);
        if (twips == null || twips < 0) return null;
        if (twips == 0) return 0;
        var eighths = (int)Math.Round(twips.Value * 0.4, MidpointRounding.AwayFromZero);
        if (eighths < 1) return 0;
        // w:sz is an eighths-of-a-point measure. Keep a sane ceiling so a
        // runaway length cannot emit a value Word's validator rejects.
        if (eighths > 255) return 255;
        return eighths;
    }

    static void FillMissing(HtmlFlowBorder dest, HtmlFlowBorder src)
    {
        dest.Top ??= src.Top;
        dest.Right ??= src.Right;
        dest.Bottom ??= src.Bottom;
        dest.Left ??= src.Left;
    }

    sealed partial class Converter
    {
        HtmlFlowBorder? ParagraphBorder(HtmlNode node, double basePt)
        {
            HtmlFlowBorder? merged = null;
            for (var n = node; n != null; n = n.Parent)
            {
                // Stop before the cell or the page. The cell border is
                // w:tcBorders; a body border is page chrome we don't draw.
                if (IsParagraphBorderBoundary(n.Tag)) break;
                var local = ResolveBorder(n, basePt, includeSpace: true);
                if (local == null) continue;
                merged ??= new HtmlFlowBorder();
                FillMissing(merged, local);
            }
            return merged != null && merged.Any ? merged : null;
        }

        static bool IsParagraphBorderBoundary(string? tag) => tag is
            "td" or "th" or "tr" or "table" or "thead" or "tbody" or "tfoot"
            or "ul" or "ol" or "body" or "html" or "#root";

        HtmlFlowBorder? CellBorder(HtmlNode cell, HtmlNode row, double basePt)
        {
            var merged = ResolveBorder(cell, basePt, includeSpace: false);
            var fromRow = ResolveBorder(row, basePt, includeSpace: false);
            if (fromRow == null) return merged;
            merged ??= new HtmlFlowBorder();
            FillMissing(merged, fromRow);
            return merged.Any ? merged : null;
        }

        HtmlFlowPadding? CellPadding(HtmlNode cell, HtmlNode row, double basePt)
        {
            var pad = ReadPadding(cell, basePt);
            if (!pad.Any) pad = ReadPadding(row, basePt);
            return pad.Any ? pad : null;
        }

        HtmlFlowBorder? ResolveBorder(HtmlNode node, double basePt, bool includeSpace)
        {
            var border = new HtmlFlowBorder
            {
                Top = ResolveSide(node, "top", basePt, includeSpace),
                Right = ResolveSide(node, "right", basePt, includeSpace),
                Bottom = ResolveSide(node, "bottom", basePt, includeSpace),
                Left = ResolveSide(node, "left", basePt, includeSpace),
            };
            return border.Any ? border : null;
        }

        HtmlFlowBorderSide? ResolveSide(HtmlNode node, string side, double basePt, bool includeSpace)
        {
            var styleRaw = Own(node, "border-" + side + "-style");
            var widthRaw = Own(node, "border-" + side + "-width");
            var colorRaw = Own(node, "border-" + side + "-color");
            if (styleRaw == null && widthRaw == null && colorRaw == null) return null;

            var style = (styleRaw ?? "none").Trim().ToLowerInvariant();
            if (style is "none" or "hidden")
                return new HtmlFlowBorderSide("nil", 0, null, null);

            if (!TryMapBorderStyle(style, out var wordStyle))
                return new HtmlFlowBorderSide("nil", 0, null, null);

            int eighths;
            if (widthRaw == null)
                eighths = BorderEighths("medium", basePt) ?? 0;
            else if (BorderEighths(widthRaw, basePt) is int parsed)
                eighths = parsed;
            else
                return null;

            if (eighths <= 0)
                return new HtmlFlowBorderSide("nil", 0, null, null);

            if (colorRaw != null && colorRaw.Trim().Equals("transparent", StringComparison.OrdinalIgnoreCase))
                return new HtmlFlowBorderSide("nil", 0, null, null);

            string? color = null;
            if (colorRaw != null
                && !colorRaw.Trim().Equals("currentcolor", StringComparison.OrdinalIgnoreCase)
                && !colorRaw.Trim().Equals("inherit", StringComparison.OrdinalIgnoreCase))
                color = ParseColor(colorRaw);

            int? space = includeSpace ? PaddingSpacePoints(node, side, basePt) : null;
            return new HtmlFlowBorderSide(wordStyle, eighths, color, space);
        }

        /// <summary>
        /// Word line styles the border helper already accepts. groove and ridge
        /// have no ST_Border token; inset and outset are the nearest.
        /// </summary>
        static bool TryMapBorderStyle(string css, out string word)
        {
            switch (css)
            {
                case "solid": word = "single"; return true;
                case "dashed": word = "dashed"; return true;
                case "dotted": word = "dotted"; return true;
                case "double": word = "double"; return true;
                case "inset": word = "inset"; return true;
                case "outset": word = "outset"; return true;
                case "groove": word = "inset"; return true;
                case "ridge": word = "outset"; return true;
                default: word = ""; return false;
            }
        }

        HtmlFlowPadding ReadPadding(HtmlNode node, double basePt)
        {
            int? Edge(string side)
            {
                var raw = Own(node, "padding-" + side);
                if (raw == null) return null;
                var twips = LengthTwips(raw, basePt);
                return twips != null && twips >= 0 ? twips : null;
            }
            return new HtmlFlowPadding(Edge("top"), Edge("right"), Edge("bottom"), Edge("left"));
        }

        int? PaddingSpacePoints(HtmlNode node, string side, double basePt)
        {
            var pad = ReadPadding(node, basePt);
            int? twips = side switch
            {
                "top" => pad.Top,
                "right" => pad.Right,
                "bottom" => pad.Bottom,
                "left" => pad.Left,
                _ => null,
            };
            if (twips == null || twips <= 0) return null;
            return (int)Math.Round(twips.Value / 20.0, MidpointRounding.AwayFromZero);
        }
    }
}
