// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using OfficeCli.Core.Html;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    /// <summary>
    /// Outcome of <see cref="MaterializeAltChunks"/>. <see cref="Unchanged"/>
    /// means the package was not modified: either there was nothing to do,
    /// or <c>--strict</c> refused because at least one chunk cannot be converted.
    /// </summary>
    public sealed class AltChunkMaterializeReport
    {
        public int Converted { get; init; }
        public int Skipped { get; init; }
        public bool Unchanged { get; init; }
        /// <summary><c>--strict</c> found a chunk it cannot convert and did not modify the file.</summary>
        public bool Refused { get; init; }
        public IReadOnlyList<string> Warnings { get; init; } = Array.Empty<string>();

        public string Summary =>
            Refused
                ? $"materialize --strict left the file unchanged: {Skipped} altChunk(s) cannot be converted."
                : Converted == 0 && Skipped == 0
                    ? "No altChunks to materialize."
                    : Converted == 0
                        ? $"Left {Skipped} altChunk(s) in place; nothing was materialized."
                        : Skipped > 0
                            ? $"Materialized {Converted} altChunk(s); {Skipped} left in place."
                            : $"Materialized {Converted} altChunk(s).";

        public string ToJson()
        {
            var sb = new StringBuilder();
            sb.Append("{\"converted\":").Append(Converted);
            sb.Append(",\"skipped\":").Append(Skipped);
            sb.Append(",\"unchanged\":").Append(Unchanged ? "true" : "false");
            sb.Append(",\"refused\":").Append(Refused ? "true" : "false");
            sb.Append(",\"warnings\":[");
            for (int i = 0; i < Warnings.Count; i++)
            {
                if (i > 0) sb.Append(',');
                sb.Append('"').Append(JsonEscape(Warnings[i])).Append('"');
            }
            sb.Append("]}");
            return sb.ToString();
        }

        static string JsonEscape(string s)
        {
            var sb = new StringBuilder(s.Length + 8);
            foreach (var c in s)
            {
                switch (c)
                {
                    case '\\': sb.Append("\\\\"); break;
                    case '"': sb.Append("\\\""); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (c < ' ') sb.Append("\\u").Append(((int)c).ToString("x4"));
                        else sb.Append(c);
                        break;
                }
            }
            return sb.ToString();
        }
    }

    /// <summary>
    /// Replace <c>w:altChunk</c> elements in the main document with native
    /// paragraphs, tables and hyperlinks. HTML, XHTML and plain text are
    /// converted in-process. RTF, MHT and any other alternative format stay
    /// in place (or, when <paramref name="strict"/> is set, block the whole
    /// call so the file is left untouched).
    ///
    /// This is a subset of what Word's HTML importer does on open+save.
    /// A <c>data:</c> image (png, jpeg, gif, bmp, tiff, emf, wmf) is embedded
    /// with <see cref="AddPicture"/> as an inline <c>w:drawing</c>; the alt
    /// attribute is the picture description. webp and svg are not embedded.
    /// http(s) and relative URLs are not downloaded or read from disk — they
    /// stay as alt text, and a data URI that fails to decode does the same.
    /// Those picture warnings do not trip <paramref name="strict"/> and do
    /// not stop the rest of the chunk. Scripts are dropped, and layout CSS
    /// (float, flex, grid, media queries) is ignored. CSS borders are not:
    /// a paragraph or callout becomes <c>w:pBdr</c>, a table cell becomes
    /// <c>w:tcBorders</c>, and a uniform cell grid also replaces
    /// <c>w:tblBorders</c>. An unsupported line style is a warning and does
    /// not trip <paramref name="strict"/>. Source formatting is written as
    /// direct formatting — the same idea as <c>w:altChunkPr/w:matchSrc</c> —
    /// rather than mapped onto the destination styles. Heading elements also
    /// reference Heading1–Heading6 so an outline still works.
    /// </summary>
    public AltChunkMaterializeReport MaterializeAltChunks(bool strict = false)
    {
        var main = _doc.MainDocumentPart;
        var body = main?.Document?.Body;
        if (main == null || body == null)
            return new AltChunkMaterializeReport();

        var chunks = body.Descendants<AltChunk>().ToList();
        if (chunks.Count == 0)
            return new AltChunkMaterializeReport();

        var jobs = new List<ChunkJob>(chunks.Count);
        var warnings = new List<string>();
        foreach (var chunk in chunks)
            jobs.Add(ClassifyChunk(chunk, warnings));

        int skipped = 0;
        foreach (var job in jobs)
            if (job.Skip) skipped++;

        if (strict && skipped > 0)
            return new AltChunkMaterializeReport { Skipped = skipped, Unchanged = true, Refused = true, Warnings = warnings };

        return MarkModified(() => ApplyJobs(main, jobs, warnings, skipped));
    }

    sealed class ChunkJob
    {
        public required AltChunk Chunk { get; init; }
        public HtmlChunkParseResult? Parsed { get; init; }
        public bool Skip { get; init; }
        public string? PartId { get; init; }
    }

    ChunkJob ClassifyChunk(AltChunk chunk, List<string> warnings)
    {
        var label = ChunkLabel(chunk);
        var parent = chunk.Parent;
        if (parent is not Body && parent is not TableCell && parent is not SdtContentBlock)
        {
            warnings.Add($"{label}: skipped — parent is not the document body or a table cell");
            return new ChunkJob { Chunk = chunk, Skip = true };
        }

        var part = TryGetAltChunkPart(chunk);
        if (part == null)
        {
            warnings.Add($"{label}: skipped — alternative-format part is missing");
            return new ChunkJob { Chunk = chunk, Skip = true, PartId = chunk.Id?.Value };
        }

        var format = AltChunkFormatOf(part.ContentType);
        if (format is not ("html" or "xhtml" or "text"))
        {
            warnings.Add($"{label}: skipped — format '{format ?? part.ContentType}' is not materialized (html, xhtml, and plain text only)");
            return new ChunkJob { Chunk = chunk, Skip = true, PartId = chunk.Id?.Value };
        }

        string text;
        try
        {
            using var stream = part.GetStream();
            text = CharsetDecoder.Decode(stream, CharsetDecoder.ParseCharset(part.ContentType));
        }
        catch (Exception ex)
        {
            warnings.Add($"{label}: skipped — could not read the chunk ({ex.Message})");
            return new ChunkJob { Chunk = chunk, Skip = true, PartId = chunk.Id?.Value };
        }

        HtmlChunkParseResult parsed;
        try
        {
            parsed = format == "text"
                ? HtmlChunkParser.ParsePlainText(text)
                : HtmlChunkParser.ParseHtml(text);
        }
        catch (Exception ex)
        {
            warnings.Add($"{label}: skipped — HTML parser failed ({ex.Message})");
            return new ChunkJob { Chunk = chunk, Skip = true, PartId = chunk.Id?.Value };
        }

        foreach (var w in parsed.Warnings)
            warnings.Add($"{label}: {w}");
        if (parsed.Blocks.Count == 0)
            warnings.Add($"{label}: chunk produced no paragraphs or tables");

        return new ChunkJob { Chunk = chunk, Parsed = parsed, PartId = chunk.Id?.Value };
    }

    static string ChunkLabel(AltChunk chunk)
    {
        var id = chunk.Id?.Value;
        return string.IsNullOrEmpty(id) ? "altChunk" : "altChunk " + id;
    }

    AltChunkMaterializeReport ApplyJobs(MainDocumentPart main, List<ChunkJob> jobs, List<string> warnings, int skipped)
    {
        int converted = 0;
        int droppedLinks = 0;
        var linkRels = new Dictionary<string, string>(StringComparer.Ordinal);
        var pictureIds = SeedPictureIds();

        foreach (var job in jobs)
        {
            if (job.Skip || job.Parsed == null) continue;
            var chunk = job.Chunk;
            var parent = chunk.Parent;
            if (parent == null) continue;
            var spacer = chunk.NextSibling();

            foreach (var block in job.Parsed.Blocks)
            {
                var el = CreateFlowBlock(block, main, linkRels, ref droppedLinks, pictureIds, warnings);
                parent.InsertBefore(el, chunk);
                if (el is Paragraph para && block is HtmlFlowParagraph flow && flow.ListKind != null)
                    ApplyListStyle(para, flow.ListKind == "bullet" ? "bullet" : "ordered", flow.ListStart, flow.ListLevel);
            }

            chunk.Remove();
            DeleteAltChunkPartIfUnreferenced(main, job.PartId);
            if (parent is TableCell cell)
                TidyCellAfterChunk(cell, spacer);
            converted++;
        }

        if (droppedLinks > 0)
            warnings.Add($"{droppedLinks} hyperlink(s) left as plain text (unsupported URI scheme)");
        // Heading styles may have been added to styles.xml. Drop the id index
        // so a later get/query in this process sees them.
        if (converted > 0) InvalidateStyleIndex();
        InvalidateBodyParaCache();

        return new AltChunkMaterializeReport
        {
            Converted = converted,
            Skipped = skipped,
            Unchanged = converted == 0,
            Warnings = warnings,
        };
    }

    void TidyCellAfterChunk(TableCell cell, OpenXmlElement? spacer)
    {
        // add htmlchunk keeps an empty paragraph after a chunk that would
        // otherwise be the cell's last child. Once the chunk is real content
        // that already ends with a paragraph, that spacer is a blank line.
        if (spacer is Paragraph pad
            && pad.Parent == cell
            && IsVisuallyEmpty(pad)
            && cell.Elements<Paragraph>().Any(p => !ReferenceEquals(p, pad)))
            pad.Remove();
        if (cell.LastChild is not Paragraph)
        {
            var trailing = new Paragraph();
            AssignParaId(trailing);
            cell.AppendChild(trailing);
        }
    }

    static bool IsVisuallyEmpty(Paragraph para)
    {
        foreach (var text in para.Descendants<Text>())
            if (!string.IsNullOrEmpty(text.Text)) return false;
        if (para.Descendants<Drawing>().Any()) return false;
        if (para.Descendants<Break>().Any()) return false;
        return true;
    }

    void DeleteAltChunkPartIfUnreferenced(MainDocumentPart main, string? partId)
    {
        if (string.IsNullOrEmpty(partId)) return;
        var body = main.Document?.Body;
        if (body == null) return;
        int refs = body.Descendants<AltChunk>().Count(ac => ac.Id?.Value == partId);
        if (refs > 0) return;
        try { main.DeletePart(partId); } catch { /* already gone */ }
    }

    OpenXmlElement CreateFlowBlock(HtmlFlowBlock block, MainDocumentPart main, Dictionary<string, string> linkRels, ref int droppedLinks, PictureIds pictureIds, List<string> warnings)
    {
        return block switch
        {
            HtmlFlowTable table => CreateTable(table, main, linkRels, ref droppedLinks, pictureIds, warnings),
            HtmlFlowParagraph para => CreateParagraph(para, main, linkRels, ref droppedLinks, pictureIds, warnings),
            _ => new Paragraph(),
        };
    }

    Paragraph CreateParagraph(HtmlFlowParagraph flow, MainDocumentPart main, Dictionary<string, string> linkRels, ref int droppedLinks, PictureIds pictureIds, List<string> warnings)
    {
        var para = new Paragraph();
        AssignParaId(para);
        var pPr = new ParagraphProperties();
        if (!string.IsNullOrEmpty(flow.StyleId))
        {
            TryMaterializeBuiltInStyle(flow.StyleId);
            pPr.ParagraphStyleId = new ParagraphStyleId { Val = flow.StyleId };
        }
        if (flow.HorizontalRule)
        {
            pPr.ParagraphBorders = new ParagraphBorders(
                new BottomBorder { Val = BorderValues.Single, Size = 12, Space = 1, Color = "auto" });
        }
        else if (ParagraphBordersFrom(flow.Border) is { } cssBorders)
            pPr.ParagraphBorders = cssBorders;
        if (!string.IsNullOrEmpty(flow.Fill))
            pPr.Shading = new Shading { Val = ShadingPatternValues.Clear, Fill = flow.Fill };
        if (flow.SpaceBeforeTwips != null || flow.SpaceAfterTwips != null)
        {
            var spacing = new SpacingBetweenLines();
            if (flow.SpaceBeforeTwips != null) spacing.Before = flow.SpaceBeforeTwips.Value.ToString();
            if (flow.SpaceAfterTwips != null) spacing.After = flow.SpaceAfterTwips.Value.ToString();
            pPr.SpacingBetweenLines = spacing;
        }
        // List markers bring their own indent from the numbering definition.
        if (flow.ListKind == null && (flow.IndentLeftTwips > 0 || flow.FirstLineTwips != null))
        {
            var ind = new Indentation();
            if (flow.IndentLeftTwips > 0) ind.Left = flow.IndentLeftTwips.ToString();
            if (flow.FirstLineTwips is int fl)
            {
                if (fl >= 0) ind.FirstLine = fl.ToString();
                else ind.Hanging = (-fl).ToString();
            }
            pPr.Indentation = ind;
        }
        if (flow.Align is { Length: > 0 } align)
        {
            pPr.Justification = new Justification
            {
                Val = align switch
                {
                    "center" => JustificationValues.Center,
                    "right" => JustificationValues.Right,
                    "both" => JustificationValues.Both,
                    _ => JustificationValues.Left,
                }
            };
        }
        // AddPicture sets this on a picture-only paragraph so a fixed Normal
        // line height does not clip the drawing. Do it here too when the
        // picture shares the paragraph with text.
        if (flow.Runs.Any(r => r.ImageSrc != null))
        {
            var spacing = pPr.SpacingBetweenLines ?? new SpacingBetweenLines();
            if (spacing.LineRule == null)
            {
                spacing.Line = "240";
                spacing.LineRule = LineSpacingRuleValues.Auto;
            }
            pPr.SpacingBetweenLines = spacing;
        }
        if (pPr.HasChildren) para.AppendChild(pPr);

        var runs = flow.Runs;
        for (int i = 0; i < runs.Count;)
        {
            if (runs[i].ImageSrc != null)
            {
                if (!TryAppendHtmlPicture(para, runs[i], pictureIds, warnings))
                    AppendImageFallback(para, runs[i]);
                i++;
                continue;
            }
            var href = runs[i].Href;
            if (string.IsNullOrWhiteSpace(href))
            {
                AppendRun(para, runs[i]);
                i++;
                continue;
            }
            int j = i;
            // An inline picture is its own run (AddPicture). Don't swallow it
            // into the surrounding hyperlink's text runs.
            while (j < runs.Count && runs[j].ImageSrc == null && runs[j].Href == href) j++;
            if (TryCreateHyperlink(main, href, linkRels, out var hyperlink))
            {
                for (int k = i; k < j; k++) AppendRun(hyperlink, runs[k]);
                if (hyperlink.HasChildren) para.AppendChild(hyperlink);
            }
            else
            {
                droppedLinks++;
                for (int k = i; k < j; k++) AppendRun(para, runs[k]);
            }
            i = j;
        }
        return para;
    }

    bool TryCreateHyperlink(MainDocumentPart main, string href, Dictionary<string, string> linkRels, out Hyperlink hyperlink)
    {
        hyperlink = new Hyperlink();
        href = href.Trim();
        if (href.StartsWith('#'))
        {
            var anchor = href[1..];
            if (anchor.Length == 0) return false;
            hyperlink.Anchor = anchor;
            return true;
        }
        try
        {
            HyperlinkUriValidator.RequireSafeScheme(href, "href");
            if (!linkRels.TryGetValue(href, out var relId))
            {
                Uri uri;
                if (Uri.TryCreate(href, UriKind.Absolute, out var abs))
                    uri = new Uri(PercentEncodeUri(href), UriKind.Absolute);
                else if (!Uri.TryCreate(href, UriKind.Relative, out uri!))
                    return false;
                relId = main.AddHyperlinkRelationship(uri, isExternal: true).Id;
                linkRels[href] = relId;
            }
            hyperlink.Id = relId;
            return true;
        }
        catch (ArgumentException)
        {
            return false;
        }
    }

    static void AppendRun(OpenXmlElement host, HtmlFlowRun flow)
    {
        if (!flow.Break && flow.Text.Length == 0) return;
        var run = new Run();
        var rPr = new RunProperties();
        if (!string.IsNullOrEmpty(flow.Font))
            rPr.RunFonts = new RunFonts { Ascii = flow.Font, HighAnsi = flow.Font, EastAsia = flow.Font };
        if (flow.Bold)
        {
            rPr.Bold = new Bold();
            rPr.BoldComplexScript = new BoldComplexScript();
        }
        if (flow.Italic)
        {
            rPr.Italic = new Italic();
            rPr.ItalicComplexScript = new ItalicComplexScript();
        }
        if (flow.Strike) rPr.Strike = new Strike();
        if (!string.IsNullOrEmpty(flow.Color))
            rPr.Color = new Color { Val = flow.Color };
        if (flow.SizePt is double pt && pt > 0)
        {
            var half = ((int)Math.Round(pt * 2, MidpointRounding.AwayFromZero)).ToString();
            rPr.FontSize = new FontSize { Val = half };
            rPr.FontSizeComplexScript = new FontSizeComplexScript { Val = half };
        }
        if (flow.Underline)
            rPr.Underline = new Underline { Val = UnderlineValues.Single };
        if (flow.VertAlign == "subscript")
            rPr.VerticalTextAlignment = new VerticalTextAlignment { Val = VerticalPositionValues.Subscript };
        else if (flow.VertAlign == "superscript")
            rPr.VerticalTextAlignment = new VerticalTextAlignment { Val = VerticalPositionValues.Superscript };
        if (!string.IsNullOrEmpty(flow.Fill))
            rPr.Shading = new Shading { Val = ShadingPatternValues.Clear, Fill = flow.Fill };
        if (rPr.HasChildren) run.AppendChild(rPr);

        if (flow.Break)
            run.AppendChild(new Break());
        if (flow.Text.Length > 0)
            AppendTextWithTabs(run, flow.Text);
        if (run.ChildElements.Count > (rPr.HasChildren ? 1 : 0))
            host.AppendChild(run);
    }

    static void AppendTextWithTabs(Run run, string text)
    {
        var parts = text.Split('\t');
        for (int p = 0; p < parts.Length; p++)
        {
            if (p > 0) run.AppendChild(new TabChar());
            if (parts[p].Length > 0)
                run.AppendChild(new Text(parts[p]) { Space = SpaceProcessingModeValues.Preserve });
        }
    }

    Table CreateTable(HtmlFlowTable flow, MainDocumentPart main, Dictionary<string, string> linkRels, ref int droppedLinks, PictureIds pictureIds, List<string> warnings)
    {
        var slots = new Dictionary<(int R, int C), CellSlot>();
        int cols = 0;
        int rows = flow.Rows.Count;
        for (int r = 0; r < flow.Rows.Count; r++)
        {
            int c = 0;
            foreach (var cell in flow.Rows[r].Cells)
            {
                while (slots.ContainsKey((r, c))) c++;
                int cs = Math.Max(1, cell.ColSpan);
                int rs = Math.Max(1, cell.RowSpan);
                for (int rr = r; rr < r + rs; rr++)
                {
                    for (int cc = c; cc < c + cs; cc++)
                    {
                        slots[(rr, cc)] = new CellSlot(cell, rr == r && cc == c, c, cs);
                    }
                }
                c += cs;
                if (c > cols) cols = c;
                if (r + rs > rows) rows = r + rs;
            }
        }
        if (cols == 0) cols = 1;

        const int contentTwips = 9360;
        int colW = Math.Max(120, contentTwips / cols);

        var table = new Table();
        var tblPr = new TableProperties();
        tblPr.TableWidth = new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct };
        tblPr.TableBorders = TableBordersFor(flow);
        table.AppendChild(tblPr);
        var grid = new TableGrid();
        for (int i = 0; i < cols; i++)
            grid.AppendChild(new GridColumn { Width = colW.ToString() });
        table.AppendChild(grid);

        for (int r = 0; r < rows; r++)
        {
            var tr = new TableRow();
            bool headerRow = r < flow.Rows.Count && flow.Rows[r].Header;
            if (headerRow)
                tr.AppendChild(new TableRowProperties(new TableHeader()));
            for (int c = 0; c < cols; c++)
            {
                if (!slots.TryGetValue((r, c), out var slot))
                {
                    tr.AppendChild(EmptyCell(colW));
                    continue;
                }
                if (slot.OriginCol != c) continue;
                tr.AppendChild(slot.IsOrigin
                    ? FilledCell(slot, colW, main, linkRels, ref droppedLinks, pictureIds, warnings)
                    : ContinueCell(slot, colW));
            }
            if (!tr.Elements<TableCell>().Any())
                tr.AppendChild(EmptyCell(colW));
            table.AppendChild(tr);
        }
        return table;
    }

    sealed record CellSlot(HtmlFlowCell Cell, bool IsOrigin, int OriginCol, int ColSpan);

    TableCell FilledCell(CellSlot slot, int colW, MainDocumentPart main, Dictionary<string, string> linkRels, ref int droppedLinks, PictureIds pictureIds, List<string> warnings)
    {
        var tc = new TableCell();
        tc.AppendChild(CellProperties(slot.Cell.ColSpan, colW, slot.Cell.Fill,
            slot.Cell.RowSpan > 1 ? MergedCellValues.Restart : null,
            slot.Cell.Border, slot.Cell.Padding));
        foreach (var block in slot.Cell.Blocks)
        {
            var el = CreateFlowBlock(block, main, linkRels, ref droppedLinks, pictureIds, warnings);
            tc.AppendChild(el);
            if (el is Paragraph para && block is HtmlFlowParagraph flow && flow.ListKind != null)
                ApplyListStyle(para, flow.ListKind == "bullet" ? "bullet" : "ordered", flow.ListStart, flow.ListLevel);
        }
        if (tc.LastChild is not Paragraph)
        {
            var trailing = new Paragraph();
            AssignParaId(trailing);
            tc.AppendChild(trailing);
        }
        return tc;
    }

    TableCell ContinueCell(CellSlot slot, int colW)
    {
        var tc = new TableCell();
        tc.AppendChild(CellProperties(slot.ColSpan, colW, slot.Cell.Fill, MergedCellValues.Continue,
            slot.Cell.Border, slot.Cell.Padding));
        var para = new Paragraph();
        AssignParaId(para);
        tc.AppendChild(para);
        return tc;
    }

    TableCell EmptyCell(int colW)
    {
        var tc = new TableCell();
        tc.AppendChild(CellProperties(1, colW, null, null));
        var para = new Paragraph();
        AssignParaId(para);
        tc.AppendChild(para);
        return tc;
    }

    static TableCellProperties CellProperties(int colSpan, int colW, string? fill, MergedCellValues? merge,
        HtmlFlowBorder? border = null, HtmlFlowPadding? padding = null)
    {
        var tcPr = new TableCellProperties();
        int span = Math.Max(1, colSpan);
        tcPr.TableCellWidth = new TableCellWidth { Width = (colW * span).ToString(), Type = TableWidthUnitValues.Dxa };
        if (span > 1) tcPr.GridSpan = new GridSpan { Val = span };
        if (merge != null) tcPr.VerticalMerge = new VerticalMerge { Val = merge.Value };
        // CT_TcPr order: tcW, gridSpan, vMerge, tcBorders, shd, tcMar.
        if (CellBordersFrom(border) is { } tcBorders)
            tcPr.TableCellBorders = tcBorders;
        if (!string.IsNullOrEmpty(fill))
            tcPr.Shading = new Shading { Val = ShadingPatternValues.Clear, Fill = fill };
        if (CellMarginFrom(padding) is { } mar)
            tcPr.TableCellMargin = mar;
        return tcPr;
    }

    /// <summary>
    /// Paragraph borders skip nil sides (no line is the same as omitting the
    /// edge). <see cref="MakeBorder"/> writes <c>w:sz</c> and <c>w:color</c>.
    /// Padding on that edge becomes <c>w:space</c> in points.
    /// </summary>
    static ParagraphBorders? ParagraphBordersFrom(HtmlFlowBorder? border)
    {
        if (border == null || !border.Any) return null;
        var children = new List<OpenXmlElement>();
        AddEdge<TopBorder>(children, border.Top, emitNil: false, withSpace: true);
        AddEdge<LeftBorder>(children, border.Left, emitNil: false, withSpace: true);
        AddEdge<BottomBorder>(children, border.Bottom, emitNil: false, withSpace: true);
        AddEdge<RightBorder>(children, border.Right, emitNil: false, withSpace: true);
        return children.Count == 0 ? null : new ParagraphBorders(children);
    }

    /// <summary>
    /// Cell borders keep an explicit nil so <c>border: none</c> covers the
    /// table grid. Unspecified sides are omitted and fall through to
    /// <c>w:tblBorders</c>.
    /// </summary>
    static TableCellBorders? CellBordersFrom(HtmlFlowBorder? border)
    {
        if (border == null || !border.Any) return null;
        var children = new List<OpenXmlElement>();
        AddEdge<TopBorder>(children, border.Top, emitNil: true, withSpace: false);
        AddEdge<LeftBorder>(children, border.Left, emitNil: true, withSpace: false);
        AddEdge<BottomBorder>(children, border.Bottom, emitNil: true, withSpace: false);
        AddEdge<RightBorder>(children, border.Right, emitNil: true, withSpace: false);
        return children.Count == 0 ? null : new TableCellBorders(children);
    }

    static TableCellMargin? CellMarginFrom(HtmlFlowPadding? padding)
    {
        if (padding is not { } pad || !pad.Any) return null;
        var children = new List<OpenXmlElement>();
        AddMargin<TopMargin>(children, pad.Top);
        AddMargin<LeftMargin>(children, pad.Left);
        AddMargin<BottomMargin>(children, pad.Bottom);
        AddMargin<RightMargin>(children, pad.Right);
        return children.Count == 0 ? null : new TableCellMargin(children);
    }

    static void AddMargin<T>(List<OpenXmlElement> children, int? twips) where T : TableWidthType, new()
    {
        if (twips == null) return;
        children.Add(new T
        {
            Width = Math.Max(0, twips.Value).ToString(CultureInfo.InvariantCulture),
            Type = TableWidthUnitValues.Dxa,
        });
    }

    /// <summary>
    /// A grid where every cell carries the same four-edge border replaces the
    /// built-in table borders (including insideH/insideV) so the CSS color is
    /// the grid. Otherwise the table element's own border paints the outside
    /// and the built-in single grid stays as the fallback.
    /// </summary>
    static TableBorders TableBordersFor(HtmlFlowTable flow)
    {
        if (TryUniformCellBorder(flow, out var side))
            return new TableBorders(
                FlowBorder<TopBorder>(side, withSpace: false),
                FlowBorder<LeftBorder>(side, withSpace: false),
                FlowBorder<BottomBorder>(side, withSpace: false),
                FlowBorder<RightBorder>(side, withSpace: false),
                FlowBorder<InsideHorizontalBorder>(side, withSpace: false),
                FlowBorder<InsideVerticalBorder>(side, withSpace: false));

        if (flow.Border == null || !flow.Border.Any)
            return DefaultTableBorders();

        return new TableBorders(
            TableEdge<TopBorder>(flow.Border.Top),
            TableEdge<LeftBorder>(flow.Border.Left),
            TableEdge<BottomBorder>(flow.Border.Bottom),
            TableEdge<RightBorder>(flow.Border.Right),
            DefaultEdge<InsideHorizontalBorder>(),
            DefaultEdge<InsideVerticalBorder>());
    }

    static bool TryUniformCellBorder(HtmlFlowTable flow, out HtmlFlowBorderSide side)
    {
        side = default;
        bool any = false;
        foreach (var row in flow.Rows)
        {
            if (row.Cells.Count == 0) return false;
            foreach (var cell in row.Cells)
            {
                if (cell.Border == null || !UniformBox(cell.Border, out var box))
                    return false;
                if (!any)
                {
                    side = box;
                    any = true;
                }
                else if (!SameBorder(side, box))
                    return false;
            }
        }
        return any;
    }

    static bool UniformBox(HtmlFlowBorder border, out HtmlFlowBorderSide side)
    {
        side = default;
        if (border.Top == null || border.Right == null || border.Bottom == null || border.Left == null)
            return false;
        if (!SameBorder(border.Top.Value, border.Right.Value)
            || !SameBorder(border.Top.Value, border.Bottom.Value)
            || !SameBorder(border.Top.Value, border.Left.Value))
            return false;
        side = border.Top.Value;
        return true;
    }

    static bool SameBorder(HtmlFlowBorderSide a, HtmlFlowBorderSide b) =>
        a.Style == b.Style && a.SizeEighths == b.SizeEighths
        && string.Equals(a.Color, b.Color, StringComparison.OrdinalIgnoreCase);

    static TableBorders DefaultTableBorders() => new(
        DefaultEdge<TopBorder>(),
        DefaultEdge<LeftBorder>(),
        DefaultEdge<BottomBorder>(),
        DefaultEdge<RightBorder>(),
        DefaultEdge<InsideHorizontalBorder>(),
        DefaultEdge<InsideVerticalBorder>());

    static T DefaultEdge<T>() where T : BorderType, new() =>
        new() { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" };

    static T TableEdge<T>(HtmlFlowBorderSide? side) where T : BorderType, new() =>
        side == null ? DefaultEdge<T>() : FlowBorder<T>(side.Value, withSpace: false);

    static void AddEdge<T>(List<OpenXmlElement> children, HtmlFlowBorderSide? side, bool emitNil, bool withSpace)
        where T : BorderType, new()
    {
        if (side == null) return;
        if (side.Value.Style == "nil" && !emitNil) return;
        children.Add(FlowBorder<T>(side.Value, withSpace));
    }

    static T FlowBorder<T>(HtmlFlowBorderSide side, bool withSpace) where T : BorderType, new()
    {
        if (side.Style == "nil")
            return MakeBorder<T>(BorderValues.Nil, 0, null, null);
        uint? space = withSpace && side.SpacePoints is int sp && sp > 0 ? (uint)sp : null;
        var color = string.IsNullOrEmpty(side.Color) ? "auto" : side.Color;
        var size = (uint)Math.Clamp(side.SizeEighths, 1, 255);
        return MakeBorder<T>(ParseBorderStyle(side.Style), size, color, space);
    }

    /// <summary>wp:docPr ids already used in the package, plus the next free id.</summary>
    sealed class PictureIds
    {
        public HashSet<uint> Used { get; } = new();
        public uint Next { get; set; } = 1;
    }

    PictureIds SeedPictureIds()
    {
        var ids = new PictureIds();
        var main = _doc.MainDocumentPart;
        uint max = 0;
        if (main != null)
        {
            foreach (var root in EnumerateContentRoots(main))
            {
                foreach (var dp in root.Descendants<DW.DocProperties>())
                {
                    if (dp.Id?.HasValue != true) continue;
                    ids.Used.Add(dp.Id.Value);
                    if (dp.Id.Value > max) max = dp.Id.Value;
                }
            }
        }
        ids.Next = max == uint.MaxValue ? 1 : max + 1;
        return ids;
    }

    /// <summary>
    /// Embed one HTML image by calling <see cref="AddPicture"/> — the same
    /// helper <c>add picture</c> uses — so data URIs, content types, and
    /// alt/description land on a normal inline drawing. The paragraph may
    /// not be in the package yet (a table cell is built before it is
    /// inserted), so docPr ids are uniqued here: <see cref="AddPicture"/>
    /// only sees drawings already in the package.
    /// </summary>
    bool TryAppendHtmlPicture(Paragraph para, HtmlFlowRun flow, PictureIds pictureIds, List<string> warnings)
    {
        if (string.IsNullOrEmpty(flow.ImageSrc)) return false;
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["src"] = flow.ImageSrc,
        };
        if (!string.IsNullOrEmpty(flow.ImageAlt))
            props["alt"] = flow.ImageAlt;
        var width = PictureLength(flow.ImageWidth);
        var height = PictureLength(flow.ImageHeight);
        if (width != null) props["width"] = width;
        if (height != null) props["height"] = height;
        // Click target. AddPicture stores it as a:hlinkClick and drops an
        // unsafe scheme without failing the picture.
        if (!string.IsNullOrWhiteSpace(flow.Href))
            props["link"] = flow.Href.Trim();

        var before = new HashSet<Run>(para.Descendants<Run>());
        try
        {
            AddPicture(para, "/body", null, props);
        }
        catch (Exception ex)
        {
            var rescued = NewPictureRun(para, before);
            if (rescued != null)
            {
                ClaimPictureRun(rescued, pictureIds);
                warnings.Add("image embedded; a picture property was dropped (" + Brief(ex) + ")");
                return true;
            }
            warnings.Add("image kept as alt text (data URI could not be embedded: " + Brief(ex) + ")");
            return false;
        }

        var run = NewPictureRun(para, before);
        if (run == null)
        {
            warnings.Add("image kept as alt text (picture helper did not insert a drawing)");
            return false;
        }
        ClaimPictureRun(run, pictureIds);
        return true;
    }

    static Run? NewPictureRun(Paragraph para, HashSet<Run> before)
    {
        foreach (var run in para.Descendants<Run>())
        {
            if (before.Contains(run)) continue;
            if (run.Descendants<Drawing>().Any()) return run;
        }
        return null;
    }

    static void ClaimPictureRun(Run run, PictureIds ids)
    {
        var dp = run.Descendants<DW.DocProperties>().FirstOrDefault();
        if (dp == null) return;
        var nv = run.Descendants<PIC.NonVisualDrawingProperties>().FirstOrDefault();
        uint id = dp.Id?.Value ?? 0;
        if (id == 0 || ids.Used.Contains(id))
        {
            while (ids.Next == 0 || ids.Used.Contains(ids.Next))
            {
                if (ids.Next == uint.MaxValue) { ids.Next = 1; break; }
                ids.Next++;
            }
            id = ids.Next;
            if (ids.Next != uint.MaxValue) ids.Next++;
            dp.Id = id;
            if (nv != null) nv.Id = id;
        }
        ids.Used.Add(id);
    }

    static void AppendImageFallback(Paragraph para, HtmlFlowRun flow)
    {
        var label = string.IsNullOrEmpty(flow.ImageAlt) ? "image" : flow.ImageAlt;
        AppendRun(para, new HtmlFlowRun { Italic = true, Text = "[image: " + label + "]" });
    }

    /// <summary>
    /// HTML width/height for <see cref="AddPicture"/>. A bare number is CSS
    /// pixels (the HTML attribute rule), not raw EMU. Percentages and
    /// font-relative units are dropped so the picture helper can use the
    /// image's own aspect ratio instead of failing the embed.
    /// </summary>
    static string? PictureLength(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return null;
        raw = raw.Trim();
        if (raw.EndsWith('%')) return null;
        if (raw.Equals("auto", StringComparison.OrdinalIgnoreCase)
            || raw.Equals("inherit", StringComparison.OrdinalIgnoreCase)
            || raw.Equals("initial", StringComparison.OrdinalIgnoreCase)
            || raw.Equals("unset", StringComparison.OrdinalIgnoreCase))
            return null;
        var token = double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out _)
            ? raw + "px"
            : raw;
        try
        {
            if (EmuConverter.ParseEmu(token) <= 0) return null;
            return token;
        }
        catch (ArgumentException)
        {
            return null;
        }
    }

    static string Brief(Exception ex)
    {
        var message = ex.Message.Replace('\r', ' ').Replace('\n', ' ').Trim();
        if (message.Length == 0) message = ex.GetType().Name;
        if (message.Length > 160) message = message[..160] + "…";
        return message;
    }
}
