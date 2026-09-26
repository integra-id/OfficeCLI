// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeCli.Core;

/// <summary>
/// Field refresh for Word documents.
///
/// <c>officecli refresh --toc</c> always rebuilds TOC entries from the current
/// headings (titles and hyperlinks) and writes PAGEREF results as the
/// placeholder <c>0</c>. It does not run Microsoft Word or a browser.
///
/// Without <c>--toc</c>, Word on Windows is tried first (real page numbers
/// and the rest of the field set). Otherwise HTML pagination fills PAGEREF
/// caches when a headless browser exists. If neither can number pages, the
/// entry rebuild is still kept and the command succeeds — page numbers stay
/// <c>0</c>.
/// </summary>
internal static class WordHtmlRefresh
{
    internal readonly record struct RefreshResult(
        bool Ok,
        string Backend,
        int TocFields,
        int Entries,
        string PageNumbers,
        string Message)
    {
        public string ToJson()
        {
            var sb = new StringBuilder();
            sb.Append("{\"backend\":\"").Append(JsonEscape(Backend)).Append('"');
            sb.Append(",\"tocFields\":").Append(TocFields);
            sb.Append(",\"entries\":").Append(Entries);
            sb.Append(",\"pageNumbers\":\"").Append(JsonEscape(PageNumbers)).Append('"');
            sb.Append(",\"message\":\"").Append(JsonEscape(Message)).Append("\"}");
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

    /// <param name="tocOnly">Rebuild TOC entries and skip Word / HTML pagination.</param>
    public static RefreshResult Refresh(string docx, bool tocOnly)
    {
        if (!tocOnly && OperatingSystem.IsWindows())
        {
            try
            {
                if (WordPdfBackend.RefreshFields(docx))
                {
                    var msg = $"Refreshed: {docx} (backend: word). Page numbers were updated by Microsoft Word.";
                    return new RefreshResult(true, "word", 0, 0, "resolved", msg);
                }
            }
            catch { /* fall through to the headless path */ }
        }

        WordTocBuilder.TocRebuildResult rebuilt;
        try
        {
            using var doc = WordprocessingDocument.Open(docx, true, new OpenSettings { AutoSave = false });
            rebuilt = WordTocBuilder.RegenerateAllTocs(doc);
            if (rebuilt.TocFields > rebuilt.Skipped)
                doc.Save();
        }
        catch
        {
            return Fail("refresh failed (could not rebuild TOC entries).");
        }

        if (tocOnly)
        {
            if (rebuilt.TocFields > 0 && rebuilt.Skipped == rebuilt.TocFields)
                return Fail("refresh --toc found TOC fields but could not rebuild them (the field result is not a contiguous block of paragraphs).");
            return EntriesResult(docx, rebuilt, tocOnly: true);
        }

        // -1: no browser / pagination failed. >=0: pagination ran, value is
        // how many PAGEREF caches were rewritten from the anchor→page map.
        int updated = -1;
        try { updated = TryApplyHtmlPageNumbers(docx); }
        catch { updated = -1; }

        if (updated > 0)
        {
            var msg = $"Refreshed: {docx} (backend: html, fields: {rebuilt.TocFields}, entries: {rebuilt.Entries}). "
                + "TOC page numbers reflect officecli's HTML pagination, which may differ from Word.";
            return new RefreshResult(true, "html", rebuilt.TocFields, rebuilt.Entries, "resolved", msg);
        }

        if (updated == 0 && rebuilt.WrotePagePlaceholders && rebuilt.TocFields > rebuilt.Skipped)
        {
            var msg = $"Rebuilt TOC entries: {docx} (backend: toc-entries, fields: {rebuilt.TocFields - rebuilt.Skipped}, entries: {rebuilt.Entries}). "
                + "HTML pagination ran but did not map any _Toc bookmark to a page, so PAGEREF results stay the placeholder 0.";
            return new RefreshResult(true, "toc-entries", rebuilt.TocFields - rebuilt.Skipped, rebuilt.Entries, "placeholder", msg);
        }

        if (updated >= 0)
        {
            var pages = rebuilt.WrotePagePlaceholders ? "placeholder" : "none";
            var msg = $"Refreshed: {docx} (backend: html, fields: {rebuilt.TocFields}, entries: {rebuilt.Entries}).";
            return new RefreshResult(true, "html", rebuilt.TocFields, rebuilt.Entries, pages, msg);
        }

        if (rebuilt.TocFields > rebuilt.Skipped)
            return EntriesResult(docx, rebuilt, tocOnly: false);

        return Fail("refresh failed (Word backend unavailable and HTML fallback failed — no headless browser found).");
    }

    static RefreshResult EntriesResult(string docx, WordTocBuilder.TocRebuildResult rebuilt, bool tocOnly)
    {
        if (rebuilt.TocFields == 0)
        {
            var none = $"No TOC field found: {docx} (backend: toc-entries). Nothing to rebuild.";
            return new RefreshResult(true, "toc-entries", 0, 0, "none", none);
        }

        var pages = rebuilt.WrotePagePlaceholders ? "placeholder" : "none";
        var why = tocOnly
            ? "Page numbers are not calculated on this path."
            : "Word and HTML pagination were unavailable, so page numbers were not calculated.";
        var placeholder = rebuilt.WrotePagePlaceholders
            ? " PAGEREF results are the placeholder 0. Update fields in Word, or run officecli refresh without --toc where Word or a headless browser is available, for real page numbers."
            : " This TOC omits page numbers.";
        var skipped = rebuilt.Skipped > 0 ? $" {rebuilt.Skipped} TOC field(s) left unchanged." : "";
        var msg = $"Rebuilt TOC entries: {docx} (backend: toc-entries, fields: {rebuilt.TocFields - rebuilt.Skipped}, entries: {rebuilt.Entries}). {why}{placeholder}{skipped}";
        return new RefreshResult(true, "toc-entries", rebuilt.TocFields - rebuilt.Skipped, rebuilt.Entries, pages, msg);
    }

    static RefreshResult Fail(string message)
        => new(false, "", 0, 0, "none", message);

    /// <returns>-1 when pagination is unavailable; otherwise the number of PAGEREF results rewritten.</returns>
    static int TryApplyHtmlPageNumbers(string docx)
    {
        string htmlSnapshot;
        using (var handler = Handlers.DocumentHandlerFactory.Open(docx, editable: false))
        {
            Handlers.Rendering.RenderingBootstrap.EnsureRegistered();
            var renderer = Rendering.RendererRegistry.Default.Resolve(
                "docx", Rendering.RenderOutputKind.Html, Rendering.RenderMode.Static);
            if (renderer == null) return -1;
            htmlSnapshot = renderer.Render(
                new Handlers.Rendering.HandlerRenderInput(handler, "docx"),
                new Rendering.RenderOptions()).Text ?? "";
        }

        var tmpHtml = Path.Combine(Path.GetTempPath(), $"officecli_refresh_{Guid.NewGuid():N}.html");
        HtmlScreenshot.PaginationResult? pagination;
        try
        {
            File.WriteAllText(tmpHtml, htmlSnapshot);
            pagination = HtmlScreenshot.GetPaginationFromDom(tmpHtml);
        }
        finally { try { File.Delete(tmpHtml); } catch { } }

        if (pagination == null) return -1;

        int updated;
        using (var doc = WordprocessingDocument.Open(docx, true, new OpenSettings { AutoSave = false }))
        {
            updated = ApplyPageNumbers(doc, pagination.AnchorPageMap);
            doc.Save();

            var part = doc.ExtendedFilePropertiesPart ?? doc.AddExtendedFilePropertiesPart();
            if (part.Properties == null)
                part.Properties = new DocumentFormat.OpenXml.ExtendedProperties.Properties();
            if (part.Properties.Pages == null)
                part.Properties.Pages = new DocumentFormat.OpenXml.ExtendedProperties.Pages();
            part.Properties.Pages.Text = pagination.TotalPages.ToString();
            part.Properties.Save();
        }
        return updated;
    }

    static int ApplyPageNumbers(WordprocessingDocument doc, Dictionary<string, int> map)
    {
        var body = doc.MainDocumentPart?.Document?.Body;
        if (body == null) return 0;
        int updated = 0;
        // Walk all PAGEREF fields. The instr text " PAGEREF _TocXXX \h "
        // identifies the bookmark; the very next Run after the separate
        // fldChar holds the cached page number Text we want to rewrite.
        foreach (var p in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
        {
            DocumentFormat.OpenXml.Wordprocessing.FieldCode? instr = null;
            foreach (var r in p.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>())
            {
                var fc = r.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.FieldChar>();
                if (fc?.FieldCharType?.Value == DocumentFormat.OpenXml.Wordprocessing.FieldCharValues.Begin)
                {
                    instr = null;
                }
                else if (instr == null)
                {
                    var ic = r.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.FieldCode>();
                    if (ic != null && ic.Text != null && ic.Text.TrimStart().StartsWith("PAGEREF", StringComparison.OrdinalIgnoreCase))
                        instr = ic;
                }
                else if (fc?.FieldCharType?.Value == DocumentFormat.OpenXml.Wordprocessing.FieldCharValues.Separate)
                {
                    var resultRun = r.NextSibling<DocumentFormat.OpenXml.Wordprocessing.Run>();
                    if (resultRun != null)
                    {
                        var anchor = ExtractPagerefAnchor(instr.Text!);
                        if (anchor != null && map.TryGetValue(anchor, out var pgNum))
                        {
                            var t = resultRun.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Text>();
                            if (t != null)
                            {
                                t.Text = pgNum.ToString();
                                updated++;
                            }
                        }
                    }
                    instr = null;
                }
            }
        }
        return updated;
    }

    static string? ExtractPagerefAnchor(string instrText)
    {
        var m = System.Text.RegularExpressions.Regex.Match(instrText, @"PAGEREF\s+(\S+)");
        return m.Success ? m.Groups[1].Value : null;
    }
}
