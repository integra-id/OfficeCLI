// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    /// <summary>
    /// `add --type htmlchunk` (aliases: html, altchunk) — embed a chunk of
    /// HTML (or another alternate format: xhtml/mht/rtf/text) into the
    /// document via Word's native "alternative format import" mechanism:
    /// the payload is stored verbatim in an <see cref="AlternativeFormatImportPart"/>
    /// and referenced from the body by <c>&lt;w:altChunk r:id="…"/&gt;</c>.
    ///
    /// Word (and LibreOffice) convert the chunk into native paragraphs,
    /// tables, lists, … the first time the file is opened and saved, so the
    /// HTML keeps its full fidelity (CSS, nested tables, colspan, …) without
    /// OfficeCLI having to translate it. Unlike `markdown` this is NOT an
    /// expansion: the chunk stays a single <c>altChunk</c> node until Word
    /// converts it (the HTML preview shows it as escaped source).
    /// <c>officecli materialize</c> can replace HTML/XHTML/plain-text chunks
    /// with native body content without Word; that converter is a subset
    /// (see <see cref="MaterializeAltChunks"/>).
    ///
    /// Input: inline via canonical <c>html</c> (aliases: content, text) or a
    /// file via <c>src</c> (alias: path). The payload format is inferred from
    /// the src extension, or forced with <c>format</c>.
    /// </summary>
    private string AddHtmlChunk(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var mainPart = _doc.MainDocumentPart
            ?? throw new InvalidOperationException("Document main part not found");

        // altChunk r:id is resolved against the part that owns the element.
        // Restrict to the main document body (and table cells in it): header,
        // footer and footnote chunks would need their own part relationship
        // and Word does not reliably import them there.
        if (parent is not Body && parent is not TableCell)
            throw new ArgumentException(
                $"Cannot add 'htmlchunk' under {parentPath}: an HTML chunk is block-level content — add it at /body or into a table cell (/body/tbl[N]/tr[M]/tc[K]).");
        if (parent is TableCell && parent.Ancestors<Body>().FirstOrDefault() == null)
            throw new ArgumentException(
                $"Cannot add 'htmlchunk' under {parentPath}: HTML chunks are only supported in the main document body (not headers, footers, footnotes or endnotes).");

        string? payload = properties.GetValueOrDefault("html")
                          ?? properties.GetValueOrDefault("content")
                          ?? properties.GetValueOrDefault("text");
        byte[]? fileBytes = null;
        string? srcFile = null;
        if (string.IsNullOrEmpty(payload)
            && (properties.TryGetValue("src", out srcFile) || properties.TryGetValue("path", out srcFile))
            && !string.IsNullOrWhiteSpace(srcFile))
        {
            if (!System.IO.File.Exists(srcFile))
                throw new ArgumentException($"htmlchunk source file not found: '{srcFile}'.");
            fileBytes = System.IO.File.ReadAllBytes(srcFile);
        }
        if (string.IsNullOrWhiteSpace(payload) && (fileBytes == null || fileBytes.Length == 0))
            throw new ArgumentException("htmlchunk requires inline 'html' content (aliases: content, text) or a 'src' file path.");

        var format = ResolveHtmlChunkFormat(properties.GetValueOrDefault("format"), srcFile);

        byte[] bytes;
        if (fileBytes != null)
        {
            // A file is stored verbatim (its own charset/BOM/<meta> apply),
            // except a bare HTML fragment, which gets the same document
            // wrapper as inline input so Word imports it as UTF-8.
            if (format.Kind is "html" && !LooksLikeHtmlDocument(DecodeForSniff(fileBytes)))
                bytes = EncodeHtmlDocument(WrapHtmlFragment(DecodeForSniff(fileBytes), properties.GetValueOrDefault("title")));
            else
                bytes = fileBytes;
        }
        else
        {
            var text = payload!;
            bytes = format.Kind switch
            {
                "html" => EncodeHtmlDocument(LooksLikeHtmlDocument(text)
                    ? text
                    : WrapHtmlFragment(text, properties.GetValueOrDefault("title"))),
                // Text-ish formats: UTF-8 with BOM so Word does not fall back
                // to the ANSI code page.
                _ => Encoding.UTF8.GetPreamble().Concat(Encoding.UTF8.GetBytes(text)).ToArray(),
            };
        }

        var chunkPart = mainPart.AddAlternativeFormatImportPart(format.PartType);
        using (var ms = new System.IO.MemoryStream(bytes))
            chunkPart.FeedData(ms);
        var relId = mainPart.GetIdOfPart(chunkPart);

        var altChunk = new AltChunk { Id = relId };
        if (properties.TryGetValue("matchSrc", out var matchSrcRaw) && IsTruthy(matchSrcRaw))
            altChunk.AltChunkProperties = new AltChunkProperties(new MatchSource());

        InsertAtIndexOrAppend(parent, altChunk, index);

        // CT_Tc must end with a paragraph — Word reports the file as corrupt
        // when a cell's last block is an altChunk. Keep an (empty) trailing
        // paragraph after a chunk that landed in last position.
        if (parent is TableCell && altChunk.NextSibling() == null)
            altChunk.InsertAfterSelf(new Paragraph());

        var siblings = parent.Elements<AltChunk>().ToList();
        var idx = siblings.IndexOf(altChunk) + 1;
        return $"{parentPath}/altChunk[{idx}]";
    }

    /// <summary>
    /// Dump support: the payload behind a body-level <c>altChunk</c> at
    /// <paramref name="path"/>, decoded to text, with the <c>format</c>
    /// value `add htmlchunk` accepts. Null when the element is missing or
    /// its payload is not a text format we can replay (e.g. an embedded
    /// .docx chunk).
    /// </summary>
    internal (string Content, string Format, bool MatchSrc)? ReadAltChunkForDump(string path)
    {
        try
        {
            if (NavigateToElement(ParsePath(path)) is not AltChunk altChunk) return null;
            var part = TryGetAltChunkPart(altChunk);
            var format = part == null ? null : AltChunkFormatOf(part.ContentType);
            if (part == null || format == null) return null;
            using var stream = part.GetStream();
            var content = OfficeCli.Core.CharsetDecoder.Decode(
                stream, OfficeCli.Core.CharsetDecoder.ParseCharset(part.ContentType));
            var matchSrc = altChunk.AltChunkProperties?.GetFirstChild<MatchSource>() is { } ms
                           && (ms.Val == null || ms.Val.Value);
            return (content, format, matchSrc);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Package URIs of the payload parts that <see cref="ReadAltChunkForDump"/>
    /// can replay (body-level chunks only). The dump's aux-part scan skips
    /// these instead of reporting them as dropped.
    /// </summary>
    internal HashSet<string> GetReplayableAltChunkPartUris()
    {
        var uris = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return uris;
        foreach (var altChunk in body.Elements<AltChunk>())
        {
            var part = TryGetAltChunkPart(altChunk);
            if (part != null && AltChunkFormatOf(part.ContentType) != null)
                uris.Add(part.Uri.OriginalString);
        }
        return uris;
    }

    /// <summary>
    /// Get readback for <c>/body/altChunk[N]</c>: relationship id, payload
    /// format / content type, size, and matchSrc.
    /// </summary>
    private DocumentNode AltChunkToNode(AltChunk altChunk, DocumentNode node)
    {
        node.Type = "altChunk";
        if (altChunk.Id?.Value is { Length: > 0 } rId)
            node.Format["id"] = rId;
        var part = TryGetAltChunkPart(altChunk);
        if (part != null)
        {
            node.Format["contentType"] = part.ContentType;
            if (AltChunkFormatOf(part.ContentType) is { } fmt)
                node.Format["format"] = fmt;
            node.Format["partUri"] = part.Uri.OriginalString;
            try
            {
                using var stream = part.GetStream();
                node.Format["size"] = stream.Length;
            }
            catch { /* unreadable payload: omit size */ }
        }
        if (altChunk.AltChunkProperties?.GetFirstChild<MatchSource>() is { } ms
            && (ms.Val == null || ms.Val.Value))
            node.Format["matchSrc"] = true;
        return node;
    }

    private AlternativeFormatImportPart? TryGetAltChunkPart(AltChunk altChunk)
    {
        var rId = altChunk.Id?.Value;
        if (string.IsNullOrEmpty(rId)) return null;
        try { return _doc.MainDocumentPart?.GetPartById(rId) as AlternativeFormatImportPart; }
        catch (ArgumentOutOfRangeException) { return null; }
    }

    private static string? AltChunkFormatOf(string? contentType)
    {
        var mediaType = (contentType ?? "").Split(';', 2)[0].Trim().ToLowerInvariant();
        return mediaType switch
        {
            "text/html" => "html",
            "application/xhtml+xml" => "xhtml",
            "message/rfc822" or "multipart/related" => "mht",
            "application/rtf" or "text/rtf" => "rtf",
            "text/plain" => "text",
            _ => null,
        };
    }

    private readonly record struct HtmlChunkFormat(string Kind, PartTypeInfo PartType);

    private static HtmlChunkFormat ResolveHtmlChunkFormat(string? format, string? srcFile)
    {
        var key = format?.Trim().ToLowerInvariant();
        if (string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(srcFile))
            key = System.IO.Path.GetExtension(srcFile).TrimStart('.').ToLowerInvariant();
        return key switch
        {
            null or "" or "html" or "htm" => new("html", AlternativeFormatImportPartType.Html),
            "xhtml" or "xht" => new("xhtml", AlternativeFormatImportPartType.Xhtml),
            "mht" or "mhtml" => new("mht", AlternativeFormatImportPartType.Mht),
            "rtf" => new("rtf", AlternativeFormatImportPartType.Rtf),
            "text" or "txt" or "plain" => new("text", AlternativeFormatImportPartType.TextPlain),
            // An unknown extension on src= is most likely still HTML; an
            // explicit unknown format= is a caller error.
            _ when format == null => new("html", AlternativeFormatImportPartType.Html),
            _ => throw new ArgumentException(
                $"Invalid htmlchunk format '{format}'. Valid values: html, xhtml, mht, rtf, text."),
        };
    }

    private static bool LooksLikeHtmlDocument(string text)
        => System.Text.RegularExpressions.Regex.IsMatch(text, @"<\s*(!doctype\s+html|html[\s>])",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

    private static string DecodeForSniff(byte[] bytes)
        => OfficeCli.Core.CharsetDecoder.Decode(new System.IO.MemoryStream(bytes), null);

    private static string WrapHtmlFragment(string fragment, string? title)
    {
        var sb = new StringBuilder();
        sb.Append("<!DOCTYPE html>\n<html>\n<head>\n<meta charset=\"utf-8\">\n");
        if (!string.IsNullOrEmpty(title))
            sb.Append("<title>").Append(System.Net.WebUtility.HtmlEncode(title)).Append("</title>\n");
        sb.Append("</head>\n<body>\n").Append(fragment).Append("\n</body>\n</html>\n");
        return sb.ToString();
    }

    // UTF-8 with BOM: Word's HTML importer honours the BOM even when the
    // payload has no <meta charset> (it otherwise assumes the ANSI code page).
    private static byte[] EncodeHtmlDocument(string html)
        => Encoding.UTF8.GetPreamble().Concat(Encoding.UTF8.GetBytes(html)).ToArray();
}
