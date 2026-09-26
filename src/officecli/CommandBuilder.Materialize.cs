// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildMaterializeCommand(Option<bool> jsonOption)
    {
        var fileArg = new Argument<FileInfo>("file") { Description = "Word document (.docx or .docm)" };
        var strictOpt = new Option<bool>("--strict")
        {
            Description = "Leave the file unchanged and exit 1 if any altChunk cannot be converted (RTF, MHT, a missing part, or a chunk outside the body). Without --strict, HTML/XHTML/plain-text chunks are still materialized and the rest stay in place. A picture that cannot be embedded is a warning in both modes: it does not trip --strict, and the rest of that chunk is still converted."
        };

        var cmd = new Command("materialize",
            "Replace Word altChunk (htmlchunk) parts with native paragraphs, lists, tables and links. Headless — does not run Microsoft Word. " +
            "Converts HTML, XHTML and plain text. Subset: headings, paragraphs, bold/italic/underline/strike, color, font size and family, sub/sup, hyperlinks, ul/ol, tables with colspan/rowspan, and inline pictures. " +
            "A data-URI image whose type the picture pipeline already accepts (png, jpeg, gif, bmp, tiff, emf, wmf) becomes an inline w:drawing; the alt attribute is the picture description. " +
            "webp, svg, video and iframe are not embedded. http(s) URLs are not downloaded and relative paths are not resolved — those images stay as alt text and do not fail the conversion. " +
            "A data URI that cannot be decoded is the same kind of warning (alt text kept); it does not trip --strict. " +
            "A small CSS subset is honored (element, class, id, descendant and child selectors; text-align, background-color, margin, font). " +
            "Scripts are dropped, and RTF/MHT chunks are left unchanged. " +
            "Formatting is written directly onto runs (matchSrc-style), and h1–h6 also reference Heading styles. " +
            "Fidelity is not Word's HTML importer — float, flex, borders and media queries are not reproduced.");
        cmd.Add(fileArg);
        cmd.Add(strictOpt);
        cmd.Add(jsonOption);

        cmd.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(fileArg)!;
            var strict = result.GetValue(strictOpt);

            var ext = Path.GetExtension(file.FullName).ToLowerInvariant();
            if (ext is not ".docx" and not ".docm")
                throw new CliException($"materialize currently only supports .docx/.docm files (got {ext}).")
                { Code = "unsupported_type" };

            if (TryResident(file.FullName, req =>
            {
                req.Command = "materialize";
                if (strict) req.Args["strict"] = "true";
            }, json) is { } rc) return rc;

            using var handler = DocumentHandlerFactory.Open(file.FullName, editable: true);
            if (handler is not WordHandler word)
                throw new CliException("materialize currently only supports .docx/.docm files.")
                { Code = "unsupported_type" };

            var report = word.MaterializeAltChunks(strict);
            // A no-op (nothing to do, or --strict refused) must not rewrite the
            // package on dispose. A real conversion is flushed by Dispose
            // because MaterializeAltChunks raised Modified.
            if (report.Converted == 0)
                word.DiscardOnDispose = true;
            if (report.Refused)
                throw new CliException(report.Summary + (report.Warnings.Count > 0 ? " " + string.Join("; ", report.Warnings) : ""))
                {
                    Code = "altchunk_skipped",
                    Help = "officecli materialize --help",
                };

            WriteMaterializeReport(report, json);
            return 0;
        }, json); });

        return cmd;
    }

    internal static void WriteMaterializeReport(WordHandler.AltChunkMaterializeReport report, bool json)
    {
        if (!json)
        {
            foreach (var w in report.Warnings)
                Console.Error.WriteLine("warning: " + w);
            Console.WriteLine(report.Summary);
            return;
        }

        List<CliWarning>? warnings = null;
        if (report.Warnings.Count > 0)
        {
            warnings = new List<CliWarning>(report.Warnings.Count);
            foreach (var w in report.Warnings)
                warnings.Add(new CliWarning { Code = "altchunk_fidelity", Message = w });
        }
        Console.WriteLine(OutputFormatter.WrapEnvelope(report.ToJson(), warnings));
    }
}
