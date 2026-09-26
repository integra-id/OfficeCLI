// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildFinalizeCommand(Option<bool> jsonOption)
    {
        var fileArg = new Argument<FileInfo>("file") { Description = "Word document (.docx or .docm)" };
        var noMaterialize = new Option<bool>("--no-materialize")
        {
            Description = "Skip altChunk materialization."
        };
        var noToc = new Option<bool>("--no-toc")
        {
            Description = "Skip the TOC entry rebuild. By default, refresh --toc runs when the file contains a TOC field. PAGEREF page numbers stay the placeholder 0."
        };
        var noPageSetup = new Option<bool>("--no-page-setup")
        {
            Description = "Skip the page-setup step. By default, pageSetup=a4-moderate is applied only to sections that have no w:pgSz. Sections that already have a page size, including the A4 size officecli create writes, are left unchanged."
        };
        var noValidate = new Option<bool>("--no-validate")
        {
            Description = "Skip the OpenXML schema check."
        };
        var strictOpt = new Option<bool>("--strict")
        {
            Description = "Passed through to materialize. If any altChunk cannot be converted (RTF, MHT, a missing part, or a chunk outside the body), leave the file unchanged, skip the later steps, and exit 1. A picture that cannot be embedded is a warning and does not trip --strict."
        };

        var cmd = new Command("finalize",
            "Finish a technical Word document in one headless pass (.docx/.docm). Default steps, in order: " +
            "(1) materialize HTML/XHTML/plain-text altChunks — RTF and MHT stay, same as materialize; " +
            "(2) refresh --toc when a TOC field exists — titles and in-document hyperlinks are rebuilt, PAGEREF page numbers stay the placeholder 0, no Microsoft Word and no browser; skipped when there is no TOC field; " +
            "(3) page setup — apply pageSetup=a4-moderate (alias tech-doc) only to sections with no w:pgSz. This does not overwrite an existing page size, so officecli create (which already stamps A4) keeps its margins; a section with margins but no w:pgSz receives the preset's Moderate margins as well; " +
            "(4) validate against the OpenXML schema. Validation errors are reported and the command exits 1; earlier steps are kept. " +
            "If materialize --strict refuses, or the TOC rebuild fails, later steps are not run and the command exits 1. " +
            "Skip a step with --no-materialize, --no-toc, --no-page-setup, or --no-validate.");
        cmd.Add(fileArg);
        cmd.Add(noMaterialize);
        cmd.Add(noToc);
        cmd.Add(noPageSetup);
        cmd.Add(noValidate);
        cmd.Add(strictOpt);
        cmd.Add(jsonOption);

        cmd.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(fileArg)!;
            var options = new WordFinalize.Options
            {
                Materialize = !result.GetValue(noMaterialize),
                Toc = !result.GetValue(noToc),
                PageSetup = !result.GetValue(noPageSetup),
                Validate = !result.GetValue(noValidate),
                Strict = result.GetValue(strictOpt),
            };

            var ext = Path.GetExtension(file.FullName).ToLowerInvariant();
            if (ext is not ".docx" and not ".docm")
                throw new CliException($"finalize currently only supports .docx/.docm files (got {ext}).")
                { Code = "unsupported_type" };

            if (TryResident(file.FullName, req =>
            {
                req.Command = "finalize";
                req.Args["materialize"] = options.Materialize ? "true" : "false";
                req.Args["toc"] = options.Toc ? "true" : "false";
                req.Args["pageSetup"] = options.PageSetup ? "true" : "false";
                req.Args["validate"] = options.Validate ? "true" : "false";
                if (options.Strict) req.Args["strict"] = "true";
            }, json) is { } rc) return rc;

            var report = WordFinalize.Run(file.FullName, options);
            WordFinalize.Write(report, json);
            return report.ExitCode;
        }, json); });

        return cmd;
    }
}
