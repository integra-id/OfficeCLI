// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildRefreshCommand(Option<bool> jsonOption)
    {
        var fileArg = new Argument<FileInfo>("file") { Description = "Office document path" };
        var tocOpt = new Option<bool>("--toc")
        {
            Description = "Rebuild TOC entries from the current headings (titles and in-document hyperlinks) without Microsoft Word or a browser. PAGEREF page numbers are written as the placeholder 0; previously resolved page numbers are not kept. Omit this flag to try Word on Windows, then HTML pagination when a headless browser exists. If neither can number pages, entries are still rebuilt and page numbers stay 0."
        };

        var cmd = new Command("refresh",
            "Recalculate derived fields. --toc rebuilds TOC entries from headings without Word (page numbers stay 0). The default tries Word on Windows, then HTML pagination; if neither can number pages, TOC entries are still rebuilt.");
        cmd.Add(fileArg);
        cmd.Add(tocOpt);
        cmd.Add(jsonOption);

        cmd.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(fileArg)!;
            var tocOnly = result.GetValue(tocOpt);

            if (TryResident(file.FullName, req =>
            {
                req.Command = "refresh";
                req.Json = json;
                if (tocOnly) req.Args["toc"] = "true";
            }, json) is { } rc) return rc;

            var ext = Path.GetExtension(file.FullName).ToLowerInvariant();
            if (ext != ".docx" && ext != ".docm")
                throw new CliException($"refresh currently only supports .docx files (got {ext}).")
                { Code = "unsupported_type" };

            var outcome = WordHtmlRefresh.Refresh(file.FullName, tocOnly);
            if (!outcome.Ok)
                throw new CliException(outcome.Message)
                { Code = "refresh_failed" };

            if (json) Console.WriteLine(OutputFormatter.WrapEnvelope(outcome.ToJson()));
            else Console.WriteLine(outcome.Message);
            return 0;
        }, json); });

        return cmd;
    }
}
