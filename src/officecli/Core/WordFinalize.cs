// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using OfficeCli.Handlers;

namespace OfficeCli.Core;

/// <summary>
/// Headless post-build pass for a technical Word document. Composes the
/// existing materialize, <c>refresh --toc</c>, page-setup preset, and
/// validate operations. It does not calculate real PAGEREF page numbers.
/// </summary>
internal static class WordFinalize
{
    internal sealed class Options
    {
        public bool Materialize { get; init; } = true;
        public bool Toc { get; init; } = true;
        public bool PageSetup { get; init; } = true;
        public bool Validate { get; init; } = true;
        public bool Strict { get; init; }
        /// <summary>Named pageSetup preset. Only <c>a4-moderate</c> / <c>tech-doc</c> exist today.</summary>
        public string PageSetupPreset { get; init; } = "a4-moderate";
    }

    internal sealed class Step
    {
        public required string Name { get; init; }
        /// <summary><c>ran</c>, <c>skipped</c>, <c>failed</c>, or <c>not-run</c>.</summary>
        public required string Status { get; init; }
        public required string Message { get; init; }
        /// <summary>Raw JSON object for this step, or null.</summary>
        public string? DetailJson { get; init; }
        public List<ValidationError>? ValidationErrors { get; init; }
    }

    internal sealed class Report
    {
        public List<Step> Steps { get; } = new();
        public List<string> Warnings { get; } = new();
        public int ExitCode { get; set; }
        public bool Success => ExitCode == 0;

        public string ToJson()
        {
            var sb = new StringBuilder();
            sb.Append("{\"steps\":[");
            for (int i = 0; i < Steps.Count; i++)
            {
                if (i > 0) sb.Append(',');
                var step = Steps[i];
                sb.Append("{\"name\":\"").Append(JsonEscape(step.Name)).Append('"');
                sb.Append(",\"status\":\"").Append(JsonEscape(step.Status)).Append('"');
                sb.Append(",\"message\":\"").Append(JsonEscape(step.Message)).Append('"');
                if (step.DetailJson != null)
                    sb.Append(",\"detail\":").Append(step.DetailJson);
                sb.Append('}');
            }
            sb.Append("]}");
            return sb.ToString();
        }
    }

    public static Report Run(string path, Options options)
    {
        var report = new Report();

        if (!options.Materialize)
            report.Steps.Add(Make("materialize", "skipped", "Skipped (--no-materialize)."));
        else
        {
            var step = RunMaterialize(path, options.Strict, report.Warnings);
            report.Steps.Add(step);
            if (step.Status == "failed")
            {
                AddRemaining(report, options, "materialize");
                report.ExitCode = 1;
                return report;
            }
        }

        if (!options.Toc)
            report.Steps.Add(Make("toc", "skipped", "Skipped (--no-toc)."));
        else
        {
            var step = RunToc(path);
            report.Steps.Add(step);
            if (step.Status == "failed")
            {
                AddRemaining(report, options, "toc");
                report.ExitCode = 1;
                return report;
            }
        }

        if (!options.PageSetup)
            report.Steps.Add(Make("pageSetup", "skipped", "Skipped (--no-page-setup)."));
        else
            report.Steps.Add(RunPageSetup(path, options.PageSetupPreset));

        if (!options.Validate)
            report.Steps.Add(Make("validate", "skipped", "Skipped (--no-validate)."));
        else
        {
            var step = RunValidate(path);
            report.Steps.Add(step);
            if (step.Status == "failed")
                report.ExitCode = 1;
        }

        return report;
    }

    public static void Write(Report report, bool json)
    {
        if (!json)
        {
            foreach (var w in report.Warnings)
                Console.Error.WriteLine("warning: " + w);
            foreach (var step in report.Steps)
                Console.WriteLine($"{step.Name}: {step.Message}");
            var validate = report.Steps.Find(s => s.Name == "validate");
            if (validate?.ValidationErrors is { Count: > 0 } errors)
            {
                Console.Error.WriteLine($"Found {errors.Count} validation error(s):");
                foreach (var err in errors)
                {
                    Console.Error.WriteLine($"  [{err.ErrorType}] {err.Description}");
                    if (err.Path != null) Console.Error.WriteLine($"    Path: {err.Path}");
                    if (err.Part != null) Console.Error.WriteLine($"    Part: {err.Part}");
                }
            }
            Console.WriteLine(report.ExitCode == 0 ? "finalize: ok" : "finalize: failed");
            return;
        }

        List<CliWarning>? warnings = null;
        if (report.Warnings.Count > 0)
        {
            warnings = new List<CliWarning>(report.Warnings.Count);
            foreach (var w in report.Warnings)
                warnings.Add(new CliWarning { Code = "altchunk_fidelity", Message = w });
        }
        Console.WriteLine(OutputFormatter.WrapEnvelope(report.ToJson(), warnings, success: report.ExitCode == 0));
    }

    static Step RunMaterialize(string path, bool strict, List<string> warnings)
    {
        using var handler = DocumentHandlerFactory.Open(path, editable: true);
        if (handler is not WordHandler word)
            throw new CliException("finalize currently only supports .docx/.docm files.")
            { Code = "unsupported_type" };

        var mat = word.MaterializeAltChunks(strict);
        // A no-op or a --strict refusal must not rewrite the package.
        if (mat.Refused || mat.Converted == 0)
            word.DiscardOnDispose = true;
        warnings.AddRange(mat.Warnings);

        if (mat.Refused)
        {
            return Make("materialize", "failed", mat.Summary, mat.ToJson());
        }
        return Make("materialize", "ran", mat.Summary, mat.ToJson());
    }

    static Step RunToc(string path)
    {
        var outcome = WordHtmlRefresh.Refresh(path, tocOnly: true);
        if (!outcome.Ok)
            return Make("toc", "failed", outcome.Message, outcome.ToJson());
        // refresh --toc succeeds with "nothing to rebuild" when the file has
        // no TOC field. finalize treats that as a skip, not as work done.
        if (outcome.TocFields == 0)
            return Make("toc", "skipped", "Skipped - no TOC field. " + outcome.Message, outcome.ToJson());
        return Make("toc", "ran", outcome.Message, outcome.ToJson());
    }

    static Step RunPageSetup(string path, string preset)
    {
        using var handler = DocumentHandlerFactory.Open(path, editable: true);
        if (handler is not WordHandler word)
            throw new CliException("finalize currently only supports .docx/.docm files.")
            { Code = "unsupported_type" };

        var ensured = word.EnsurePageSetupWhereMissing(preset);
        if (!ensured.Changed)
            word.DiscardOnDispose = true;

        var detail = PageSetupDetailJson(ensured);
        if (!ensured.Changed)
        {
            return Make("pageSetup", "skipped",
                $"Skipped - every section already has a page size (w:pgSz). {preset} was not applied.",
                detail);
        }

        var msg = $"Applied {preset} to {string.Join(", ", ensured.Applied)} (no w:pgSz).";
        if (ensured.Unchanged.Count > 0)
            msg += $" Left {string.Join(", ", ensured.Unchanged)} unchanged (page size already set).";
        return Make("pageSetup", "ran", msg, detail);
    }

    static Step RunValidate(string path)
    {
        using var handler = DocumentHandlerFactory.Open(path);
        var errors = handler.Validate();
        var detail = ValidationDetailJson(errors);
        if (errors.Count == 0)
            return Make("validate", "ran", "Validation passed: no errors found.", detail);
        return new Step
        {
            Name = "validate",
            Status = "failed",
            Message = $"Found {errors.Count} validation error(s).",
            DetailJson = detail,
            ValidationErrors = errors,
        };
    }

    static void AddRemaining(Report report, Options options, string stoppedAfter)
    {
        void Rest(string name, bool enabled, string skipMessage)
        {
            if (report.Steps.Exists(s => s.Name == name)) return;
            report.Steps.Add(enabled
                ? Make(name, "not-run", $"Not run - stopped after {stoppedAfter}.")
                : Make(name, "skipped", skipMessage));
        }
        Rest("toc", options.Toc, "Skipped (--no-toc).");
        Rest("pageSetup", options.PageSetup, "Skipped (--no-page-setup).");
        Rest("validate", options.Validate, "Skipped (--no-validate).");
    }

    static Step Make(string name, string status, string message, string? detailJson = null)
        => new() { Name = name, Status = status, Message = message, DetailJson = detailJson };

    static string PageSetupDetailJson(WordHandler.PageSetupEnsureReport ensured)
    {
        var sb = new StringBuilder();
        sb.Append("{\"preset\":\"").Append(JsonEscape(ensured.Preset)).Append('"');
        sb.Append(",\"applied\":").Append(JsonStringArray(ensured.Applied));
        sb.Append(",\"unchanged\":").Append(JsonStringArray(ensured.Unchanged));
        sb.Append('}');
        return sb.ToString();
    }

    static string ValidationDetailJson(List<ValidationError> errors)
    {
        var sb = new StringBuilder();
        sb.Append("{\"count\":").Append(errors.Count).Append(",\"errors\":[");
        for (int i = 0; i < errors.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var e = errors[i];
            sb.Append("{\"type\":\"").Append(JsonEscape(e.ErrorType)).Append('"');
            sb.Append(",\"description\":\"").Append(JsonEscape(e.Description)).Append('"');
            if (e.Path != null) sb.Append(",\"path\":\"").Append(JsonEscape(e.Path)).Append('"');
            if (e.Part != null) sb.Append(",\"part\":\"").Append(JsonEscape(e.Part)).Append('"');
            sb.Append('}');
        }
        sb.Append("]}");
        return sb.ToString();
    }

    static string JsonStringArray(IReadOnlyList<string> values)
    {
        var sb = new StringBuilder();
        sb.Append('[');
        for (int i = 0; i < values.Count; i++)
        {
            if (i > 0) sb.Append(',');
            sb.Append('"').Append(JsonEscape(values[i])).Append('"');
        }
        sb.Append(']');
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
