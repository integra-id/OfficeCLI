// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    /// <summary>
    /// Outcome of <see cref="EnsurePageSetupWhereMissing"/>.
    /// <see cref="PageSetupEnsureReport.Changed"/> is false when every section
    /// already has a <c>w:pgSz</c> — including the A4 size <c>officecli create</c>
    /// stamps — so margins on those sections are left alone.
    /// </summary>
    public sealed class PageSetupEnsureReport
    {
        public bool Changed { get; init; }
        public string Preset { get; init; } = "";
        /// <summary>1-based <c>/section[N]</c> paths that received the preset.</summary>
        public IReadOnlyList<string> Applied { get; init; } = Array.Empty<string>();
        /// <summary>Sections skipped because <c>w:pgSz</c> was already present.</summary>
        public IReadOnlyList<string> Unchanged { get; init; } = Array.Empty<string>();
    }

    /// <summary>
    /// Apply a named <c>pageSetup</c> preset to sections that have no page size.
    /// Sections that already declare <c>w:pgSz</c> (width or height) are not
    /// modified. The preset writes both paper size and its body margins, so a
    /// section that had margins but no <c>w:pgSz</c> receives the preset margins
    /// too. Uses <see cref="ApplyNamedPagePresets"/> — the same path as
    /// <c>set /section[N] --prop pageSetup=…</c>.
    /// </summary>
    public PageSetupEnsureReport EnsurePageSetupWhereMissing(string preset)
    {
        if (!WordPageDefaults.TryGetPageSetupPreset(preset, out _, out _))
            throw new ArgumentException(
                $"Unknown pageSetup preset: '{preset}'. Valid: {WordPageDefaults.PageSetupPresetList}.");

        var sections = FindSectionProperties();
        var applied = new List<string>();
        var unchanged = new List<string>();
        var props = new Dictionary<string, string> { ["pageSetup"] = preset };
        for (int i = 0; i < sections.Count; i++)
        {
            var path = $"/section[{i + 1}]";
            if (SectionHasPageSize(sections[i]))
            {
                unchanged.Add(path);
                continue;
            }
            ApplyNamedPagePresets(sections[i], props);
            applied.Add(path);
        }

        if (applied.Count == 0)
            return new PageSetupEnsureReport { Changed = false, Preset = preset, Applied = applied, Unchanged = unchanged };

        Modified = true;
        SaveDoc();
        return new PageSetupEnsureReport { Changed = true, Preset = preset, Applied = applied, Unchanged = unchanged };
    }

    /// <summary>True when the section declares a paper width or height.</summary>
    static bool SectionHasPageSize(SectionProperties sectPr)
    {
        var pgSz = sectPr.GetFirstChild<PageSize>();
        if (pgSz == null) return false;
        return pgSz.Width != null || pgSz.Height != null;
    }
}
