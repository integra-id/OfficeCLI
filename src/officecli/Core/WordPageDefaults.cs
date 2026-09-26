// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Single source of truth for Word default page geometry (twips).
/// Used as fallback when a section's pgSz/pgMar is missing — callers
/// must always read the source <c>SectionProperties</c> first and only
/// drop to these defaults when the value is genuinely absent.
/// </summary>
public static class WordPageDefaults
{
    // A4: 210mm × 297mm at 1440 twips/inch (= 567 twips/cm).
    public const uint A4WidthTwips = 11906;
    public const uint A4HeightTwips = 16838;

    // OOXML legal range for w:pgSz/@w:w and @w:h. Word's UI clamps roughly to
    // ~0.4cm–55.9cm; the EcmaSpec defines 1..31680 (22"). Use 240 (1/6") as the
    // lower bound — anything smaller will not produce a renderable page in Word.
    public const uint PageDimMinTwips = 240;
    public const uint PageDimMaxTwips = 31680;

    public static void ValidatePageDim(long twips, string keyName)
    {
        if (twips < PageDimMinTwips || twips > PageDimMaxTwips)
            throw new ArgumentException(
                $"{keyName} must be in range {PageDimMinTwips}–{PageDimMaxTwips} twips " +
                $"(~0.4cm–55.9cm), got {twips}.");
    }

    /// <summary>Named page size in twips (portrait). <c>a4</c> matches <see cref="A4WidthTwips"/>.</summary>
    public readonly record struct PageSizePreset(uint WidthTwips, uint HeightTwips);

    /// <summary>
    /// Named body-margin preset in twips. Header/footer distances are not part of
    /// the preset — Word's margin gallery only changes the four body edges.
    /// </summary>
    public readonly record struct MarginPreset(int TopTwips, int BottomTwips, uint LeftTwips, uint RightTwips);

    // US Letter: 8.5in × 11in. 1 inch = 1440 twips.
    public const uint LetterWidthTwips = 12240;
    public const uint LetterHeightTwips = 15840;

    // Word ribbon "Moderate": top/bottom 1in (25.4mm), left/right 0.75in (19.05mm).
    public const int ModerateMarginVerticalTwips = 1440;
    public const uint ModerateMarginHorizontalTwips = 1080;

    public const string PageSizePresetList = "a4, letter";
    public const string MarginPresetList = "normal, narrow, moderate, wide";
    public const string PageSetupPresetList = "a4-moderate (aliases: tech-doc, techdoc)";

    public static bool TryGetPageSizePreset(string name, out PageSizePreset preset)
    {
        switch (NormalizePresetKey(name))
        {
            case "a4":
                preset = new PageSizePreset(A4WidthTwips, A4HeightTwips);
                return true;
            case "letter":
                preset = new PageSizePreset(LetterWidthTwips, LetterHeightTwips);
                return true;
            default:
                preset = default;
                return false;
        }
    }

    public static bool TryGetMarginPreset(string name, out MarginPreset preset)
    {
        switch (NormalizePresetKey(name))
        {
            case "normal":
                preset = new MarginPreset(1440, 1440, 1440, 1440); // 1in all
                return true;
            case "narrow":
                preset = new MarginPreset(720, 720, 720, 720); // 0.5in all
                return true;
            case "moderate":
                preset = new MarginPreset(
                    ModerateMarginVerticalTwips, ModerateMarginVerticalTwips,
                    ModerateMarginHorizontalTwips, ModerateMarginHorizontalTwips);
                return true;
            case "wide":
                preset = new MarginPreset(1440, 1440, 2880, 2880); // 1in top/bottom, 2in left/right
                return true;
            default:
                preset = default;
                return false;
        }
    }

    /// <summary>
    /// Combined page-setup preset: paper size plus body margins.
    /// <c>a4-moderate</c> / <c>tech-doc</c> is A4 + Word Moderate margins.
    /// </summary>
    public static bool TryGetPageSetupPreset(string name, out PageSizePreset size, out MarginPreset margins)
    {
        switch (NormalizePresetKey(name))
        {
            case "a4moderate":
            case "techdoc":
                size = new PageSizePreset(A4WidthTwips, A4HeightTwips);
                margins = new MarginPreset(
                    ModerateMarginVerticalTwips, ModerateMarginVerticalTwips,
                    ModerateMarginHorizontalTwips, ModerateMarginHorizontalTwips);
                return true;
            default:
                size = default;
                margins = default;
                return false;
        }
    }

    /// <summary>Case-insensitive preset key with hyphens, underscores, and spaces stripped.</summary>
    public static string NormalizePresetKey(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return "";
        var buffer = new char[name.Length];
        var n = 0;
        foreach (var c in name.Trim())
        {
            if (c is '-' or '_' or ' ') continue;
            buffer[n++] = char.ToLowerInvariant(c);
        }
        return new string(buffer, 0, n);
    }
}
