// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // ID-column width rule shared with the technical-document skill:
    // max(18mm, longestLine × 2.0mm + 8mm). 2.0mm is about one mono glyph
    // plus a chip; 8mm covers cell padding; 18mm is the floor for a short
    // header such as "ID". A Latin character counts as 1; a full-width
    // character counts as ~1.82 (EstimateTextWidthInChars). 6 units → 1134 twips.
    private const double IdColumnTwipsPerMm = 1440.0 / 25.4;

    private readonly record struct ColumnIdFlags(bool? NoWrap, bool Fit, int? ExplicitTwips);

    /// <summary>
    /// Virtual <c>/…/tbl[N]/col[C]</c> path. OOXML has no <c>w:col</c>, so
    /// the generic navigator cannot resolve it. Returns false when
    /// <paramref name="path"/> is not a column path, or when the parent
    /// exists but is not a table (the caller keeps the generic error).
    /// Throws when the path is a column path and that table is missing.
    /// </summary>
    private bool TryResolveVirtualTableColumn(string path, out Table table, out int column1Based)
    {
        table = null!;
        column1Based = 0;
        var match = Regex.Match(
            path,
            @"^(?<parent>.+)/col\[(?<idx>\d+)\]$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        if (!match.Success) return false;

        var parent = NavigateToElement(ParsePath(match.Groups["parent"].Value), out var ctx);
        if (parent is not Table found)
        {
            if (parent == null)
                throw new ArgumentException($"Path not found: {path}" + (ctx != null ? $". {ctx}" : ""));
            return false;
        }

        table = found;
        column1Based = int.Parse(match.Groups["idx"].Value, CultureInfo.InvariantCulture);
        return true;
    }

    private List<string> SetTableColumn(Table table, int column1Based, Dictionary<string, string> properties)
    {
        var flags = ReadColumnIdFlags(properties);
        var unsupported = new List<string>();
        foreach (var key in properties.Keys)
        {
            var lower = key.ToLowerInvariant();
            if (lower is "idcolumn" or "nowrap" or "width") continue;
            unsupported.Add(key);
        }

        if (flags.NoWrap == null && !flags.Fit && flags.ExplicitTwips == null)
            return unsupported;

        var backup = (Table)table.CloneNode(true);
        try
        {
            ApplyColumnIdTreatment(table, column1Based, flags);
            var affected = table.Ancestors<Paragraph>().FirstOrDefault();
            if (affected != null)
                affected.TextId = GenerateParaId();
            SaveDoc();
            return unsupported;
        }
        catch
        {
            table.Parent?.ReplaceChild(backup, table);
            throw;
        }
    }

    private DocumentNode GetTableColumn(Table table, int column1Based, string path)
    {
        var cols = RequireGridColumns(table);
        if (column1Based < 1 || column1Based > cols.Count)
            throw new ArgumentException($"Column {column1Based} not found (total: {cols.Count})");

        int slot = column1Based - 1;
        var addressable = AddressableCells(table, slot);
        var node = new DocumentNode
        {
            Path = path,
            Type = "column",
        };
        int gridTwips = 0;
        if (int.TryParse(cols[slot].Width?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out gridTwips)
            && gridTwips > 0)
            node.Format["width"] = gridTwips.ToString(CultureInfo.InvariantCulture) + "dxa";

        if (addressable.Count > 0 && addressable.All(c => IsToggleOn(c.TableCellProperties?.NoWrap)))
        {
            node.Format["nowrap"] = true;
            if (gridTwips > 0 && gridTwips == FitIdColumnWidthTwips(addressable.Max(LongestLineUnits)))
                node.Format["idColumn"] = true;
        }
        return node;
    }

    /// <summary>
    /// Table-level <c>idColumn</c> / <c>idColumns</c>: a 1-based index or a
    /// comma-separated list. Each listed column gets <c>w:noWrap</c> and a
    /// width that fits its longest line. A boolean is rejected here — that
    /// form belongs on the column path.
    /// </summary>
    private void ApplyTableIdColumnSpec(Table table, string spec)
    {
        var columns = ParseIdColumnList(table, spec);
        foreach (var column in columns)
            ApplyColumnIdTreatment(table, column, new ColumnIdFlags(NoWrap: true, Fit: true, ExplicitTwips: null));
    }

    private ColumnIdFlags ReadColumnIdFlags(Dictionary<string, string> properties)
    {
        bool sawId = false;
        bool idOn = false;
        if (properties.TryGetValue("idColumn", out var idRaw))
        {
            if (!ParseHelpers.IsValidBooleanString(idRaw))
                throw new ArgumentException(
                    $"Invalid 'idColumn' value: '{idRaw}'. On a column path, idColumn is true or false. " +
                    "To mark a column by index, set the table --prop idColumn=N (or idColumns=1,3).");
            sawId = true;
            idOn = IsTruthy(idRaw);
        }

        bool wrapSet = false;
        bool wrap = false;
        if (properties.TryGetValue("noWrap", out var noWrapRaw)
            || properties.TryGetValue("nowrap", out noWrapRaw))
        {
            wrapSet = true;
            wrap = IsTruthy(noWrapRaw);
        }

        bool fit = false;
        int? explicitTwips = null;
        if (properties.TryGetValue("width", out var widthRaw))
        {
            if (IsColumnFitWidthToken(widthRaw))
                fit = true;
            else
            {
                var parsed = ParseTwips(widthRaw);
                if (parsed == 0)
                    throw new ArgumentException(
                        $"Invalid 'width' value: '{widthRaw}'. Column width must be a positive length, or 'fit'.");
                explicitTwips = checked((int)parsed);
            }
        }

        bool? noWrap = wrapSet ? wrap : (sawId ? idOn : null);
        bool doFit = explicitTwips == null && (fit || (sawId && idOn));
        return new ColumnIdFlags(noWrap, doFit, explicitTwips);
    }

    private static bool IsColumnFitWidthToken(string value)
    {
        var token = value.Trim();
        return token.Equals("fit", StringComparison.OrdinalIgnoreCase)
            || token.Equals("auto", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsIdColumnWordBoolean(string value)
    {
        var token = value.Trim();
        return token.Equals("true", StringComparison.OrdinalIgnoreCase)
            || token.Equals("false", StringComparison.OrdinalIgnoreCase)
            || token.Equals("yes", StringComparison.OrdinalIgnoreCase)
            || token.Equals("no", StringComparison.OrdinalIgnoreCase)
            || token.Equals("on", StringComparison.OrdinalIgnoreCase)
            || token.Equals("off", StringComparison.OrdinalIgnoreCase);
    }

    private List<int> ParseIdColumnList(Table table, string spec)
    {
        // "1" and "0" are boolean tokens elsewhere (IsTruthy), but on a table
        // they are column indexes. Only the word forms are the boolean mistake.
        if (IsIdColumnWordBoolean(spec))
            throw new ArgumentException(
                $"Invalid 'idColumn' value: '{spec}'. On a table, idColumn is a 1-based column index " +
                "or a comma-separated list (idColumn=1 or idColumns=1,3). " +
                "On a column path, use --prop idColumn=true.");

        var cols = RequireGridColumns(table);
        var seen = new HashSet<int>();
        var ordered = new List<int>();
        foreach (var part in spec.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (!int.TryParse(part, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index)
                || index < 1 || index > cols.Count)
                throw new ArgumentException(
                    $"Invalid 'idColumn' value: '{spec}'. Expected 1-based column index(es) between 1 and {cols.Count}.");
            if (seen.Add(index))
                ordered.Add(index);
        }
        if (ordered.Count == 0)
            throw new ArgumentException(
                "Invalid 'idColumn' value: empty. Pass a 1-based column index or a comma-separated list.");
        return ordered;
    }

    /// <summary>
    /// Writes <c>w:noWrap</c> and/or a fitted (or explicit) width onto one
    /// grid column. Single-span cells that start in the column are updated.
    /// A cell that only covers the column (a merge that starts earlier, or
    /// a merge that starts here and spans further) keeps its own wrap and
    /// has its dxa width shifted by the same delta so the row still adds up.
    /// Other columns are left alone. A dxa table width moves by that delta;
    /// a pct/auto table width is left as the page proportion.
    /// </summary>
    private void ApplyColumnIdTreatment(Table table, int column1Based, ColumnIdFlags flags)
    {
        var cols = RequireGridColumns(table);
        if (column1Based < 1 || column1Based > cols.Count)
            throw new ArgumentException($"Column {column1Based} not found (total: {cols.Count})");

        int slot = column1Based - 1;
        var addressable = new List<TableCell>();
        var covering = new List<TableCell>();
        foreach (var row in table.Elements<TableRow>())
        {
            if (!TryCellAtSlot(row, slot, out var cell, out var start, out var span))
                continue;
            if (start == slot && span == 1)
                addressable.Add(cell);
            else
                covering.Add(cell);
        }
        if (addressable.Count == 0)
            throw new ArgumentException(
                $"Cannot mark column {column1Based}: no single-span cell starts in that column (merged cells). Unmerge first.");

        int? target = flags.ExplicitTwips;
        if (flags.Fit)
        {
            double units = addressable.Count == 0 ? 0 : addressable.Max(LongestLineUnits);
            target = FitIdColumnWidthTwips(units);
        }

        if (target is int widthTwips)
        {
            int old = 0;
            int.TryParse(cols[slot].Width?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out old);
            int delta = widthTwips - old;
            cols[slot].Width = widthTwips.ToString(CultureInfo.InvariantCulture);
            foreach (var cell in addressable)
                SetCellDxaWidth(cell, widthTwips);
            if (delta != 0)
            {
                foreach (var cell in covering)
                    ShiftDxaWidth(cell, delta);
                ShiftDxaTableWidth(table, delta);
            }
        }

        if (flags.NoWrap is bool wrap)
        {
            foreach (var cell in addressable)
                SetCellNoWrap(cell, wrap);
        }
    }

    private static List<TableCell> AddressableCells(Table table, int slot)
    {
        var cells = new List<TableCell>();
        foreach (var row in table.Elements<TableRow>())
        {
            if (TryCellAtSlot(row, slot, out var cell, out var start, out var span)
                && start == slot && span == 1)
                cells.Add(cell);
        }
        return cells;
    }

    private static List<GridColumn> RequireGridColumns(Table table)
    {
        var grid = table.GetFirstChild<TableGrid>()
            ?? throw new InvalidOperationException("Table has no <w:tblGrid>");
        var cols = grid.Elements<GridColumn>().ToList();
        if (cols.Count == 0)
            throw new InvalidOperationException("Table has an empty <w:tblGrid>");
        return cols;
    }

    private static bool TryCellAtSlot(TableRow row, int slot, out TableCell cell, out int start, out int span)
    {
        int acc = row.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value ?? 0;
        foreach (var candidate in row.Elements<TableCell>())
        {
            int cellSpan = candidate.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
            if (cellSpan < 1) cellSpan = 1;
            if (slot >= acc && slot < acc + cellSpan)
            {
                cell = candidate;
                start = acc;
                span = cellSpan;
                return true;
            }
            acc += cellSpan;
        }
        cell = null!;
        start = 0;
        span = 0;
        return false;
    }

    private static double LongestLineUnits(TableCell cell)
    {
        double max = 0;
        foreach (var para in cell.Elements<Paragraph>())
        {
            var sb = new StringBuilder();
            void Flush()
            {
                if (sb.Length == 0) return;
                var units = ParseHelpers.EstimateTextWidthInChars(sb.ToString());
                if (units > max) max = units;
                sb.Clear();
            }
            foreach (var child in para.Descendants())
            {
                switch (child)
                {
                    case Text text:
                        sb.Append(text.Text);
                        break;
                    case Break:
                    case CarriageReturn:
                        Flush();
                        break;
                }
            }
            Flush();
        }
        return max;
    }

    internal static int FitIdColumnWidthTwips(double maxCharUnits)
    {
        if (maxCharUnits < 0) maxCharUnits = 0;
        double fitted = maxCharUnits * 2.0 * IdColumnTwipsPerMm + 8.0 * IdColumnTwipsPerMm;
        double floor = 18.0 * IdColumnTwipsPerMm;
        int twips = (int)Math.Ceiling(Math.Max(floor, fitted));
        // A pasted paragraph must not blow the grid out to an unbounded width.
        if (twips > 31680) twips = 31680; // 22 inches
        if (twips < 1) twips = 1;
        return twips;
    }

    private static void SetCellDxaWidth(TableCell cell, int twips)
    {
        var tcPr = cell.TableCellProperties ?? cell.PrependChild(new TableCellProperties());
        tcPr.TableCellWidth = new TableCellWidth
        {
            Width = twips.ToString(CultureInfo.InvariantCulture),
            Type = TableWidthUnitValues.Dxa,
        };
    }

    private static void ShiftDxaWidth(TableCell cell, int delta)
    {
        var tcW = cell.TableCellProperties?.TableCellWidth;
        if (tcW == null) return;
        if (tcW.Type?.Value is not null && tcW.Type.Value != TableWidthUnitValues.Dxa) return;
        if (!int.TryParse(tcW.Width?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var current))
            return;
        int next = current + delta;
        if (next < 1) next = 1;
        tcW.Width = next.ToString(CultureInfo.InvariantCulture);
        tcW.Type ??= TableWidthUnitValues.Dxa;
    }

    private static void ShiftDxaTableWidth(Table table, int delta)
    {
        var tblW = table.GetFirstChild<TableProperties>()?.TableWidth;
        if (tblW == null) return;
        if (tblW.Type?.Value is not null && tblW.Type.Value != TableWidthUnitValues.Dxa) return;
        if (!long.TryParse(tblW.Width?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var current))
            return;
        long next = current + delta;
        if (next < 1) next = 1;
        tblW.Width = next.ToString(CultureInfo.InvariantCulture);
        tblW.Type ??= TableWidthUnitValues.Dxa;
    }

    private static void SetCellNoWrap(TableCell cell, bool on)
    {
        var tcPr = cell.TableCellProperties;
        if (!on)
        {
            if (tcPr != null) tcPr.NoWrap = null;
            return;
        }
        tcPr ??= cell.PrependChild(new TableCellProperties());
        // Same assignment the cell setter uses. The SDK property inserts
        // <w:noWrap/> in CT_TcPr order (after shd, before tcMar).
        if (tcPr.NoWrap == null)
            tcPr.NoWrap = new NoWrap();
    }
}
