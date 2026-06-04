// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // Theme color map (lazy-initialized from theme1.xml)
    private Dictionary<string, string>? _excelThemeColors;
    // Indexed color palette (default 64 + custom overrides from styles.xml)
    private string[]? _resolvedIndexedColors;

    private Dictionary<string, string> GetExcelThemeColors()
    {
        if (_excelThemeColors != null) return _excelThemeColors;
        var colorScheme = _doc.WorkbookPart?.ThemePart?.Theme?.ThemeElements?.ColorScheme;
        _excelThemeColors = Core.ThemeColorResolver.BuildColorMap(colorScheme);
        return _excelThemeColors;
    }

    /// <summary>
    /// Excel theme color index mapping:
    /// 0=lt1, 1=dk1, 2=lt2, 3=dk2, 4=accent1, 5=accent2, 6=accent3, 7=accent4, 8=accent5, 9=accent6
    /// </summary>
    private static readonly string[] ThemeIndexToName =
        ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];

    private string? ResolveThemeColor(uint themeIndex, double? tintValue = null)
    {
        if (themeIndex >= (uint)ThemeIndexToName.Length) return null;
        var themeColors = GetExcelThemeColors();
        if (!themeColors.TryGetValue(ThemeIndexToName[themeIndex], out var hex)) return null;

        if (tintValue.HasValue && Math.Abs(tintValue.Value) > 0.001)
        {
            // Excel tint: positive = tint toward white, negative = shade toward black
            // Convert to OOXML 0-100000 range
            var t = tintValue.Value;
            if (t > 0)
                return Core.ColorMath.ApplyTransforms(hex, tint: (int)((1 - t) * 100000));
            else
                return Core.ColorMath.ApplyTransforms(hex, shade: (int)((1 + t) * 100000));
        }

        return $"#{hex}";
    }

    private string[] GetResolvedIndexedColors()
    {
        if (_resolvedIndexedColors != null) return _resolvedIndexedColors;

        // Start with default palette
        _resolvedIndexedColors = (string[])DefaultIndexedColors.Clone();

        // Check for custom overrides in styles.xml
        var stylesheet = _doc.WorkbookPart?.WorkbookStylesPart?.Stylesheet;
        var colors = stylesheet?.GetFirstChild<Colors>();
        var indexedColors = colors?.GetFirstChild<IndexedColors>();
        if (indexedColors != null)
        {
            int idx = 0;
            foreach (var rgbColor in indexedColors.Elements<RgbColor>())
            {
                if (idx < _resolvedIndexedColors.Length && rgbColor.Rgb?.Value != null)
                {
                    var raw = rgbColor.Rgb.Value;
                    _resolvedIndexedColors[idx] = FormatColorForCss(raw);
                }
                idx++;
            }
        }
        return _resolvedIndexedColors;
    }

    /// <summary>
    /// Generate a self-contained HTML file that previews all sheets as spreadsheet tables.
    /// Supports cell formatting (font, fill, borders, alignment), merged cells,
    /// column widths, row heights, frozen panes, and sheet tab switching.
    /// </summary>
    public string ViewAsHtml()
    {
        using var _cul = InvariantCultureScope.Enter();
        var sb = new StringBuilder();
        var sheets = GetWorksheets();
        var wbStylesPart = _doc.WorkbookPart?.WorkbookStylesPart;
        var stylesheet = wbStylesPart?.Stylesheet;

        // If any sheet has a pivot table, build an editable in-memory copy so
        // we can re-materialize cells from the pivot cache without mutating
        // the live _doc. The copy's WorksheetParts replace the originals for
        // rendering; styles/theme come from _doc (identical).
        //
        // CONSISTENCY(pivot-clone-in-memory): we clone _doc directly instead of
        // re-opening _filePath from disk. The earlier "read the file back via
        // FileStream(FileShare.ReadWrite)" approach races the handler's still-
        // held editable handle on macOS and throws IOException despite the
        // share-mode hint — the error surfaces as a trailing "process cannot
        // access" stderr after every add pivot/slicer command, and worse, on
        // every SUBSEQUENT command once the file has a pivot part at all (the
        // `sheets.Any(...PivotTableParts...)` branch fires on every ViewAsHtml
        // from the NotifyWatch path). SpreadsheetDocument.Clone(Stream, bool)
        // serialises the already-loaded package into the MemoryStream without
        // touching disk, so there is no second file handle to race.
        MemoryStream? pivotMs = null;
        SpreadsheetDocument? pivotDoc = null;
        List<(string Name, WorksheetPart Part)>? pivotSheets = null;
        if (sheets.Any(s => s.Part.PivotTableParts.Any()))
        {
            pivotMs = new MemoryStream();
            pivotDoc = (SpreadsheetDocument)_doc.Clone(pivotMs, isEditable: true);
            pivotSheets = GetWorksheets(pivotDoc);

            foreach (var (_, wsPart) in pivotSheets)
            {
                if (wsPart.PivotTableParts.Any())
                    OfficeCli.Core.PivotTableHelper.RefreshPivotCellsForView(wsPart);
            }

            // Use the copy's stylesheet so new indent styles created by the
            // pivot refresh are visible to the HTML renderer.
            stylesheet = pivotDoc.WorkbookPart?.WorkbookStylesPart?.Stylesheet;
        }

        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html>");
        sb.AppendLine("<head>");
        sb.AppendLine("<meta charset=\"UTF-8\">");
        sb.AppendLine("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
        sb.AppendLine($"<title>{HtmlEncode(Path.GetFileName(_filePath))}</title>");
        sb.AppendLine("<style>");
        sb.AppendLine(GenerateExcelCss());
        sb.AppendLine("</style>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");

        // File title
        sb.AppendLine($"<div class=\"file-title\">{HtmlEncode(Path.GetFileName(_filePath))}</div>");

        // Sheet content areas (tabs moved to bottom)
        sb.AppendLine("<div class=\"sheet-slider\">");
        for (int sheetIdx = 0; sheetIdx < sheets.Count; sheetIdx++)
        {
            var (sheetName, worksheetPart) = sheets[sheetIdx];
            // Use the pivot-refreshed copy's WorksheetPart when available
            var renderPart = pivotSheets != null && sheetIdx < pivotSheets.Count
                ? pivotSheets[sheetIdx].Part : worksheetPart;
            var activeClass = sheetIdx == 0 ? " active" : "";
            // Check if sheet is RTL
            var sheetView = GetSheet(renderPart).GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>();
            var isRtl = sheetView?.RightToLeft?.Value == true;
            var dirAttr = isRtl ? " dir=\"rtl\"" : "";
            sb.AppendLine($"<div class=\"sheet-content{activeClass}\" data-sheet=\"{sheetIdx}\"{dirAttr}>");
            var charts = CollectSheetCharts(worksheetPart, sheetName);
            // Shapes and textboxes (xdr:sp). Reuses the chart overlay
            // positioning pipeline — same (fromRow,toRow,fromCol,toCol,html)
            // tuple is consumed by RenderSheetTable to emit an absolutely-
            // positioned overlay over the sheet grid.
            var shapes = CollectSheetShapes(worksheetPart);
            if (shapes.Count > 0)
                charts.AddRange(shapes);
            RenderSheetTable(sb, sheetName, renderPart, stylesheet, charts, sheetIdx);
            sb.AppendLine("</div>");
        }
        sb.AppendLine("</div>");

        // Sheet tabs at bottom (like real Excel)
        sb.AppendLine("<div class=\"sheet-tabs\" role=\"tablist\">");
        for (int i = 0; i < sheets.Count; i++)
        {
            var activeClass = i == 0 ? " active" : "";
            var tabColorStyle = "";
            var sheetProps = GetSheet(sheets[i].Part).GetFirstChild<SheetProperties>();
            var tabColorEl = sheetProps?.TabColor;
            if (tabColorEl?.Rgb?.Value != null)
            {
                var rgb = tabColorEl.Rgb.Value;
                if (rgb.Length > 6) rgb = rgb[^6..];
                // Hex-gate before inline style interpolation — unchecked
                // raw value would break out of the style attribute.
                if (rgb.Length == 6
                    && rgb.All(c => (c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f')))
                    tabColorStyle = $" style=\"--tab-color:#{rgb}\"";
            }
            sb.AppendLine($"  <div class=\"sheet-tab{activeClass}\"{tabColorStyle} data-sheet=\"{i}\" role=\"tab\" tabindex=\"0\" onclick=\"switchSheet({i})\" onkeydown=\"if(event.key==='Enter'||event.key===' ')switchSheet({i})\">{HtmlEncode(sheets[i].Name)}</div>");
        }
        sb.AppendLine("</div>");

        // Sheet switching JavaScript
        sb.AppendLine("<script>");
        sb.AppendLine(GenerateExcelJs());
        sb.AppendLine("</script>");
        // CONSISTENCY(excel-virt): private virt script injected after standard overlay.
        // Open-source GetVirtScript() returns empty; private override loads watch-overlay-virt.js.
        var virtScript = GetVirtScript();
        if (virtScript.Length > 0)
        {
            sb.AppendLine("<script>");
            sb.AppendLine(virtScript);
            sb.AppendLine("</script>");
        }

        sb.AppendLine("</body>");
        sb.AppendLine("</html>");

        pivotDoc?.Dispose();
        pivotMs?.Dispose();

        return sb.ToString();
    }

    /// <summary>
    /// Get the number of sheets (for watch notifications).
    /// </summary>
    public int GetSheetCount() => GetWorksheets().Count;

    /// <summary>Get the 0-based index of a sheet by name, or -1 if not found.</summary>
    public int GetSheetIndex(string sheetName)
    {
        var sheets = GetWorksheets();
        for (int i = 0; i < sheets.Count; i++)
            if (string.Equals(sheets[i].Name, sheetName, System.StringComparison.OrdinalIgnoreCase))
                return i;
        return -1;
    }

    // ==================== Sheet Rendering ====================

    private void RenderSheetTable(StringBuilder sb, string sheetName, WorksheetPart worksheetPart, Stylesheet? stylesheet,
        List<(int fromRow, int toRow, int fromCol, int toCol, string html)>? charts = null, int sheetIdx = 0)
    {
        var ws = GetSheet(worksheetPart);
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData == null && (charts == null || charts.Count == 0))
        {
            if (worksheetPart.DrawingsPart?.WorksheetDrawing == null)
                sb.AppendLine("<div class=\"empty-sheet\">Empty sheet</div>");
            return;
        }

        // Read default dimensions from sheetFormatPr
        var sheetFmtPr = ws.GetFirstChild<SheetFormatProperties>();
        // Excel column width → pixels: chars * 7.0017 (DEFAULT_CHARACTER_WIDTH for Calibri 11)
        // pt = px * 0.75
        var defaultColWidthPt = sheetFmtPr?.DefaultColumnWidth?.Value != null
            ? sheetFmtPr.DefaultColumnWidth.Value * 7.0017 * 0.75 : 8.43 * 7.0017 * 0.75;
        var defaultRowHeightPt = sheetFmtPr?.DefaultRowHeight?.Value ?? 15.0;

        // Read default font size from stylesheet
        var defaultFontPt = 11.0;
        if (stylesheet?.Fonts != null && stylesheet.Fonts.Elements<Font>().Any())
        {
            var defFont = stylesheet.Fonts.Elements<Font>().First();
            defaultFontPt = defFont.FontSize?.Val?.Value ?? 11.0;
        }

        // Create formula evaluator for this sheet to compute uncached formula values
        var evaluator = sheetData != null ? new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart) : null;

        // Collect merge info
        var mergeMap = BuildMergeMap(ws);

        // Build conditional formatting CSS overrides (skip if no cell data)
        var cfMap = sheetData != null ? BuildConditionalFormatMap(ws, stylesheet, sheetData, _doc.WorkbookPart) : new Dictionary<string, string>();
        var dataBarMap = sheetData != null ? BuildDataBarMap(ws, sheetData) : new Dictionary<string, string>();
        var iconSetMap = sheetData != null ? BuildIconSetMap(ws, sheetData) : new Dictionary<string, string>();
        // R12a: sparklines live in cells that often have no CellValue. Build the
        // host-cell → SVG map now; maxCol/maxRow are extended to cover those cells
        // below, after they're computed from cell data.
        var sparklineMap = BuildSparklineMap(ws);

        // Collect column widths
        var colWidths = GetColumnWidths(ws);

        // Detect frozen panes
        var (frozenRows, frozenCols) = GetFrozenPanes(ws);

        // Compute cumulative left offsets for frozen columns (for sticky positioning)
        // Index 0 = row header width (30pt), index 1 = col 1 left offset, etc.
        var frozenLeftOffsets = new Dictionary<int, double>();
        if (frozenCols > 0)
        {
            double cumLeft = 30; // row header width in pt
            for (int fc = 1; fc <= frozenCols; fc++)
            {
                frozenLeftOffsets[fc] = cumLeft;
                cumLeft += colWidths.TryGetValue(fc, out var w) ? w : defaultColWidthPt;
            }
        }

        // Determine grid dimensions. Count all cells that exist in SheetData —
        // every Cell element with a CellReference contributes to maxRow/maxCol,
        // even if the cell is empty (no value, no formula). Empty cells are
        // explicitly created by the user or by Excel; either way they should
        // render so the grid matches the actual data range.
        var rows = sheetData?.Elements<Row>().ToList() ?? new List<Row>();
        int maxCol = 0;
        int maxRow = 0;
        foreach (var row in rows)
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 0);
            bool rowHasCells = false;
            foreach (var cell in row.Elements<Cell>())
            {
                var cellRef = cell.CellReference?.Value;
                if (cellRef == null) continue;
                var (colName, _) = ParseCellReference(cellRef);
                var colIdx = ColumnNameToIndex(colName);
                if (colIdx > maxCol) maxCol = colIdx;
                rowHasCells = true;
            }
            if (rowHasCells && rowIdx > maxRow) maxRow = rowIdx;
        }

        // Extend maxRow/maxCol from chart anchors even when no cell data
        if (charts != null)
        {
            foreach (var (fromRow, toRow, fromCol, toCol, _) in charts)
            {
                if (toRow > maxRow) maxRow = toRow;
                if (toCol > maxCol) maxCol = toCol;
            }
        }

        // Extend maxCol to cover columns that carry an explicit <col> width but no
        // cell data. Excel renders those columns (the user sized them on purpose),
        // and they are valid spill targets for a long text cell to their left.
        foreach (var widthCol in colWidths.Keys)
            if (widthCol > maxCol) maxCol = widthCol;

        // R12a: extend maxRow/maxCol to cover sparkline host cells (which often
        // have no CellValue and would otherwise be cropped out of the grid).
        foreach (var hostRef in sparklineMap.Keys)
        {
            var (hc, hr) = ParseCellReference(hostRef);
            var hcIdx = ColumnNameToIndex(hc);
            if (hcIdx > maxCol) maxCol = hcIdx;
            if (hr > maxRow) maxRow = hr;
        }

        // Empty sheet (no cells and no charts)
        if (maxRow == 0 || maxCol == 0)
        {
            if (worksheetPart.DrawingsPart?.WorksheetDrawing == null)
                sb.AppendLine("<div class=\"empty-sheet\">Empty sheet</div>");
            return;
        }

        // Extend maxRow/maxCol to include chart anchor ranges
        if (charts != null)
            foreach (var (_, toRow, fromCol, toCol, _) in charts)
            {
                if (toCol > maxCol) maxCol = toCol;
                if (toRow > maxRow) maxRow = toRow;
            }

        // Column cap: >200 cols is unusable in a browser table regardless of rendering mode.
        // Row cap: default 5000; overridable via OnGetHtmlRowCap when the rendering backend
        // keeps DOM node count bounded independently of sheet size.
        var actualRow = maxRow;
        var actualCol = maxCol;
        maxRow = Math.Min(maxRow, GetHtmlRowCap());
        maxCol = Math.Min(maxCol, 200);
        var truncated = actualRow > maxRow || actualCol > maxCol;

        // Build cell lookup: (row, col) → Cell
        var cellMap = new Dictionary<(int row, int col), Cell>();
        foreach (var row in rows)
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 0);
            if (rowIdx > maxRow) break;
            foreach (var cell in row.Elements<Cell>())
            {
                var cellRef = cell.CellReference?.Value;
                if (cellRef == null) continue;
                var (colName, _) = ParseCellReference(cellRef);
                var colIdx = ColumnNameToIndex(colName);
                if (colIdx <= maxCol)
                    cellMap[(rowIdx, colIdx)] = cell;
            }
        }

        // Row height and hidden row lookup
        var rowHeights = new Dictionary<int, double>();
        var hiddenRows = new HashSet<int>();
        foreach (var row in rows)
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 0);
            if (row.CustomHeight?.Value == true && row.Height?.Value != null)
                rowHeights[rowIdx] = row.Height.Value;
            if (row.Hidden?.Value == true)
                hiddenRows.Add(rowIdx);
        }

        // Compute cumulative top offsets for frozen rows (for sticky positioning)
        // Includes thead height (~24pt for column headers)
        var frozenTopOffsets = new Dictionary<int, double>();
        if (frozenRows > 0)
        {
            double cumTop = 24; // approximate thead (column header) height
            for (int fr = 1; fr <= frozenRows; fr++)
            {
                frozenTopOffsets[fr] = cumTop;
                if (rowHeights.TryGetValue(fr, out var rh))
                    cumTop += rh;
                else
                {
                    // Estimate row height from max font size in the row's cells
                    double maxFontPt = defaultFontPt;
                    foreach (var cell in cellMap.Where(kv => kv.Key.row == fr).Select(kv => kv.Value))
                    {
                        var si = cell.StyleIndex?.Value ?? 0;
                        if (stylesheet?.CellFormats != null && si < (uint)stylesheet.CellFormats.Elements<CellFormat>().Count())
                        {
                            var xf = stylesheet.CellFormats.Elements<CellFormat>().ElementAt((int)si);
                            var fontId = xf.FontId?.Value ?? 0;
                            if (stylesheet.Fonts != null && fontId < (uint)stylesheet.Fonts.Elements<Font>().Count())
                            {
                                var font = stylesheet.Fonts.Elements<Font>().ElementAt((int)fontId);
                                var sz = font.FontSize?.Val?.Value ?? defaultFontPt;
                                if (sz > maxFontPt) maxFontPt = sz;
                            }
                        }
                    }
                    cumTop += maxFontPt * 1.4 + 4; // font height + padding
                }
            }
        }

        // Collect hidden columns
        var hiddenCols = new HashSet<int>();
        foreach (var (colIdx, widthPx) in colWidths)
        {
            if (widthPx <= 0) hiddenCols.Add(colIdx);
        }

        // Columns without an explicit OOXML <col> width fall through to
        // defaultColWidthPt (Excel's default ~8.43 chars ≈ 44pt). We do NOT
        // auto-fit to content width: real Excel keeps the default column width
        // and lets long text spill into empty right-neighbour cells (see the
        // spill handling in the cell render below). Auto-fitting would grow the
        // column to hold the text, defeating both Excel fidelity and spill.

        // Build chart lookup: fromRow → chart info for inline insertion
        var chartAtRow = new Dictionary<int, (int toRow, int fromCol, int toCol, string html)>();
        if (charts != null)
            foreach (var (fromRow, toRow, fromCol, toCol, html) in charts)
                chartAtRow[fromRow] = (toRow, fromCol, toCol, html);

        // Compute total table width so the table sizes to its content (not the wrapper).
        // Without an explicit width, table-layout:fixed inside a flex wrapper shrinks columns
        // proportionally to fit the viewport, ignoring declared col widths.
        double totalTableWidthPt = 30; // row-header-col width
        for (int c = 1; c <= maxCol; c++)
        {
            if (hiddenCols.Contains(c)) continue;
            totalTableWidthPt += colWidths.TryGetValue(c, out var cw) ? cw : defaultColWidthPt;
        }

        // Start table (position:relative for chart overlays)
        sb.AppendLine("<div class=\"table-wrapper\" style=\"position:relative\">");
        sb.AppendLine($"<table style=\"width:{totalTableWidthPt:0.##}pt\">");
        sb.AppendLine($"<caption class=\"sr-only\">{HtmlEncode(sheetName)}</caption>");

        // Colgroup for column widths + header column (skip hidden columns to match td count)
        sb.Append("<colgroup><col class=\"row-header-col\">");
        for (int c = 1; c <= maxCol; c++)
        {
            if (hiddenCols.Contains(c)) continue; // skip hidden cols — tds are also skipped
            var width = colWidths.TryGetValue(c, out var w) ? w : defaultColWidthPt;
            sb.Append($"<col style=\"width:{width:0.##}pt\">");
        }
        sb.AppendLine("</colgroup>");

        // Column header row
        sb.Append("<thead><tr><th class=\"corner-cell\"");
        if (frozenRows > 0 || frozenCols > 0) sb.Append(" style=\"position:sticky;top:0;left:0;z-index:4\"");
        sb.Append("></th>");
        for (int c = 1; c <= maxCol; c++)
        {
            if (hiddenCols.Contains(c)) continue;
            var colName = IndexToColumnName(c);
            var isFrozenColHeader = frozenCols > 0 && c <= frozenCols;
            string stickyStyle;
            if (frozenRows > 0 && isFrozenColHeader)
            {
                var leftPt = frozenLeftOffsets.TryGetValue(c, out var lf) ? lf : 0;
                stickyStyle = $" style=\"position:sticky;top:0;left:{leftPt:0.##}pt;z-index:4\"";
            }
            else if (frozenRows > 0)
                stickyStyle = " style=\"position:sticky;top:0;z-index:3\"";
            else if (isFrozenColHeader)
            {
                var leftPt = frozenLeftOffsets.TryGetValue(c, out var lf2) ? lf2 : 0;
                stickyStyle = $" style=\"position:sticky;left:{leftPt:0.##}pt;z-index:3\"";
            }
            else
                stickyStyle = "";
            sb.Append($"<th class=\"col-header\" data-path=\"/{HtmlEncode(sheetName)}/col[{colName}]\"{stickyStyle}>{colName}</th>");
        }
        sb.AppendLine("</tr></thead>");

        // chartAtRow and sideCharts already built above

        // Visible column count for chart colspan
        var visibleColCount = Enumerable.Range(1, maxCol).Count(c => !hiddenCols.Contains(c));

        // CONSISTENCY(excel-virt): Extension point — private override in
        // ExcelHandler.HtmlPreview.Virt.cs replaces the full static tbody with a
        // JSON-data tbody + JS virtual renderer. BuildRowInnerHtml is shared for
        // cell rendering; open-source RenderTbody emits static <tr> elements.
        var ctx = new SheetRenderContext(sheetName, sheetIdx, cellMap, maxRow, maxCol,
            rowHeights, hiddenRows, hiddenCols, mergeMap, frozenRows, frozenCols,
            frozenLeftOffsets, frozenTopOffsets, cfMap, dataBarMap, iconSetMap, sparklineMap,
            stylesheet, evaluator, defaultColWidthPt, defaultRowHeightPt, colWidths);
        RenderTbody(sb, ctx);
        sb.AppendLine("</table>");

        // Render charts as absolute-positioned overlays on top of the table grid.
        // Position is computed from anchor row/col using column widths and row heights.
        if (charts != null)
        {
            var rowHeaderWidthPt = 30.0; // matches .row-header-col CSS
            foreach (var (fromRow, toRow, fromCol, toCol, html) in charts)
            {
                // Compute left position: sum of column widths from col 1 to fromCol + row header
                double leftPt = rowHeaderWidthPt;
                for (int c = 1; c <= fromCol && c <= maxCol; c++)
                {
                    if (hiddenCols.Contains(c)) continue;
                    leftPt += colWidths.TryGetValue(c, out var cw) ? cw : defaultColWidthPt;
                }
                // Compute top position: sum of row heights from row 1 to fromRow + header row (~24px)
                double topPt = 24.0 * 0.75; // header row height in pt
                for (int r = 1; r <= fromRow && r <= maxRow; r++)
                {
                    if (hiddenRows.Contains(r)) continue;
                    topPt += rowHeights.TryGetValue(r, out var rh) ? rh : defaultRowHeightPt;
                }
                // Compute width/height from anchor span
                double widthPt = 0;
                for (int c = fromCol + 1; c <= toCol && c <= maxCol; c++)
                {
                    if (hiddenCols.Contains(c)) continue;
                    widthPt += colWidths.TryGetValue(c, out var cw2) ? cw2 : defaultColWidthPt;
                }
                double heightPt = 0;
                for (int r = fromRow + 1; r <= toRow && r <= maxRow; r++)
                {
                    if (hiddenRows.Contains(r)) continue;
                    heightPt += rowHeights.TryGetValue(r, out var rh2) ? rh2 : defaultRowHeightPt;
                }
                if (widthPt < 100) widthPt = 400; // fallback min size
                if (heightPt < 50) heightPt = 250;

                sb.AppendLine($"<div style=\"position:absolute;left:{leftPt:0.##}pt;top:{topPt:0.##}pt;width:{widthPt:0.##}pt;height:{heightPt:0.##}pt;z-index:10;pointer-events:auto\" data-from-col=\"{fromCol}\" data-from-row=\"{fromRow}\">");
                sb.Append(html);
                sb.AppendLine("</div>");
            }
        }

        // Truncation warning
        if (truncated)
            sb.AppendLine($"<div class=\"truncation-warning\">Showing {maxRow} of {actualRow} rows, {maxCol} of {actualCol} columns</div>");
        sb.AppendLine("</div>"); // close table-wrapper
    }

    // ==================== Merge Map ====================

    internal record struct MergeInfo(bool IsAnchor, int RowSpan, int ColSpan);

    // CONSISTENCY(excel-virt): Packages all sheet-level computed data needed to render
    // tbody rows. Passed to RenderTbody so the private virt override can serialise all
    // cell HTML to JSON without re-running the data-collection logic.
    internal record SheetRenderContext(
        string SheetName,
        int SheetIdx,
        Dictionary<(int row, int col), Cell> CellMap,
        int MaxRow, int MaxCol,
        Dictionary<int, double> RowHeights,
        HashSet<int> HiddenRows,
        HashSet<int> HiddenCols,
        Dictionary<string, MergeInfo> MergeMap,
        int FrozenRows, int FrozenCols,
        Dictionary<int, double> FrozenLeftOffsets,
        Dictionary<int, double> FrozenTopOffsets,
        Dictionary<string, string> CfMap,
        Dictionary<string, string> DataBarMap,
        Dictionary<string, string> IconSetMap,
        Dictionary<string, string> SparklineMap,
        Stylesheet? Stylesheet,
        Core.FormulaEvaluator? Evaluator,
        double DefaultColWidthPt,
        double DefaultRowHeightPt,
        Dictionary<int, double> ColWidths);

    // CONSISTENCY(excel-virt): Private ExcelHandler.HtmlPreview.Virt.cs implements
    // OnRenderTbody to emit virtualised rows (JSON data + empty tbody) and sets
    // handled=true to skip the default. When no private implementation exists the
    // partial call is removed by the compiler and the default static rendering runs.
    partial void OnRenderTbody(StringBuilder sb, SheetRenderContext ctx, ref bool handled);

    // CONSISTENCY(excel-virt): default 5000-row cap for HTML preview; backend can
    // override via OnGetHtmlRowCap when DOM node count is bounded independently.
    partial void OnGetHtmlRowCap(ref int cap);
    internal int GetHtmlRowCap()
    {
        var cap = 5000;
        OnGetHtmlRowCap(ref cap);
        return cap;
    }

    internal void RenderTbody(StringBuilder sb, SheetRenderContext ctx)
    {
        bool handled = false;
        OnRenderTbody(sb, ctx, ref handled);
        if (handled) return;
        // Default: render all rows as static <tr> elements.
        sb.AppendLine("<tbody>");
        for (int r = 1; r <= ctx.MaxRow; r++)
        {
            if (ctx.HiddenRows.Contains(r)) { sb.AppendLine($"<tr data-row=\"{ctx.SheetIdx}-{r}\" style=\"display:none\"></tr>"); continue; }
            bool isRowFrozen = ctx.FrozenRows > 0 && r <= ctx.FrozenRows;
            var rowStyles = new List<string>();
            if (ctx.RowHeights.TryGetValue(r, out var rh)) rowStyles.Add($"height:{rh:0.##}pt");
            if (isRowFrozen) rowStyles.Add("background:#fff");
            var rowStyle = rowStyles.Count > 0 ? $" style=\"{string.Join(";", rowStyles)}\"" : "";
            var frozenAttr = isRowFrozen ? " data-frozen=\"1\"" : "";
            sb.Append($"<tr data-row=\"{ctx.SheetIdx}-{r}\"{rowStyle}{frozenAttr}>");
            sb.Append(BuildRowInnerHtml(ctx, r, isRowFrozen));
            sb.AppendLine("</tr>");
        }
        sb.AppendLine("</tbody>");
    }

    // CONSISTENCY(excel-virt): Shared row-cell renderer used by RenderTbody (open-source
    // static rendering) and ExcelHandler.HtmlPreview.Virt.cs (JSON serialisation).
    // Returns the <tr> inner content: row-header <th> + all cell <td> elements,
    // without the <tr> wrapper.
    internal string BuildRowInnerHtml(SheetRenderContext ctx, int r, bool isRowFrozen)
    {
        var rowSb = new StringBuilder();
        string rowHeaderStyle;
        if (isRowFrozen)
            rowHeaderStyle = " style=\"position:sticky;top:0;left:0;z-index:3\"";
        else if (ctx.FrozenCols > 0)
            rowHeaderStyle = " style=\"position:sticky;left:0;z-index:2\"";
        else
            rowHeaderStyle = "";
        rowSb.Append($"<th class=\"row-header\" data-path=\"/{HtmlEncode(ctx.SheetName)}/row[{r}]\"{rowHeaderStyle}>{r}</th>");

        for (int c = 1; c <= ctx.MaxCol; c++)
        {
            if (ctx.HiddenCols.Contains(c)) continue;
            var cellRef = $"{IndexToColumnName(c)}{r}";
            if (ctx.MergeMap.TryGetValue(cellRef, out var mergeInfo))
            {
                if (!mergeInfo.IsAnchor) continue;
                var cell = ctx.CellMap.TryGetValue((r, c), out var mc) ? mc : null;
                // Merged-region perimeter borders come from the perimeter member cells:
                // right edge from the right-column member, bottom edge from the bottom-row
                // member. The anchor's own right/bottom edges are interior to the merge.
                var rightMember = ctx.CellMap.TryGetValue((r, c + mergeInfo.ColSpan - 1), out var rmc) ? rmc : null;
                var bottomMember = ctx.CellMap.TryGetValue((r + mergeInfo.RowSpan - 1, c), out var bmc) ? bmc : null;
                var style = GetCellStyleCss(cell, ctx.Stylesheet, ctx.FrozenRows, ctx.FrozenCols, r, c, ctx.FrozenLeftOffsets, ctx.FrozenTopOffsets, ctx.CfMap, ctx.DataBarMap, ctx.IconSetMap, mergePerimeter: true, rightBorderCell: rightMember, bottomBorderCell: bottomMember);
                var value = cell != null ? GetFormattedCellValue(cell, ctx.Stylesheet, ctx.Evaluator) : "";
                var richHtml = cell != null ? TryBuildRichTextHtml(cell) : null;
                var adjColSpan = mergeInfo.ColSpan;
                if (adjColSpan > 1 && ctx.HiddenCols.Count > 0)
                    for (int hc = c + 1; hc < c + mergeInfo.ColSpan; hc++)
                        if (ctx.HiddenCols.Contains(hc)) adjColSpan--;
                var spanAttrs = "";
                if (adjColSpan > 1) spanAttrs += $" colspan=\"{adjColSpan}\"";
                if (mergeInfo.RowSpan > 1) spanAttrs += $" rowspan=\"{mergeInfo.RowSpan}\"";
                // Rich-text runs render as pre-built spans (already encoded); the
                // bar/icon overlay path is mutually exclusive with rich text here.
                var content = richHtml != null && !ctx.DataBarMap.ContainsKey(cellRef) && !ctx.IconSetMap.ContainsKey(cellRef)
                    ? richHtml
                    : BuildCellContent(cellRef, value, ctx.DataBarMap, ctx.IconSetMap);
                content = WrapVerticalAlign(content, GetCellVerticalAlign(cell, ctx.Stylesheet), richHtml);
                if (ctx.SparklineMap.TryGetValue(cellRef, out var spkSvg)) content = spkSvg + content;
                var diagSvg = TryBuildCellDiagonalSvg(cell, ctx.Stylesheet) ?? "";
                rowSb.Append($"<td data-path=\"/{HtmlEncode(ctx.SheetName)}/{cellRef}\"{GetFormulaAttr(cell)}{spanAttrs}{style}>{diagSvg}{content}</td>");
            }
            else
            {
                var cell = ctx.CellMap.TryGetValue((r, c), out var nc) ? nc : null;
                var style = GetCellStyleCss(cell, ctx.Stylesheet, ctx.FrozenRows, ctx.FrozenCols, r, c, ctx.FrozenLeftOffsets, ctx.FrozenTopOffsets, ctx.CfMap, ctx.DataBarMap, ctx.IconSetMap);
                var value = cell != null ? GetFormattedCellValue(cell, ctx.Stylesheet, ctx.Evaluator) : "";
                var richHtml = cell != null ? TryBuildRichTextHtml(cell) : null;
                var content = richHtml != null && !ctx.DataBarMap.ContainsKey(cellRef) && !ctx.IconSetMap.ContainsKey(cellRef)
                    ? richHtml
                    : BuildCellContent(cellRef, value, ctx.DataBarMap, ctx.IconSetMap);
                content = WrapVerticalAlign(content, GetCellVerticalAlign(cell, ctx.Stylesheet), richHtml);
                if (ctx.SparklineMap.TryGetValue(cellRef, out var spkSvg)) content = spkSvg + content;
                var diagSvg = TryBuildCellDiagonalSvg(cell, ctx.Stylesheet) ?? "";
                // Text-spill emulation (Excel-fidelity): a non-wrapped left/general
                // aligned text cell with empty right-neighbours paints its overflow
                // across those neighbours, clipping at the first occupied cell. The
                // <td> stays 1 column wide (preserving borders/gridlines/merges); the
                // text lives in an inline span that overflows visibly up to the summed
                // empty-neighbour width.
                var spillWidth = GetSpillWidthPt(ctx, cell, value, r, c);
                string spillClass = "";
                if (spillWidth > 0)
                {
                    spillClass = " class=\"spill\"";
                    content = $"<span class=\"spill-text\" style=\"max-width:{spillWidth:0.##}pt\">{content}</span>";
                }
                rowSb.Append($"<td data-path=\"/{HtmlEncode(ctx.SheetName)}/{cellRef}\"{GetFormulaAttr(cell)}{spillClass}{style}>{diagSvg}{content}</td>");
            }
        }
        return rowSb.ToString();
    }

    // ==================== Text spill (Excel overflow into empty neighbours) ====================
    //
    // Real Excel renders a long text cell's overflow across adjacent empty cells to
    // the right (for left/general alignment), clipping only at the first occupied
    // right-neighbour or the sheet edge. Returns the EXTRA pt budget (own column NOT
    // included — that is the td's normal width) the inline span may overflow into, or
    // 0 when the cell must clip at its own boundary (number, wrapText, occupied
    // neighbour, non-left alignment, etc.).
    private double GetSpillWidthPt(SheetRenderContext ctx, Cell? cell, string value, int r, int c)
    {
        if (cell == null || string.IsNullOrEmpty(value)) return 0;

        // (a) Must be text/general — NOT a number and NOT a numeric formula result.
        // Excel right-aligns numbers and never spills them.
        var dt = cell.DataType?.Value;
        bool isText = dt == CellValues.SharedString || dt == CellValues.InlineString || dt == CellValues.String;
        bool isBoolOrError = dt == CellValues.Boolean || dt == CellValues.Error;
        // Formula or general number: if it has a CellValue and isn't a string type,
        // treat as numeric (Excel does not spill numbers). Booleans/errors centre and
        // also don't spill.
        if (!isText || isBoolOrError) return 0;

        // (b) Resolve alignment + wrapText from the cell's xf.
        bool wrapText = false;
        string? hAlign = null;
        if (ctx.Stylesheet?.CellFormats != null)
        {
            var si = (int)(cell.StyleIndex?.Value ?? 0);
            var xfs = ctx.Stylesheet.CellFormats.Elements<CellFormat>().ToList();
            if (si >= 0 && si < xfs.Count)
            {
                var al = xfs[si].Alignment;
                wrapText = al?.WrapText?.Value == true;
                if (al?.Horizontal?.HasValue == true) hAlign = al.Horizontal.InnerText;
            }
        }
        if (wrapText) return 0; // wrapped cells clip/wrap, never spill

        // (c) Only left/general aligned text spills to the right (the common case).
        // right/center/justify/fill keep clipping (handled as follow-up).
        if (hAlign != null && hAlign != "left" && hAlign != "general" && hAlign != "fill")
            return 0;

        // (d) Sum widths of the run of EMPTY right-neighbours, stopping at the first
        // occupied cell, a merged region, a hidden column, or the sheet edge.
        double extra = 0;
        for (int nc = c + 1; nc <= ctx.MaxCol; nc++)
        {
            if (ctx.HiddenCols.Contains(nc)) break;
            var neighbourRef = $"{IndexToColumnName(nc)}{r}";
            if (ctx.MergeMap.ContainsKey(neighbourRef)) break;
            bool occupied = ctx.CellMap.TryGetValue((r, nc), out var ncell)
                            && ncell != null
                            && !string.IsNullOrEmpty(GetCellDisplayValue(ncell));
            if (occupied) break;
            extra += ctx.ColWidths.TryGetValue(nc, out var w) ? w : ctx.DefaultColWidthPt;
        }
        if (extra <= 0) return 0;

        // Own column width + the empty-neighbour run = the span's max overflow budget.
        double ownWidth = ctx.ColWidths.TryGetValue(c, out var ow) ? ow : ctx.DefaultColWidthPt;
        return ownWidth + extra;
    }

    // CONSISTENCY(excel-virt): Private ExcelHandler.HtmlPreview.Virt.cs implements
    // OnGetVirtScript to load watch-overlay-virt.js from embedded resources.
    // When no private implementation exists the partial call is removed and result
    // stays empty (no virtualisation script injected).
    partial void OnGetVirtScript(ref string result);

    internal string GetVirtScript()
    {
        var result = string.Empty;
        OnGetVirtScript(ref result);
        return result;
    }

    private Dictionary<string, MergeInfo> BuildMergeMap(Worksheet ws)
    {
        var map = new Dictionary<string, MergeInfo>(StringComparer.OrdinalIgnoreCase);
        var mergeCells = ws.GetFirstChild<MergeCells>();
        if (mergeCells == null) return map;

        foreach (var mc in mergeCells.Elements<MergeCell>())
        {
            var rangeRef = mc.Reference?.Value;
            if (string.IsNullOrEmpty(rangeRef) || !rangeRef.Contains(':')) continue;

            var parts = rangeRef.Split(':');
            var (startCol, startRow) = ParseCellReference(parts[0]);
            var (endCol, endRow) = ParseCellReference(parts[1]);
            var startColIdx = ColumnNameToIndex(startCol);
            var endColIdx = ColumnNameToIndex(endCol);
            // Clamp merge range to rendering limits to prevent memory explosion
            var clampedEndRow = Math.Min(endRow, 5000);
            var clampedEndCol = Math.Min(endColIdx, 200);
            var rowSpan = clampedEndRow - startRow + 1;
            var colSpan = clampedEndCol - startColIdx + 1;

            for (int r = startRow; r <= clampedEndRow; r++)
            {
                for (int ci = startColIdx; ci <= clampedEndCol; ci++)
                {
                    var cellRef = $"{IndexToColumnName(ci)}{r}";
                    bool isAnchor = (r == startRow && ci == startColIdx);
                    map[cellRef] = new MergeInfo(isAnchor, isAnchor ? rowSpan : 0, isAnchor ? colSpan : 0);
                }
            }
        }

        return map;
    }

    // ==================== Column Widths ====================

    private static Dictionary<int, double> GetColumnWidths(Worksheet ws)
    {
        var result = new Dictionary<int, double>();
        var columns = ws.GetFirstChild<Columns>();
        if (columns == null) return result;

        foreach (var col in columns.Elements<Column>())
        {
            if (col.Width?.Value == null) continue;
            var min = (int)(col.Min?.Value ?? 1u);
            var max = (int)(col.Max?.Value ?? (uint)min);
            // Hidden columns get width 0
            // Excel column width → pixels: chars * 7.0017; pt = px * 0.75
            var widthPt = col.Hidden?.Value == true ? 0 : (col.Width.Value == 0 ? 0 : col.Width.Value * 7.0017 * 0.75);
            for (int c = min; c <= max; c++)
                result[c] = widthPt;
        }

        return result;
    }

    // ==================== Frozen Panes ====================

    private static (int frozenRows, int frozenCols) GetFrozenPanes(Worksheet ws)
    {
        var sheetViews = ws.GetFirstChild<SheetViews>();
        var sheetView = sheetViews?.GetFirstChild<SheetView>();
        var pane = sheetView?.GetFirstChild<Pane>();
        if (pane == null) return (0, 0);

        // Only handle frozen panes (not split panes)
        if (pane.State?.Value != PaneStateValues.Frozen && pane.State?.Value != PaneStateValues.FrozenSplit)
            return (0, 0);

        var frozenRows = (int)(pane.VerticalSplit?.Value ?? 0);
        var frozenCols = (int)(pane.HorizontalSplit?.Value ?? 0);
        return (frozenRows, frozenCols);
    }

    // ==================== Conditional Formatting ====================

    /// <summary>
    /// Evaluate conditional formatting rules and return CSS overrides per cell.
    /// </summary>
    private Dictionary<string, string> BuildConditionalFormatMap(
        Worksheet ws, Stylesheet? stylesheet, SheetData sheetData, WorkbookPart? workbookPart)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        var cfElements = ws.Elements<ConditionalFormatting>().ToList();
        if (cfElements.Count == 0) return result;

        // Color-scale rules carry inline stop colors (no dxfId, no stylesheet);
        // render them in a separate pass so they work even when the workbook has
        // no <dxfs> or no stylesheet at all.
        AddColorScaleBackgrounds(cfElements, sheetData, result);

        // The remaining dxf-indexed rules need the stylesheet's <dxfs> catalogue.
        var dxfs = stylesheet?.DifferentialFormats?.Elements<DifferentialFormat>().ToArray();
        if (dxfs == null || dxfs.Length == 0) return result;

        var evaluator = new Core.FormulaEvaluator(sheetData, workbookPart);

        foreach (var cf in cfElements)
        {
            var sqref = cf.SequenceOfReferences?.Items?.ToList();
            if (sqref == null || sqref.Count == 0) continue;

            foreach (var rule in cf.Elements<ConditionalFormattingRule>())
            {
                var dxfId = rule.FormatId?.Value;
                if (dxfId == null || dxfId >= dxfs.Length) continue;
                var dxf = dxfs[(int)dxfId];

                // Extract CSS from dxf
                var cssParts = new List<string>();
                var fill = dxf.Fill?.PatternFill;
                if (fill != null)
                {
                    var bgColor = fill.BackgroundColor?.Rgb?.Value ?? fill.ForegroundColor?.Rgb?.Value;
                    if (bgColor != null)
                    {
                        if (bgColor.Length > 6) bgColor = bgColor[^6..];
                        cssParts.Add($"background:#{bgColor}");
                    }
                }
                var font = dxf.Font;
                if (font != null)
                {
                    var fontColor = font.Color?.Rgb?.Value;
                    if (fontColor != null)
                    {
                        if (fontColor.Length > 6) fontColor = fontColor[^6..];
                        cssParts.Add($"color:#{fontColor}");
                    }
                }
                if (cssParts.Count == 0) continue;
                var cssOverride = string.Join(";", cssParts);

                // Expand sqref and evaluate each cell
                foreach (var rangeStr in sqref)
                {
                    var cells = ExpandSqref(rangeStr.Value ?? "");
                    foreach (var (cellRef, row, col) in cells)
                    {
                        if (result.ContainsKey(cellRef)) continue; // first matching rule wins

                        bool matches = EvaluateCfRule(rule, cellRef, row, col, sheetData, evaluator);
                        if (matches)
                            result[cellRef] = cssOverride;
                    }
                }
            }
        }
        return result;
    }

    /// <summary>
    /// Color-scale CF rules (type="colorScale") carry inline &lt;cfvo&gt; stops and a
    /// matching list of &lt;color&gt; children; they never use a dxfId. Interpolate a
    /// per-cell background between the stop colors (2-stop min→max, or 3-stop
    /// min→mid→max) based on each cell's value, and write "background:#RRGGBB" into
    /// the shared CF map so the existing per-cell apply path picks it up.
    /// First matching rule wins, consistent with the dxf loop.
    /// </summary>
    private void AddColorScaleBackgrounds(
        List<ConditionalFormatting> cfElements, SheetData sheetData, Dictionary<string, string> result)
    {
        foreach (var cf in cfElements)
        {
            var sqref = cf.SequenceOfReferences?.Items?.ToList();
            if (sqref == null || sqref.Count == 0) continue;

            foreach (var rule in cf.Elements<ConditionalFormattingRule>())
            {
                var colorScale = rule.GetFirstChild<ColorScale>();
                if (colorScale == null) continue;

                var stops = colorScale.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>()
                    .Select(c => NormalizeScaleColor(c.Rgb?.Value))
                    .ToList();
                if (stops.Count < 2) continue;

                // Collect numeric cell values in range to derive min/max anchors.
                var cells = new List<(string cellRef, double value)>();
                foreach (var rangeStr in sqref)
                {
                    foreach (var (cellRef, row, col) in ExpandSqref(rangeStr.Value ?? ""))
                    {
                        var cell = sheetData.Descendants<Cell>()
                            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));
                        if (cell?.CellValue != null && double.TryParse(cell.CellValue.Text,
                            System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v))
                            cells.Add((cellRef, v));
                    }
                }
                if (cells.Count == 0) continue;

                double minVal = cells.Min(c => c.value);
                double maxVal = cells.Max(c => c.value);
                if (maxVal <= minVal) maxVal = minVal + 1;

                foreach (var (cellRef, value) in cells)
                {
                    if (result.ContainsKey(cellRef)) continue; // first matching rule wins
                    var t = (value - minVal) / (maxVal - minVal);
                    var rgb = InterpolateColorScale(stops, t);
                    result[cellRef] = $"background:#{rgb}";
                }
            }
        }
    }

    /// <summary>Strip an optional 8-hex ARGB prefix down to 6-hex RRGGBB; default black.</summary>
    private static string NormalizeScaleColor(string? argb)
    {
        if (string.IsNullOrEmpty(argb)) return "000000";
        return argb.Length > 6 ? argb[^6..] : argb;
    }

    /// <summary>
    /// Linearly interpolate across an ordered list of RRGGBB stops by fraction
    /// t in [0,1]. Two stops → single segment; three stops → min/mid(0.5)/max.
    /// </summary>
    private static string InterpolateColorScale(List<string> stops, double t)
    {
        t = Math.Max(0, Math.Min(1, t));
        // Map t onto the segment between stop[i] and stop[i+1].
        int segCount = stops.Count - 1;
        double scaled = t * segCount;
        int lo = Math.Min((int)scaled, segCount - 1);
        double frac = scaled - lo;
        var (r1, g1, b1) = HexToRgb(stops[lo]);
        var (r2, g2, b2) = HexToRgb(stops[lo + 1]);
        int r = (int)Math.Round(r1 + (r2 - r1) * frac);
        int g = (int)Math.Round(g1 + (g2 - g1) * frac);
        int b = (int)Math.Round(b1 + (b2 - b1) * frac);
        return $"{r:X2}{g:X2}{b:X2}";
    }

    private static (int r, int g, int b) HexToRgb(string hex)
    {
        int r = Convert.ToInt32(hex.Substring(0, 2), 16);
        int g = Convert.ToInt32(hex.Substring(2, 2), 16);
        int b = Convert.ToInt32(hex.Substring(4, 2), 16);
        return (r, g, b);
    }

    /// <summary>
    /// Build data bar info per cell: returns HTML for the bar overlay.
    /// </summary>
    private Dictionary<string, string> BuildDataBarMap(Worksheet ws, SheetData sheetData)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var cf in ws.Elements<ConditionalFormatting>())
        {
            foreach (var rule in cf.Elements<ConditionalFormattingRule>())
            {
                var dataBar = rule.GetFirstChild<DataBar>();
                if (dataBar == null) continue;

                var sqref = cf.SequenceOfReferences?.Items?.ToList();
                if (sqref == null || sqref.Count == 0) continue;

                // Get bar color
                var barColorEl = dataBar.GetFirstChild<Color>();
                var barColor = barColorEl?.Rgb?.Value ?? "FF4472C4";
                if (barColor.Length > 6) barColor = barColor[^6..];

                // Collect all cell values in range
                var cells = new List<(string cellRef, double value)>();
                foreach (var rangeStr in sqref)
                {
                    foreach (var (cellRef, row, col) in ExpandSqref(rangeStr.Value ?? ""))
                    {
                        var cell = sheetData.Descendants<Cell>()
                            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));
                        if (cell?.CellValue != null && double.TryParse(cell.CellValue.Text,
                            System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v))
                            cells.Add((cellRef, v));
                    }
                }
                if (cells.Count == 0) continue;

                // Determine min/max from cfvo elements or from data
                var cfvos = dataBar.Elements<ConditionalFormatValueObject>().ToList();
                double minVal, maxVal;
                var dataMin = cells.Min(c => c.value);
                if (cfvos.Count >= 2 && cfvos[0].Type?.Value == ConditionalFormatValueObjectValues.Number
                    && double.TryParse(cfvos[0].Val?.Value, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out var explicitMin))
                    minVal = explicitMin;
                else
                    // R17a: Excel anchors bars at 0 ONLY when all values are
                    // non-negative. With negatives present, the axis floor is the
                    // actual data minimum so negative magnitudes get proportional
                    // (left-extending) bars instead of clamping to 0% width.
                    minVal = Math.Min(0, dataMin);

                if (cfvos.Count >= 2 && cfvos[1].Type?.Value == ConditionalFormatValueObjectValues.Number
                    && double.TryParse(cfvos[1].Val?.Value, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out var explicitMax))
                    maxVal = explicitMax;
                else
                    maxVal = cells.Max(c => c.value);

                if (maxVal <= minVal) maxVal = minVal + 1;

                // Read bar length bounds (Excel defaults: min=10%, max=90%)
                var minLength = dataBar.MinLength?.Value ?? 10U;
                var maxLength = dataBar.MaxLength?.Value ?? 90U;
                var showValue = dataBar.ShowValue?.Value ?? true;

                // R17a: when the range straddles zero, draw a zero-axis and split
                // bars left/right of it. zeroPct is the axis position (0–100%).
                bool hasNegative = minVal < 0;
                var zeroPct = hasNegative ? (0 - minVal) / (maxVal - minVal) * 100 : 0;

                foreach (var (cellRef, value) in cells)
                {
                    string barDiv;
                    if (hasNegative)
                    {
                        // Width proportional to |value| over the full span; positive
                        // bars extend right from the zero-axis, negative bars left.
                        var wPct = Math.Min(100, Math.Abs(value) / (maxVal - minVal) * 100);
                        if (value >= 0)
                            barDiv = $"<div style=\"position:absolute;left:{zeroPct:0.#}%;top:1px;bottom:1px;width:{wPct:0.#}%;background:linear-gradient(to right,#{barColor},#{barColor}40);border-radius:1px\"></div>";
                        else
                            barDiv = $"<div style=\"position:absolute;left:{Math.Max(0, zeroPct - wPct):0.#}%;top:1px;bottom:1px;width:{wPct:0.#}%;background:linear-gradient(to left,#{barColor},#{barColor}40);border-radius:1px\"></div>";
                        // Zero-axis marker (thin line) — drawn once-style per cell is fine.
                        barDiv += $"<div style=\"position:absolute;left:{zeroPct:0.#}%;top:0;bottom:0;border-left:1px dashed #c0504d\"></div>";
                    }
                    else
                    {
                        var rawPct = (value - minVal) / (maxVal - minVal) * 100;
                        // Scale to minLength..maxLength range
                        var pct = Math.Max(0, Math.Min(100, minLength + rawPct / 100 * (maxLength - minLength)));
                        barDiv = $"<div style=\"position:absolute;left:0;top:1px;bottom:1px;width:{pct:0.#}%;background:linear-gradient(to right,#{barColor},#{barColor}40);border-radius:1px\"></div>";
                    }
                    // Store bar HTML + showValue flag (prefixed with "0|" or "1|")
                    result[cellRef] = $"{(showValue ? "1" : "0")}|{barDiv}";
                }
            }
        }
        return result;
    }

    /// <summary>
    /// Build icon set info per cell: returns HTML for the icon.
    /// </summary>
    private Dictionary<string, string> BuildIconSetMap(Worksheet ws, SheetData sheetData)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var cf in ws.Elements<ConditionalFormatting>())
        {
            foreach (var rule in cf.Elements<ConditionalFormattingRule>())
            {
                var iconSet = rule.GetFirstChild<IconSet>();
                if (iconSet == null) continue;

                var sqref = cf.SequenceOfReferences?.Items?.ToList();
                if (sqref == null || sqref.Count == 0) continue;

                var iconSetName = iconSet.IconSetValue?.Value ?? IconSetValues.ThreeTrafficLights1;
                var showValue = iconSet.ShowValue?.Value ?? true;
                var reverse = iconSet.Reverse?.Value ?? false;

                // Collect all cell values in range
                var cells = new List<(string cellRef, double value)>();
                foreach (var rangeStr in sqref)
                {
                    foreach (var (cellRef, row, col) in ExpandSqref(rangeStr.Value ?? ""))
                    {
                        var cell = sheetData.Descendants<Cell>()
                            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));
                        if (cell?.CellValue != null && double.TryParse(cell.CellValue.Text,
                            System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v))
                            cells.Add((cellRef, v));
                    }
                }
                if (cells.Count == 0) continue;

                // Parse cfvo thresholds
                var cfvos = iconSet.Elements<ConditionalFormatValueObject>().ToList();
                var allValues = cells.Select(c => c.value).OrderBy(v => v).ToList();
                double minVal = allValues.First(), maxVal = allValues.Last();
                var range = maxVal - minVal;
                if (range == 0) range = 1;

                // Resolve thresholds (skip first cfvo which is the base)
                var thresholds = new List<double>();
                for (int i = 1; i < cfvos.Count; i++)
                {
                    var cfvo = cfvos[i];
                    var type = cfvo.Type?.Value ?? ConditionalFormatValueObjectValues.Percent;
                    double.TryParse(cfvo.Val?.Value, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out var tv);
                    if (type == ConditionalFormatValueObjectValues.Number)
                        thresholds.Add(tv);
                    else if (type == ConditionalFormatValueObjectValues.Percent)
                        thresholds.Add(minVal + range * tv / 100);
                    else if (type == ConditionalFormatValueObjectValues.Percentile)
                    {
                        var idx = (int)Math.Round(tv / 100.0 * (allValues.Count - 1));
                        thresholds.Add(allValues[Math.Clamp(idx, 0, allValues.Count - 1)]);
                    }
                    else
                        thresholds.Add(minVal + range * tv / 100);
                }

                foreach (var (cellRef, value) in cells)
                {
                    // Determine which bucket the value falls into
                    int bucket = 0;
                    for (int i = 0; i < thresholds.Count; i++)
                    {
                        if (value >= thresholds[i]) bucket = i + 1;
                    }
                    if (reverse) bucket = cfvos.Count - 1 - bucket;
                    var icon = GetIconHtml(iconSetName, bucket, cfvos.Count);
                    // Prefix with showValue flag: "0|" = hide value, "1|" = show value
                    result[cellRef] = $"{(showValue ? "1" : "0")}|{icon}";
                }
            }
        }
        return result;
    }

    /// <summary>
    /// R12a: build a cellRef → inline-SVG map for sparklines. Sparkline groups
    /// live in the worksheet extension list (x14:sparklineGroups); each
    /// x14:sparkline carries a Formula (data range) and a ReferenceSequence
    /// (the host cell). We render a small SVG into that host cell — column =
    /// proportional bars, line/stacked = polyline — so the cell is no longer
    /// blank. Lightweight approximation of the native in-cell mini chart.
    /// </summary>
    private Dictionary<string, string> BuildSparklineMap(Worksheet ws)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var extList = ws.GetFirstChild<WorksheetExtensionList>();
        if (extList == null) return result;
        foreach (var group in extList.Descendants<X14.SparklineGroup>())
        {
            var type = group.Type?.Value;
            var kind = type == X14.SparklineTypeValues.Column ? "column"
                     : type == X14.SparklineTypeValues.Stacked ? "stacked"
                     : "line";
            foreach (var spk in group.Descendants<X14.Sparkline>())
            {
                var dataRange = spk.Formula?.Text;
                var hostCell = spk.ReferenceSequence?.Text;
                if (string.IsNullOrEmpty(dataRange) || string.IsNullOrEmpty(hostCell)) continue;
                var values = ReadCellRangeAsDoubles(dataRange);
                if (values == null || values.Length == 0) continue;
                // host cell may be a range like "F1:F1" — take the first ref.
                var host = hostCell.Contains(':') ? hostCell.Split(':')[0] : hostCell;
                host = host.Contains('!') ? host.Split('!')[1] : host;
                result[host] = BuildSparklineSvg(values, kind);
            }
        }
        return result;
    }

    /// <summary>Render a sparkline's values as a small inline SVG (~80x20px).</summary>
    private static string BuildSparklineSvg(double[] values, string kind)
    {
        const double w = 80, h = 18;
        var min = values.Min();
        var max = values.Max();
        var range = max - min;
        if (range == 0) range = 1;
        // Baseline at zero when data straddles it, else at the value floor.
        double zeroFloor = Math.Min(min, 0);
        double zeroRange = Math.Max(max, 0) - zeroFloor;
        if (zeroRange == 0) zeroRange = 1;
        var sb = new StringBuilder();
        sb.Append($"<svg class=\"sparkline\" width=\"{w:0}\" height=\"{h:0}\" viewBox=\"0 0 {w:0} {h:0}\" preserveAspectRatio=\"none\" style=\"vertical-align:middle\">");
        if (kind == "column" || kind == "stacked")
        {
            int n = values.Length;
            double bw = w / n;
            for (int i = 0; i < n; i++)
            {
                var v = values[i];
                // bar from the zero line; positive up, negative down.
                var zeroY = h - (0 - zeroFloor) / zeroRange * h;
                var valY = h - (v - zeroFloor) / zeroRange * h;
                var top = Math.Min(zeroY, valY);
                var bh = Math.Max(1, Math.Abs(valY - zeroY));
                var color = v < 0 ? "#C0504D" : "#4472C4";
                sb.Append($"<rect x=\"{i * bw + 0.5:0.#}\" y=\"{top:0.#}\" width=\"{Math.Max(1, bw - 1):0.#}\" height=\"{bh:0.#}\" fill=\"{color}\"/>");
            }
        }
        else
        {
            int n = values.Length;
            var pts = new List<string>();
            for (int i = 0; i < n; i++)
            {
                var x = n > 1 ? (double)i / (n - 1) * w : w / 2;
                var y = h - (values[i] - min) / range * h;
                pts.Add($"{x:0.#},{y:0.#}");
            }
            sb.Append($"<polyline points=\"{string.Join(" ", pts)}\" fill=\"none\" stroke=\"#4472C4\" stroke-width=\"1\"/>");
        }
        sb.Append("</svg>");
        return sb.ToString();
    }

    private static string GetIconHtml(IconSetValues iconSetName, int bucket, int totalBuckets)
    {
        // Traffic lights: red=0, yellow=1, green=2
        if (iconSetName == IconSetValues.ThreeTrafficLights1 || iconSetName == IconSetValues.ThreeTrafficLights2)
        {
            var color = bucket switch { 0 => "#C00000", 1 => "#FFC000", _ => "#00B050" };
            return $"<span style=\"display:inline-block;width:10px;height:10px;border-radius:50%;background:{color};margin-right:4px;vertical-align:middle\"></span>";
        }
        // Arrows
        if (iconSetName == IconSetValues.ThreeArrows || iconSetName == IconSetValues.ThreeArrowsGray)
        {
            return bucket switch
            {
                0 => "<span style=\"color:#C00000;margin-right:4px;vertical-align:middle\">&#x25BC;</span>",
                1 => "<span style=\"color:#FFC000;margin-right:4px;vertical-align:middle\">&#x25B6;</span>",
                _ => "<span style=\"color:#00B050;margin-right:4px;vertical-align:middle\">&#x25B2;</span>",
            };
        }
        // R17b: directional arrow icon sets (4/5 arrows). Native renders
        // graduated arrows ↓↘→↗↑; previously these fell through to the default
        // colored circle. Glyphs: ↓ U+2193, ↘ U+2198, → U+2192, ↗ U+2197, ↑ U+2191.
        if (iconSetName == IconSetValues.FiveArrows || iconSetName == IconSetValues.FiveArrowsGray)
        {
            var gray = iconSetName == IconSetValues.FiveArrowsGray;
            var glyph = bucket switch { 0 => "&#x2193;", 1 => "&#x2198;", 2 => "&#x2192;", 3 => "&#x2197;", _ => "&#x2191;" };
            var color = gray ? "#808080" : bucket switch { 0 => "#C00000", 1 => "#E08000", 2 => "#FFC000", 3 => "#92D050", _ => "#00B050" };
            return $"<span style=\"color:{color};margin-right:4px;vertical-align:middle\">{glyph}</span>";
        }
        if (iconSetName == IconSetValues.FourArrows || iconSetName == IconSetValues.FourArrowsGray)
        {
            var gray = iconSetName == IconSetValues.FourArrowsGray;
            var glyph = bucket switch { 0 => "&#x2193;", 1 => "&#x2198;", 2 => "&#x2197;", _ => "&#x2191;" };
            var color = gray ? "#808080" : bucket switch { 0 => "#C00000", 1 => "#E08000", 2 => "#92D050", _ => "#00B050" };
            return $"<span style=\"color:{color};margin-right:4px;vertical-align:middle\">{glyph}</span>";
        }
        // 4-icon traffic lights
        if (iconSetName == IconSetValues.FourTrafficLights)
        {
            var color = bucket switch { 0 => "#C00000", 1 => "#FFC000", 2 => "#92D050", _ => "#00B050" };
            return $"<span style=\"display:inline-block;width:10px;height:10px;border-radius:50%;background:{color};margin-right:4px;vertical-align:middle\"></span>";
        }
        // Default: colored circles
        if (totalBuckets <= 3)
        {
            var color = bucket switch { 0 => "#C00000", 1 => "#FFC000", _ => "#00B050" };
            return $"<span style=\"display:inline-block;width:10px;height:10px;border-radius:50%;background:{color};margin-right:4px;vertical-align:middle\"></span>";
        }
        else
        {
            var pct = totalBuckets > 1 ? (double)bucket / (totalBuckets - 1) : 1;
            var r = (int)(0xC0 * (1 - pct));
            var g = (int)(0xB0 * pct);
            var color = $"#{r:X2}{g:X2}00";
            return $"<span style=\"display:inline-block;width:10px;height:10px;border-radius:50%;background:{color};margin-right:4px;vertical-align:middle\"></span>";
        }
    }

    /// <summary>Evaluate whether a conditional formatting rule matches a specific cell.</summary>
    private bool EvaluateCfRule(ConditionalFormattingRule rule, string cellRef, int row, int col,
        SheetData sheetData, Core.FormulaEvaluator evaluator)
    {
        var ruleType = rule.Type?.Value;

        // Get cell value for comparison
        double? cellValue = null;
        var cell = sheetData.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));
        if (cell != null)
        {
            if (double.TryParse(cell.CellValue?.Text, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var v))
                cellValue = v;
        }

        if (ruleType == ConditionalFormatValues.Expression)
        {
            // Formula-based rule: evaluate with cell reference adjustment
            var formula = rule.Elements<Formula>().FirstOrDefault()?.Text;
            if (string.IsNullOrEmpty(formula)) return false;

            // Adjust formula references relative to the first cell in sqref
            // The formula is written for the top-left cell; adjust for current cell
            var adjusted = AdjustCfFormula(formula, row, col, rule);
            var result = evaluator.TryEvaluateFull(adjusted);
            return result?.BoolValue == true || (result?.NumericValue != null && result.NumericValue != 0);
        }

        if (ruleType == ConditionalFormatValues.CellIs && cellValue.HasValue)
        {
            var op = rule.Operator?.Value;
            var f1 = rule.Elements<Formula>().FirstOrDefault()?.Text;
            var f2 = rule.Elements<Formula>().Skip(1).FirstOrDefault()?.Text;
            double? v1 = f1 != null ? evaluator.TryEvaluate(f1) ?? (double.TryParse(f1, out var p1) ? p1 : null) : null;
            double? v2 = f2 != null ? evaluator.TryEvaluate(f2) ?? (double.TryParse(f2, out var p2) ? p2 : null) : null;
            if (v1 == null) return false;
            if (op == ConditionalFormattingOperatorValues.GreaterThan) return cellValue > v1;
            if (op == ConditionalFormattingOperatorValues.LessThan) return cellValue < v1;
            if (op == ConditionalFormattingOperatorValues.GreaterThanOrEqual) return cellValue >= v1;
            if (op == ConditionalFormattingOperatorValues.LessThanOrEqual) return cellValue <= v1;
            if (op == ConditionalFormattingOperatorValues.Equal) return cellValue == v1;
            if (op == ConditionalFormattingOperatorValues.NotEqual) return cellValue != v1;
            if (op == ConditionalFormattingOperatorValues.Between) return v2.HasValue && cellValue >= v1 && cellValue <= v2;
            if (op == ConditionalFormattingOperatorValues.NotBetween) return v2.HasValue && (cellValue < v1 || cellValue > v2);
            return false;
        }

        if (ruleType == ConditionalFormatValues.Top10 && cellValue.HasValue)
        {
            // Top/bottom N (or N%) of the numeric values in the rule's range.
            // Mirrors Excel: rank=N, percent=true means N% of cells, bottom=true
            // ranks ascending. A cell matches if it falls within that slice.
            var nums = CollectCfRangeNumbers(rule, sheetData);
            if (nums.Count == 0) return false;
            var rank = (int)(rule.Rank?.Value ?? 10);
            var bottom = rule.Bottom?.Value == true;
            var percent = rule.Percent?.Value == true;
            var count = percent
                ? (int)Math.Ceiling(nums.Count * (rank / 100.0))
                : rank;
            if (count < 1) count = 1;
            if (count > nums.Count) count = nums.Count;
            var ordered = bottom
                ? nums.OrderBy(n => n).ToList()
                : nums.OrderByDescending(n => n).ToList();
            // Threshold = the value at the Nth position; include ties (Excel
            // colors all cells at/beyond the cutoff value, like rank ties).
            var threshold = ordered[count - 1];
            return bottom ? cellValue <= threshold : cellValue >= threshold;
        }

        if (ruleType == ConditionalFormatValues.DuplicateValues || ruleType == ConditionalFormatValues.UniqueValues)
        {
            // Color cells whose value appears more than once (duplicateValues)
            // or exactly once (uniqueValues) within the rule's range. Compare
            // on the raw cell text so non-numeric values are handled too.
            var thisText = cell?.CellValue?.Text;
            if (string.IsNullOrEmpty(thisText)) return false;
            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var t in CollectCfRangeTexts(rule, sheetData))
                counts[t] = counts.TryGetValue(t, out var c) ? c + 1 : 1;
            var occurrences = counts.TryGetValue(thisText, out var n) ? n : 0;
            return ruleType == ConditionalFormatValues.DuplicateValues
                ? occurrences > 1
                : occurrences == 1;
        }

        return false;
    }

    /// <summary>Collect the numeric values of every cell in a CF rule's sqref range.</summary>
    private List<double> CollectCfRangeNumbers(ConditionalFormattingRule rule, SheetData sheetData)
    {
        var result = new List<double>();
        foreach (var c in CollectCfRangeCells(rule, sheetData))
        {
            if (double.TryParse(c.CellValue?.Text, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var v))
                result.Add(v);
        }
        return result;
    }

    /// <summary>Collect the raw text of every non-empty cell in a CF rule's sqref range.</summary>
    private List<string> CollectCfRangeTexts(ConditionalFormattingRule rule, SheetData sheetData)
    {
        var result = new List<string>();
        foreach (var c in CollectCfRangeCells(rule, sheetData))
        {
            var t = c.CellValue?.Text;
            if (!string.IsNullOrEmpty(t)) result.Add(t);
        }
        return result;
    }

    /// <summary>Enumerate the cells that exist within a CF rule's sqref range(s).</summary>
    private IEnumerable<Cell> CollectCfRangeCells(ConditionalFormattingRule rule, SheetData sheetData)
    {
        var cf = rule.Parent as ConditionalFormatting;
        var sqrefs = cf?.SequenceOfReferences?.Items;
        if (sqrefs == null) yield break;
        foreach (var sqItem in sqrefs)
        {
            var sqref = sqItem?.Value;
            if (string.IsNullOrEmpty(sqref)) continue;
            var start = sqref.Contains(':') ? sqref.Split(':')[0] : sqref;
            var end = sqref.Contains(':') ? sqref.Split(':')[1] : sqref;
            var (startColName, startRow) = ParseCellReference(start);
            var (endColName, endRow) = ParseCellReference(end);
            int c1 = ColumnNameToIndex(startColName), c2 = ColumnNameToIndex(endColName);
            int r1 = Math.Min(startRow, endRow), r2 = Math.Max(startRow, endRow);
            int cMin = Math.Min(c1, c2), cMax = Math.Max(c1, c2);
            foreach (var cell in sheetData.Descendants<Cell>())
            {
                var refv = cell.CellReference?.Value;
                if (string.IsNullOrEmpty(refv)) continue;
                var (colName, rowNum) = ParseCellReference(refv);
                var colIdx = ColumnNameToIndex(colName);
                if (rowNum >= r1 && rowNum <= r2 && colIdx >= cMin && colIdx <= cMax)
                    yield return cell;
            }
        }
    }

    /// <summary>Adjust a CF formula's cell references from the anchor cell to the target cell.</summary>
    private string AdjustCfFormula(string formula, int targetRow, int targetCol, ConditionalFormattingRule rule)
    {
        // Find the anchor cell from the parent ConditionalFormatting sqref
        var cf = rule.Parent as ConditionalFormatting;
        var sqref = cf?.SequenceOfReferences?.Items?.FirstOrDefault()?.Value;
        if (string.IsNullOrEmpty(sqref)) return formula;

        // Extract anchor from sqref (e.g. "E7:E21" → anchor is E7)
        var anchorRef = sqref.Contains(':') ? sqref.Split(':')[0] : sqref;
        var (anchorColName, anchorRow) = ParseCellReference(anchorRef);
        var anchorCol = ColumnNameToIndex(anchorColName);

        var rowDelta = targetRow - anchorRow;
        var colDelta = targetCol - anchorCol;
        if (rowDelta == 0 && colDelta == 0) return formula;

        // Replace cell references in formula, adjusting by delta
        return Regex.Replace(formula, @"(\$?)([A-Z]+)(\$?)(\d+)", m =>
        {
            var colAbsolute = m.Groups[1].Value == "$";
            var rowAbsolute = m.Groups[3].Value == "$";
            var refCol = ColumnNameToIndex(m.Groups[2].Value);
            var refRow = int.Parse(m.Groups[4].Value);

            var newCol = colAbsolute ? refCol : refCol + colDelta;
            var newRow = rowAbsolute ? refRow : refRow + rowDelta;
            if (newCol < 1) newCol = 1;
            if (newRow < 1) newRow = 1;
            return $"{(colAbsolute ? "$" : "")}{IndexToColumnName(newCol)}{(rowAbsolute ? "$" : "")}{newRow}";
        });
    }

    /// <summary>Expand a sqref string like "E7:E21" into individual cell references.</summary>
    private List<(string cellRef, int row, int col)> ExpandSqref(string sqref)
    {
        var result = new List<(string, int, int)>();
        foreach (var part in sqref.Split(' '))
        {
            if (part.Contains(':'))
            {
                var sides = part.Split(':');
                var (startColName, startRow) = ParseCellReference(sides[0]);
                var (endColName, endRow) = ParseCellReference(sides[1]);
                var startCol = ColumnNameToIndex(startColName);
                var endCol = ColumnNameToIndex(endColName);
                for (int r = startRow; r <= endRow; r++)
                    for (int c = startCol; c <= endCol; c++)
                        result.Add(($"{IndexToColumnName(c)}{r}", r, c));
            }
            else
            {
                var (colName, row) = ParseCellReference(part);
                result.Add((part, row, ColumnNameToIndex(colName)));
            }
        }
        return result;
    }

    // ==================== Cell Style to CSS ====================

    // CONSISTENCY(excel-merge-border): for a merged anchor td, the right/bottom CSS
    // edges of the td are the merged region PERIMETER, which Excel sources from the
    // perimeter member cells — not the anchor (whose right/bottom are interior to the
    // merge). Callers pass rightBorderCell (right-column member) and bottomBorderCell
    // (bottom-row member); default null = same as the anchor (non-merge path unchanged).
    private string GetCellStyleCss(Cell? cell, Stylesheet? stylesheet, int frozenRows, int frozenCols, int row, int col,
        Dictionary<int, double>? frozenLeftOffsets = null, Dictionary<int, double>? frozenTopOffsets = null,
        Dictionary<string, string>? cfMap = null, Dictionary<string, string>? dataBarMap = null,
        Dictionary<string, string>? iconSetMap = null,
        bool mergePerimeter = false, Cell? rightBorderCell = null, Cell? bottomBorderCell = null)
    {
        var styles = new List<string>();

        // Frozen pane sticky positioning
        bool isFrozenRow = frozenRows > 0 && row <= frozenRows;
        bool isFrozenCol = frozenCols > 0 && col <= frozenCols;
        // z-index layering: corner-cell=4, col-header=3, frozen-row+col=2, frozen-col=1
        var frozenLeft = frozenLeftOffsets?.TryGetValue(col, out var fl) == true ? fl : 0;
        var frozenTop = frozenTopOffsets?.TryGetValue(row, out var ft) == true ? ft : 0;
        if (isFrozenRow && isFrozenCol)
            styles.Add($"position:sticky;top:0;left:{frozenLeft:0.##}pt;z-index:2");
        else if (isFrozenRow)
            styles.Add("position:sticky;top:0;z-index:1");
        else if (isFrozenCol)
            styles.Add($"position:sticky;left:{frozenLeft:0.##}pt;z-index:1");

        if (cell == null || stylesheet == null)
        {
            // CF color-scale backgrounds carry inline colors and don't need the
            // stylesheet — apply them here too so a workbook with no <dxfs>/styles
            // (e.g. a freshly-created file) still renders the gradient.
            var cfRefEarly = $"{IndexToColumnName(col)}{row}";
            if (cfMap != null && cfMap.TryGetValue(cfRefEarly, out var cfCssEarly))
            {
                foreach (var cfPart in cfCssEarly.Split(';'))
                    styles.RemoveAll(s => s.StartsWith(cfPart.Split(':')[0].Trim() + ":"));
                styles.Add(cfCssEarly);
            }
            // Data bar / icon set need position:relative on the TD so their inner
            // absolutely-positioned div anchors to the cell, not the sheet wrapper.
            // This must run here too — a freshly-created xlsx has no stylesheet and
            // every cell hits this early-return branch.
            if ((dataBarMap != null && dataBarMap.ContainsKey(cfRefEarly)) ||
                (iconSetMap != null && iconSetMap.ContainsKey(cfRefEarly)))
            {
                styles.Add("position:relative");
            }
            // Frozen rows need opaque background so scrolling content doesn't show through
            // Use actual cell fill if available; fallback to white for cells with no explicit fill
            if (isFrozenRow && !styles.Any(s => s.StartsWith("background")))
                styles.Add("background:#fff");
            return styles.Count > 0 ? $" style=\"{string.Join(";", styles)}\"" : "";
        }

        var styleIndex = cell.StyleIndex?.Value ?? 0;

        {
            var cellFormats = stylesheet.CellFormats;
            if (cellFormats != null && styleIndex < (uint)cellFormats.Elements<CellFormat>().Count())
            {
                var xf = cellFormats.Elements<CellFormat>().ElementAt((int)styleIndex);
                BuildFontCss(xf, stylesheet, styles);
                BuildFillCss(xf, stylesheet, styles);
                BuildBorderCss(xf, stylesheet, styles, mergePerimeter, rightBorderCell, bottomBorderCell);
                BuildAlignmentCss(xf, styles, cell);

                // Number-format [Color] section (e.g. "$#,##0.00;[Red](...)" colors
                // negatives red). Applies to numeric cells only; the section is
                // chosen by the cell's value sign. Overrides the font color.
                var numFmtColor = GetCellNumberFormatColor(cell, xf, stylesheet);
                if (numFmtColor != null)
                {
                    styles.RemoveAll(s => s.StartsWith("color:"));
                    styles.Add($"color:{numFmtColor}");
                }

                // Diagonal border needs the TD to be a positioning context for the
                // inline SVG overlay emitted into the cell content (BuildRowInnerHtml).
                if (TryBuildCellDiagonalSvg(cell, stylesheet) != null
                    && !styles.Any(s => s.StartsWith("position:")))
                    styles.Add("position:relative");
            }
        }

        // Conditional formatting overrides (background, color)
        var cfCellRef = $"{IndexToColumnName(col)}{row}";
        if (cfMap != null && cfMap.TryGetValue(cfCellRef, out var cfCss))
        {
            // CF overrides existing background/color — remove conflicting base styles
            foreach (var cfPart in cfCss.Split(';'))
            {
                var prop = cfPart.Split(':')[0].Trim();
                styles.RemoveAll(s => s.StartsWith(prop + ":"));
            }
            styles.Add(cfCss);
        }

        // Data bar or icon set: add position:relative so inner elements can be absolutely positioned
        if ((dataBarMap != null && dataBarMap.ContainsKey(cfCellRef)) ||
            (iconSetMap != null && iconSetMap.ContainsKey(cfCellRef)))
        {
            styles.Add("position:relative");
        }

        // Frozen rows need opaque background so scrolling content doesn't show through
        if (isFrozenRow && !styles.Any(s => s.StartsWith("background:")))
            styles.Add("background:#fff");

        return styles.Count > 0 ? $" style=\"{string.Join(";", styles)}\"" : "";
    }

    private void BuildFontCss(CellFormat xf, Stylesheet stylesheet, List<string> styles)
    {
        var fontId = xf.FontId?.Value ?? 0;
        var fonts = stylesheet.Fonts;
        if (fonts == null || fontId >= (uint)fonts.Elements<Font>().Count()) return;

        var font = fonts.Elements<Font>().ElementAt((int)fontId);

        if (font.Bold != null && font.Bold.Val?.Value != false) styles.Add("font-weight:bold");
        if (font.Italic != null && font.Italic.Val?.Value != false) styles.Add("font-style:italic");
        if (font.Strike != null && font.Strike.Val?.Value != false) styles.Add("text-decoration:line-through");
        if (font.Underline != null)
        {
            var existing = styles.FindIndex(s => s.StartsWith("text-decoration:"));
            if (existing >= 0)
                styles[existing] = styles[existing] + " underline";
            else
                styles.Add("text-decoration:underline");
            // Render double / doubleAccounting as a true double underline.
            var ulVal = font.Underline.Val?.Value;
            if (ulVal == UnderlineValues.Double || ulVal == UnderlineValues.DoubleAccounting)
                styles.Add("text-decoration-style:double");
        }

        // Superscript/Subscript: handled by wrapping cell content in <sup>/<sub>
        // (see GetCellVerticalAlign + the <td> content path). vertical-align on a
        // <td> only controls cell-content block alignment (top/middle/bottom) — it
        // does NOT raise/lower the baseline of inline text, and the font-size:smaller
        // it paired with was overwritten by the full font-size:Npt added below.
        // R10a: emit nothing here; the <sup>/<sub> wrapper provides both the
        // raised baseline and the size reduction.

        if (font.FontSize?.Val?.Value != null)
            styles.Add($"font-size:{font.FontSize.Val.Value:0.##}pt");

        if (font.FontName?.Val?.Value != null)
            styles.Add($"font-family:'{CssSanitize(font.FontName.Val.Value)}'");

        var color = ResolveFontColor(font);
        if (color != null) styles.Add($"color:{color}");
    }

    /// <summary>
    /// R10a: returns "super" / "sub" / null for a cell's font vertical alignment.
    /// Used to wrap cell content in a <sup>/<sub> inline element — vertical-align
    /// on the <td> itself has no baseline-shifting effect on cell content.
    /// </summary>
    private static string? GetCellVerticalAlign(Cell? cell, Stylesheet? stylesheet)
    {
        if (cell == null || stylesheet?.CellFormats == null || stylesheet.Fonts == null) return null;
        var styleIndex = cell.StyleIndex?.Value ?? 0;
        if (styleIndex >= (uint)stylesheet.CellFormats.Elements<CellFormat>().Count()) return null;
        var xf = stylesheet.CellFormats.Elements<CellFormat>().ElementAt((int)styleIndex);
        var fontId = xf.FontId?.Value ?? 0;
        if (fontId >= (uint)stylesheet.Fonts.Elements<Font>().Count()) return null;
        var font = stylesheet.Fonts.Elements<Font>().ElementAt((int)fontId);
        var v = font.GetFirstChild<VerticalTextAlignment>()?.Val?.Value;
        if (v == VerticalAlignmentRunValues.Superscript) return "super";
        if (v == VerticalAlignmentRunValues.Subscript) return "sub";
        return null;
    }

    /// <summary>
    /// R10a: wrap cell content in a <sup>/<sub> element when the cell font is
    /// super/subscript. Skipped for rich text (its runs carry their own
    /// formatting) and when there's no content. <sup>/<sub> give both the
    /// raised/lowered baseline and the ~0.83em size reduction natively.
    /// </summary>
    private static string WrapVerticalAlign(string content, string? vAlign, string? richHtml)
    {
        if (vAlign == null || richHtml != null || string.IsNullOrEmpty(content)) return content;
        var tag = vAlign == "super" ? "sup" : "sub";
        return $"<{tag}>{content}</{tag}>";
    }

    private void BuildFillCss(CellFormat xf, Stylesheet stylesheet, List<string> styles)
    {
        var fillId = xf.FillId?.Value ?? 0;
        if (fillId <= 1) return; // 0=none, 1=gray125 pattern (default)

        var fills = stylesheet.Fills;
        if (fills == null || fillId >= (uint)fills.Elements<Fill>().Count()) return;

        var fill = fills.Elements<Fill>().ElementAt((int)fillId);

        // Gradient fill
        var gf = fill.GetFirstChild<GradientFill>();
        if (gf != null)
        {
            var stops = gf.Elements<GradientStop>().ToList();
            if (stops.Count >= 2)
            {
                var colors = stops
                    .Select(s => ResolveColorRgb(s.Color))
                    .Where(c => c != null)
                    .ToList();
                if (colors.Count >= 2)
                {
                    var deg = (int)(gf.Degree?.Value ?? 0);
                    styles.Add($"background:linear-gradient({(deg + 90) % 360}deg,{string.Join(",", colors)})");
                    return;
                }
            }
        }

        // Pattern fill
        var pf = fill.PatternFill;
        if (pf != null)
        {
            var bgColor = ResolveColorRgb(pf.ForegroundColor);
            if (bgColor != null) styles.Add($"background:{bgColor}");
        }
    }

    private void BuildBorderCss(CellFormat xf, Stylesheet stylesheet, List<string> styles,
        bool mergePerimeter = false, Cell? rightBorderCell = null, Cell? bottomBorderCell = null)
    {
        var borderId = xf.BorderId?.Value ?? 0;
        var borders = stylesheet.Borders;
        Border? border = (borderId != 0 && borders != null && borderId < (uint)borders.Elements<Border>().Count())
            ? borders.Elements<Border>().ElementAt((int)borderId)
            : null;

        // top/left always come from the anchor cell (top-left member of the merge).
        AddBorderSideCss(border?.TopBorder, "top", styles);
        AddBorderSideCss(border?.LeftBorder, "left", styles);

        // right/bottom: for a merged anchor the td edge is the region PERIMETER, which
        // Excel sources from the perimeter member cell — not the (interior) anchor edge.
        // If the member cell is absent (or carries no border) the region has NO border on
        // that edge. Non-merge path (mergePerimeter=false) keeps the anchor's own edges.
        var rightBorder = mergePerimeter ? (rightBorderCell != null ? GetCellBorder(rightBorderCell, stylesheet) : null) : border;
        var bottomBorder = mergePerimeter ? (bottomBorderCell != null ? GetCellBorder(bottomBorderCell, stylesheet) : null) : border;
        AddBorderSideCss(rightBorder?.RightBorder, "right", styles);
        AddBorderSideCss(bottomBorder?.BottomBorder, "bottom", styles);
    }

    // Resolve a cell's <border> element via its style index, or null if none.
    private Border? GetCellBorder(Cell cell, Stylesheet stylesheet)
    {
        var styleIndex = cell.StyleIndex?.Value ?? 0;
        var cellFormats = stylesheet.CellFormats;
        if (cellFormats == null || styleIndex >= (uint)cellFormats.Elements<CellFormat>().Count()) return null;
        var xf = cellFormats.Elements<CellFormat>().ElementAt((int)styleIndex);
        var borderId = xf.BorderId?.Value ?? 0;
        var borders = stylesheet.Borders;
        if (borderId == 0 || borders == null || borderId >= (uint)borders.Elements<Border>().Count()) return null;
        return borders.Elements<Border>().ElementAt((int)borderId);
    }

    /// <summary>
    /// Look up the cell's &lt;border&gt; and, if it carries a diagonal
    /// (diagonalDown / diagonalUp with a styled &lt;diagonal&gt; child), return an
    /// absolutely-positioned inline SVG that draws the diagonal line(s) inside the
    /// TD. Mirrors the PPTX table diagonal-overlay idiom (cell-diag SVG). Returns
    /// null when the cell has no diagonal border. The TD must be position:relative
    /// for the overlay to anchor to the cell (added in GetCellStyleCss).
    /// </summary>
    private string? TryBuildCellDiagonalSvg(Cell? cell, Stylesheet? stylesheet)
    {
        if (cell == null || stylesheet == null) return null;
        var styleIndex = cell.StyleIndex?.Value ?? 0;
        var cellFormats = stylesheet.CellFormats;
        if (cellFormats == null || styleIndex >= (uint)cellFormats.Elements<CellFormat>().Count())
            return null;
        var xf = cellFormats.Elements<CellFormat>().ElementAt((int)styleIndex);
        var borderId = xf.BorderId?.Value ?? 0;
        var borders = stylesheet.Borders;
        if (borders == null || borderId == 0 || borderId >= (uint)borders.Elements<Border>().Count())
            return null;
        var border = borders.Elements<Border>().ElementAt((int)borderId);

        bool down = border.DiagonalDown?.Value == true;
        bool up = border.DiagonalUp?.Value == true;
        if (!down && !up) return null;

        var diag = border.DiagonalBorder;
        if (diag?.Style?.Value == null || diag.Style.Value == BorderStyleValues.None) return null;

        var bsv = diag.Style.Value;
        double widthPx = bsv == BorderStyleValues.Thick ? 3 : bsv == BorderStyleValues.Medium ? 2 : 1;
        var color = ResolveColorRgb(diag.Color) ?? "#000";

        var lines = new StringBuilder();
        // diagonalDown = top-left → bottom-right; diagonalUp = bottom-left → top-right.
        if (down)
            lines.Append($"<line x1=\"0\" y1=\"0\" x2=\"100%\" y2=\"100%\" stroke=\"{color}\" stroke-width=\"{widthPx:0.##}\"/>");
        if (up)
            lines.Append($"<line x1=\"0\" y1=\"100%\" x2=\"100%\" y2=\"0\" stroke=\"{color}\" stroke-width=\"{widthPx:0.##}\"/>");

        return $"<svg class=\"cell-diag\" width=\"100%\" height=\"100%\" style=\"position:absolute;inset:0;pointer-events:none;overflow:visible\" preserveAspectRatio=\"none\">{lines}</svg>";
    }

    private void AddBorderSideCss(BorderPropertiesType? bp, string side, List<string> styles)
    {
        if (bp?.Style?.Value == null || bp.Style.Value == BorderStyleValues.None) return;

        var bsv = bp.Style.Value;
        var width = "1px";
        if (bsv == BorderStyleValues.Medium) width = "2px";
        else if (bsv == BorderStyleValues.Thick) width = "3px";
        else if (bsv == BorderStyleValues.Double) width = "3px";

        var cssStyle = "solid";
        if (bsv == BorderStyleValues.Dashed || bsv == BorderStyleValues.MediumDashed) cssStyle = "dashed";
        else if (bsv == BorderStyleValues.Dotted) cssStyle = "dotted";
        else if (bsv == BorderStyleValues.Double) cssStyle = "double";

        var color = ResolveColorRgb(bp.Color);
        color ??= "#000";

        styles.Add($"border-{side}:{width} {cssStyle} {color}");
    }

    private void BuildAlignmentCss(CellFormat xf, List<string> styles, Cell? cell = null)
    {
        var alignment = xf.Alignment;
        bool hasExplicitHAlign = alignment?.Horizontal?.HasValue == true;

        if (hasExplicitHAlign)
        {
            var h = alignment!.Horizontal!.InnerText;
            var cssAlign = h switch
            {
                "center" => "center",
                "right" => "right",
                "left" => "left",
                "justify" => "justify",
                "fill" => "left",
                "general" => (string?)null, // fall through to auto-detect
                _ => null
            };
            if (cssAlign != null) { styles.Add($"text-align:{cssAlign}"); hasExplicitHAlign = true; }
            else hasExplicitHAlign = false;
        }

        // Excel default: numbers right-aligned, text left-aligned (General alignment)
        if (!hasExplicitHAlign && cell != null)
        {
            var dt = cell.DataType?.Value;
            bool isText = dt == CellValues.SharedString || dt == CellValues.InlineString || dt == CellValues.String;
            if (!isText && cell.CellValue != null)
                styles.Add("text-align:right");
        }

        if (alignment == null) return;

        if (alignment.Vertical?.HasValue == true)
        {
            var v = alignment.Vertical.InnerText;
            var cssVAlign = v switch
            {
                "top" => "top",
                "center" => "middle",
                "bottom" => "bottom",
                _ => null
            };
            if (cssVAlign != null) styles.Add($"vertical-align:{cssVAlign}");
        }

        if (alignment.WrapText?.Value == true)
            styles.Add("white-space:pre-wrap;word-wrap:break-word");

        if (alignment.TextRotation?.HasValue == true && alignment.TextRotation.Value != 0)
        {
            var rot = alignment.TextRotation.Value;
            if (rot == 255)
            {
                // 255 = stacked vertical text (each char on its own line)
                styles.Add("writing-mode:vertical-rl;text-orientation:upright;letter-spacing:-2px");
            }
            else
            {
                // Excel: 0-90 = counter-clockwise, 91-180 = clockwise (91=1°CW, 180=90°CW)
                // Excel: 1-90 = CCW (CSS negative), 91-180 = CW (CSS positive, 91=1°, 180=90°)
                int cssDeg = rot <= 90 ? -(int)rot : (int)rot - 90;
                styles.Add($"transform:rotate({cssDeg}deg);white-space:nowrap");
            }
        }

        if (alignment.Indent?.HasValue == true && alignment.Indent.Value > 0)
        {
            // 1 indent level ≈ width of "0" in default font ≈ fontSize × 0.6
            var defFontSz = _doc.WorkbookPart?.WorkbookStylesPart?.Stylesheet
                ?.Fonts?.Elements<Font>().FirstOrDefault()?.FontSize?.Val?.Value ?? 11.0;
            var indentPt = alignment.Indent.Value * defFontSz * 0.6;
            styles.Add($"padding-left:{indentPt:0.#}pt");
        }

        // Reading order: 1=LTR, 2=RTL (for mixed-direction content)
        if (alignment.ReadingOrder?.HasValue == true)
        {
            var ro = alignment.ReadingOrder.Value;
            if (ro == 2) styles.Add("direction:rtl;unicode-bidi:embed");
            else if (ro == 1) styles.Add("direction:ltr;unicode-bidi:embed");
        }
    }

    // ==================== Color Resolution ====================

    private string? ResolveFontColor(Font font)
    {
        if (font.Color?.Rgb?.Value != null)
        {
            var raw = font.Color.Rgb.Value;
            return FormatColorForCss(raw);
        }
        if (font.Color?.Theme?.Value != null)
        {
            var tint = font.Color.Tint?.Value;
            return ResolveThemeColor(font.Color.Theme.Value, tint);
        }
        return null;
    }

    // Standard Excel indexed color palette (first 64 colors) — can be overridden by styles.xml
    private static readonly string[] DefaultIndexedColors = [
        "#000000","#FFFFFF","#FF0000","#00FF00","#0000FF","#FFFF00","#FF00FF","#00FFFF",
        "#000000","#FFFFFF","#FF0000","#00FF00","#0000FF","#FFFF00","#FF00FF","#00FFFF",
        "#800000","#008000","#000080","#808000","#800080","#008080","#C0C0C0","#808080",
        "#9999FF","#993366","#FFFFCC","#CCFFFF","#660066","#FF8080","#0066CC","#CCCCFF",
        "#000080","#FF00FF","#FFFF00","#00FFFF","#800080","#800000","#008080","#0000FF",
        "#00CCFF","#CCFFFF","#CCFFCC","#FFFF99","#99CCFF","#FF99CC","#CC99FF","#FFCC99",
        "#3366FF","#33CCCC","#99CC00","#FFCC00","#FF9900","#FF6600","#666699","#969696",
        "#003366","#339966","#003300","#333300","#993300","#993366","#333399","#333333"
    ];

    private string? ResolveColorRgb(ColorType? color)
    {
        if (color?.Rgb?.Value != null)
            return FormatColorForCss(color.Rgb.Value);
        if (color?.Indexed?.Value != null)
        {
            var idx = (int)color.Indexed.Value;
            var palette = GetResolvedIndexedColors();
            if (idx >= 0 && idx < palette.Length)
                return palette[idx];
            if (idx == 64) return null; // system foreground (context dependent)
            if (idx == 65) return null; // system background
        }
        if (color?.Theme?.Value != null)
        {
            var tint = color.Tint?.Value;
            return ResolveThemeColor(color.Theme.Value, tint);
        }
        return null;
    }

    private static string FormatColorForCss(string raw)
    {
        // Reject non-hex raw values before interpolating into inline CSS —
        // styles.xml / indexedColors attrs are attacker-controlled, and an
        // unvalidated raw flows into `color:#{raw}` / `background:#{raw}`
        // as an XSS sink.
        static bool isHex(string s) =>
            s.All(c => (c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f'));
        if (raw.Length == 8 && isHex(raw)) return "#" + raw[2..];
        if (raw.Length is 6 or 3 && isHex(raw)) return "#" + raw;
        return "#000";
    }

    // ==================== Formatted Cell Value ====================

    /// <summary>
    /// Get cell display value with number formatting applied for HTML preview.
    /// Handles common formats: percentage, thousands separator, decimal places, dates.
    /// </summary>
    private string GetFormattedCellValue(Cell cell, Stylesheet? stylesheet, Core.FormulaEvaluator? evaluator = null)
    {
        var rawValue = GetCellDisplayValue(cell);

        // If the cell has a formula, always try to evaluate (cached values may be stale)
        if (cell.CellFormula?.Text != null && evaluator != null)
        {
            var result = evaluator.TryEvaluateFull(cell.CellFormula.Text);
            if (result != null)
            {
                if (result.IsError) return result.ErrorValue!;
                rawValue = result.ToCellValueText();
                if (result.IsString) return rawValue;
                if (result.IsBool) return result.BoolValue!.Value ? "TRUE" : "FALSE";
            }
            // If evaluation fails (null), fall through to use cached value / raw display
        }

        if (string.IsNullOrEmpty(rawValue)) return rawValue;

        // Boolean: convert 1/0 to TRUE/FALSE
        if (cell.DataType?.Value == CellValues.Boolean)
            return rawValue == "1" ? "TRUE" : "FALSE";

        // Only format numeric values (not strings, shared strings, etc.)
        if (cell.DataType?.Value == CellValues.SharedString ||
            cell.DataType?.Value == CellValues.InlineString ||
            cell.DataType?.Value == CellValues.String ||
            cell.DataType?.Value == CellValues.Error)
            return rawValue;

        if (!double.TryParse(rawValue, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var numVal))
        {
            // GetCellDisplayValue pre-formats date serials to ISO (e.g. "2023-03-15"),
            // which fails double.TryParse. For numeric cells with a date format code,
            // recover the raw serial and let ApplyNumberFormat honour the date code
            // (HTML preview only — get/query keeps the canonical ISO form).
            var serialText = cell.CellValue?.Text;
            if (cell.CellFormula?.Text == null && serialText != null &&
                double.TryParse(serialText, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out var serial))
            {
                var dateFmt = ResolveCellFormatCode(cell, stylesheet);
                if (dateFmt != null && ContainsDateTokenOutsideQuotes(dateFmt))
                    return ApplyNumberFormat(serial, dateFmt);
            }
            return rawValue;
        }

        // Clean up floating point artifacts for display (e.g. 25300000.000000004 → 25300000)
        var cleanVal = numVal;
        var rounded = Math.Round(numVal, 10);
        if (Math.Abs(rounded - Math.Round(rounded)) < 1e-9)
            cleanVal = Math.Round(rounded);
        rawValue = cleanVal == numVal ? rawValue
            : cleanVal.ToString(System.Globalization.CultureInfo.InvariantCulture);

        // Look up number format
        var fmtCode = ResolveCellFormatCode(cell, stylesheet);
        if (fmtCode == null) return rawValue;

        return ApplyNumberFormat(numVal, fmtCode);
    }

    /// <summary>
    /// Resolve a cell's number format code (custom &lt;numFmt&gt; first, then built-in).
    /// Returns null when the cell has no explicit (non-General) format.
    /// </summary>
    private static string? ResolveCellFormatCode(Cell cell, Stylesheet? stylesheet)
    {
        var styleIndex = cell.StyleIndex?.Value ?? 0;
        if (styleIndex == 0 || stylesheet == null) return null;

        var cellFormats = stylesheet.CellFormats;
        if (cellFormats == null || styleIndex >= (uint)cellFormats.Elements<CellFormat>().Count())
            return null;

        var xf = cellFormats.Elements<CellFormat>().ElementAt((int)styleIndex);
        var numFmtId = xf.NumberFormatId?.Value ?? 0;
        if (numFmtId == 0) return null;

        var customFmt = stylesheet.NumberingFormats?.Elements<NumberingFormat>()
            .FirstOrDefault(nf => nf.NumberFormatId?.Value == numFmtId);
        if (customFmt?.FormatCode?.Value != null)
            return customFmt.FormatCode.Value;
        return ResolveBuiltInFormat(numFmtId);
    }

    /// <summary>
    /// Rich-text cells (shared-string items with multiple &lt;r&gt; runs carrying
    /// per-run &lt;rPr&gt;) flatten to plain text via GetCellDisplayValue's InnerText.
    /// Build per-run &lt;span&gt; HTML instead so color/bold/italic/font survive.
    /// Returns pre-encoded HTML (run text already HtmlEncoded) when the cell is a
    /// shared string with at least one run that has run-properties; otherwise null
    /// so the caller falls back to the flat CellHtml path.
    /// </summary>
    private string? TryBuildRichTextHtml(Cell cell)
    {
        if (cell.DataType?.Value != CellValues.SharedString) return null;
        var value = cell.CellValue?.Text;
        if (value == null || !int.TryParse(value, out int idx)) return null;

        var sst = _doc.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
        var item = sst?.SharedStringTable?.Elements<SharedStringItem>().ElementAtOrDefault(idx);
        if (item == null) return null;

        var runs = item.Elements<Run>().ToList();
        // Only worth wrapping when at least one run carries explicit run-properties;
        // a single plain run is identical to flat text — let the normal path handle it.
        if (runs.Count == 0 || !runs.Any(r => r.RunProperties != null)) return null;

        var sb = new StringBuilder();
        foreach (var run in runs)
        {
            var rPr = run.RunProperties;
            var style = new StringBuilder();
            if (rPr != null)
            {
                // Shared-string run properties expose children only via GetFirstChild;
                // <b/>/<i/>/<u/> are presence-flags (a Val=false would disable, mirror that).
                var bold = rPr.GetFirstChild<Bold>();
                if (bold != null && bold.Val?.Value != false)
                    style.Append("font-weight:bold;");
                var italic = rPr.GetFirstChild<Italic>();
                if (italic != null && italic.Val?.Value != false)
                    style.Append("font-style:italic;");
                var underline = rPr.GetFirstChild<Underline>();
                if (underline != null && underline.Val?.Value != UnderlineValues.None)
                    style.Append("text-decoration:underline;");
                var colorHex = ResolveRunColorHex(rPr.GetFirstChild<Color>());
                if (colorHex != null) style.Append($"color:{colorHex};");
                if (rPr.GetFirstChild<FontSize>()?.Val?.Value is double fs)
                    style.Append($"font-size:{fs:0.##}pt;");
                var fontName = rPr.GetFirstChild<RunFont>()?.Val?.Value;
                if (!string.IsNullOrEmpty(fontName))
                    style.Append($"font-family:'{fontName}';");
            }
            var text = HtmlEncode(run.Text?.Text ?? "");
            sb.Append(style.Length > 0 ? $"<span style=\"{style}\">{text}</span>" : $"<span>{text}</span>");
        }
        return sb.ToString();
    }

    /// <summary>Resolve a shared-string run &lt;color&gt; to a #RRGGBB hex (rgb or theme), else null.</summary>
    private string? ResolveRunColorHex(Color? color)
    {
        if (color == null) return null;
        var rgb = color.Rgb?.Value;
        if (!string.IsNullOrEmpty(rgb))
        {
            if (rgb.Length > 6) rgb = rgb[^6..];   // strip ARGB alpha
            return $"#{rgb}";
        }
        if (color.Theme?.Value is uint themeIdx)
        {
            var tint = color.Tint?.Value;
            return ResolveThemeColor(themeIdx, tint);
        }
        return null;
    }

    private static string? ResolveBuiltInFormat(uint numFmtId) => numFmtId switch
    {
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        14 => "m/d/yy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yy h:mm",
        37 => "#,##0 ;(#,##0)",
        38 => "#,##0 ;(#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;(#,##0.00)",
        49 => "@",
        _ => null
    };

    // Excel number-format [Color] names → CSS hex (the named palette only; the
    // [Color N] indexed form maps to the indexed-color table, omitted here as a
    // known limitation — named colors cover the common $;[Red](…) negative case).
    private static readonly Dictionary<string, string> NumFmtColorNames =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["Black"] = "#000000", ["White"] = "#FFFFFF", ["Red"] = "#FF0000",
            ["Green"] = "#008000", ["Blue"] = "#0000FF", ["Yellow"] = "#FFFF00",
            ["Magenta"] = "#FF00FF", ["Cyan"] = "#00FFFF",
        };

    /// <summary>
    /// Resolve the CSS color implied by the number format's [Color] tag for the
    /// section that applies to <paramref name="value"/> (positive;negative;zero).
    /// Returns null when the active section carries no [Color] marker.
    /// </summary>
    private static string? GetNumberFormatColor(double value, string fmtCode)
    {
        string section;
        if (fmtCode.Contains(';'))
        {
            var sections = fmtCode.Split(';');
            if (value < 0 && sections.Length >= 2) section = sections[1];
            else if (value == 0 && sections.Length >= 3) section = sections[2];
            else section = sections[0];
        }
        else
        {
            section = fmtCode;
        }

        var m = System.Text.RegularExpressions.Regex.Match(
            section, @"\[(Black|White|Red|Green|Blue|Yellow|Magenta|Cyan)\]",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        return m.Success && NumFmtColorNames.TryGetValue(m.Groups[1].Value, out var hex) ? hex : null;
    }

    /// <summary>
    /// Resolve the number-format [Color] CSS for a numeric cell, or null. Mirrors
    /// the format-code lookup in GetFormattedCellValue (custom &lt;numFmt&gt; then
    /// built-in id), then delegates section/value selection to GetNumberFormatColor.
    /// </summary>
    private string? GetCellNumberFormatColor(Cell cell, CellFormat xf, Stylesheet stylesheet)
    {
        // Only numeric cells carry value-driven format sections.
        var dt = cell.DataType?.Value;
        if (dt == CellValues.SharedString || dt == CellValues.InlineString
            || dt == CellValues.String || dt == CellValues.Boolean || dt == CellValues.Error)
            return null;
        if (!double.TryParse(cell.CellValue?.Text, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var numVal))
            return null;

        var numFmtId = xf.NumberFormatId?.Value ?? 0;
        if (numFmtId == 0) return null;

        var customFmt = stylesheet.NumberingFormats?.Elements<NumberingFormat>()
            .FirstOrDefault(nf => nf.NumberFormatId?.Value == numFmtId);
        var fmtCode = customFmt?.FormatCode?.Value ?? ResolveBuiltInFormat(numFmtId);
        if (fmtCode == null) return null;

        return GetNumberFormatColor(numVal, fmtCode);
    }

    private static string ApplyNumberFormat(double value, string fmtCode)
    {
        // Handle multi-section format codes: positive;negative;zero
        if (fmtCode.Contains(';'))
        {
            var sections = fmtCode.Split(';');
            if (value < 0 && sections.Length >= 2)
            {
                var negFmt = sections[1].Trim();
                // If format already handles negative (has parens or minus), don't add extra minus
                return ApplyNumberFormat(Math.Abs(value), negFmt);
            }
            if (value == 0 && sections.Length >= 3)
            {
                var zeroFmt = sections[2].Trim();
                // Quoted literal for zero section: "zero" → zero
                if (zeroFmt.StartsWith('"') && zeroFmt.EndsWith('"'))
                    return zeroFmt[1..^1];
                return ApplyNumberFormat(value, zeroFmt);
            }
            fmtCode = sections[0].Trim();
        }

        // Strip [Color] markers: [Red], [Blue], [Green], [Color N], etc.
        fmtCode = System.Text.RegularExpressions.Regex.Replace(fmtCode, @"\[(Red|Blue|Green|Yellow|White|Black|Cyan|Magenta|Color\s*\d+)\]", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Trim();

        // Strip [$...] locale/currency specifiers (e.g. [$-409], [$€-407], [$¥-411])
        fmtCode = System.Text.RegularExpressions.Regex.Replace(fmtCode, @"\[\$[^\]]*\]", "").Trim();

        // Strip Excel numfmt special characters:
        // _X = space placeholder, *X = fill character, \X = literal character escape
        fmtCode = System.Text.RegularExpressions.Regex.Replace(fmtCode, @"_.", "").Trim();
        fmtCode = System.Text.RegularExpressions.Regex.Replace(fmtCode, @"\*.", "").Trim();
        fmtCode = System.Text.RegularExpressions.Regex.Replace(fmtCode, @"\\(.)", "$1").Trim();

        // Strip condition markers: [>100], [<=0], etc.
        fmtCode = System.Text.RegularExpressions.Regex.Replace(fmtCode, @"\[[<>=!]+\d+\.?\d*\]", "").Trim();

        // Handle parenthesis wrapping: ($#,##0.00) → prefix="(" suffix=")"
        if (fmtCode.StartsWith('(') && fmtCode.EndsWith(')'))
        {
            var inner = fmtCode[1..^1];
            return "(" + ApplyNumberFormat(value, inner) + ")";
        }

        var fmt = fmtCode.ToLowerInvariant();

        // Date/time formats may contain quoted literals (e.g. "D"d"D").
        // Skip prefix/suffix extraction for these — the date handler in
        // ApplyNumberFormatCore processes quotes via NormalizeDateFormatCase.
        if (ContainsDateTokenOutsideQuotes(fmtCode))
            return ApplyNumberFormatCore(value, fmtCode);

        // Extract currency/text prefix and suffix (e.g. "$", "€", "¥", or quoted strings like "USD ")
        var prefix = "";
        var suffix = "";
        var cleanFmt = fmtCode;
        // Handle literal characters: $, ¥, €, £
        foreach (var sym in new[] { "$", "¥", "€", "£", "₹" })
        {
            if (cleanFmt.Contains(sym))
            {
                var idx = cleanFmt.IndexOf(sym);
                var hashIdx = cleanFmt.IndexOf('#');
                var zeroIdx = cleanFmt.IndexOf('0');
                var firstDigit = (hashIdx >= 0 && zeroIdx >= 0) ? Math.Min(hashIdx, zeroIdx)
                    : Math.Max(hashIdx, zeroIdx);
                if (firstDigit < 0 || idx <= firstDigit)
                    prefix = sym;
                else
                    suffix = sym;
                cleanFmt = cleanFmt.Replace(sym, "");
            }
        }
        // Handle quoted prefix/suffix: "USD "
        var quoteMatch = System.Text.RegularExpressions.Regex.Match(cleanFmt, "^\"([^\"]+)\"");
        if (quoteMatch.Success) { prefix += quoteMatch.Groups[1].Value; cleanFmt = cleanFmt[quoteMatch.Length..]; }
        var quoteSuffix = System.Text.RegularExpressions.Regex.Match(cleanFmt, "\"([^\"]+)\"$");
        if (quoteSuffix.Success) { suffix = quoteSuffix.Groups[1].Value + suffix; cleanFmt = cleanFmt[..^quoteSuffix.Length]; }

        // Re-check for a paren-wrapped numeric pattern after the _X/*X/\X strips and
        // currency-symbol extraction. Accounting negative sections such as
        // "_($* (#,##0.00_)" lose their "_(" wrappers to the strip passes, leaving a
        // bare leading "(" (and possibly trailing ")") that the early line-2225 check
        // (which ran on the pre-strip fmtCode) never saw. Native Excel keeps the paren.
        cleanFmt = cleanFmt.Trim();
        var parenOpen = false;
        var parenClose = false;
        if (cleanFmt.StartsWith('('))
        {
            parenOpen = true;
            cleanFmt = cleanFmt[1..];
            if (cleanFmt.EndsWith(')')) { parenClose = true; cleanFmt = cleanFmt[..^1]; }
        }

        // Handle +/- prefix in format (e.g. "+0.0%", "-#,##0")
        cleanFmt = cleanFmt.Trim();
        if (cleanFmt.StartsWith('+'))
        { prefix += "+"; cleanFmt = cleanFmt[1..]; }
        else if (cleanFmt.StartsWith('-'))
        { prefix += "-"; cleanFmt = cleanFmt[1..]; }

        // Pure text format (only quoted prefix/suffix, no numeric pattern)
        if (string.IsNullOrEmpty(cleanFmt.Trim()))
            return prefix + suffix;

        var formatted = ApplyNumberFormatCore(value, cleanFmt.Trim());
        // Accounting paren wraps the numeric core (and trailing currency), keeping
        // "(1,234.56" contiguous even when a left-aligned "$" prefix is present.
        if (parenOpen) formatted = "(" + formatted + suffix + (parenClose ? ")" : "");
        else formatted += suffix;
        // For single-section formats with currency prefix, negative sign goes before the prefix
        if (value < 0 && prefix.Length > 0 && formatted.StartsWith('-'))
            return "-" + prefix + formatted[1..];
        return prefix + formatted;
    }

    private static int Gcd(int a, int b)
    {
        a = Math.Abs(a); b = Math.Abs(b);
        while (b != 0) { var t = b; b = a % b; a = t; }
        return a == 0 ? 1 : a;
    }

    private static string ApplyNumberFormatCore(double value, string fmtCode)
    {
        var fmt = fmtCode.ToLowerInvariant();

        // Percentage formats
        if (fmt.Contains('%'))
        {
            var pctVal = value * 100;
            var decimals = CountDecimalPlaces(fmtCode);
            return pctVal.ToString($"F{decimals}") + "%";
        }

        // Fraction formats: "# ?/?", "# ??/??", "?/?" etc.
        // Denominator-digit count = number of '?' after the slash → max denominator
        // (1→9, 2→99, 3→999). Find the best rational approximation within that limit.
        var fracMatch = System.Text.RegularExpressions.Regex.Match(fmtCode, @"\?+\s*/\s*(\?+)");
        if (fracMatch.Success)
        {
            int denomDigits = fracMatch.Groups[1].Value.Count(c => c == '?');
            int maxDenom = (int)Math.Pow(10, denomDigits) - 1;
            if (maxDenom < 1) maxDenom = 9;
            bool neg = value < 0;
            double abs = Math.Abs(value);
            long whole = (long)Math.Floor(abs);
            double frac = abs - whole;

            // Continued-fraction convergents give the simplest fraction within the
            // denominator limit (matches Excel, e.g. 0.14159 → 1/7 not 14/99).
            long h0 = 0, h1 = 1, k0 = 1, k1 = 0;
            double b = frac;
            for (int guard = 0; guard < 64; guard++)
            {
                long a = (long)Math.Floor(b);
                long h2 = a * h1 + h0, k2 = a * k1 + k0;
                if (k2 > maxDenom || k2 == 0) break;
                h0 = h1; h1 = h2; k0 = k1; k1 = k2;
                double rem = b - a;
                if (rem < 1e-12) break;
                b = 1 / rem;
            }
            int bestNum = (int)h1, bestDen = (int)Math.Max(k1, 1);
            // Reduce the chosen fraction
            if (bestNum != 0)
            {
                int g = Gcd(bestNum, bestDen);
                bestNum /= g; bestDen /= g;
            }
            // Numerator rounded up to a whole unit (e.g. 0.999 → 1/1) folds into whole part
            if (bestNum == bestDen && bestNum != 0) { whole += 1; bestNum = 0; }

            var sb = new StringBuilder();
            if (neg) sb.Append('-');
            if (bestNum == 0)
                sb.Append(whole.ToString(System.Globalization.CultureInfo.InvariantCulture));
            else if (whole == 0)
                sb.Append($"{bestNum}/{bestDen}");
            else
                sb.Append($"{whole} {bestNum}/{bestDen}");
            return sb.ToString();
        }

        // Elapsed time format: [h]:mm:ss or [mm]:ss (total hours/minutes, can exceed 24/60)
        var elapsedMatch = System.Text.RegularExpressions.Regex.Match(fmtCode, @"\[(h+)\]:?(mm)?:?(ss)?");
        if (elapsedMatch.Success)
        {
            var totalHours = (int)(value * 24);
            var totalMinutes = (int)(value * 24 * 60) % 60;
            var totalSeconds = (int)(value * 24 * 3600) % 60;
            var parts = new List<string> { totalHours.ToString() };
            if (elapsedMatch.Groups[2].Success) parts.Add(totalMinutes.ToString("D2"));
            if (elapsedMatch.Groups[3].Success) parts.Add(totalSeconds.ToString("D2"));
            return string.Join(":", parts);
        }

        // Date formats (serial number → DateTime)
        if (fmt.Contains('y') || fmt.Contains('m') || fmt.Contains('d') || fmt.Contains('h'))
        {
            try
            {
                var dt = DateTime.FromOADate(value);
                // Context-sensitive m/mm: after h → minute, otherwise → month
                // Strategy: mark minute 'm' as '\x01' placeholder, then convert remaining m→M
                var dotnetFmt = NormalizeDateFormatCase(fmtCode);
                // Step 1: Replace h:mm and h:m patterns → mark minutes as placeholder
                dotnetFmt = System.Text.RegularExpressions.Regex.Replace(dotnetFmt, @"([hH]+)([:.])(mm?)", m =>
                    m.Groups[1].Value + m.Groups[2].Value + new string('\x01', m.Groups[3].Value.Length));
                // Also handle mm:ss (mm before ss is also minutes)
                dotnetFmt = System.Text.RegularExpressions.Regex.Replace(dotnetFmt, @"(mm?)([:.])(ss?)", m =>
                    new string('\x01', m.Groups[1].Value.Length) + m.Groups[2].Value + m.Groups[3].Value);
                // Step 2: Convert remaining m/mm to M/MM (month)
                dotnetFmt = dotnetFmt.Replace("mmmm", "MMMM").Replace("mmm", "MMM")
                    .Replace("mm", "MM").Replace("m", "M");
                // Step 3: Restore minute placeholders
                dotnetFmt = dotnetFmt.Replace("\x01\x01", "mm").Replace("\x01", "m");
                // Step 4: Other conversions
                // If AM/PM format (has 't' outside quotes), use h (12h); otherwise use H (24h)
                if (!ContainsCharOutsideQuotes(dotnetFmt, 't'))
                    dotnetFmt = dotnetFmt.Replace("hh", "HH").Replace("h", "H");
                dotnetFmt = dotnetFmt.Replace("dddd", "dddd").Replace("ddd", "ddd").Replace("dd", "dd");
                return dt.ToString(dotnetFmt, System.Globalization.CultureInfo.InvariantCulture);
            }
            catch { return value.ToString(); }
        }

        // Scientific notation
        if (fmt.Contains("e+") || fmt.Contains("e-"))
        {
            var decimals = CountDecimalPlaces(fmtCode);
            if (value == 0) return decimals > 0 ? $"0.{new string('0', decimals)}E+00" : "0E+00";
            var eIdx = fmt.IndexOf("e+", StringComparison.Ordinal);
            if (eIdx < 0) eIdx = fmt.IndexOf("e-", StringComparison.Ordinal);
            var expDigits = eIdx >= 0 ? fmtCode[(eIdx + 2)..].Count(c => c == '0') : 2;
            var exp = (int)Math.Floor(Math.Log10(Math.Abs(value)));
            var mantissa = value / Math.Pow(10, exp);
            var expStr = exp >= 0 ? $"+{exp.ToString().PadLeft(expDigits, '0')}" : $"-{Math.Abs(exp).ToString().PadLeft(expDigits, '0')}";
            return $"{mantissa.ToString($"F{decimals}")}E{expStr}";
        }

        // Trailing comma scaling: each trailing comma divides value by 1000
        // e.g. "#," = ÷1000, "#,," = ÷1000000, "#,##0," = thousands + ÷1000
        var trailingCommas = 0;
        var fmtTrimmed = fmtCode.TrimEnd();
        while (fmtTrimmed.EndsWith(',')) { trailingCommas++; fmtTrimmed = fmtTrimmed[..^1]; }
        if (trailingCommas > 0)
        {
            value /= Math.Pow(1000, trailingCommas);
            fmtCode = fmtTrimmed;
        }

        // Numeric with thousands separator and/or decimals
        bool hasThousands = fmtCode.Contains(',') && (fmtCode.Contains('#') || fmtCode.Contains('0'));
        var numDecimals = CountDecimalPlaces(fmtCode);

        if (hasThousands)
            return value.ToString($"N{numDecimals}", System.Globalization.CultureInfo.InvariantCulture);
        if (numDecimals > 0)
            return value.ToString($"F{numDecimals}");

        // @ = text format — return raw
        if (fmt == "@") return value.ToString();

        // Integer format "0"
        if (fmtCode.Trim() == "0") return ((long)Math.Round(value)).ToString();

        return value.ToString();
    }

    private static int CountDecimalPlaces(string fmtCode)
    {
        var dotIdx = fmtCode.IndexOf('.');
        if (dotIdx < 0) return 0;
        int count = 0;
        for (int i = dotIdx + 1; i < fmtCode.Length; i++)
        {
            if (fmtCode[i] == '0' || fmtCode[i] == '#') count++;
            else break;
        }
        return count;
    }

    /// <summary>
    /// Returns true if fmtCode contains date/time tokens (y, m, d, h, s) outside
    /// double-quoted strings. Used to route date formats past prefix/suffix extraction.
    /// </summary>
    private static bool ContainsDateTokenOutsideQuotes(string fmtCode)
    {
        bool inQuote = false;
        foreach (var ch in fmtCode)
        {
            if (ch == '"') { inQuote = !inQuote; continue; }
            if (!inQuote)
            {
                var lower = char.ToLowerInvariant(ch);
                if (lower is 'y' or 'm' or 'd' or 'h' or 's') return true;
            }
        }
        return false;
    }

    /// <summary>
    /// Returns true if ch appears outside double-quoted strings in fmtCode.
    /// </summary>
    private static bool ContainsCharOutsideQuotes(string fmtCode, char target)
    {
        bool inQuote = false;
        foreach (var ch in fmtCode)
        {
            if (ch == '"') { inQuote = !inQuote; continue; }
            if (!inQuote && ch == target) return true;
        }
        return false;
    }

    /// <summary>
    /// Normalize Excel date/time format specifiers to .NET-compatible case
    /// and replace AM/PM → tt, A/P → t outside quoted strings.
    /// </summary>
    private static string NormalizeDateFormatCase(string fmtCode)
    {
        var sb = new StringBuilder(fmtCode.Length);
        bool inQuote = false;
        for (int i = 0; i < fmtCode.Length; i++)
        {
            var ch = fmtCode[i];
            if (ch == '"') { inQuote = !inQuote; sb.Append(ch); continue; }
            if (inQuote) { sb.Append(ch); continue; }
            // AM/PM → tt (check before single-char A/P)
            if ((ch == 'A' || ch == 'a') && i + 4 < fmtCode.Length
                && (fmtCode[i + 1] == 'M' || fmtCode[i + 1] == 'm')
                && fmtCode[i + 2] == '/'
                && (fmtCode[i + 3] == 'P' || fmtCode[i + 3] == 'p')
                && (fmtCode[i + 4] == 'M' || fmtCode[i + 4] == 'm'))
            {
                sb.Append("tt"); i += 4; continue;
            }
            // A/P → t
            if ((ch == 'A' || ch == 'a') && i + 2 < fmtCode.Length
                && fmtCode[i + 1] == '/'
                && (fmtCode[i + 2] == 'P' || fmtCode[i + 2] == 'p'))
            {
                sb.Append('t'); i += 2; continue;
            }
            sb.Append(ch switch { 'Y' => 'y', 'D' => 'd', 'S' => 's', 'M' => 'm', 'H' => 'h', _ => ch });
        }
        return sb.ToString();
    }

    // ==================== CSS ====================

    private string GenerateExcelCss()
    {
        // Read default font from workbook styles (font index 0)
        var defFontName = OfficeDefaultFonts.MinorLatin;
        var defFontSize = OfficeDefaultFonts.ExcelBodySizePt;
        var stylesheet = _doc.WorkbookPart?.WorkbookStylesPart?.Stylesheet;
        if (stylesheet?.Fonts != null && stylesheet.Fonts.Elements<Font>().Any())
        {
            var f0 = stylesheet.Fonts.Elements<Font>().First();
            if (f0.FontName?.Val?.Value != null) defFontName = CssSanitize(f0.FontName.Val.Value);
            if (f0.FontSize?.Val?.Value != null) defFontSize = f0.FontSize.Val.Value.ToString("0.##");
        }
        return $$"""
        * { margin: 0; padding: 0; box-sizing: border-box; }
        html, body { height: 100%; }
        body {
            font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif;
            background: #f0f0f0;
            color: #333;
            display: flex;
            flex-direction: column;
            min-height: 100vh;
        }
        .file-title {
            padding: 12px 20px;
            font-size: 14px;
            font-weight: 600;
            background: #217346;
            color: #fff;
        }
        .sheet-tabs {
            display: flex;
            background: #e0e0e0;
            border-top: 1px solid #ccc;
            overflow-x: auto;
            padding: 0 8px;
            flex-shrink: 0;
            position: sticky;
            bottom: 0;
            z-index: 10;
        }
        .sheet-tab {
            --tab-color: #e8e8e8;
            padding: 8px 16px;
            font-size: 12px;
            cursor: pointer;
            border: 1px solid #bbb;
            border-top: none;
            background: var(--tab-color);
            color: #fff;
            margin-bottom: 0;
            border-radius: 0 0 3px 3px;
            white-space: nowrap;
            user-select: none;
            position: relative;
            transition: background 0.15s, color 0.15s;
        }
        .sheet-tab[style*="--tab-color:#e8e8e8"], .sheet-tab:not([style*="--tab-color"]) {
            color: #333;
        }
        .sheet-tab:hover { opacity: 0.85; }
        .sheet-tab.active {
            background: linear-gradient(to bottom, #fff 60%, color-mix(in srgb, var(--tab-color) 30%, #fff)) !important;
            color: #333 !important;
            border-color: #aaa;
            border-bottom: 3px solid var(--tab-color);
            font-weight: 600;
        }
        .sheet-slider { flex: 1; position: relative; overflow: hidden; display: flex; flex-direction: column; min-height: 0; }
        .sheet-content { background: #fff; display: none; flex: 1; min-height: 0; }
        .sheet-content.active { display: flex; flex-direction: column; }
        .table-wrapper {
            flex: 1;
            overflow: auto;
            min-height: 0;
            background: #fff;
        }
        table {
            border-collapse: collapse;
            font-size: {{defFontSize}}px;
            font-family: '{{defFontName}}', 'Segoe UI', sans-serif;
            table-layout: fixed;
        }
        .row-header-col { width: 30pt; }
        th {
            background: #f8f8f8;
            border: 1px solid #e0e0e0;
            font-weight: normal;
            color: #666;
            font-size: 10px;
            text-align: center;
            padding: 2px 4px;
        }
        .corner-cell { background: #f0f0f0; z-index: 4; }
        .col-header {
            position: sticky;
            top: 0;
            z-index: 3;
            background: #f8f8f8;
            min-width: 50px;
            cursor: s-resize;
        }
        .row-header {
            position: sticky;
            left: 0;
            z-index: 2;
            background: #f8f8f8;
            min-width: 40px;
            cursor: e-resize;
            /* Drop right border so the data cell's own (often darker) left border shows through.
               Otherwise, with border-collapse, the row-header's light grey right border can win
               the collapse contest and erase the merged-cell left border (rowspan cells especially). */
            border-right: none;
        }
        td {
            /* Default gridlines are painted with inset box-shadow instead of
               border, so they do NOT participate in border-collapse tie-breaking.
               Explicit OOXML borders (rendered as inline border styles on cells
               with an OOXML style) always win at cell boundaries; missing cells
               / style-0 cells no longer erase neighbours' black borders via the
               CSS position-based tie-break. Right+bottom gridlines are owned by
               each cell; first-row top and first-col left gridlines are added
               via the :first-child rules below. */
            box-shadow: inset -1px -1px 0 #e0e0e0;
            padding: 2px 4px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            vertical-align: bottom;
            max-width: 500px;
            word-break: break-all; /* CJK text wrapping support */
        }
        /* Text spill: a left/general text cell with empty right-neighbours paints
           its overflow across them (Excel fidelity). The td stays 1 column wide so
           borders/gridlines/merges are unaffected; overflow:visible lets the inner
           span bleed into the (empty) neighbour cells. The span clips at max-width =
           own width + summed empty-neighbour widths, stopping before the first
           occupied cell — matching real Excel. */
        td.spill { overflow: visible; }
        td.spill .spill-text {
            display: inline-block;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: clip;
            vertical-align: bottom;
        }
        tbody tr:first-child td { box-shadow: inset -1px -1px 0 #e0e0e0, inset 0 1px 0 #e0e0e0; }
        tr td:first-of-type { box-shadow: inset -1px -1px 0 #e0e0e0, inset 1px 0 0 #e0e0e0; }
        tbody tr:first-child td:first-of-type { box-shadow: inset -1px -1px 0 #e0e0e0, inset 1px 1px 0 #e0e0e0; }
        .empty-sheet {
            padding: 40px;
            text-align: center;
            color: #999;
            font-size: 14px;
        }
        /* Chart containers */
        .chart-container {
            margin: 16px auto;
            background: #fff;
            border: 1px solid #e0e0e0;
            border-radius: 6px;
            padding: 12px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        }
        .chart-container svg { display: block; }
        /* Truncation warning */
        .truncation-warning {
            padding: 8px 16px;
            background: #FFF3CD;
            color: #856404;
            border: 1px solid #FFEEBA;
            font-size: 12px;
            text-align: center;
            margin: 4px 0;
        }
        /* Screen reader only */
        .sr-only { position:absolute; clip:rect(0 0 0 0); width:1px; height:1px; overflow:hidden; }
        /* Print styles */
        @media print {
            .file-title, .sheet-tabs { display: none !important; }
            .table-wrapper { max-height: none !important; overflow: visible !important; flex: none !important; }
            body { background: #fff !important; min-height: auto !important; }
            .sheet-content { display: block !important; flex: none !important; }
            td { max-width: none !important; white-space: normal !important; overflow: visible !important; }
        }
        """;
    }

    // ==================== JavaScript ====================

    private static string GenerateExcelJs() => """
        function switchSheet(idx) {
            document.querySelectorAll('.sheet-tab').forEach(function(t) {
                t.classList.toggle('active', parseInt(t.getAttribute('data-sheet')) === idx);
            });
            document.querySelectorAll('.sheet-content').forEach(function(c) {
                c.classList.toggle('active', parseInt(c.getAttribute('data-sheet')) === idx);
            });
            window.scrollTo(0, 0);
        }
        // Fix frozen row sticky top values using actual rendered heights
        document.querySelectorAll('.table-wrapper table').forEach(function(table) {
            var thead = table.querySelector('thead');
            if (!thead) return;
            var theadH = thead.offsetHeight;
            var cumTop = theadH;
            var frozen = table.querySelectorAll('tr[data-frozen]');
            frozen.forEach(function(tr) {
                tr.querySelectorAll('th, td').forEach(function(cell) {
                    if (cell.style.position === 'sticky') cell.style.top = cumTop + 'px';
                });
                cumTop += tr.offsetHeight;
            });
        });
        """;

    // ==================== Utility ====================

    // CONSISTENCY(html-encode): shared plain entity-encoder lives in Core/HtmlPreviewHelper.
    private static string HtmlEncode(string text) => HtmlPreviewHelper.HtmlEncode(text);

    /// <summary>HtmlEncode + convert newlines to br for cell display</summary>
    private static string CellHtml(string text)
    {
        var encoded = HtmlEncode(text);
        return encoded.Contains('\n') ? encoded.Replace("\n", "<br>") : encoded;
    }

    /// <summary>Get data-formula attribute for cells with formulas (for inline editing).</summary>
    private static string GetFormulaAttr(Cell? cell)
    {
        var formula = cell?.CellFormula?.Text;
        if (string.IsNullOrEmpty(formula)) return "";
        return $" data-formula=\"={HtmlEncode(formula)}\"";
    }

    private static string BuildCellContent(string cellRef, string value,
        Dictionary<string, string> dataBarMap, Dictionary<string, string> iconSetMap)
    {
        var hasBar = dataBarMap.TryGetValue(cellRef, out var barEntry);
        var hasIcon = iconSetMap.TryGetValue(cellRef, out var iconEntry);
        if (!hasBar && !hasIcon) return CellHtml(value);

        // Parse "showValue|html" format
        var barShowValue = true;
        var barHtml = "";
        if (hasBar && barEntry != null)
        {
            var sep = barEntry.IndexOf('|');
            barShowValue = sep < 0 || barEntry[0] != '0';
            barHtml = sep >= 0 ? barEntry[(sep + 1)..] : barEntry;
        }
        var iconShowValue = true;
        var iconHtml = "";
        if (hasIcon && iconEntry != null)
        {
            var sep = iconEntry.IndexOf('|');
            iconShowValue = sep < 0 || iconEntry[0] != '0';
            iconHtml = sep >= 0 ? iconEntry[(sep + 1)..] : iconEntry;
        }
        var showValue = barShowValue && iconShowValue;

        var sb = new StringBuilder();
        if (hasBar) sb.Append(barHtml);
        if (hasIcon) sb.Append($"<span style=\"position:absolute;left:4px;top:50%;transform:translateY(-50%);z-index:1\">{iconHtml}</span>");
        if (showValue)
            sb.Append($"<span style=\"position:relative;z-index:1\">{CellHtml(value)}</span>");
        return sb.ToString();
    }

    private static string CssSanitize(string value)
    {
        // Strip characters that could break CSS context
        return Regex.Replace(value, @"[;:{}()\\""']", "");
    }

}
