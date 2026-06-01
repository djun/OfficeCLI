// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // ==================== Column-predicate row matching ====================
    //
    // `row[Salary>5000]` matches the DATA rows of a table by a column's cell
    // value, addressing the column by its header NAME (or column letter). This
    // is the human/agent-natural "where" form for a normal table — you think in
    // column names, not B/E. It reuses the shared AttributeFilter operator
    // engine for comparison and the ListObject column metadata (names + range +
    // header/totals flags) already surfaced by TableToNode, so no header
    // sniffing or new comparison logic is introduced. P1 scope: real
    // ListObjects only; auto-binds to the single table that owns the
    // referenced column(s). Returns row nodes (`/Sheet/row[N]`) so the result
    // feeds Set/Remove for "match rows then operate".

    // Bracket keys that filter row PROPERTIES (height/hidden/...). Any other key
    // in a row selector is treated as a table COLUMN reference.
    private static readonly HashSet<string> RowAttributeKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "height", "hidden", "outlineLevel", "collapsed", "customHeight",
    };

    // [ "quoted name" | 'quoted name' | bareKey ] (op) (value)
    // Quoted keys carry spaces ("Total Amount"); bare keys (Salary, 单价) do not.
    private static readonly Regex RowColPredicateRegex = new(
        @"\[\s*(?:""([^""]+)""|'([^']+)'|([\w.]+))\s*(>=|<=|!=|~=|=|>|<)\s*([^\]]*)\]",
        RegexOptions.Compiled);

    // Parse the column predicates from a row selector, dropping row-attribute
    // and numeric-index brackets. Empty result → caller falls back to the plain
    // row-attribute branch.
    private static List<AttributeFilter.Condition> ParseRowColumnPredicates(string selector)
    {
        var conds = new List<AttributeFilter.Condition>();
        foreach (Match m in RowColPredicateRegex.Matches(selector))
        {
            var key = m.Groups[1].Success ? m.Groups[1].Value
                    : m.Groups[2].Success ? m.Groups[2].Value
                    : m.Groups[3].Value;
            if (RowAttributeKeys.Contains(key)) continue;   // row property, not a column
            if (int.TryParse(key, out _)) continue;          // row[2] index, not a predicate
            conds.Add(new AttributeFilter.Condition(key, MapRowPredicateOp(m.Groups[4].Value), m.Groups[5].Value.Trim()));
        }
        return conds;
    }

    private static AttributeFilter.FilterOp MapRowPredicateOp(string op) => op switch
    {
        ">=" => AttributeFilter.FilterOp.GreaterOrEqual,
        "<=" => AttributeFilter.FilterOp.LessOrEqual,
        ">"  => AttributeFilter.FilterOp.GreaterThan,
        "<"  => AttributeFilter.FilterOp.LessThan,
        "!=" => AttributeFilter.FilterOp.NotEqual,
        "~=" => AttributeFilter.FilterOp.Contains,
        _    => AttributeFilter.FilterOp.Equal,
    };

    // Match table data rows whose cells satisfy every column predicate. Auto-
    // binds to the single ListObject that owns all referenced columns; throws on
    // no-match or cross-table ambiguity rather than silently picking one.
    private List<DocumentNode> QueryRowsByColumnPredicate(string? sheetFilter, List<AttributeFilter.Condition> colConds)
    {
        var candidates = new List<(string sheetName, WorksheetPart part, TableDefinitionPart tdp,
            int dataR1, int dataR2, Dictionary<string, int> colAbsIndex)>();

        foreach (var (sheetName, worksheetPart) in GetWorksheets())
        {
            if (sheetFilter != null && !sheetName.Equals(sheetFilter, StringComparison.OrdinalIgnoreCase))
                continue;

            foreach (var tdp in worksheetPart.TableDefinitionParts)
            {
                var tbl = tdp.Table;
                if (tbl?.Reference?.Value == null) continue;
                if (!TryParseRange(tbl.Reference.Value, out var rng)) continue;

                var colNames = tbl.GetFirstChild<TableColumns>()?.Elements<TableColumn>()
                    .Select(c => c.Name?.Value ?? "").ToList() ?? new List<string>();

                var resolved = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                bool allResolved = true;
                foreach (var cond in colConds)
                {
                    if (!TryResolveTableColumn(cond.Key, colNames, rng.c1, rng.c2, out var absCol))
                    { allResolved = false; break; }
                    resolved[cond.Key] = absCol;
                }
                if (!allResolved) continue;

                bool headerRow = (tbl.HeaderRowCount?.Value ?? 1) != 0;
                bool totalRow = (tbl.TotalsRowCount?.Value ?? 0) > 0 || (tbl.TotalsRowShown?.Value ?? false);
                candidates.Add((sheetName, worksheetPart, tdp,
                    rng.r1 + (headerRow ? 1 : 0), rng.r2 - (totalRow ? 1 : 0), resolved));
            }
        }

        if (candidates.Count == 0)
        {
            var cols = string.Join(", ", colConds.Select(c => $"'{c.Key}'"));
            var scope = sheetFilter == null ? "any sheet" : $"sheet '{sheetFilter}'";
            throw new ArgumentException(
                $"row[col op val] found no table on {scope} with column(s) {cols}. " +
                "Column predicates resolve header names (or column letters) against a table (ListObject).");
        }
        if (candidates.Count > 1)
        {
            var where = string.Join(", ", candidates.Select(c => $"{c.sheetName}!{c.tdp.Table!.Name?.Value}"));
            throw new ArgumentException(
                $"row[col op val] is ambiguous — column(s) exist in {candidates.Count} tables ({where}). " +
                "Scope by sheet, e.g. /SheetName/row[...].");
        }

        var cand = candidates[0];
        var results = new List<DocumentNode>();
        var sheetData = GetSheet(cand.part).GetFirstChild<SheetData>();
        if (sheetData == null) return results;
        var eval = new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart);

        for (int r = cand.dataR1; r <= cand.dataR2; r++)
        {
            // Probe node carries each predicate column's cell value under its
            // key so AttributeFilter evaluates all operators with one engine.
            var probe = new DocumentNode { Type = "cell" };
            foreach (var cond in colConds)
            {
                var cell = FindCell(sheetData, $"{IndexToColumnName(cand.colAbsIndex[cond.Key])}{r}");
                probe.Format[cond.Key] = cell != null ? GetCellDisplayValue(cell, eval) : "";
            }
            if (!AttributeFilter.MatchAll(probe, colConds)) continue;

            var rowNode = new DocumentNode
            {
                Path = $"/{cand.sheetName}/row[{r}]",
                Type = "row",
                Preview = r.ToString(),
            };
            // Carry each predicate column's value under its column key so the
            // row is self-describing AND the CLI post-filter (which re-applies
            // the same [col op val] conditions on top of these results) resolves
            // the key and re-confirms the match, instead of dropping every row
            // because "Salary" is absent from a plain row node's Format.
            foreach (var cond in colConds)
                rowNode.Format[cond.Key] = probe.Format[cond.Key];
            rowNode.ChildCount = sheetData.Elements<Row>()
                .FirstOrDefault(rw => rw.RowIndex?.Value == (uint)r)?.Elements<Cell>().Count() ?? 0;
            results.Add(rowNode);
        }
        return results;
    }

    // Resolve a predicate key to an ABSOLUTE column index within a table's
    // column span [c1..c2]. Header name (case-insensitive) wins over a column
    // letter so a header literally named "B" stays reachable by name.
    private static bool TryResolveTableColumn(string key, List<string> colNames, int c1, int c2, out int absCol)
    {
        absCol = 0;
        var nameIdx = colNames.FindIndex(n => n.Equals(key, StringComparison.OrdinalIgnoreCase));
        if (nameIdx >= 0) { absCol = c1 + nameIdx; return true; }
        if (Regex.IsMatch(key, @"^[A-Za-z]{1,3}$"))
        {
            var letterIdx = ColumnNameToIndex(key.ToUpperInvariant());
            if (letterIdx >= c1 && letterIdx <= c2) { absCol = letterIdx; return true; }
        }
        return false;
    }
}
