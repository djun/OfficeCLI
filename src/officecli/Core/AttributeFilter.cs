// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeCli.Core;

/// <summary>
/// Parses CSS-like attribute filters from query selectors and matches them against DocumentNode.
/// Supports operators: = (exact), != (not equal), ~= (contains), >= (greater or equal), <= (less or equal).
/// Example: "shape[fill=#FF0000][size>=24pt][text~=报告]"
/// </summary>
internal static class AttributeFilter
{
    public enum FilterOp { Equal, NotEqual, Contains, GreaterOrEqual, LessOrEqual, GreaterThan, LessThan, Exists }

    public record Condition(string Key, FilterOp Op, string Value);

    // Regex: [key op value] where op is ~=, >=, <=, !=, =, >, or <.
    // The leading '@' is an optional XPath-style attribute prefix accepted
    // for round-trip parity with Get/Add output (e.g. `/slide[1]/shape[@id=10000]`
    // pastes back into query unchanged). Stripped from the captured key
    // group via the non-capturing prefix.
    // Order matters: multi-char operators before single-char to avoid partial match
    private static readonly Regex AttrRegex = new(
        @"\[@?([\w.]+)\s*(~=|>=|<=|\\?!=|=|>|<)\s*([^\]]*)\]",
        RegexOptions.Compiled);

    // Regex: [key] (has-attribute, no operator). Optional '@' prefix mirrors AttrRegex.
    private static readonly Regex HasAttrRegex = new(
        @"\[@?([\w.]+)\]",
        RegexOptions.Compiled);

    // Regex to find any [...] block (for validation)
    private static readonly Regex BracketBlockRegex = new(
        @"\[([^\]]*)\]",
        RegexOptions.Compiled);

    /// <summary>
    /// Parse all [key op value] conditions from a selector string.
    /// Throws CliException for malformed selectors.
    /// </summary>
    public static List<Condition> Parse(string selector)
    {
        // Check for unclosed brackets
        var openCount = selector.Count(c => c == '[');
        var closeCount = selector.Count(c => c == ']');
        if (openCount != closeCount)
            throw new CliException($"Malformed selector: unclosed bracket in \"{selector}\"")
            {
                Code = "invalid_selector",
                Suggestion = "Ensure every '[' has a matching ']'. Example: paragraph[style=Heading 1]"
            };

        var conditions = new List<Condition>();
        var matchedSpans = new HashSet<(int Start, int End)>();

        foreach (Match m in AttrRegex.Matches(selector))
        {
            var key = m.Groups[1].Value;
            var opStr = m.Groups[2].Value.Replace("\\", "");
            var rawVal = m.Groups[3].Value;
            // CONSISTENCY(find-regex): preserve quotes when the value is the
            // `r"..."` / `r'...'` regex form so MatchOne can detect it. Trim
            // would otherwise eat the surrounding quote that marks the prefix.
            var isRegexForm = rawVal.Length >= 3 && rawVal[0] == 'r'
                && (rawVal[1] == '"' || rawVal[1] == '\'');
            var val = isRegexForm ? rawVal : rawVal.Trim('\'', '"');

            // Detect corrupted values from mis-parsed operators (e.g. === parsed as = with value ==X)
            if (val.StartsWith("=") || val.StartsWith("~") || val.StartsWith("!"))
                throw new CliException($"Malformed selector: invalid operator in \"[{m.Groups[0].Value.Trim('[', ']')}]\". Supported operators: =, !=, ~=, >=, <=, >, <")
                {
                    Code = "invalid_selector",
                    Suggestion = $"Did you mean [{key}={val.TrimStart('=', '~', '!')}]?"
                };

            var op = opStr switch
            {
                "~=" => FilterOp.Contains,
                ">=" => FilterOp.GreaterOrEqual,
                "<=" => FilterOp.LessOrEqual,
                ">" => FilterOp.GreaterThan,
                "<" => FilterOp.LessThan,
                "!=" => FilterOp.NotEqual,
                _ => FilterOp.Equal
            };

            // BUG-R10-01: wildcard '*' in attribute value silently returned 0
            // matches. Users tried e.g. `ole[progId=Excel*]` expecting a
            // contains-like match. Fail fast with a clear error pointing to
            // the right operator rather than quietly mis-filtering.
            if (val.Contains('*'))
                throw new CliException(
                    $"Wildcards (*) are not supported in attribute filters. " +
                    $"Use ~= for contains, e.g. {key}~={val.Trim('*')}.")
                {
                    Code = "invalid_selector",
                    Suggestion = $"Did you mean [{key}~={val.Trim('*')}]?"
                };

            conditions.Add(new Condition(key, op, val));
            matchedSpans.Add((m.Index, m.Index + m.Length));
        }

        // Find [...] blocks that weren't matched by the key=value regex
        foreach (Match block in BracketBlockRegex.Matches(selector))
        {
            if (matchedSpans.Any(s => s.Start == block.Index)) continue;
            var content = block.Groups[1].Value;
            if (string.IsNullOrWhiteSpace(content))
                throw new CliException($"Malformed selector: empty brackets \"[]\" in \"{selector}\"")
                {
                    Code = "invalid_selector",
                    Suggestion = "Use [key=value] or [key] syntax. Example: paragraph[style=Heading 1]"
                };
            // Index like [1] — valid path syntax, skip
            if (int.TryParse(content, out _)) continue;
            // [key] with no operator — "has attribute" filter (CSS [attr] syntax)
            var hasAttrMatch = HasAttrRegex.Match(block.Value);
            if (hasAttrMatch.Success)
            {
                conditions.Add(new Condition(hasAttrMatch.Groups[1].Value, FilterOp.Exists, ""));
                matchedSpans.Add((block.Index, block.Index + block.Length));
                continue;
            }
            // Unrecognized bracket content
            throw new CliException($"Malformed selector: cannot parse \"[{content}]\". Expected [key=value] with operator =, !=, ~=, >=, <=, >, or <")
            {
                Code = "invalid_selector",
                Suggestion = "Example: paragraph[style=Heading 1], shape[fill!=#FF0000], cell[formula]"
            };
        }

        return conditions;
    }

    /// <summary>
    /// Filter a list of DocumentNodes by the given conditions.
    /// All operators (=, !=, ~=, >=, <=) are applied as a post-filter.
    /// This is safe even when handler selectors already pre-filter = and !=,
    /// since filtering is idempotent.
    /// </summary>
    public static List<DocumentNode> Apply(List<DocumentNode> nodes, List<Condition> conditions, bool applyAll = true)
    {
        if (conditions.Count == 0) return nodes;

        var toApply = applyAll
            ? conditions
            : conditions.Where(c => c.Op is FilterOp.Contains or FilterOp.GreaterOrEqual or FilterOp.LessOrEqual or FilterOp.GreaterThan or FilterOp.LessThan or FilterOp.Exists).ToList();

        if (toApply.Count == 0) return nodes;

        return nodes.Where(n => MatchAll(n, toApply)).ToList();
    }

    /// <summary>
    /// Filter nodes and collect diagnostic warnings.
    /// Warns when: a filter key doesn't exist in ANY node's Format,
    /// or when >= / <= / > / < is used on a non-numeric value.
    /// </summary>
    /// <summary>
    /// Rewrite conditions' keys through <paramref name="keyResolver"/>. Used so
    /// handler-level alias maps (e.g. Excel cell: bold -> font.bold) also apply
    /// when AttributeFilter post-filters against DocumentNode.Format in the CLI
    /// query pipeline.
    /// </summary>
    public static List<Condition> NormalizeKeys(List<Condition> conditions, Func<string, string> keyResolver)
    {
        if (conditions.Count == 0) return conditions;
        return conditions.Select(c => new Condition(keyResolver(c.Key), c.Op, c.Value)).ToList();
    }

    public static (List<DocumentNode> Results, List<string> Warnings) ApplyWithWarnings(
        List<DocumentNode> nodes, List<Condition> conditions, bool applyAll = true)
    {
        var warnings = new List<string>();
        if (conditions.Count == 0) return (nodes, warnings);

        var toApply = applyAll
            ? conditions
            : conditions.Where(c => c.Op is FilterOp.Contains or FilterOp.GreaterOrEqual or FilterOp.LessOrEqual or FilterOp.GreaterThan or FilterOp.LessThan or FilterOp.Exists).ToList();

        if (toApply.Count == 0) return (nodes, warnings);

        // Check for missing keys: if a filter key doesn't exist in ANY node, warn
        foreach (var cond in toApply)
        {
            if (cond.Op == FilterOp.NotEqual) continue; // missing key is valid for !=
            bool anyHasKey = nodes.Any(n => ResolveValue(n, cond.Key).HasKey);
            if (!anyHasKey && nodes.Count > 0)
            {
                warnings.Add($"Warning: filter key '{cond.Key}' not found in any result's Format. " +
                    $"Available keys: {string.Join(", ", GetAllFormatKeys(nodes))}");
            }
        }

        // Check for non-numeric values on >= / <= / > / <
        foreach (var cond in toApply.Where(c => c.Op is FilterOp.GreaterOrEqual or FilterOp.LessOrEqual or FilterOp.GreaterThan or FilterOp.LessThan))
        {
            if (ExtractNumber(cond.Value) == null && !EmuConverter.TryParseEmu(cond.Value, out _))
            {
                warnings.Add($"Warning: '{cond.Value}' in [{cond.Key}{OpToString(cond.Op)}{cond.Value}] " +
                    $"is not numeric — comparison may produce unexpected results");
            }
            // Also check actual values in nodes
            foreach (var node in nodes)
            {
                var (hasKey, actual) = ResolveValue(node, cond.Key);
                if (hasKey && ExtractNumber(actual) == null && !EmuConverter.TryParseEmu(actual, out _))
                {
                    warnings.Add($"Warning: value '{actual}' for key '{cond.Key}' at {node.Path} " +
                        $"is not numeric — {OpToString(cond.Op)} comparison may be unreliable");
                    break; // one warning per condition is enough
                }
            }
        }

        var results = nodes.Where(n => MatchAll(n, toApply)).ToList();
        return (results, warnings);
    }

    private static string OpToString(FilterOp op) => op switch
    {
        FilterOp.Equal => "=",
        FilterOp.NotEqual => "!=",
        FilterOp.Contains => "~=",
        FilterOp.GreaterOrEqual => ">=",
        FilterOp.LessOrEqual => "<=",
        FilterOp.GreaterThan => ">",
        FilterOp.LessThan => "<",
        FilterOp.Exists => "(exists)",
        _ => "?"
    };

    private static HashSet<string> GetAllFormatKeys(List<DocumentNode> nodes)
    {
        var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var node in nodes)
        {
            foreach (var key in node.Format.Keys)
                keys.Add(key);
            if (node.Text != null) keys.Add("text");
            if (!string.IsNullOrEmpty(node.Type)) keys.Add("type");
        }
        return keys;
    }

    /// <summary>
    /// Check if a DocumentNode matches all conditions.
    /// </summary>
    public static bool MatchAll(DocumentNode node, List<Condition> conditions)
    {
        foreach (var cond in conditions)
        {
            if (!MatchOne(node, cond)) return false;
        }
        return true;
    }

    /// <summary>
    /// CONSISTENCY(find-regex): shared text-match used by both the `~=` Contains
    /// operator and the CLI `--find` post-filter. Mirrors Word/Pptx Set's
    /// `r"..."` / `r'...'` regex prefix — without it, `--find r"Bullet"`
    /// literally looked for the string `r"Bullet"` (quotes included) and always
    /// returned 0. A plain (non-prefixed) value still does a case-insensitive
    /// contains; a malformed regex falls back to literal contains.
    /// </summary>
    public static bool MatchesTextFilter(string text, string find)
    {
        if (TryParseRegexPrefix(find, out var pattern))
        {
            try
            {
                return Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase);
            }
            catch (System.ArgumentException)
            {
                // Malformed regex — fall through to literal contains so the
                // user still gets usable behavior, never an opaque exception.
            }
        }
        return text.Contains(find, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// CONSISTENCY(find-regex): canonical parser for the `r"..."` / `r'...'`
    /// raw-string regex prefix shared by every find vocabulary (query `~=`,
    /// CLI `--find`, Word/Pptx/Excel Set find/replace). Returns true and the
    /// inner pattern when <paramref name="find"/> is r-prefixed; false for a
    /// plain literal. Centralizing the parse here keeps one source of truth —
    /// do not re-implement the prefix scan anywhere else.
    /// </summary>
    public static bool TryParseRegexPrefix(string find, out string pattern)
    {
        pattern = "";
        if (find.Length >= 3 && find[0] == 'r'
            && (find[1] == '"' || find[1] == '\''))
        {
            var quote = find[1];
            var endIdx = find.LastIndexOf(quote);
            if (endIdx > 1)
            {
                pattern = find[2..endIdx];
                return true;
            }
        }
        return false;
    }

    private static bool MatchOne(DocumentNode node, Condition cond)
    {
        // Resolve actual value from node
        var (hasKey, actualStr) = ResolveValue(node, cond.Key);

        // CONSISTENCY(style-dual-key): paragraph `style` has two surfacings —
        // OOXML styleId (Format["style"]/["styleId"], e.g. "H5") and the
        // user-facing display name (node.Style/Format["styleName"], e.g.
        // "H正文"). The Word handler-level selector matches either; the CLI
        // post-filter must mirror that, otherwise `[style=H正文]` returns the
        // 3 handler-matched paragraphs only to have the post-filter drop them
        // because Format["style"] holds the styleId. styleId= / styleName=
        // are precise keys with no fallback.
        if ((cond.Op == FilterOp.Equal || cond.Op == FilterOp.NotEqual)
            && string.Equals(cond.Key, "style", StringComparison.OrdinalIgnoreCase))
        {
            bool dualHit = StringEquals(node.Style ?? "", cond.Value)
                || (node.Format.TryGetValue("style", out var sid) && StringEquals(sid?.ToString() ?? "", cond.Value))
                || (node.Format.TryGetValue("styleName", out var sname) && StringEquals(sname?.ToString() ?? "", cond.Value));
            return cond.Op == FilterOp.Equal ? dualHit : !dualHit;
        }

        switch (cond.Op)
        {
            case FilterOp.Exists:
                return hasKey && !string.IsNullOrEmpty(actualStr);

            case FilterOp.Equal:
                if (!hasKey) return false;
                return StringEquals(actualStr, cond.Value)
                    || DimensionEquals(actualStr, cond.Value);

            case FilterOp.NotEqual:
                if (!hasKey) return true; // key absent → not equal
                return !StringEquals(actualStr, cond.Value)
                    && !DimensionEquals(actualStr, cond.Value);

            case FilterOp.Contains:
                if (!hasKey) return false;
                return MatchesTextFilter(actualStr, cond.Value);

            case FilterOp.GreaterOrEqual:
                if (!hasKey) return false;
                return CompareNumeric(actualStr, cond.Value) is int ge && ge >= 0;

            case FilterOp.LessOrEqual:
                if (!hasKey) return false;
                return CompareNumeric(actualStr, cond.Value) is int le && le <= 0;

            case FilterOp.GreaterThan:
                if (!hasKey) return false;
                return CompareNumeric(actualStr, cond.Value) is int gt && gt > 0;

            case FilterOp.LessThan:
                if (!hasKey) return false;
                return CompareNumeric(actualStr, cond.Value) is int lt && lt < 0;

            default:
                return true;
        }
    }

    private static (bool HasKey, string Value) ResolveValue(DocumentNode node, string key)
    {
        // Case-insensitive Format key lookup (highest priority)
        var matchedKey = node.Format.Keys.FirstOrDefault(k =>
            string.Equals(k, key, StringComparison.OrdinalIgnoreCase));

        if (matchedKey != null)
        {
            var val = node.Format[matchedKey];
            return (true, val?.ToString() ?? "");
        }

        // "text" falls back to node.Text if not in Format
        if (string.Equals(key, "text", StringComparison.OrdinalIgnoreCase))
        {
            return (node.Text != null, node.Text ?? "");
        }

        // "type" falls back to node.Type if not in Format
        if (string.Equals(key, "type", StringComparison.OrdinalIgnoreCase))
        {
            return (!string.IsNullOrEmpty(node.Type), node.Type ?? "");
        }

        // Excel cells expose their displayed value as node.Text, not
        // Format["value"]. Map the user-facing `value` filter key onto Text so
        // numeric/equality post-filters (e.g. cell[value>100]) resolve. Gated on
        // the cell node type so a future `value` Format key on other node kinds
        // is not shadowed.
        if (string.Equals(key, "value", StringComparison.OrdinalIgnoreCase)
            && string.Equals(node.Type, "cell", StringComparison.OrdinalIgnoreCase))
        {
            return (node.Text != null, node.Text ?? "");
        }

        // BUG-BT-R6-01: "style" falls back to node.Style if not in Format.
        // Word/PPT handlers populate the top-level DocumentNode.Style property
        // (serialized as the top-level "style" key in JSON output) but do NOT
        // duplicate it into Format. Without this fallback, query selectors
        // like `paragraph[style=Normal]` returned 0 results even though every
        // paragraph in the document literally had style="Normal".
        if (string.Equals(key, "style", StringComparison.OrdinalIgnoreCase))
        {
            return (!string.IsNullOrEmpty(node.Style), node.Style ?? "");
        }

        return (false, "");
    }

    private static bool StringEquals(string a, string b)
    {
        if (string.Equals(a, b, StringComparison.OrdinalIgnoreCase))
            return true;
        // Normalize color hex: "#FF0000" matches "FF0000" and vice versa
        var aNorm = a.TrimStart('#');
        var bNorm = b.TrimStart('#');
        if (aNorm != a || bNorm != b)
            return string.Equals(aNorm, bNorm, StringComparison.OrdinalIgnoreCase);
        return false;
    }

    private static bool DimensionEquals(string actual, string expected)
    {
        // Fuzzy EMU equality (±500) is for unit-qualified dimensions so that
        // e.g. "2cm" matches a stored "2.0cm". It must NOT fire for unitless
        // numbers: EmuConverter parses a bare "50" as 50 EMU, so "50" and "150"
        // fall within the ±500 tolerance and would be judged equal — wrong for
        // cell values (cell[value!=150] then drops every cell). Require an
        // explicit unit on at least one side; bare numbers compare exactly via
        // StringEquals instead.
        if (string.IsNullOrEmpty(ExtractUnit(actual)) && string.IsNullOrEmpty(ExtractUnit(expected)))
            return false;
        if (EmuConverter.TryParseEmu(actual, out var a) && EmuConverter.TryParseEmu(expected, out var b))
            return Math.Abs(a - b) <= 500;
        return false;
    }

    /// <summary>
    /// Compare two values numerically. Supports:
    /// - Plain numbers: "24", "1.5"
    /// - pt-suffixed: "24pt", "10.5pt"
    /// - EMU/dimension values: "2cm", "1in"
    /// Returns negative if actual &lt; expected, 0 if equal, positive if actual &gt; expected.
    /// Returns <c>null</c> when the values are not both numerically/dimensionally
    /// comparable. The &gt;/&lt;/&gt;=/&lt;= operators treat null as "no match" so a
    /// numeric filter never matches non-numeric text via a string comparison —
    /// e.g. cell[value&gt;5000] must NOT match the text cell "张三" (whose code
    /// points would otherwise sort above "5000").
    /// </summary>
    private static int? CompareNumeric(string actual, string expected)
    {
        // Try plain decimal comparison (handles "24", "1.5", "24pt" vs "20pt", etc.)
        var actualNum = ExtractNumber(actual);
        var expectedNum = ExtractNumber(expected);

        if (actualNum.HasValue && expectedNum.HasValue)
        {
            // If both have the same unit suffix (or none), compare directly
            var actualUnit = ExtractUnit(actual);
            var expectedUnit = ExtractUnit(expected);
            if (string.Equals(actualUnit, expectedUnit, StringComparison.OrdinalIgnoreCase)
                || string.IsNullOrEmpty(actualUnit) || string.IsNullOrEmpty(expectedUnit))
            {
                return actualNum.Value.CompareTo(expectedNum.Value);
            }
        }

        // Try EMU-based dimension comparison (handles mixed units: "2cm" vs "1in")
        if (EmuConverter.TryParseEmu(actual, out var actualEmu) && EmuConverter.TryParseEmu(expected, out var expectedEmu))
        {
            return actualEmu.CompareTo(expectedEmu);
        }

        // Fallback: plain number comparison (mixed units, both unitless numbers)
        if (actualNum.HasValue && expectedNum.HasValue)
            return actualNum.Value.CompareTo(expectedNum.Value);

        // Not numerically comparable — no string fallback. A numeric operator on
        // a non-numeric value is "no match", not a lexical comparison.
        return null;
    }

    private static decimal? ExtractNumber(string value)
    {
        if (string.IsNullOrEmpty(value)) return null;

        // Strip known unit suffixes
        var trimmed = value.TrimEnd();
        foreach (var suffix in new[] { "pt", "px", "cm", "in", "em", "%" })
        {
            if (trimmed.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
            {
                trimmed = trimmed[..^suffix.Length];
                break;
            }
        }

        return decimal.TryParse(trimmed, NumberStyles.Any, CultureInfo.InvariantCulture, out var n) ? n : null;
    }

    private static string ExtractUnit(string value)
    {
        if (string.IsNullOrEmpty(value)) return "";
        foreach (var suffix in new[] { "pt", "px", "cm", "in", "em", "%" })
        {
            if (value.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                return suffix;
        }
        return "";
    }
}
