// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Handlers;

public static partial class WordBatchEmitter
{

    private static Dictionary<string, string> BuildChartProps(ChartSpec spec)
    {
        // AddChart ingests data series via a single `data="Name1:v1,v2;Name2:v1,v2"`
        // string. Reconstruct that string from the series children Get
        // exposes; categories come from the chart's own Format key.
        var props = FilterEmittableProps(spec.Format);
        // Strip Get-only / SDK-managed keys that AddChart neither expects
        // nor accepts.
        props.Remove("id");
        props.Remove("seriesCount");

        // Build data="Name:v1,v2;..." from series children, plus the per-series
        // dotted props AddChart honors (series{N}.color, series{N}.dataLabels).
        // N is 1-based over the EMITTED (non-reference-line) series, matching how
        // AddChart maps the dotted keys onto the series it parses from `data=`.
        // Without this, a bar/column chart whose series carry explicit fill
        // colors (Q1 blue / Q2 orange / Q3 green) round-tripped to Word's default
        // single-color palette (all blue), and per-series data labels (showVal)
        // were dropped. Reader already surfaces series Format["color"] and
        // Format["dataLabels"]; only this emit was missing.
        var seriesParts = new List<string>();
        int seriesIdx = 0;
        bool emittedPerSeriesFill = false;
        foreach (var s in spec.Series)
        {
            if (s.Type != "series") continue;
            // Skip reference-line series: AddReferenceLine re-creates the Target
            // series from `referenceLine=...` props. Including its values in the
            // data string would duplicate the series on replay.
            if (s.Format.TryGetValue("refLine", out var rl) && rl?.ToString() == "true") continue;
            if (!s.Format.TryGetValue("name", out var nObj) || nObj == null) continue;
            if (!s.Format.TryGetValue("values", out var vObj) || vObj == null) continue;
            var name = nObj.ToString() ?? "";
            var vals = vObj.ToString() ?? "";
            if (name.Length == 0 || vals.Length == 0) continue;
            seriesParts.Add($"{name}:{vals}");
            seriesIdx++;

            // Per-series explicit fill color (srgbClr / "none"). Reader only sets
            // this key when the series carries an explicit <c:spPr> fill, so its
            // presence is meaningful — forward it verbatim.
            if (s.Format.TryGetValue("color", out var cObj) && cObj != null)
            {
                var c = cObj.ToString() ?? "";
                if (c.Length > 0) { props[$"series{seriesIdx}.color"] = c; emittedPerSeriesFill = true; }
            }
            // Per-series gradient fill (e.g. "5B9BD5-2E75B6:90:s0"). Distinct
            // from a chart-level `gradient=` which applies one gradient to ALL
            // series — here each series can carry its own (Q1 blue / Q2 orange).
            else if (s.Format.TryGetValue("gradient", out var gObj) && gObj != null)
            {
                var g = gObj.ToString() ?? "";
                if (g.Length > 0) { props[$"series{seriesIdx}.gradient"] = g; emittedPerSeriesFill = true; }
            }
            // NOTE: per-series data-label show flags (Format["dataLabels"]) are
            // NOT emitted yet — AddChart has no per-series data-label setter, and
            // routing them to the chart-level datalabels.show* keys tripped a
            // schema-order bug in that handler (<c:showVal> emitted as an
            // unexpected child of an existing <c:dLbls>). Deferred: data-label
            // round-trip needs the chart-level dLbls builder fixed first.
        }
        if (seriesParts.Count > 0)
        {
            props["data"] = string.Join(";", seriesParts);
        }
        // The chart-level `gradient` (and `color`) key the Reader surfaces is a
        // chart-scope REPRESENTATIVE — it reports the FIRST series' fill. When
        // series carry distinct per-series fills, replaying that chart-level key
        // would paint EVERY series with series-1's fill (the "all blue" bug).
        // Drop it once per-series fills are emitted.
        if (emittedPerSeriesFill)
        {
            props.Remove("gradient");
            props.Remove("color");
        }
        return props;
    }
}
