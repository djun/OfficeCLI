// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Handlers;

/// <summary>
/// Auxiliary-parts scan — surfaces a warning per package part that the dump
/// surface does NOT round-trip, so silent data loss becomes visible.
///
/// <para>
/// Rationale: <see cref="WordBatchEmitter.EmitWordWithWarnings(WordHandler)"/>
/// emits a curated set of parts (body, styles, numbering, theme, settings,
/// headers/footers, comments, footnotes, endnotes, charts, media). Every
/// other part in the OPC package is silently dropped on dump∘replay —
/// templates with content-control bindings (customXml), AutoText repositories
/// (glossary/document.xml), modern-comment metadata (people/commentsExtended/
/// commentsIds), browser-optimised templates (webSettings), embedded fonts
/// (fontTable + word/fonts/*.odttf), and user-defined custom document
/// properties (docProps/custom.xml entries outside the OfficeCLI.* namespace)
/// all vanish without trace.
/// </para>
///
/// <para>
/// This pass walks the package once and emits a
/// <see cref="WordBatchEmitter.DocxUnsupportedWarning"/> per unrecognised
/// part. The dump command then mirrors these into envelope.warnings (with
/// a stderr "warning:" line for human consumption) — same channel as the
/// OLE / shape unsupported warnings introduced in earlier rounds.
/// </para>
///
/// <para>
/// Allowlist (not denylist) — every part with a curated emit path is listed
/// here. When a new emitter lands (e.g. real glossary support) the matching
/// prefix migrates from <see cref="UnsupportedReasons"/> to
/// <see cref="KnownEmittedPrefixes"/>/<see cref="KnownEmittedExact"/> so the
/// warning disappears.
/// </para>
/// </summary>
public static partial class WordBatchEmitter
{
    // Part URIs (or URI prefixes) that the curated emit path round-trips.
    // OPC package "auto-managed" parts (Content_Types, top-level rels,
    // docProps/core+app — restamped on save) are listed alongside the
    // semantic parts the dump actually emits.
    private static readonly HashSet<string> KnownEmittedExact = new(StringComparer.OrdinalIgnoreCase)
    {
        "/word/document.xml",
        "/word/styles.xml",
        "/word/stylesWithEffects.xml",   // legacy compat, restamped by SDK
        "/word/numbering.xml",
        "/word/settings.xml",
        "/word/comments.xml",
        "/word/footnotes.xml",
        "/word/endnotes.xml",
        // OPC auto-managed
        "/docProps/core.xml",            // restamped by OfficeCliMetadata
        "/docProps/app.xml",             // restamped by OfficeCliMetadata
        "/[Content_Types].xml",
        "/_rels/.rels",
    };

    private static readonly string[] KnownEmittedPrefixes = new[]
    {
        "/word/theme/",            // theme1.xml etc — EmitThemeRaw
        "/word/header",            // header1.xml, header2.xml... EmitHeadersFooters
        "/word/footer",            // footer1.xml etc
        "/word/media/",            // images — picture run emit
        "/word/charts/",           // chart XML + embedded xlsx — chart run emit
        "/word/embeddings/",       // OLE payloads — warning already raised per-run
        "/word/diagrams/",         // SmartArt — partial coverage via shape emit
        "/word/activeX/",          // ActiveX controls (form-control aux)
        "/word/printerSettings/",  // SDK strips on save
        "/word/customizations.xml", // legacy customizations
    };

    // Maps an unsupported part URI (or URI prefix) → human-readable reason.
    // Order matters: prefixes are tested in declaration order; first match wins.
    private static readonly (string Prefix, string Element, string Reason)[] UnsupportedReasons = new[]
    {
        ("/customXml/itemProps",       "customXmlProps",         "customXml schema-store reference dropped on dump"),
        ("/customXml/item",            "customXml",              "customXml data store (SDT/content-control bindings) dropped on dump"),
        ("/customXml/",                "customXml",              "customXml part dropped on dump"),
        ("/word/glossary/",            "glossary",               "Building Blocks / AutoText repository dropped on dump"),
        ("/word/people.xml",           "people",                 "modern-comment author metadata dropped on dump"),
        ("/word/commentsExtended.xml", "commentsExtended",       "modern-comment threading metadata dropped on dump"),
        ("/word/commentsIds.xml",      "commentsIds",            "modern-comment durable-id metadata dropped on dump"),
        ("/word/commentsExtensible.xml","commentsExtensible",    "modern-comment extension metadata dropped on dump"),
        ("/word/webSettings.xml",      "webSettings",            "web-publishing settings dropped on dump"),
        ("/word/fontTable.xml",        "fontTable",              "embedded-font table dropped on dump"),
        ("/word/fonts/",               "embeddedFont",           "embedded font binary (.odttf) dropped on dump"),
        ("/word/vbaProject.bin",       "vbaProject",             "VBA macro project dropped on dump"),
        ("/word/vbaData.xml",          "vbaData",                "VBA macro metadata dropped on dump"),
    };

    /// <summary>
    /// Walk the package once, comparing each part's zip-URI against the
    /// emitted allowlist; record a warning for every part the dump surface
    /// does NOT round-trip. Also probes <c>docProps/custom.xml</c> for
    /// user-defined properties (any name outside the <c>OfficeCLI.*</c>
    /// namespace is silently dropped on save).
    /// </summary>
    internal static void EmitAuxiliaryPartsScan(WordHandler word, List<DocxUnsupportedWarning> warnings)
    {
        IEnumerable<string> parts;
        try { parts = word.EnumeratePartUris(); }
        catch { return; }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var uri in parts)
        {
            if (!seen.Add(uri)) continue;
            // Relationship parts (.rels) are auto-managed by the SDK alongside
            // their owning part — skip uniformly. Without this, every emitted
            // part's `_rels/<name>.xml.rels` would also surface as a warning.
            if (uri.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) continue;

            if (KnownEmittedExact.Contains(uri)) continue;
            if (KnownEmittedPrefixes.Any(p => uri.StartsWith(p, StringComparison.OrdinalIgnoreCase))) continue;

            // Special-case: docProps/custom.xml — OfficeCliMetadata always
            // restamps OfficeCLI.* entries; user-authored entries are silently
            // dropped on save. Warn only if the part carries non-OfficeCLI
            // properties; the auto-stamped pair (OfficeCLI.Version + .LastModified)
            // is expected and not a loss.
            if (string.Equals(uri, "/docProps/custom.xml", StringComparison.OrdinalIgnoreCase))
            {
                var userProps = word.EnumerateCustomDocPropertyNames()
                    .Where(n => !n.StartsWith("OfficeCLI.", StringComparison.Ordinal))
                    .ToList();
                if (userProps.Count > 0)
                {
                    warnings.Add(new DocxUnsupportedWarning(
                        Element: "customDocProperty",
                        Path: uri,
                        Reason: $"user-defined custom document properties dropped on dump ({string.Join(", ", userProps)})"));
                }
                continue;
            }

            // Look up the catalogued reason; if none matches, emit a generic
            // "unknown part" warning so silent loss never goes unreported.
            string element = "auxiliaryPart";
            string reason = "package part dropped on dump (no curated emit path)";
            foreach (var (prefix, elt, why) in UnsupportedReasons)
            {
                if (uri.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    element = elt;
                    reason = why;
                    break;
                }
            }

            warnings.Add(new DocxUnsupportedWarning(
                Element: element,
                Path: uri,
                Reason: reason));
        }
    }
}
