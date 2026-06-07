// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class WordBatchEmitter
{

    // RawXmlHelper.Execute propagates the root's xmlns declarations onto every
    // direct child so the SDK's InnerXml setter can resolve prefixes (SDK does
    // not inherit root xmlns scope when parsing inner content). After replay,
    // the part's XML carries redundant xmlns attrs on each child, which the
    // next dump reads back verbatim — phantom bloat that breaks idempotency.
    //
    // Canonicalize on emit: parse the part's XML, drop child-element xmlns
    // declarations that match the root's declarations, re-serialize. The
    // first-pass emit (source's clean XML) and second-pass emit (target's
    // bloated XML) both collapse to the same canonical shape.
    // Schema order for <w:ind>'s attributes. SDK serialises in this order
    // on write, so source-side OuterXml (which mirrors the on-disk order
    // from the original producer) and replay-target OuterXml (SDK's
    // canonical) can disagree on attribute order alone. Re-sort to a fixed
    // canonical order so both passes emit identical bytes.
    private static readonly string[] s_indAttrOrder =
    [
        "start", "end", "left", "right",
        "hanging", "firstLine",
        "startChars", "endChars", "leftChars", "rightChars",
        "hangingChars", "firstLineChars",
    ];

    private static void SortIndAttrs(System.Xml.Linq.XElement ind)
    {
        var attrs = ind.Attributes().ToList();
        // Keep xmlns declarations first (in original order), then sort
        // typed attrs by the schema-order table, then unknown attrs by name.
        var nsDecls = attrs.Where(a => a.IsNamespaceDeclaration).ToList();
        var typed = attrs.Where(a => !a.IsNamespaceDeclaration).ToList();
        int OrderKey(System.Xml.Linq.XAttribute a)
        {
            var idx = Array.IndexOf(s_indAttrOrder, a.Name.LocalName);
            return idx < 0 ? 99 : idx;
        }
        var sorted = typed.OrderBy(OrderKey).ThenBy(a => a.Name.LocalName, StringComparer.Ordinal).ToList();
        ind.RemoveAttributes();
        foreach (var a in nsDecls) ind.Add(a);
        foreach (var a in sorted) ind.Add(a);
    }

    private static void RenameAttr(System.Xml.Linq.XElement el, string fromLocal, string toLocal, string ns)
    {
        var fromName = System.Xml.Linq.XName.Get(fromLocal, ns);
        var toName = System.Xml.Linq.XName.Get(toLocal, ns);
        var src = el.Attribute(fromName);
        if (src == null) return;
        if (el.Attribute(toName) != null) { src.Remove(); return; }
        // Preserve attribute order: re-build the attribute list with the
        // rename applied in-place. SetAttributeValue(newName) by itself would
        // append the new attr at the tail and shift byte order.
        var rebuilt = el.Attributes()
            .Select(a => a.Name == fromName
                ? new System.Xml.Linq.XAttribute(toName, a.Value)
                : new System.Xml.Linq.XAttribute(a.Name, a.Value))
            .ToList();
        el.RemoveAttributes();
        foreach (var a in rebuilt) el.Add(a);
    }

    private static string CanonicalizeRawXml(string xml)
    {
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) return xml;
        try
        {
            var doc = System.Xml.Linq.XDocument.Parse(xml);
            if (doc.Root == null) return xml;
            var rootNsAttrs = doc.Root.Attributes()
                .Where(a => a.IsNamespaceDeclaration)
                .ToDictionary(a => a.Name, a => a.Value);
            foreach (var desc in doc.Root.Descendants())
            {
                var toRemove = desc.Attributes()
                    .Where(a => a.IsNamespaceDeclaration
                                && rootNsAttrs.TryGetValue(a.Name, out var v)
                                && v == a.Value)
                    .ToList();
                foreach (var a in toRemove) a.Remove();
            }
            // SDK normalises bidi-aware <w:ind w:start="…"> ↔ <w:ind w:left="…">
            // (and end ↔ right) on serialisation depending on the document's
            // bidi state. The two forms are byte-different but semantically
            // equivalent in non-bidi documents. Canonicalise to the bidi-
            // aware names AND fix the attribute order so the dump pair emits
            // identical bytes regardless of SDK's choice.
            var wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            foreach (var ind in doc.Descendants(System.Xml.Linq.XName.Get("ind", wNs)))
            {
                RenameAttr(ind, "left", "start", wNs);
                RenameAttr(ind, "right", "end", wNs);
                // BIDI-aware character-count variants also drift through SDK
                // normalisation. proof_fixed family: <w:ind … w:leftChars="0" …>
                // → SDK rewrites as <w:ind … w:startChars="0" …>.
                RenameAttr(ind, "leftChars", "startChars", wNs);
                RenameAttr(ind, "rightChars", "endChars", wNs);
                SortIndAttrs(ind);
            }
            // Stabilise root attribute order: SDK serialises xmlns attrs in
            // an internal order that can shift when mc:Ignorable / other
            // typed attrs change, so byte-equal round-trip needs a canonical
            // ordering. Emit xmlns attrs first (sorted by prefix; default
            // xmlns first if any), then non-xmlns attrs (sorted by name).
            var root = doc.Root;
            var allAttrs = root.Attributes().ToList();
            foreach (var a in allAttrs) a.Remove();
            var nsAttrs = allAttrs.Where(a => a.IsNamespaceDeclaration)
                .OrderBy(a => a.Name == System.Xml.Linq.XNamespace.Xmlns + "xmlns" ? "" : a.Name.LocalName,
                         StringComparer.Ordinal)
                .ToList();
            var otherAttrs = allAttrs.Where(a => !a.IsNamespaceDeclaration)
                .OrderBy(a => a.Name.NamespaceName, StringComparer.Ordinal)
                .ThenBy(a => a.Name.LocalName, StringComparer.Ordinal)
                .ToList();
            foreach (var a in nsAttrs) root.Add(a);
            foreach (var a in otherAttrs) root.Add(a);
            return root.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        }
        catch
        {
            // Malformed XML — leave as-is rather than corrupting.
            return xml;
        }
    }

    // <w:footnotePr>/<w:endnotePr> in settings.xml carry <w:footnote>/<w:endnote>
    // child refs pointing at the separator + continuationSeparator notes
    // (typically id="-1" and id="0") that live in footnotes.xml / endnotes.xml.
    // The dump round-trips note CONTENT via typed `add footnote`/`add endnote`
    // (body-referenced notes only) — it never recreates a separator-only notes
    // part. So when a source carries those separator refs but no body-referenced
    // notes (a footnotes.xml holding ONLY the id -1/0 separators, which every
    // Word doc has), replaying the settings raw-set leaves the refs pointing at
    // notes parts that don't exist in the blank target, and the referential
    // validator rejects the result ("w:footnote … does not exist in part
    // /MainDocumentPart/FootnotesPart"). Word auto-manages separators, so
    // dropping these refs is lossless; strip them while keeping footnotePr/
    // endnotePr and any real config children (pos, numFmt, numStart, …).
    //
    // ALSO strips any settings element carrying a relationship reference
    // (an attribute in the `r:` namespace — r:id / r:embed / r:link). The dump
    // never recreates settings.xml.rels, so e.g. <w:attachedTemplate r:id="…"/>
    // (the pointer to Normal.dotm that nearly every real Word document carries)
    // dangles on replay — the OOXML validator NRE'd "before producing results"
    // and real Word refused to open the file. Dropping the pointer is lossless
    // (Word falls back to the Normal template). Same family as <w:mailMerge>'s
    // data-source r:id etc.; remove the whole referencing element.
    private static string StripDanglingNoteSeparatorRefs(string settingsXml)
    {
        if (string.IsNullOrEmpty(settingsXml) || !settingsXml.StartsWith("<")) return settingsXml;
        // Fast path: nothing to strip unless a note-properties block or a
        // relationship-bearing element (r: prefixed attribute) is present.
        if (!settingsXml.Contains("notePr") && !settingsXml.Contains(":id=")
            && !settingsXml.Contains(":embed=") && !settingsXml.Contains(":link="))
            return settingsXml;
        try
        {
            var doc = System.Xml.Linq.XDocument.Parse(settingsXml);
            if (doc.Root == null) return settingsXml;
            var wNs = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var rNs = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            var removed = false;
            foreach (var pr in doc.Descendants(wNs + "footnotePr")
                         .Concat(doc.Descendants(wNs + "endnotePr")).ToList())
            {
                foreach (var sep in pr.Elements(wNs + "footnote")
                             .Concat(pr.Elements(wNs + "endnote")).ToList())
                {
                    sep.Remove();
                    removed = true;
                }
            }
            // Drop any element with a dangling relationship reference (settings
            // .xml.rels is not round-tripped). attachedTemplate is the common one.
            foreach (var el in doc.Descendants().ToList())
            {
                if (el.Attributes().Any(a => a.Name.Namespace == rNs))
                {
                    el.Remove();
                    removed = true;
                }
            }
            if (!removed) return settingsXml;
            return doc.Root.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        }
        catch
        {
            return settingsXml;
        }
    }

    // <w:numPicBullet> defines a picture (image) list bullet; a level opts into
    // it with <w:lvlPicBulletId>. The picture lives in word/media/* referenced
    // by numbering.xml.rels (r:id inside the numPicBullet's VML/drawing). The
    // dump round-trips numbering.xml verbatim via raw-set but never recreates
    // the numbering part's rels or the media binary, so the r:id dangles on
    // replay — real Word then refuses to open the file ("may be corrupt") and
    // the SDK validator NREs walking the broken numPicBullet. Strip the
    // numPicBullet definitions AND the lvlPicBulletId opt-ins so the level
    // falls back to its own <w:lvlText> glyph (already round-tripped). Lossy
    // for picture bullets only; mirrors the dangling footnote/endnote separator
    // ref strip and the external-rel SDT fallback.
    private static string StripDanglingPicBullets(string numberingXml)
    {
        if (string.IsNullOrEmpty(numberingXml) || !numberingXml.StartsWith("<")) return numberingXml;
        if (!numberingXml.Contains("PicBullet")) return numberingXml; // fast path
        try
        {
            var doc = System.Xml.Linq.XDocument.Parse(numberingXml);
            if (doc.Root == null) return numberingXml;
            var wNs = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var removed = false;
            foreach (var el in doc.Descendants(wNs + "numPicBullet")
                         .Concat(doc.Descendants(wNs + "lvlPicBulletId")).ToList())
            {
                el.Remove();
                removed = true;
            }
            if (!removed) return numberingXml;
            return doc.Root.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        }
        catch
        {
            return numberingXml;
        }
    }

    // Round-trip the source's <w:docDefaults> (the document-wide rPr/pPr
    // baseline inside styles.xml) VERBATIM via raw-set replace. This is the
    // root fix for "blank-default pollution": BlankDocCreator stamps an
    // opinionated docDefaults (Calibri, sz=22/11pt, szCs=22) that a source
    // omitting a slot — calibre/pandoc exports routinely carry only szCs, or
    // only a complex-script font, leaving the Latin size/lang/textAlignment to
    // Word's application default — would otherwise inherit on replay, rendering
    // at the wrong size/font. Per-property emits (docDefaults.font.latin,
    // docDefaults.fontSize, …) only covered the slots the source set
    // EXPLICITLY and could not express "this slot is absent", so the blank's
    // value leaked through. Replacing the whole block makes the rebuilt
    // docDefaults byte-identical to the source — including its absences — so
    // Word applies the same defaults to both. Mirrors the theme/settings/
    // numbering raw-emit rationale (structured XML edited as a block).
    private static void EmitDocDefaultsRaw(WordHandler word, List<BatchItem> items)
    {
        string stylesXml;
        try { stylesXml = word.Raw("/styles"); }
        catch { return; }
        if (string.IsNullOrEmpty(stylesXml) || !stylesXml.StartsWith("<")) return;
        string? dd;
        try
        {
            var doc = System.Xml.Linq.XDocument.Parse(stylesXml);
            var wNs = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var el = doc.Root?.Element(wNs + "docDefaults");
            // Source carries no docDefaults — leave the blank target's in place
            // (removing it risks a docDefaults-less styles.xml some consumers
            // dislike; absence here is rare and low-impact).
            dd = el?.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        }
        catch { return; }
        if (string.IsNullOrEmpty(dd)) return;

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/styles",
            Xpath = "/w:styles/w:docDefaults",
            Action = "replace",
            Xml = dd
        });
    }

    private static void EmitThemeRaw(WordHandler word, List<BatchItem> items)
    {
        // Theme carries clrScheme + fontScheme + fmtScheme — pure structured
        // XML that users rarely modify property-by-property; the natural
        // operation is "swap the entire theme block". Raw-set replace fits
        // that model exactly. Word.Raw returns the literal string
        // "(no theme)" when the part is missing.
        //
        // ALWAYS emit, even for source docs that have no theme part. The
        // blank target auto-stamps theme1.xml (for Word render
        // parity), so silently skipping the emit caused dump∘replay∘dump
        // to drift by +1 item every pass: dump-1 saw no theme and
        // emitted nothing; replay left blank's theme in place; dump-2
        // saw blank's theme and emitted it. Dump-1 now emits an empty
        // <a:theme/> placeholder for theme-less sources, which the apply
        // path overwrites blank's seeded theme with — making dump-2 see
        // the same empty theme and emit the same placeholder. Fixed point.
        string xml;
        try { xml = word.Raw("/theme"); }
        catch { xml = ""; }
        xml = CanonicalizeRawXml(xml);
        // A bare <a:theme/> (or <a:theme name="Office Theme"/>) is schema-INVALID:
        // <a:theme> requires a child <a:themeElements> (clrScheme + fontScheme +
        // fmtScheme). Replaying that placeholder over the blank target's valid
        // theme1.xml produced a file real Word refuses to open ("file may be
        // corrupt"); the source docx that triggered this carried a 0-byte
        // theme1.xml, which Word tolerates but the SDK read back as an empty
        // theme. Emit the SAME complete theme a blank doc stamps instead: the
        // result is Word-openable AND keeps the dump→replay→dump item count
        // stable (replay writes a real theme, dump-2 reads it back and emits
        // one theme item, same as dump-1) — the original reason this site
        // always emits something rather than skipping theme-less sources.
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<") || !xml.Contains("themeElements"))
            xml = BlankDocCreator.BuildDefaultTheme(null, null).OuterXml;

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/theme",
            Xpath = "/a:theme",
            Action = "replace",
            Xml = xml
        });
    }

    private static void EmitSettingsRaw(WordHandler word, List<BatchItem> items)
    {
        // Settings carries dozens of feature flags + compat shims that
        // surface on root.Format only piecemeal — and not all of them are
        // wired through Set's case table. Wholesale raw-set is the simplest
        // way to keep Word feature toggles (evenAndOddHeaders, mirrorMargins,
        // schema-pegged compat options, …) round-tripped without
        // per-property allowlisting.
        //
        // ALWAYS emit, even for source docs without a settings part. The
        // blank target auto-stamps a settings.xml (characterSpacingControl
        // + compat block), so silently skipping the emit caused the same
        // idempotency drift as EmitThemeRaw: dump-1 saw no settings and
        // emitted nothing, dump-2 saw blank's leftover and emitted it.
        // Empty placeholder clears blank's seeded settings so dump-2
        // reads the same empty state and emits the same placeholder.
        string xml;
        try { xml = word.Raw("/settings"); }
        catch { xml = ""; }
        xml = CanonicalizeRawXml(xml);
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<"))
            xml = "<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" />";
        xml = StripDanglingNoteSeparatorRefs(xml);

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/settings",
            Xpath = "/w:settings",
            Action = "replace",
            Xml = xml
        });
    }

    private static void EmitNumberingRaw(WordHandler word, List<BatchItem> items)
    {
        // Numbering models list templates (abstractNum + num pairs, each
        // abstractNum holds 9 levels with their own pPr / numFmt / lvlText).
        // Reconstructing this through typed Add would mean another emitter
        // in itself; for v0.5 we ship the entire <w:numbering> XML wholesale
        // via raw-set. The blank document creates an empty numbering part,
        // so a single replace on the part root is sufficient.
        string xml;
        try { xml = word.Raw("/numbering"); }
        catch { return; }
        xml = CanonicalizeRawXml(xml);
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) return;
        // Skip when numbering is empty (just `<w:numbering/>` with no children).
        if (!xml.Contains("<w:abstractNum") && !xml.Contains("<w:num "))
            return;
        xml = StripDanglingPicBullets(xml);

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/numbering",
            Xpath = "/w:numbering",
            Action = "replace",
            Xml = xml
        });
    }

    private static void EmitHeadersFooters(WordHandler word, List<BatchItem> items)
    {
        var root = word.Get("/");
        if (root.Children == null) return;
        // BUG-X4-T2: header/footer parts carry no `type` key on Get; the
        // section's `headerRef.default|first|even` (and `footerRef.*`)
        // entries are the only place the part's role is recorded. Build a
        // reverse lookup so EmitHeaderFooterPart can emit the right
        // `type` prop (default/first/even) instead of always emitting
        // "default" — which on a doc with both default + first headers
        // throws "Header of type 'default' already exists" on replay.
        // In addition to (path → type), track which section's headerRef /
        // footerRef points at the part. Multi-section docs with per-section
        // default headers used to all emit `add header parent="/"` —
        // AddHeader resolves "/" to a single sectPr, so the 2nd-and-later
        // default headers tripped "Header of type 'default' already exists"
        // on replay. Emit `parent=/section[N]` so each header targets its
        // true owning section (mirrors ResolveTargetSectPrForHeaderFooter's
        // /section[N] resolver).
        var headerPathInfo = new Dictionary<string, (string Type, string? SectionPath)>(StringComparer.OrdinalIgnoreCase);
        var footerPathInfo = new Dictionary<string, (string Type, string? SectionPath)>(StringComparer.OrdinalIgnoreCase);
        // headerRef.<type> / footerRef.<type> live on **section** nodes
        // (see WordHandler.Query.cs:902), not on root. An earlier fix
        // scanned root.Format and silently found nothing, so every emitted
        // header/footer was typed "default" — round-trip failed when a doc
        // had both default + first headers. Walk all section children to
        // build the path→type map.
        void HarvestRefs(DocumentNode node, string? sectionPath)
        {
            foreach (var (key, val) in node.Format)
            {
                if (val == null) continue;
                var s = val.ToString();
                if (string.IsNullOrEmpty(s)) continue;
                if (key.StartsWith("headerRef.", StringComparison.OrdinalIgnoreCase))
                {
                    var t = key["headerRef.".Length..];
                    if (!headerPathInfo.ContainsKey(s))
                        headerPathInfo[s] = (t, sectionPath);
                }
                else if (key.StartsWith("footerRef.", StringComparison.OrdinalIgnoreCase))
                {
                    var t = key["footerRef.".Length..];
                    if (!footerPathInfo.ContainsKey(s))
                        footerPathInfo[s] = (t, sectionPath);
                }
            }
        }
        // Harvest sections FIRST so real section attribution wins over the
        // body-level sectPr fallback. Then harvest root: root.Format mirrors
        // the body-level <w:sectPr> (which OOXML treats as the FINAL section's
        // properties), so any ref that only appears at root belongs to the
        // last section. Emitting `parent="/"` for those is fine because
        // AddHeader's `/` resolver also falls back to the body-level sectPr —
        // both ends agree on the same section. Earlier the order was reversed
        // (root first, sections second) and the `if !ContainsKey` guard
        // wrongly let root entries shadow real section attribution.
        List<DocumentNode> sectionList = new();
        try
        {
            var sections = word.Query("section");
            if (sections != null) sectionList = sections.ToList();
        }
        catch { /* missing section info — fall through with default typing */ }
        foreach (var sec in sectionList) HarvestRefs(sec, sec.Path);
        var rootFallbackSection = sectionList.Count > 0 ? sectionList[^1].Path : null;
        HarvestRefs(root, rootFallbackSection);

        int hIdx = 0, fIdx = 0;
        foreach (var child in root.Children)
        {
            if (child.Type == "header")
            {
                // Skip orphaned header parts (present in the package but
                // not referenced by any section's w:headerReference). Re-
                // emitting them as `add header type=default` collides with
                // the real default header on batch replay ("Header of type
                // 'default' already exists"). Only re-emit parts that a
                // section actually links to.
                if (!headerPathInfo.TryGetValue(child.Path, out var hi)) continue;
                hIdx++;
                EmitHeaderFooterPart(word, child.Path, "header", hIdx, items, hi.Type, hi.SectionPath);
            }
            else if (child.Type == "footer")
            {
                // Same orphan guard as header above.
                if (!footerPathInfo.TryGetValue(child.Path, out var fi)) continue;
                fIdx++;
                EmitHeaderFooterPart(word, child.Path, "footer", fIdx, items, fi.Type, fi.SectionPath);
            }
        }
    }

    private static void EmitHeaderFooterPart(WordHandler word, string sourcePath, string kind,
                                             int targetIndex, List<BatchItem> items,
                                             string subTypeOverride = "default",
                                             string? sectionParent = null)
    {
        var partNode = word.Get(sourcePath);
        // BUG-DUMP9-08: tables are valid block-level OOXML inside hdr/ftr
        // (same schema as body) and Navigation surfaces them as `table`-typed
        // children, but the previous filter only kept paragraphs and silently
        // dropped tables. Iterate in source order, tracking per-type indices
        // so paragraph and table paths line up with replay output.
        var blockChildren = (partNode.Children ?? new List<DocumentNode>())
            .Where(c => c.Type == "paragraph" || c.Type == "p"
                     || c.Type == "table" || c.Type == "tbl")
            .ToList();
        // partNode.Format does not expose `type`; the caller resolves the
        // role (default/first/even) from the section's headerRef.* / footerRef.*
        // map and passes it via subTypeOverride.
        var subType = subTypeOverride;

        // Create the part with just its role (default/first/even). AddHeader/
        // AddFooter seed an empty auto paragraph; EmitParagraph(autoPresent:
        // true) on paras[0] then routes through CollapseFieldChains so a
        // PAGE-field header (the canonical case) round-trips as a typed
        // `add field` row instead of being baked into static "1" text on the
        // seed paragraph (BUG-X4-T3). Run-level formatting on multi-run
        // first paragraphs is preserved by the per-run emit path below.
        var addHeaderProps = new Dictionary<string, string> { ["type"] = subType };
        // First-page header auto-stamps <w:titlePg/> on its section (UX:
        // without titlePg, Word silently ignores type="first" headerRef).
        // Source may have headerRef-first WITHOUT titlePg — preserve that
        // shape by passing noTitlePg=true so AddHeader skips the auto-stamp.
        // Otherwise the next dump would emit a phantom `titlePage=true` key.
        if ((kind == "header" || kind == "footer")
            && string.Equals(subType, "first", StringComparison.OrdinalIgnoreCase)
            && sectionParent != null)
        {
            try
            {
                var sectionNode = word.Get(sectionParent);
                bool sourceHadTitlePg = sectionNode.Format.TryGetValue("titlePage", out var tpv)
                                     && tpv is bool b && b;
                if (!sourceHadTitlePg)
                    addHeaderProps["noTitlePg"] = "true";
            }
            catch { /* section path unresolved — fall through with auto-stamp */ }
        }
        // CONSISTENCY(headerfooter-noEvenAndOdd-opt-out): even-{header,footer}
        // auto-stamps <w:evenAndOddHeaders/> in /settings. Source whose settings
        // lacks the toggle (rare but real — Word renders inconsistently across
        // versions) gets a phantom toggle injected on replay. Suppress by
        // surfacing `noEvenAndOddHeaders=true` so AddHeader/AddFooter skip the
        // stamp. The settings raw-set already replaced /settings with the
        // source xml before this add executes.
        if ((kind == "header" || kind == "footer")
            && string.Equals(subType, "even", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                // BUG-DUMP-R4-02: `Get("/settings")` returns a node whose
                // Format dict is empty — PopulateDocSettings is only called
                // by GetRootNode, not when /settings is resolved directly.
                // Reading `Format["evenAndOddHeaders"]` off the settings node
                // therefore always returned false, so dump emitted a phantom
                // `noEvenAndOddHeaders=true` even when the source's
                // settings.xml carried the toggle. Read from root, which IS
                // populated, mirroring the `titlePage` check above (that one
                // reads off /section[N] which also runs its own populator).
                var rootNode = word.Get("/");
                bool sourceHadToggle = rootNode.Format.TryGetValue("evenAndOddHeaders", out var ev)
                                     && ev is bool eb && eb;
                if (!sourceHadToggle)
                    addHeaderProps["noEvenAndOddHeaders"] = "true";
            }
            catch { /* settings unreadable — fall through */ }
        }
        items.Add(new BatchItem
        {
            Command = "add",
            // Route per-section headers/footers to their owning section
            // (e.g. /section[2]) instead of root "/", so multi-section docs
            // that carry one default header per section don't collide on
            // replay. Falls back to "/" when the part is not owned by any
            // section in the harvested map (defensive — EmitHeadersFooters
            // already filters orphans before reaching here).
            Parent = sectionParent ?? "/",
            Type = kind,
            Props = addHeaderProps
        });

        var partTargetPath = $"/{kind}[{targetIndex}]";
        int pIdx = 0, tblIdx = 0;
        bool sawFirstPara = false;
        foreach (var child in blockChildren)
        {
            if (child.Type == "table" || child.Type == "tbl")
            {
                tblIdx++;
                EmitTable(word, child.Path, tblIdx, items, ctx: null,
                          parentTablePath: null, containerPath: partTargetPath);
            }
            else
            {
                pIdx++;
                EmitParagraph(word, child.Path, partTargetPath, pIdx, items,
                              autoPresent: !sawFirstPara);
                sawFirstPara = true;
            }
        }
    }

    private static void EmitComments(WordHandler word, List<BatchItem> items,
                                     Dictionary<string, int> paraIdToTargetIdx)
    {
        var comments = word.Query("comment");
        foreach (var c in comments)
        {
            var props = FilterEmittableProps(c.Format);
            if (!string.IsNullOrEmpty(c.Text))
                props["text"] = c.Text!;
            // Map anchoredTo (source paraId path) -> target paragraph index.
            // anchoredTo looks like "/body/p[@paraId=00100000]"; parse and
            // resolve via the paraId map we built during EmitBody.
            string parentTarget = "/body/p[1]";  // safe fallback to first body para
            if (props.TryGetValue("anchoredTo", out var anchor))
            {
                var pid = ExtractParaId(anchor);
                if (pid != null && paraIdToTargetIdx.TryGetValue(pid, out var idx))
                    parentTarget = $"/body/p[{idx}]";
                props.Remove("anchoredTo");
            }
            // BUG-DUMP4-03: emit the 1-based run index where the source
            // CommentRangeStart sits inside its paragraph so replay can
            // narrow the anchor instead of widening to the entire para.
            // 0 means "before all runs" (paragraph start); >=1 means
            // "after run N". AddComment already accepts a run-targeted
            // parent path (/body/p[N]/r[M]), but we keep the prop on the
            // paragraph-level emit so the wire format stays uniform with
            // the existing parent-resolution logic — replay can switch on
            // runStart later without changing the schema.
            if (c.Format.TryGetValue("id", out var cid) && cid != null)
            {
                var runStart = word.FindCommentAnchorRunIndex(cid.ToString()!);
                // 0 = before all runs (paragraph start); always emit so
                // replay knows the anchor is positional, not whole-paragraph.
                props["runStart"] = runStart.ToString();
            }
            // The comment id is allocated by AddComment on the target side;
            // do not propagate the source id (would conflict on replay).
            props.Remove("id");
            // BUG-X7-04 (T-4): previously dropped `date` so dump→replay always
            // re-stamped the comment with the SDK's "now". That breaks
            // archival / audit-trail use cases where the source timestamp is
            // load-bearing. Preserve it; AddComment accepts an explicit
            // ISO-8601 date and the SDK will use it instead of stamping.

            items.Add(new BatchItem
            {
                Command = "add",
                Parent = parentTarget,
                Type = "comment",
                Props = props
            });
        }
    }

    // Emit a body-level SDT (Content Control). Simple SDTs (a single text run,
    // dropdown/combobox/date pickers) round-trip as a typed `add /body --type
    // sdt` carrying type/alias/tag/items/format + the visible text — all of
    // which AddSdt rebuilds. Without this, SDTs were silently dropped from dump
    // output (BUG-X2-06 / X2-3).
    //
    // Rich BLOCK SDTs are different: a Table of Contents, or any content control
    // wrapping multiple paragraphs / hyperlinks / fields / a table, carries
    // block structure the text-only path cannot express — it concatenates every
    // inner paragraph into one `text` run, collapsing a multi-line TOC into a
    // single line. Round-trip the whole <w:sdt> verbatim via raw-set instead,
    // inserted before the body's trailing sectPr so it lands at the same spot
    // the sequential `add /body` items build up to (AppendToParent inserts body
    // children before that sectPr). Same rationale as the theme/settings/
    // numbering raw emits: structured XML edited as a block, not per-property.
    private static void EmitSdt(WordHandler word, string sourcePath, List<BatchItem> items, BodyEmitContext ctx)
    {
        var rawXml = word.RawElementXml(sourcePath);
        if (!string.IsNullOrEmpty(rawXml) && IsRichBlockSdt(rawXml!))
        {
            // External relationship references (hyperlink r:id, image r:embed/
            // r:link) would dangle in the blank target — raw injection does not
            // recreate the matching rels. Fall back to the text emit and surface
            // the loss rather than producing a file with broken references.
            if (HasExternalRelRef(rawXml!))
            {
                ctx.Warnings.Add(new DocxUnsupportedWarning(
                    Element: "sdt.richContent",
                    Path: sourcePath,
                    Reason: "content control with rich block content AND external relationship references (hyperlinks/images) flattened to text on dump"));
            }
            else
            {
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = "/document",
                    Xpath = "//w:body/w:sectPr",
                    Action = "insertbefore",
                    Xml = rawXml
                });
                return;
            }
        }

        DocumentNode sdt;
        try { sdt = word.Get(sourcePath); }
        catch { return; }

        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        // Whitelist Get-canonical keys that AddSdt consumes. `editable` is a
        // Get readback (negation of `lock`), the source-side `id` is allocated
        // at creation, so neither is forwarded.
        foreach (var key in new[] { "type", "alias", "tag", "items", "format" })
        {
            if (sdt.Format.TryGetValue(key, out var v) && v != null)
            {
                var s = v.ToString() ?? "";
                if (s.Length > 0) props[key] = s;
            }
        }
        if (!string.IsNullOrEmpty(sdt.Text))
            props["text"] = sdt.Text!;

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = "/body",
            Type = "sdt",
            Props = props
        });
    }

    // A block SDT is "rich" when its content carries structure the text-only
    // typed emit cannot reproduce: more than one paragraph, or any hyperlink /
    // complex field / table / drawing. Such SDTs round-trip verbatim via
    // raw-set; everything else (single text run, form-control pickers) stays on
    // the introspectable typed `add sdt` path.
    private static bool IsRichBlockSdt(string sdtXml)
    {
        // <w:p> / <w:p attr...> — but not <w:pPr>, <w:pict>, <w:proofErr> (the
        // char after "w:p" must be a space or '>').
        if (System.Text.RegularExpressions.Regex.Matches(sdtXml, "<w:p[ >]").Count > 1)
            return true;
        return sdtXml.Contains("<w:hyperlink", StringComparison.Ordinal)
            || sdtXml.Contains("<w:fldChar", StringComparison.Ordinal)
            || sdtXml.Contains("w:instrText", StringComparison.Ordinal)
            || sdtXml.Contains("<w:fldSimple", StringComparison.Ordinal)
            || sdtXml.Contains("<w:tbl", StringComparison.Ordinal)
            || sdtXml.Contains("<w:drawing", StringComparison.Ordinal);
    }

    // Raw injection of an <w:sdt> into the blank target preserves the element
    // verbatim but cannot recreate the package relationships its r:id/r:embed/
    // r:link attributes point at — those would dangle. Detect them so the
    // caller can fall back to the (lossy but valid) text emit.
    private static bool HasExternalRelRef(string xml)
        => xml.Contains("r:id=", StringComparison.Ordinal)
        || xml.Contains("r:embed=", StringComparison.Ordinal)
        || xml.Contains("r:link=", StringComparison.Ordinal);

    private static void EmitSection(WordHandler word, List<BatchItem> items)
    {
        var root = word.Get("/");
        // protectionEnforced has no Set case in WordHandler — `set / protectionEnforced=...`
        // emits a WARNING on every replay regardless of protection state.
        // Enforcement is implicit in any non-"none" protection value (the
        // `protection` Set handler stamps w:enforcement=1 itself), so the
        // separate flag is dump-only metadata with no replay path. Drop it
        // unconditionally; for protection="none" also drop the noisy
        // protection key so round-trips stay clean.
        root.Format.Remove("protectionEnforced");
        if (root.Format.TryGetValue("protection", out var protVal)
            && string.Equals(protVal?.ToString(), "none", StringComparison.OrdinalIgnoreCase))
        {
            root.Format.Remove("protection");
        }
        var blankBaseline = _blankRootBaseline.Value;
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (k, v) in root.Format)
        {
            bool include = RootScalarKeys.Contains(k);
            if (!include)
            {
                foreach (var pref in RootPrefixGroups)
                {
                    if (k.StartsWith(pref, StringComparison.OrdinalIgnoreCase))
                    {
                        include = true;
                        break;
                    }
                }
            }
            if (!include) continue;
            // docDefaults round-trips verbatim via EmitDocDefaultsRaw now —
            // skip the per-property emit here so the two paths don't fight
            // (and so source-absent slots aren't re-stamped from the blank).
            if (k.StartsWith("docDefaults.", StringComparison.OrdinalIgnoreCase)) continue;
            if (v == null) continue;
            var s = v switch { bool b => b ? "true" : "false", _ => v.ToString() ?? "" };
            if (s.Length == 0) continue;
            // Skip when the source's value already matches what BlankDocCreator
            // would stamp. Otherwise dump-then-replay leaves blank's value on
            // the target unchanged, but the SECOND dump picks it up (because
            // the value is now explicit in the part) and emits a `set /` row
            // dump-1 had skipped — losing idempotency. Symmetry: dump-2
            // applies the same rule and also skips. The existing
            // docDefaults.font.latin="" clear below is the inverse case
            // (blank's value is undesirable — actively clear it).
            if (blankBaseline.TryGetValue(k, out var blankVal)
                && string.Equals(blankVal, s, StringComparison.Ordinal))
            {
                continue;
            }
            props[k] = s;
        }
        // NOTE: docDefaults (fonts, size, lang, spacing, …) is no longer
        // emitted property-by-property here — EmitDocDefaultsRaw round-trips
        // the whole <w:docDefaults> block verbatim, which also handles the
        // "source omits a slot the blank stamped" pollution the old
        // per-property clears (bare-font rewrite, the BUG-X6-05 font.latin=""
        // clear) were patching one slot at a time.
        if (props.Count == 0) return;
        items.Add(new BatchItem
        {
            Command = "set",
            Path = "/",
            Props = props
        });
    }

    private static void EmitStyles(WordHandler word, List<BatchItem> items)
    {
        // Use query() rather than walking Get("/styles").Children — the
        // positional /styles/style[N] children Get returns are not
        // addressable on the Get side (style paths resolve by id, not by
        // index). Query produces id-based paths and excludes docDefaults.
        var styles = word.Query("style");
        // Blank-baseline cleanup: BlankDocCreator always stamps a Normal
        // style (for Word render parity — Calibri 11pt, 1.08x
        // line). When the source has no entry for styleId="Normal",
        // skipping the emit leaks the blank's stamped Normal into the
        // replay target — dump-2's dump then emits it as a phantom
        // `add /styles Normal`, breaking idempotency. Always prepend a
        // remove-Normal so target's styles end up matching source's
        // (idempotent: Remove of a missing style is a soft success).
        // When source HAS Normal, EmitStyles below recreates it via the
        // builtin-name upsert path; the redundant remove is harmless and
        // keeps the wire format independent of source/blank divergence.
        bool sourceHasNormal = styles.Any(s =>
            string.Equals(s.Format.TryGetValue("id", out var v) ? v?.ToString() : null,
                          "Normal", StringComparison.Ordinal));
        if (!sourceHasNormal)
        {
            items.Add(new BatchItem
            {
                Command = "remove",
                Path = "/styles/Normal",
            });
        }
        // Dedupe by styleId. A styleId is effectively a key — OOXML requires
        // it unique — but real-world sources (LibreOffice / merged docs) carry
        // duplicates (e.g. 88 <w:style> elements, 58 unique ids). Word itself
        // tolerates this by keeping the FIRST occurrence and ignoring the rest
        // (it opens the file fine). Mirror that: emit each styleId once. Without
        // this, every duplicate replayed as an `add style` that failed the id
        // uniqueness check. (Get-by-id also resolves all duplicates to the first
        // style anyway, so the extra emits were redundant copies.)
        var seenStyleIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var stub in styles)
        {
            // CONSISTENCY(slash-in-style-id): style ids/names containing '/'
            // produce paths like /styles/Style/With/Slash that the path
            // parser splits on. Get fails. Fall back to the Query stub —
            // we lose pPr/rPr details but at least the style stub
            // (id/name/type/basedOn) round-trips, instead of dropping the
            // style entirely (BUG BT-3).
            DocumentNode full;
            try { full = word.Get(stub.Path); }
            catch { full = stub; }
            var props = FilterEmittableProps(full.Format);
            // Ensure id is present (Add requires it for /styles target).
            if (!props.ContainsKey("id") && !props.ContainsKey("styleId"))
            {
                if (props.TryGetValue("name", out var n)) props["id"] = n;
                else continue;
            }
            var emitId = props.GetValueOrDefault("id") ?? props.GetValueOrDefault("styleId");
            if (!string.IsNullOrEmpty(emitId) && !seenStyleIds.Add(emitId))
                continue; // duplicate styleId — keep first, skip the rest (Word's behavior)
            // BUG-X6-03: built-in style ids (Normal / Heading1-9 / Title /
            // …) collide with the blank template's reservations on a
            // fresh batch target. AddStyle is now idempotent for those
            // specific ids (upsert: drop existing + re-add). For non-
            // built-in ids the strict "already exists" check still
            // applies. Emit `add` uniformly so the wire format stays a
            // simple `add`-only stream regardless of style provenance.
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = "/styles",
                Type = "style",
                Props = props
            });
            // BUG-X4-T1: FilterEmittableProps drops the `tabs` scalar (it's a
            // List<Dict>, not stringable). EmitParagraph compensates by
            // emitting per-stop `add tab` rows; EmitStyles must do the same
            // or paragraph-level custom tab stops on a style (Heading TOC
            // leader tabs, etc.) silently disappear on round-trip.
            var styleId = props.TryGetValue("id", out var sid) ? sid
                : props.TryGetValue("styleId", out sid) ? sid : null;
            if (styleId != null && full.Format.TryGetValue("tabs", out var styleTabs))
            {
                EmitTabStops($"/styles/{styleId}", styleTabs, items);
            }
        }
    }

    private sealed class NoteCursor { public int Index; }
}
