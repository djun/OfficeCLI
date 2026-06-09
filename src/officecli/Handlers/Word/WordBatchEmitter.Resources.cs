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

    // BUG-R4B(BUG6): the theme part (word/theme/theme1.xml) can carry a
    // <a:blipFill><a:blip r:embed="rIdN"/> referencing an image relationship in
    // theme1.xml.rels (a custom fmtScheme bg fill). The dump round-trips the
    // theme XML verbatim via raw-set but never recreates the theme part's rels
    // or the media binary, so the r:embed dangles on replay and the rebuilt
    // theme1.xml fails validation ("relationship 'rId1' ... does not exist").
    // Theme-image round-trip is a separate feature; here we just ensure the
    // rebuilt theme has NO dangling reference: strip the unreconstructable
    // blip/blipFill cleanly (falling back to the previous fill in the same
    // *StyleLst, or dropping the entry) and signal the loss via a warning.
    // Mirrors StripDanglingNoteSeparatorRefs / StripDanglingPicBullets.
    private static string StripDanglingThemeBlipRefs(string themeXml, out bool stripped)
    {
        stripped = false;
        if (string.IsNullOrEmpty(themeXml) || !themeXml.StartsWith("<")) return themeXml;
        // Fast path: only act when a relationship-bearing attribute is present.
        if (!themeXml.Contains(":embed=") && !themeXml.Contains(":link="))
            return themeXml;
        try
        {
            var doc = System.Xml.Linq.XDocument.Parse(themeXml);
            if (doc.Root == null) return themeXml;
            var aNs = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/drawingml/2006/main";
            var rNs = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            var removed = false;
            // Any blipFill whose blip carries a relationship reference cannot be
            // reconstructed (image binary + rels not round-tripped). Remove the
            // whole <a:blipFill> so the parent fill list entry is replaced with a
            // schema-valid neutral fill rather than a dangling one.
            foreach (var blipFill in doc.Descendants(aNs + "blipFill").ToList())
            {
                var blip = blipFill.Element(aNs + "blip");
                var hasRel = blip != null
                    && blip.Attributes().Any(a => a.Name.Namespace == rNs);
                if (!hasRel) continue;
                // Replace the unreconstructable blipFill with a neutral
                // <a:solidFill><a:schemeClr val="phClr"/></a:solidFill> — the
                // canonical placeholder fill used throughout fmtScheme style
                // lists — so the surrounding *StyleLst keeps a valid child count
                // and Word still renders a (plain) fill.
                var placeholder = new System.Xml.Linq.XElement(aNs + "solidFill",
                    new System.Xml.Linq.XElement(aNs + "schemeClr",
                        new System.Xml.Linq.XAttribute("val", "phClr")));
                blipFill.ReplaceWith(placeholder);
                removed = true;
            }
            // Defensive: drop any other element still carrying an r:embed/r:link
            // (e.g. an a:blip that is not inside an a:blipFill).
            foreach (var el in doc.Descendants().ToList())
            {
                if (el.Attributes().Any(a => a.Name.Namespace == rNs
                    && (a.Name.LocalName == "embed" || a.Name.LocalName == "link")))
                {
                    el.Remove();
                    removed = true;
                }
            }
            if (!removed) return themeXml;
            stripped = true;
            return doc.Root.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        }
        catch
        {
            return themeXml;
        }
    }

    private static void EmitThemeRaw(WordHandler word, List<BatchItem> items,
                                     List<DocxUnsupportedWarning>? warnings = null)
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

        // BUG-R4B(BUG6): scrub dangling theme image references (the image binary
        // and theme1.xml.rels are not round-tripped) so the rebuilt theme
        // validates clean. Warn so the loss is visible.
        xml = StripDanglingThemeBlipRefs(xml, out var blipStripped);
        if (blipStripped)
            warnings?.Add(new DocxUnsupportedWarning(
                Element: "theme.blipFill",
                Path: "/theme",
                Reason: "theme image fill (a:blipFill r:embed) dropped — theme-part image round-trip is not supported; the fill was replaced with a neutral placeholder to keep the rebuilt theme valid"));

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

    private static void EmitHeadersFooters(WordHandler word, List<BatchItem> items,
                                           List<DocxUnsupportedWarning>? warnings = null)
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
                EmitHeaderFooterPart(word, child.Path, "header", hIdx, items, hi.Type, hi.SectionPath, warnings);
            }
            else if (child.Type == "footer")
            {
                // Same orphan guard as header above.
                if (!footerPathInfo.TryGetValue(child.Path, out var fi)) continue;
                fIdx++;
                EmitHeaderFooterPart(word, child.Path, "footer", fIdx, items, fi.Type, fi.SectionPath, warnings);
            }
        }
    }

    private static void EmitHeaderFooterPart(WordHandler word, string sourcePath, string kind,
                                             int targetIndex, List<BatchItem> items,
                                             string subTypeOverride = "default",
                                             string? sectionParent = null,
                                             List<DocxUnsupportedWarning>? warnings = null)
    {
        var partNode = word.Get(sourcePath);
        // BUG-DUMP9-08: tables are valid block-level OOXML inside hdr/ftr
        // (same schema as body) and Navigation surfaces them as `table`-typed
        // children, but the previous filter only kept paragraphs and silently
        // dropped tables. Iterate in source order, tracking per-type indices
        // so paragraph and table paths line up with replay output.
        // BUG-R11A(BUG3): include block-SDT children. A header/footer body can be
        // wrapped in (possibly nested) <w:sdt><w:sdtContent>; without `sdt` here
        // the walk produced zero content ops and the entire part body (PAGE/
        // NUMPAGES fields and all) was dropped on dump → batch.
        var blockChildren = (partNode.Children ?? new List<DocumentNode>())
            .Where(c => c.Type == "paragraph" || c.Type == "p"
                     || c.Type == "table" || c.Type == "tbl"
                     || c.Type == "sdt")
            .ToList();
        // partNode.Format does not expose `type`; the caller resolves the
        // role (default/first/even) from the section's headerRef.* / footerRef.*
        // map and passes it via subTypeOverride.
        var subType = subTypeOverride;

        // BUG-R6B(BUG2): a non-standard w:type (e.g. "odd", not in ST_HdrFtr
        // {even,default,first}) is pre-existing source rot. validate/get/dump
        // now all degrade gracefully on it; AddHeader/AddFooter on replay would
        // still reject "odd", so normalize the emitted op to "default" and warn
        // rather than emit a self-unreplayable script. Strict round-trip
        // fidelity isn't possible for a value the schema doesn't recognise.
        if (!string.Equals(subType, "default", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(subType, "first", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(subType, "even", StringComparison.OrdinalIgnoreCase))
        {
            warnings?.Add(new DocxUnsupportedWarning(
                Element: kind,
                Path: sourcePath,
                Reason: $"non-standard {kind}Reference w:type '{subType}' (not in {{default, first, even}}); emitted as 'default'"));
            subType = "default";
        }

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
        // BUG-R5B(BUG1): a header/footer body can host a textbox-bearing run
        // (e.g. a centered page-number textbox in the footer). TryEmitTextbox —
        // which AddTextbox supports for /header[N] and /footer[N] hosts — bails
        // out when ctx is null, and EmitParagraph was previously called with no
        // ctx here, so the textbox (and the PAGE field inside it) was silently
        // dropped. Build a part-scoped ctx so the textbox emit path fires and
        // unsupported-content warnings surface. Footnote/endnote/chart cursors
        // are part-local (header/footer rarely carry them, and the body ctx's
        // cursors must not be consumed from here).
        var hfCtx = new BodyEmitContext(
            FootnoteTexts: new List<string>(),
            EndnoteTexts: new List<string>(),
            FootnoteCursor: new NoteCursor(),
            EndnoteCursor: new NoteCursor(),
            ChartSpecs: new List<ChartSpec>(),
            ChartCursor: new NoteCursor(),
            ParaIdToTargetIdx: null,
            DeferredBookmarks: new List<BatchItem>(),
            TextboxCounters: new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase),
            TableOrdinalBox: new int[1],
            Warnings: warnings ?? new List<DocxUnsupportedWarning>());
        int pIdx = 0, tblIdx = 0;
        bool sawFirstPara = false;
        // BUG-DUMP-R2-NESTED-LEAD (header/footer site): a header/footer body
        // may begin with a table (CT_HdrFtr allows it). `add header`/`add footer`
        // auto-seeds an empty leading paragraph; when the first source child is a
        // table that seed has no source counterpart. Suppress the seed-reuse so
        // any later paragraph adds AFTER the table instead of overwriting the
        // leading seed, then drop the phantom seed below.
        //
        // BUG-R11A(BUG3): an SDT-wrapped header/footer body (the whole body is a
        // block <w:sdt>) is the same shape — its first child is neither a
        // paragraph the seed can host. Generalize the leading-seed opt-out to
        // "first child is a table OR a block-SDT".
        bool firstChildIsNonPara = blockChildren.Count > 0
            && (blockChildren[0].Type == "table" || blockChildren[0].Type == "tbl"
                || blockChildren[0].Type == "sdt");
        // The host part for raw-set: /document for body, otherwise the part path.
        var hfRawPart = partTargetPath;
        var hfRootXPath = kind == "header" ? "/w:hdr" : "/w:ftr";
        foreach (var child in blockChildren)
        {
            if (child.Type == "table" || child.Type == "tbl")
            {
                tblIdx++;
                // BUG-R11A(BUG1): pass the part-scoped hfCtx (was ctx: null) so a
                // block <w:sdt> nested in a header/footer table cell is emitted
                // via EmitTable's cell-SDT branch. hfCtx's note/chart cursors are
                // fresh and part-local, so threading it here consumes nothing the
                // body walk relies on.
                EmitTable(word, child.Path, tblIdx, items, ctx: hfCtx,
                          parentTablePath: null, containerPath: partTargetPath);
            }
            else if (child.Type == "sdt")
            {
                // BUG-R11A(BUG3): a block <w:sdt> that is a direct child of the
                // header/footer body. Reuse the cell-SDT machinery: rich block
                // content (the canonical PAGE/NUMPAGES footer, and the nested
                // <w:sdt><w:sdtContent><w:sdt> shape) round-trips verbatim via a
                // raw-set into the part root — injecting the OUTER sdt preserves
                // any nesting. Text-shaped controls go through the typed
                // `add sdt` path targeting the part. `cellHasContent: sawFirstPara`
                // chooses prepend (lands ahead of the auto-seed when this SDT is
                // the leading body content) vs append (after preceding paragraphs).
                EmitCellSdt(word, child.Path, partTargetPath, hfRootXPath, hfRawPart,
                            cellHasContent: sawFirstPara, items, hfCtx);
            }
            else
            {
                pIdx++;
                EmitParagraph(word, child.Path, partTargetPath, pIdx, items,
                              autoPresent: !sawFirstPara && !firstChildIsNonPara, hfCtx);
                sawFirstPara = true;
            }
        }
        // Remove the unconsumed auto-seeded leading paragraph (see above).
        if (firstChildIsNonPara)
        {
            items.Add(new BatchItem
            {
                Command = "remove",
                Path = $"{partTargetPath}/p[1]",
            });
        }
    }

    private static void EmitComments(WordHandler word, List<BatchItem> items,
                                     Dictionary<string, int> paraIdToTargetIdx)
    {
        var comments = word.Query("comment");
        int targetCommentIdx = 0;  // 1-based index of the comment as it will be rebuilt
        int sourceCommentIdx = 0;  // 1-based positional index in the source comments part
        foreach (var c in comments)
        {
            targetCommentIdx++;
            sourceCommentIdx++;
            var props = FilterEmittableProps(c.Format);

            // BUG-R9A(BUG1): emit the comment body STRUCTURALLY instead of
            // flattening it to a single `text` prop. A comment body may carry
            // multiple runs (each with its own rPr) and multiple paragraphs;
            // the old flatten-to-`text` path discarded all per-run formatting
            // and any paragraph beyond the first (silent data loss). Strategy:
            //   - `add comment` carries the FIRST paragraph's FIRST run text +
            //     that run's rPr (ApplyCommentFormatKeys applies them to the
            //     lone run present at creation time).
            //   - remaining runs in the first paragraph -> `add run` into
            //     /comments/comment[N]/p[1].
            //   - additional paragraphs -> `add paragraph` into
            //     /comments/comment[N], then `add run` per run.
            // Mirrors how /body content runs/paragraphs are emitted. Plain and
            // empty comments still round-trip (single run / no run).
            //
            // Enumerate the source comment's paragraphs WITH their run children.
            // Use the positional comment index (word.Query("comment") returns
            // comments in source order, so loop position == positional index)
            // and Get(path, depth:2) so each paragraph node carries populated
            // run Children — word.Query enumerates collection children at
            // depth 0 (empty Children), which would silently re-flatten.
            var bodyParas = new List<DocumentNode>();
            for (int pIdx = 1; ; pIdx++)
            {
                DocumentNode? para;
                try { para = word.Get($"/comments/comment[{sourceCommentIdx}]/p[{pIdx}]", depth: 2); }
                catch { break; }
                if (para == null) break;
                bodyParas.Add(para);
            }

            var firstParaRuns = bodyParas.Count > 0
                ? bodyParas[0].Children.Where(IsRoundTrippableCommentRun).ToList()
                : new List<DocumentNode>();

            // BUG-R6B(BUG1): always emit `text`, even when empty. An empty
            // comment (no inline text, or only an empty table) is valid OOXML;
            // omitting `text` produced a dump op that AddComment refused to
            // replay ("'text' property is required"), silently dropping the
            // comment on round-trip. AddComment now accepts text="".
            // The first run's text + rPr ride on `add comment`; if there is no
            // first run (empty comment) fall back to empty text.
            if (firstParaRuns.Count > 0)
            {
                var firstRun = firstParaRuns[0];
                props["text"] = firstRun.Text ?? string.Empty;
                MergeRunFormatProps(props, firstRun);
            }
            else
            {
                props["text"] = c.Text ?? string.Empty;
            }
            // Map anchoredTo (source paraId path) -> target paragraph index.
            // anchoredTo looks like "/body/p[@paraId=00100000]"; parse and
            // resolve via the paraId map we built during EmitBody.
            string parentTarget = "/body/p[1]";  // safe fallback to first body para
            if (props.TryGetValue("anchoredTo", out var anchor))
            {
                // BUG-R4 (DBF-R4-01): a comment anchored inside a table cell
                // resolves to "/body/tbl[N]/tr[M]/tc[K]/p[J]" — a positional
                // path that is structurally stable across dump→batch (the table
                // is re-created with fresh paraIds, so the body paraId map can't
                // help). Pass it through verbatim so the comment re-anchors in
                // the cell instead of falling back to /body/p[1].
                if (anchor.Contains("/tbl[", StringComparison.OrdinalIgnoreCase))
                {
                    parentTarget = anchor;
                }
                else
                {
                    var pid = ExtractParaId(anchor);
                    if (pid != null && paraIdToTargetIdx.TryGetValue(pid, out var idx))
                        parentTarget = $"/body/p[{idx}]";
                }
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

            // BUG-R9A(BUG1): structural emit of the remainder of the comment
            // body. The target comment is rebuilt at /comments/comment[N]
            // where N == targetCommentIdx (comments replay in source order).
            string targetCommentPath = $"/comments/comment[{targetCommentIdx}]";

            // Remaining runs of the first paragraph (run [1] already rode on
            // `add comment`). p[1] always exists after `add comment`.
            // BUG-R13A: coalesce hyperlink runs so a hyperlink in the comment
            // body round-trips as a typed `add hyperlink` (was dropped as a
            // flat `add r` with unsupported url/isHyperlink props).
            EmitContainerBodyRuns(firstParaRuns.Skip(1).ToList(),
                $"{targetCommentPath}/p[1]", items);

            // Additional paragraphs (paragraph [1] is the `add comment` body).
            for (int pi = 1; pi < bodyParas.Count; pi++)
            {
                var para = bodyParas[pi];
                var paraProps = FilterEmittableProps(para.Format);
                paraProps.Remove("text");  // text is carried by per-run emits below
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = targetCommentPath,
                    Type = "paragraph",
                    Props = paraProps.Count > 0 ? paraProps : null
                });
                var runs = para.Children.Where(IsRoundTrippableCommentRun).ToList();
                // AddParagraph with no `text` produces an empty paragraph; emit
                // each run so per-run formatting survives. The new paragraph is
                // the (pi+1)-th paragraph of the comment.
                EmitContainerBodyRuns(runs, $"{targetCommentPath}/p[{pi + 1}]", items);
            }
        }
    }

    // BUG-R9A(BUG1): a comment-body run is round-trippable as a plain `add run`
    // only when it is text-carrying (no drawing / field / break / footnote-ref
    // structure). Comment bodies in practice hold plain text runs; richer
    // structure inside a comment is rare and out of scope here — skip such runs
    // rather than mis-emit them as plain text.
    private static bool IsRoundTrippableCommentRun(DocumentNode run)
    {
        return run.Type == "run" || run.Type == "r";
    }

    // BUG-R9A(BUG1): emit one comment-body run as `add run`, carrying its text
    // and rPr (italic/bold/color/size/font/…). Mirrors EmitPlainOrHyperlinkRun
    // for /body runs, minus the hyperlink/revision special-casing (comment
    // bodies don't carry those in the supported round-trip).
    private static void EmitCommentRun(DocumentNode run, string paraTargetPath, List<BatchItem> items, int hlBaseline = 0)
    {
        // BUG-R13A: a run flattened out of a <w:hyperlink> wrapper carries
        // url/anchor/isHyperlink (and _hyperlinkParent) Format keys that
        // `add r` does not understand — emitting it as a flat `add r` silently
        // dropped the hyperlink wrapper + URL on replay (only the link text
        // survived as a plain run). Route such runs through the body walker's
        // EmitPlainOrHyperlinkRun, which emits a proper typed `add hyperlink`
        // op (rebuilding the <w:hyperlink> + rel relationship). Plain runs fall
        // through to the flat `add r` path unchanged. Multi-run hyperlinks are
        // coalesced upstream in EmitContainerBodyRuns; this single-run guard
        // covers the in-loop callers that emit one run at a time.
        if (run.Format.ContainsKey("url") || run.Format.ContainsKey("anchor")
            || run.Format.ContainsKey("isHyperlink"))
        {
            EmitPlainOrHyperlinkRun(run, paraTargetPath, items, null, hlBaseline);
            return;
        }
        var rProps = FilterEmittableProps(run.Format);
        if (!string.IsNullOrEmpty(run.Text))
            rProps["text"] = run.Text!;
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = paraTargetPath,
            Type = "r",
            Props = rProps.Count > 0 ? rProps : null
        });
    }

    // BUG-R13A: emit a sequence of container-body (comment / footnote / endnote)
    // runs, coalescing consecutive runs that share the same <w:hyperlink>
    // wrapper into one structured `add hyperlink` (+ per-run `add r`) so a
    // multi-run formatted hyperlink survives the round-trip with every run's
    // rPr intact. Reuses the body-paragraph walker's CoalesceHyperlinkRuns /
    // EmitPlainOrHyperlinkRun machinery (single source of truth for hyperlink
    // emit). Non-hyperlink runs pass through EmitCommentRun unchanged.
    private static void EmitContainerBodyRuns(List<DocumentNode> runs, string paraTargetPath, List<BatchItem> items)
    {
        // BUG-R14B: capture the hyperlink baseline ONCE for this container body
        // so multi-run hyperlinks re-index from 1 within it (mirrors the body
        // walker; per-run capture would mis-reset a 2nd hyperlink to index 1).
        int hlBaseline = items.Count(it => it.Type == "hyperlink"
            && string.Equals(it.Parent, paraTargetPath, StringComparison.Ordinal));
        foreach (var run in CoalesceHyperlinkRuns(runs))
            EmitCommentRun(run, paraTargetPath, items, hlBaseline);
    }

    // BUG-R9A(BUG1): fold a run's rPr format keys into the `add comment` prop
    // bag so ApplyCommentFormatKeys applies them to the comment's first run.
    // `text` is set separately by the caller; never copy paragraph/comment-level
    // keys here (the run node carries only run-level format).
    private static void MergeRunFormatProps(Dictionary<string, string> props, DocumentNode run)
    {
        var rProps = FilterEmittableProps(run.Format);
        foreach (var (k, v) in rProps)
        {
            if (string.Equals(k, "text", StringComparison.OrdinalIgnoreCase)) continue;
            props[k] = v;
        }
    }

    // BUG-R12A(BUG3): emit a footnote/endnote STRUCTURALLY (mirrors the R9
    // comment-body fix EmitComments above) instead of flattening the note body
    // to a single `text` prop. A note body may carry multiple runs (each with
    // its own rPr) and multiple paragraphs; the old flatten-to-`text` path
    // (BodyEmitContext.FootnoteTexts/EndnoteTexts) discarded all per-run
    // formatting and any paragraph beyond the first (silent data loss).
    //
    // Strategy (identical to EmitComments):
    //   - `add footnote`/`add endnote` (anchored on the body carrier paragraph)
    //     carries the FIRST content paragraph's FIRST content run text + that
    //     run's rPr (ApplyFootnoteEndnoteFormatKeys applies them to the lone
    //     authored run AddFootnote/AddEndnote seeds at creation time).
    //   - remaining runs in the first paragraph -> `add r` into
    //     /footnote[N]/p[1] (or /endnote[N]/p[1]).
    //   - additional paragraphs -> `add paragraph` into /footnote[N], then
    //     `add r` per run.
    //
    // <paramref name="kind"/> is "footnote" or "endnote"; <paramref
    // name="sourceNoteIdx"/> is the 1-based positional index in the source note
    // part (== document-order reference cursor + 1, since references walk in
    // order). <paramref name="targetNoteIdx"/> is the 1-based index of the note
    // as it will be rebuilt — equal to the number of `add footnote`/`add
    // endnote` ops already emitted (including this one). The note reference mark
    // run (footnoteRef/endnoteRef, empty text) is skipped: AddFootnote/AddEndnote
    // recreates it on replay.
    private static void EmitNoteReference(WordHandler word, string kind, int sourceNoteIdx,
                                          int targetNoteIdx, string carrierPath, List<BatchItem> items)
    {
        // BUG-DUMP-ENDNOTE-ID: the source-side `/{kind}[N]` path resolves by
        // note Id (== N), NOT by ordinal position among the user notes —
        // /endnote[2] means "endnote whose w:id=2", not "the 2nd endnote". The
        // 1-based document-order reference cursor (sourceNoteIdx) only equals the
        // Id when the part's user notes start at id 1 (the convention Word and
        // our own AddFootnote/AddEndnote use: separators at id -1/0, first user
        // note at id 1). LibreOffice numbers endnote separators at id 0/1, so the
        // first user endnote is id 2 and /endnote[1] resolves to the
        // continuationSeparator (empty body) — every endnote body was silently
        // dropped while the footnote path round-tripped by coincidence of id
        // convention. Translate the ordinal cursor to the real source note Id by
        // enumerating user notes (id > 0) in document order, then address the
        // source by id-qualified path. The rebuilt-side targetNotePath stays
        // positional: AddFootnote/AddEndnote always allocate ids 1..N, so on the
        // rebuilt part ordinal == id.
        int sourceNoteId = ResolveUserNoteId(word, kind, sourceNoteIdx);

        // Count the note's paragraphs from its raw XML (deterministic). A
        // depth-N note Get returns EMPTY children — it does not enumerate its
        // <w:p> grandchildren — and, inside the dump session, out-of-range
        // /<kind>[N]/p[K] does NOT reliably throw (it clamped, producing a flood
        // of empty paragraphs), so neither the children list nor a Get-until-
        // throw loop is a safe bound. The raw XML <w:p…> open-tag count is.
        string sourceNotePath = $"/{kind}[@{kind}Id={sourceNoteId}]";
        int paraCount = 0;
        var noteXml = word.GetElementXml(sourceNotePath);
        if (!string.IsNullOrEmpty(noteXml))
            paraCount = System.Text.RegularExpressions.Regex.Matches(noteXml, "<w:p[ >]").Count;

        // Enumerate exactly that many paragraphs WITH their run children.
        // Get(path, depth:2) per paragraph so each node carries populated run
        // Children (the per-paragraph Get is the only way to reach the runs).
        var bodyParas = new List<DocumentNode>();
        for (int pIdx = 1; pIdx <= paraCount; pIdx++)
        {
            DocumentNode? para;
            try { para = word.Get($"{sourceNotePath}/p[{pIdx}]", depth: 2); }
            catch { break; }
            if (para == null) break;
            bodyParas.Add(para);
        }

        var firstParaRuns = bodyParas.Count > 0
            ? bodyParas[0].Children.Where(c => IsRoundTrippableNoteRun(word, c)).ToList()
            : new List<DocumentNode>();

        // `add footnote`/`add endnote` requires a non-empty `text` (AddFootnote/
        // AddEndnote throw without it). Carry the first content run's text + rPr;
        // fall back to the concatenated note text only when no content run
        // resolves (degenerate/empty note).
        var noteProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (firstParaRuns.Count > 0)
        {
            var firstRun = firstParaRuns[0];
            // TrimStart the FIRST content run's text only: AddFootnote/AddEndnote
            // prepend a single space between the superscript reference mark and
            // the authored text, and GetFootnoteText trims it back on readback —
            // so the source first run carries that leading space. Re-feeding it
            // verbatim to AddFootnote would prepend ANOTHER space every round
            // (dump fixed-point never converges; R7B exact-text assertions fail).
            // Trimming here mirrors GetFootnoteText exactly: a genuinely
            // authored leading space (no preceding refmark space, e.g. the
            // hand-authored "Plain " case) has none to trim, so it is preserved.
            noteProps["text"] = (firstRun.Text ?? string.Empty).TrimStart();
            MergeRunFormatProps(noteProps, firstRun);
        }
        else
        {
            string fallback = "";
            try { fallback = word.Get(sourceNotePath).Text ?? ""; }
            catch { /* leave empty */ }
            noteProps["text"] = fallback;
        }

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = carrierPath,
            Type = kind,
            Props = noteProps
        });

        // Structural emit of the remainder. The target note is rebuilt at
        // /<kind>[targetNoteIdx] (notes replay in reference order; the reserved
        // separator / continuationSeparator notes, ids -1/0, are excluded by
        // Query and by the positional /<kind>[N] path index). The first authored
        // run + p[1] already exist after the `add <kind>` above.
        string targetNotePath = $"/{kind}[{targetNoteIdx}]";

        // BUG-R13A: coalesce hyperlink runs so a hyperlink inside a footnote/
        // endnote body round-trips as a typed `add hyperlink` (was dropped as a
        // flat `add r` carrying unsupported url/isHyperlink props).
        EmitContainerBodyRuns(firstParaRuns.Skip(1).ToList(),
            $"{targetNotePath}/p[1]", items);

        for (int pi = 1; pi < bodyParas.Count; pi++)
        {
            var para = bodyParas[pi];
            var paraProps = FilterEmittableProps(para.Format);
            paraProps.Remove("text"); // text carried by per-run emits below
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = targetNotePath,
                Type = "paragraph",
                Props = paraProps.Count > 0 ? paraProps : null
            });
            var runs = para.Children.Where(c => IsRoundTrippableNoteRun(word, c)).ToList();
            EmitContainerBodyRuns(runs, $"{targetNotePath}/p[{pi + 1}]", items);
        }
    }

    // BUG-DUMP-ENDNOTE-ID: map a 1-based document-order user-note ordinal to the
    // real OOXML note Id. `query footnote`/`query endnote` returns user notes
    // (id > 0, separators excluded) in document order with id-qualified paths
    // (/endnote[@endnoteId=2]); this is the same set the reference cursor counts.
    // The Nth reference therefore corresponds to the Nth entry's Id. Falls back
    // to the ordinal itself (legacy id==ordinal assumption) when the path can't
    // be parsed or the ordinal is out of range — preserves the prior behaviour
    // for the well-formed Word-convention case rather than throwing.
    private static int ResolveUserNoteId(WordHandler word, string kind, int ordinal)
    {
        var notes = word.Query(kind);
        if (ordinal >= 1 && ordinal <= notes.Count)
        {
            var m = System.Text.RegularExpressions.Regex.Match(
                notes[ordinal - 1].Path, $@"@{kind}Id=(-?\d+)");
            if (m.Success && int.TryParse(m.Groups[1].Value, out var id))
                return id;
        }
        return ordinal;
    }

    // BUG-R12A(BUG3): a note-body run is round-trippable as a plain `add r` only
    // when it is a text-carrying run that is NOT the note reference mark
    // (<w:footnoteRef/> / <w:endnoteRef/>, which renders the superscript marker
    // and is recreated by AddFootnote/AddEndnote). Richer structure (drawings,
    // fields, nested notes) inside a note body is rare and out of scope — skip
    // such runs rather than mis-emit them as plain text. Mirrors
    // IsRoundTrippableCommentRun, plus the ref-mark exclusion.
    //
    // BUG-DUMP-ENDNOTE-ID: the ref-mark exclusion must reject only a *pure*
    // ref-mark run (the <w:*Ref/> with no body text). Word emits the ref mark
    // and the note text in SEPARATE runs, but LibreOffice fuses them into a
    // single <w:r><w:*Ref/><w:t>body</w:t></w:r>. Rejecting any run that merely
    // *contains* the ref child dropped that fused run's entire body text — the
    // root of "endnote bodies silently dropped". Get's .Text already excludes
    // the ref mark (it contributes no <w:t>), and AddFootnote/AddEndnote rebuilds
    // the ref mark from scratch, so a fused run round-trips correctly as a plain
    // text run; only a text-less ref mark is dropped.
    private static bool IsRoundTrippableNoteRun(WordHandler word, DocumentNode run)
    {
        if (run.Type != "run" && run.Type != "r") return false;
        var raw = word.GetElementXml(run.Path);
        if (!string.IsNullOrEmpty(raw)
            && (raw.Contains("footnoteRef", StringComparison.Ordinal)
                || raw.Contains("endnoteRef", StringComparison.Ordinal))
            && string.IsNullOrEmpty(run.Text))
            return false; // a pure reference mark — recreated by AddFootnote/AddEndnote
        return true;
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

        EmitSdtTyped(word, sourcePath, "/body", items);
    }

    // BUG-R11A(BUG1): block <w:sdt> that is a DIRECT CHILD of a table cell.
    // Mirrors body-level EmitSdt: rich block content (multi-paragraph / field /
    // table / drawing / line-break) round-trips verbatim via raw-set appended
    // into the just-emitted cell; everything text-shaped goes through the typed
    // `add sdt` path targeting the cell. Without this, the cell child walk in
    // EmitTable enumerated only paragraphs and nested tables, so a cell-nested
    // SDT (and its content) was silently dropped on dump → round-trip data loss.
    //
    // <paramref name="cellXPath"/> is the `(//w:tbl)[N]/w:tr[r]/w:tc[c]` selector
    // that resolves to the target cell at replay time (built from the document-
    // order table ordinal so it is stable regardless of later tables / nesting).
    // <paramref name="rawPart"/> is the host part ("/document" for body tables,
    // "/header[N]" / "/footer[N]" otherwise). <paramref name="cellHasContent"/>
    // decides prepend vs append so the SDT keeps document order relative to the
    // cell's auto-seeded leading paragraph.
    private static void EmitCellSdt(WordHandler word, string sourcePath, string cellTargetPath,
                                    string cellXPath, string rawPart, bool cellHasContent,
                                    List<BatchItem> items, BodyEmitContext ctx)
    {
        var rawXml = word.RawElementXml(sourcePath);
        if (!string.IsNullOrEmpty(rawXml) && IsRichBlockSdt(rawXml!))
        {
            if (HasExternalRelRef(rawXml!))
            {
                ctx.Warnings.Add(new DocxUnsupportedWarning(
                    Element: "sdt.richContent",
                    Path: sourcePath,
                    Reason: "content control in a table cell with rich block content AND external relationship references (hyperlinks/images) flattened to text on dump"));
            }
            else
            {
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = rawPart,
                    Xpath = cellXPath,
                    Action = cellHasContent ? "append" : "prepend",
                    Xml = rawXml
                });
                return;
            }
        }

        EmitSdtTyped(word, sourcePath, cellTargetPath, items);
    }

    // Shared typed `add sdt` emit. Whitelists the Get-canonical keys AddSdt
    // consumes plus the visible text; targets <paramref name="parentPath"/>
    // (/body for body-level SDTs, a cell path for cell-nested ones). AddSdt
    // accepts both Body and TableCell parents, so the same emit serves both.
    private static void EmitSdtTyped(WordHandler word, string sourcePath, string parentPath,
                                     List<BatchItem> items)
    {
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
            Parent = parentPath,
            Type = "sdt",
            Props = props
        });
    }

    // A block SDT is "rich" when its content carries structure the text-only
    // typed emit cannot reproduce: more than one paragraph, more than one run,
    // any run-level rPr, or any hyperlink / complex field / table / drawing /
    // break / tab. Such SDTs round-trip verbatim via raw-set; everything else
    // (single plain text run, form-control pickers) stays on the introspectable
    // typed `add sdt` path.
    private static bool IsRichBlockSdt(string sdtXml)
    {
        // <w:p> / <w:p attr...> — but not <w:pPr>, <w:pict>, <w:proofErr> (the
        // char after "w:p" must be a space or '>').
        if (System.Text.RegularExpressions.Regex.Matches(sdtXml, "<w:p[ >]").Count > 1)
            return true;
        // BUG-R12A(BUG1b): a single-paragraph block SDT whose content carries
        // multiple runs or any run-level formatting (bold/color/size/font/…)
        // cannot round-trip through the flat `add sdt text=` path — AddSdt seeds
        // one unformatted run from the concatenated text, so "FIRST"+"SECOND"
        // (2nd bold/red) comes back as a single plain "FIRSTSECOND" run. The
        // run-level richness check here was previously only applied to inline
        // (run-level) SDTs (IsRichInlineSdt); body/cell/header/footer BLOCK
        // SDTs flattened. Raw-set the SDT verbatim (no rels) so per-run rPr
        // survives. Restrict the run-count probe to CONTENT runs by counting
        // <w:r> opens — sdtPr/sdtEndPr carry no <w:r>, so no false positives.
        if (System.Text.RegularExpressions.Regex.Matches(sdtXml, "<w:r[ >]").Count > 1)
            return true;
        // Any run-level rPr inside a content run (the rPr sits under <w:r>; a
        // pPr's <w:rPr> paragraph-mark formatting is matched too, which is also
        // worth preserving verbatim and the typed path can't express it).
        if (sdtXml.Contains("<w:rPr", StringComparison.Ordinal))
            return true;
        return sdtXml.Contains("<w:hyperlink", StringComparison.Ordinal)
            || sdtXml.Contains("<w:fldChar", StringComparison.Ordinal)
            || sdtXml.Contains("w:instrText", StringComparison.Ordinal)
            || sdtXml.Contains("<w:fldSimple", StringComparison.Ordinal)
            || sdtXml.Contains("<w:tbl", StringComparison.Ordinal)
            || sdtXml.Contains("<w:drawing", StringComparison.Ordinal)
            // BUG-DUMPR2: a single-paragraph SDT can still carry intra-run
            // structure the text-only typed emit can't reproduce — a line break
            // or tab. sdt.Text concatenates run text and drops <w:br/>/<w:tab/>,
            // so "a<w:br/>b" flattened to "ab". Treat their presence as rich so
            // the SDT round-trips verbatim via raw-set (no rels involved).
            || sdtXml.Contains("<w:br", StringComparison.Ordinal)
            || sdtXml.Contains("<w:tab", StringComparison.Ordinal)
            || sdtXml.Contains("<w:cr", StringComparison.Ordinal);
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
        //
        // Page-geometry absence: when the source body sectPr OMITS <w:pgSz>
        // (and/or <w:pgMar>), Get returns no pageWidth/pageHeight/marginTop/…
        // keys, so nothing above stamps them — but the rebuild target is a
        // blank doc whose sectPr already carries the template's A4 pgSz +
        // default pgMar. Left untouched, the rebuild renders A4 while real
        // Word renders the pgSz-less source as its application default (US
        // Letter) → whole-document re-wrap. Emit an explicit remove signal
        // (`pageSize=none` / `pageMargin=none`; "none" is the established
        // sectPr-child remove sentinel) so the rebuilt sectPr also defers to
        // the app default. Independent per element — a source with pgSz but
        // no pgMar (or vice versa) is handled correctly. When the source HAS
        // the element the normal pageWidth/marginTop emit above carries it,
        // and no remove signal is emitted.
        // Page geometry round-trip fix: the loop above sourced pageWidth/
        // pageHeight/marginTop/… from Get's canonical cm strings, which round
        // twips to 2 decimals (1418 twips → "2.5cm"). Replaying that through
        // ParseTwips yields 1417 — a ±1-twip drift on every dump→batch cycle.
        // Overwrite each PRESENT geometry key with its native-twip integer
        // (bare numbers parse back as exact twips), so the rebuild's pgSz/pgMar
        // match the source byte-for-byte. Only keys already in `props` are
        // overwritten — the blank-baseline skip above and the pageSize=none /
        // pageMargin=none sentinels below stay in force.
        var rawTwips = word.BodySectionPageGeometryTwips();
        foreach (var (gk, gv) in rawTwips)
        {
            if (props.ContainsKey(gk)) props[gk] = gv;
        }
        // pgBorders fold: Get emits pgBorders.<side> + pgBorders.<side>.sz/
        // .color/.space as separate keys (mirrors pbdr.* / border.*). Set's
        // pgborders.<side> case parses a single semicolon-encoded
        // STYLE;SIZE;COLOR;SPACE value, so fold the sub-keys into the bare
        // side key and drop them. pgBorders.offsetFrom passes through verbatim
        // (it's a standalone Set key). Without folding, the 3-segment sub-keys
        // hit UNSUPPORTED on replay and the per-side weight/color/space were
        // lost — the page border collapsed to the box default.
        FoldPgBordersProps(props);
        var (hasPgSz, hasPgMar) = word.BodySectionPageGeometryPresence();
        if (!hasPgSz) props["pageSize"] = "none";
        if (!hasPgMar) props["pageMargin"] = "none";
        if (props.Count == 0) return;
        items.Add(new BatchItem
        {
            Command = "set",
            Path = "/",
            Props = props
        });
    }

    // Fold pgBorders.<side>.sz/.color/.space sub-keys into the bare
    // pgBorders.<side> key as a STYLE;SIZE;COLOR;SPACE value (the form Set's
    // pgborders.<side> case parses via ParseBorderValue). Mirrors the pbdr.* /
    // border.* fold in WordBatchEmitter.Filters.cs. pgBorders.offsetFrom is a
    // standalone Set key and is left untouched.
    private static void FoldPgBordersProps(Dictionary<string, string> props)
    {
        var fold = new Dictionary<string, (string? style, string? sz, string? color, string? space)>(
            StringComparer.OrdinalIgnoreCase);
        var subKeys = new List<string>();
        foreach (var (key, val) in props)
        {
            if (!key.StartsWith("pgBorders.", StringComparison.OrdinalIgnoreCase)) continue;
            var parts = key.Split('.');
            // parts[0]=pgBorders, parts[1]=side|offsetFrom
            if (parts.Length < 2) continue;
            // offsetFrom is a flat key — not a per-side border. Leave it alone.
            if (parts.Length == 2 &&
                string.Equals(parts[1], "offsetFrom", StringComparison.OrdinalIgnoreCase))
                continue;
            var side = $"{parts[0]}.{parts[1]}"; // pgBorders.top
            fold.TryGetValue(side, out var cur);
            if (parts.Length == 2)
            {
                cur.style = val;
            }
            else if (parts.Length == 3)
            {
                switch (parts[2].ToLowerInvariant())
                {
                    case "sz": cur.sz = val; break;
                    case "color": cur.color = val; break;
                    case "space": cur.space = val; break;
                }
                subKeys.Add(key); // 3-segment sub-keys get dropped after folding
            }
            fold[side] = cur;
        }
        foreach (var sk in subKeys) props.Remove(sk);
        foreach (var (side, folded) in fold)
        {
            if (folded.style == null) continue;
            var sz = folded.sz ?? "";
            var col = folded.color ?? "";
            var sp = folded.space ?? "";
            var v = folded.style;
            if (folded.sz != null || folded.color != null || folded.space != null)
                v += ";" + sz;
            if (folded.color != null || folded.space != null)
                v += ";" + col;
            if (folded.space != null)
                v += ";" + sp;
            props[side] = v;
        }
    }

    private static void EmitStyles(WordHandler word, List<BatchItem> items)
    {
        // Use query() rather than walking Get("/styles").Children — the
        // positional /styles/style[N] children Get returns are not
        // addressable on the Get side (style paths resolve by id, not by
        // index). Query produces id-based paths and excludes docDefaults.
        var styles = word.Query("style");
        // STYLE-RAW-FALLBACK: scalar Format keys (basedOn / spaceAfter /
        // font / size / …) cannot express a TABLE style's visual formatting:
        // its style-level <w:tblPr> (borders, band sizes, cell margins), its
        // <w:tblStylePr> conditional-formatting blocks (firstRow / lastRow /
        // band1Vert / …), and table-level <w:shd>/<w:tcPr>/<w:trPr>. A table
        // that draws its borders/shading/banding from a table style (no inline
        // <w:tblBorders> on the table itself) therefore lost ALL visual
        // formatting on round-trip — the rebuilt style emitted only scalars,
        // dropping tblBorders/tblStylePr/shd, and Word rendered it as plain
        // borderless text. Give table styles a raw-set replace fallback that
        // round-trips the whole <w:style> element verbatim — exactly the
        // pattern docDefaults / theme / settings already use. The scalar `add
        // style` still runs first (creating the style + handling id collisions
        // via AddStyle's upsert/suffix path); the raw-set then swaps the
        // freshly-added <w:style> for the source's verbatim copy, so no
        // double-apply and no scalar/raw drift. Mirrors EmitDocDefaultsRaw.
        var rawStyleByMatchAttr = BuildRawTableStyleMap(word);
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
            // STYLE-RAW-FALLBACK: if this style is a table style whose verbatim
            // XML we captured, replace the just-added <w:style> wholesale so
            // its tblPr / tblStylePr / shd / trPr / tcPr survive. Keyed by the
            // id the `add` actually used (emitId) so an id collision/suffix on
            // the target still lands on the right element. The raw XML's
            // w:styleId is normalized to emitId by BuildRawTableStyleMap.
            if (!string.IsNullOrEmpty(emitId)
                && rawStyleByMatchAttr.TryGetValue(emitId, out var rawStyleXml))
            {
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = "/styles",
                    Xpath = $"/w:styles/w:style[@w:styleId='{emitId}']",
                    Action = "replace",
                    Xml = rawStyleXml
                });
            }
        }
    }

    // STYLE-RAW-FALLBACK helper: parse the source styles.xml once and return a
    // map from styleId → verbatim <w:style> XML, restricted to TABLE styles
    // (w:type="table"). Only table styles need this fallback today: their
    // <w:tblPr>/<w:tblStylePr>/<w:shd>/<w:trPr>/<w:tcPr> formatting has no
    // scalar Format representation, unlike paragraph/character styles whose
    // pPr/rPr round-trips through the scalar emit path. Keeping the scope to
    // table styles avoids re-clobbering the (correct) scalar emit for the far
    // more numerous paragraph/character styles. The keying id is each style's
    // own w:styleId — callers match it against the id the `add` step used.
    private static Dictionary<string, string> BuildRawTableStyleMap(WordHandler word)
    {
        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        string stylesXml;
        try { stylesXml = word.Raw("/styles"); }
        catch { return map; }
        if (string.IsNullOrEmpty(stylesXml) || !stylesXml.StartsWith("<")) return map;
        try
        {
            var doc = System.Xml.Linq.XDocument.Parse(stylesXml);
            var wNs = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            foreach (var styleEl in doc.Root?.Elements(wNs + "style") ?? Enumerable.Empty<System.Xml.Linq.XElement>())
            {
                var type = styleEl.Attribute(wNs + "type")?.Value;
                if (!string.Equals(type, "table", StringComparison.Ordinal)) continue;
                var idAttr = styleEl.Attribute(wNs + "styleId");
                var styleId = idAttr?.Value;
                if (string.IsNullOrEmpty(styleId)) continue;
                // Dedupe: keep the first occurrence, matching EmitStyles' own
                // first-wins styleId dedup (Word tolerates duplicate ids).
                if (map.ContainsKey(styleId)) continue;
                map[styleId] = styleEl.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
            }
        }
        catch { return new Dictionary<string, string>(StringComparer.Ordinal); }
        return map;
    }

    private sealed class NoteCursor { public int Index; }
}
