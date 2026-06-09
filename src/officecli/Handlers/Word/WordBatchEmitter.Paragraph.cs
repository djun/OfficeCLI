// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class WordBatchEmitter
{

    /// <summary>
    /// Emit a paragraph at the target index under <paramref name="parentPath"/>.
    /// When <paramref name="autoPresent"/> is true, the parent already has a
    /// pre-existing paragraph at that index (e.g. an auto-created table cell
    /// paragraph); we issue a `set` instead of a fresh `add` so the existing
    /// paragraph gets reused rather than duplicated.
    /// </summary>
    private static void EmitParagraph(WordHandler word, string sourcePath, string parentPath,
                                      int targetIndex, List<BatchItem> items, bool autoPresent,
                                      BodyEmitContext? ctx = null)
    {
        var pNode = word.Get(sourcePath);

        if (TryEmitDisplayEquation(pNode, parentPath, autoPresent, items)) return;

        // Track source paraId -> target index BEFORE any early-return path
        // (section break, TOC, …). Comments anchored on a section-break or
        // TOC paragraph would otherwise miss the mapping and fall back to
        // /body/p[1], silently retargeting the comment.
        if (ctx?.ParaIdToTargetIdx != null && parentPath == "/body" &&
            pNode.Format.TryGetValue("paraId", out var earlyParaId) && earlyParaId != null)
        {
            ctx.ParaIdToTargetIdx[earlyParaId.ToString()!] = targetIndex;
        }

        if (TryEmitInlineSectionBreak(word, pNode, parentPath, items, ctx)) return;
        if (TryEmitTocParagraph(pNode, parentPath, items)) return;
        if (TryEmitTextboxOnlyParagraph(word, pNode, parentPath, autoPresent, items, ctx)) return;

        var props = FilterEmittableProps(pNode.Format);
        // paraMarkIns.* → AddParagraph's bare trackChange.author/date form.
        // A bare trackChange.author (no `trackChange=<kind>` literal) on
        // `add p` produces both <w:pPr><w:rPr><w:ins/></w:rPr></w:pPr> and
        // wraps each auto-created content run in <w:ins>. We pass NO `text=`
        // here so step 1 only stamps the paragraph mark; subsequent `add r`
        // steps (which already carry their own trackChange=ins) wrap the
        // content. Guarded against clashing with a sibling pPrChange — if
        // both surface, paraMarkIns wins (its absence on round-trip is more
        // visually obvious in Word's revision UI).
        if (props.Remove("paraMarkIns.author", out var pmiAuthor))
        {
            props["revision.author"] = pmiAuthor;
            // Strip revision.type=format / revision.date from a sibling pPrChange
            // so the bare-attribution path on AddParagraph fires instead of the
            // pPrChange path. pPrChange round-trip on paragraphs that are
            // ALSO newly inserted is a corner case we accept losing for now
            // (rare in practice; pPrChange semantics overlap with paraMarkIns
            // on a fresh paragraph anyway).
            props.Remove("revision.type");
        }
        if (props.Remove("paraMarkIns.date", out var pmiDate))
        {
            if (!props.ContainsKey("revision.date"))
                props["revision.date"] = pmiDate;
        }
        props.Remove("paraMarkIns.id");
        // revision.beforeLost — Get-side flag set when pPrChange's
        // PreviousParagraphProperties carries a real pre-change pPr snapshot
        // that Add v1 can't reconstruct. Strip the key (Add wouldn't accept
        // it) and emit a dump warning so callers know the Reject-Change
        // restoration state will not survive replay.
        if (props.Remove("revision.beforeLost", out var _) && ctx != null)
        {
            ctx.Warnings.Add(new DocxUnsupportedWarning(
                Element: "revision.before",
                Path: pNode.Path,
                Reason: "pPrChange previous-properties snapshot dropped on dump→batch round-trip; Word's Reject-Change restoration state will be lost on replay"));
        }
        // paraMarkDel.* — surfaces the dump path through AddParagraph's
        // paraMarkDel block (added alongside this readback). Pass the keys
        // through unchanged; AddParagraph allowlists the prefix and consumes
        // them at end-of-function. Don't fold into revision.* — that
        // namespace already routes to ins/format paths and a paragraph can
        // legitimately carry BOTH paraMarkDel (¶ join) and revision.type=
        // format (pPrChange).
        // (No remove here — let the props pass through to AddParagraph.)
        // BUG-DUMP26-01: numId/numLevel that came from style inheritance
        // (ResolveNumPrFromStyle, no direct w:numPr on the paragraph) must
        // not ride on `add p` — the style already supplies them, and emitting
        // them would semantically promote inherited→explicit on replay.
        // Mirrors the first-run hoist precedent for run-character props
        // inherited from styles.
        bool numInherited = pNode.Format.TryGetValue("numInherited", out var niVal)
            && string.Equals(niVal?.ToString(), "true", StringComparison.OrdinalIgnoreCase);
        if (numInherited)
        {
            props.Remove("numId");
            props.Remove("numLevel");
            props.Remove("numFmt");
            props.Remove("listStyle");
            props.Remove("start");
        }
        // When a paragraph carries numId, the abstractNum/num pair is already
        // in /numbering (raw-set wholesale by EmitNumberingRaw). Forwarding
        // numFmt/listStyle/start to AddParagraph triggers ad-hoc
        // numbering-definition creation in WordHandler.Add — Word allocates
        // a fresh numId (1→9, 2→16, …) and the paragraph references the
        // new one, orphaning the original abstract numbering's level rPr
        // (color, bold, custom marker text). Drop those keys so the
        // paragraph just attaches by numId+numLevel to the existing def.
        if (props.ContainsKey("numId"))
        {
            props.Remove("numFmt");
            props.Remove("listStyle");
            props.Remove("start");
        }
        // BUG-R4F-02: a paragraph may carry a numId that does not resolve to any
        // <w:num> in /numbering (dangling reference). This is valid OOXML — Word
        // renders the paragraph, just without a list marker — but the Add-side
        // dangling-numId guard (WordHandler.Add.Text.cs) rejects it, and because
        // `add p` is atomic the whole paragraph (TEXT included) is lost on replay.
        // Drop the numbering props so the `add p` succeeds with its text, and
        // surface a warning so the dropped numbering is visible. The numbering is
        // kept whenever the numId IS defined (the common case). Mirrors the
        // out-of-window font-size / zero-width-column clamps in Filters.cs.
        if (props.TryGetValue("numId", out var numIdStr)
            && int.TryParse(numIdStr, out var numIdInt)
            && numIdInt > 0
            && !word.IsNumIdDefined(numIdInt))
        {
            props.Remove("numId");
            props.Remove("numLevel");
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "numId",
                Path: pNode.Path,
                Reason: $"paragraph references numId={numIdInt} which is not defined in /numbering (dangling reference); the numbering was dropped so the paragraph text survives dump→batch round-trip"));
        }
        // Collapse non-TOC field chains (fldChar(begin) + instrText(" PAGE ")
        // + fldChar(separate) + display run(s) + fldChar(end)) into a single
        // synthetic "field" entry. Without this collapse, the subsequent
        // `runs` filter sees only the cached display run and emits the field
        // value as static text — PAGE/REF/SEQ/HYPERLINK/NUMPAGES degrade to
        // their evaluated string and stop auto-updating (BUG-X2-05 / X2-1).
        var fieldEntries = CollapseFieldChains(pNode.Children ?? new List<DocumentNode>());
        // R14-bug1+2: each legacy form field embeds a BookmarkStart/End of its
        // own name. AddFormField recreates that bookmark internally — if the
        // emit pipeline also drops an `add bookmark name=X` row before the
        // `add formfield name=X`, AddFormField throws on the duplicate. Filter
        // bookmarks whose name matches a sibling formfield synth's ffName.
        var formFieldNames = fieldEntries
            .Where(e => e.Type == "formfield" && e.Format.TryGetValue("ffName", out _))
            .Select(e => e.Format["ffName"]?.ToString() ?? "")
            .Where(n => !string.IsNullOrEmpty(n))
            .ToHashSet(StringComparer.Ordinal);
        if (formFieldNames.Count > 0)
        {
            fieldEntries = fieldEntries
                .Where(e => !(e.Type == "bookmark"
                    && e.Format.TryGetValue("name", out var bn)
                    && bn != null
                    && formFieldNames.Contains(bn.ToString() ?? "")))
                .ToList();
        }
        // BUG-DUMP5-01/02: include break-typed children in the same ordered
        // list as runs so document-order is preserved on emit.
        var runs = fieldEntries
            .Where(c => c.Type == "run" || c.Type == "r" || c.Type == "picture" || c.Type == "field" || c.Type == "formfield" || c.Type == "ptab" || c.Type == "break"
                || c.Type == "equation"
                || c.Type == "tab"
                || c.Type == "bookmark"
                // R10-bug1: include ole children so TryEmitOleRun can fire
                // a warning instead of letting them be silently filtered
                // out of the run list (full round-trip is a backlog item).
                || c.Type == "ole")
            .ToList();
        var breaks = runs.Where(c => c.Type == "break").ToList();
        var bookmarks = (pNode.Children ?? new List<DocumentNode>())
            .Where(c => c.Type == "bookmark"
                && !(c.Format.TryGetValue("name", out var bn) && bn != null
                    && formFieldNames.Contains(bn.ToString() ?? "")))
            .ToList();
        var inlineSdts = (pNode.Children ?? new List<DocumentNode>())
            .Where(c => c.Type == "sdt")
            .ToList();

        bool collapseSingleRun = ShouldCollapseSingleRun(word, runs, breaks.Count, bookmarks.Count, inlineSdts.Count);
        pNode.Format.TryGetValue("tabs", out var pTabs);

        if (collapseSingleRun)
        {
            if (runs.Count == 1)
            {
                var runProps = FilterEmittableProps(runs[0].Format);
                foreach (var (k, v) in runProps)
                {
                    if (!props.ContainsKey(k)) props[k] = v;
                }
                if (!string.IsNullOrEmpty(runs[0].Text))
                    props["text"] = runs[0].Text!;
            }

            if (autoPresent)
            {
                if (props.Count > 0)
                {
                    items.Add(new BatchItem
                    {
                        Command = "set",
                        Path = $"{parentPath}/p[last()]",
                        Props = props
                    });
                }
            }
            else
            {
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = parentPath,
                    Type = "p",
                    Props = props.Count > 0 ? props : null
                });
            }
            EmitTabStops($"{parentPath}/p[last()]", pTabs, items);
            return;
        }

        // Multi-run paragraph: emit the paragraph empty first, then add each
        // run as an explicit child. See BUG-DUMP-HOIST in
        // StripRunCharacterPropsFromParagraph — for multi-run paragraphs the
        // firstRun hoist would re-apply formatting to every sibling on
        // replay, so strip run-level keys before emit.
        StripRunCharacterPropsFromParagraph(props);
        if (autoPresent)
        {
            if (props.Count > 0)
            {
                items.Add(new BatchItem
                {
                    Command = "set",
                    Path = $"{parentPath}/p[last()]",
                    Props = props
                });
            }
        }
        else
        {
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = parentPath,
                Type = "p",
                Props = props.Count > 0 ? props : null
            });
        }

        var paraTargetPath = $"{parentPath}/p[last()]";
        EmitTabStops(paraTargetPath, pTabs, items);

        // BUG-DUMP4-06: emit inline SdtRun children before the runs loop.
        foreach (var sdt in inlineSdts)
        {
            // BUG-R12A(BUG1): an inline/run-level <w:sdt> whose content carries
            // more than one run OR any run-level rPr (bold/color/size/…) cannot
            // round-trip through the flat `add sdt text=` path — AddSdt seeds a
            // single unformatted run from `text`, so "FIRSTSECOND" comes back as
            // one plain run and the bold+red is lost. Mirror the R11 block-SDT
            // fix: raw-set the <w:sdt> verbatim into the just-emitted host
            // paragraph (a run-level SDT has no inner <w:p>, so the rich-BLOCK
            // detector doesn't apply; use a run-level richness check). Restricted
            // to /body hosts + no external rels (same constraints as the inline
            // textbox raw-set) so dangling r:id/r:embed can't be produced; other
            // hosts fall back to the flat text emit.
            if (parentPath == "/body" && TryEmitRichInlineSdt(word, sdt, items, ctx))
                continue;
            var sdtProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var key in new[] { "type", "alias", "tag", "items", "format" })
            {
                if (sdt.Format.TryGetValue(key, out var v) && v != null)
                {
                    var s = v.ToString() ?? "";
                    if (s.Length > 0) sdtProps[key] = s;
                }
            }
            if (!string.IsNullOrEmpty(sdt.Text))
                sdtProps["text"] = sdt.Text!;
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "sdt",
                Props = sdtProps
            });
        }

        // BUG-DUMP6-05: collapse N runs that share a hyperlink wrapper into
        // one synthetic hyperlink-typed entry — see CoalesceHyperlinkRuns.
        runs = CoalesceHyperlinkRuns(runs);
        // BUG-D1-MULTIDRAWING-HOST: when this paragraph hosts ≥2 drawing-
        // bearing runs (side-by-side card layout), every textbox must attach
        // to the SAME host paragraph just emitted by the `add p` above so the
        // side-by-side relationship survives round-trip. A single drawing
        // either reached the wrapper-coalesce shortcut (children.Count == 1)
        // OR shares its source paragraph with sibling text/runs — in the
        // latter case attaching to the host paragraph is still correct
        // (preserves the inline relationship that the source had).
        int drawingBearingCount = runs.Count(r =>
        {
            if (r.Type == "picture") return true;
            if (r.Type != "run" && r.Type != "r") return false;
            var probe = word.GetElementXml(r.Path);
            return !string.IsNullOrEmpty(probe) && IsTextboxDrawing(probe);
        });
        string? sharedAttachPara = drawingBearingCount >= 2 ? paraTargetPath : null;
        foreach (var run in runs)
        {
            if (TryEmitBookmarkRun(run, paraTargetPath, items, ctx)) continue;
            if (TryEmitBreakRun(run, paraTargetPath, items)) continue;
            if (TryEmitTabRun(run, paraTargetPath, items)) continue;
            if (TryEmitPtabRun(run, paraTargetPath, items)) continue;
            if (TryEmitEquationRun(word, run, paraTargetPath, items)) continue;
            if (TryEmitFormFieldRun(run, paraTargetPath, items)) continue;
            if (TryEmitFieldRun(run, paraTargetPath, items, ctx)) continue;
            // OLE/embedded-object runs surface as type="ole" (see CreateOleNode
            // in WordHandler.ImageHelpers.cs). TryEmitOleRun base64-inlines the
            // embedded payload + icon and the VML frame metadata into a
            // self-contained `add ole` (picture-run style), so the object
            // round-trips with no external file; it warns only when the payload
            // can't be resolved.
            if (TryEmitOleRun(run, paraTargetPath, items, ctx, word)) continue;
            if (TryEmitPictureRun(word, run, paraTargetPath, parentPath, targetIndex, items, ctx, sharedAttachPara)) continue;
            if (TryEmitNoteRefRun(word, run, paraTargetPath, items, ctx)) continue;
            EmitPlainOrHyperlinkRun(run, paraTargetPath, items, ctx);
        }
    }

    // ── Extracted helpers (behavior unchanged from inline original) ──

    private static bool TryEmitDisplayEquation(DocumentNode pNode, string parentPath, bool autoPresent, List<BatchItem> items)
    {
        // Display-mode equations (<m:oMathPara>) surface in EmitBody's
        // bodyNode.Children as type=paragraph, but a direct Get on the
        // path returns type=equation with the LaTeX-ish formula in
        // DocumentNode.Text. EmitParagraph would otherwise emit an empty
        // `add p` and lose the entire formula. Route to typed
        // `add /body --type equation` instead.
        if (pNode.Type != "equation" || parentPath != "/body" || autoPresent) return false;
        var mode = pNode.Format.TryGetValue("mode", out var m) ? m?.ToString() : "display";
        var eqProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["mode"] = string.IsNullOrEmpty(mode) ? "display" : mode
        };
        if (!string.IsNullOrEmpty(pNode.Text))
            eqProps["formula"] = pNode.Text!;
        // BUG-DUMP19-02: forward block-equation alignment.
        if (pNode.Format.TryGetValue("align", out var eqAlign)
            && eqAlign != null && !string.IsNullOrEmpty(eqAlign.ToString()))
            eqProps["align"] = eqAlign.ToString()!;
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = "/body",
            Type = "equation",
            Props = eqProps
        });
        return true;
    }

    private static bool TryEmitInlineSectionBreak(WordHandler word, DocumentNode pNode, string parentPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // Inline section break: a paragraph carrying <w:sectPr> is the
        // OOXML representation of a mid-document section boundary.
        // AddSection on /body produces this same shape, so we emit
        // `add /body --type section` (which creates a fresh break paragraph)
        // rather than emitting a regular `add p`. The companion
        // sectionBreak.* keys map back to AddSection's prop vocabulary.
        if (parentPath != "/body" ||
            !pNode.Format.TryGetValue("sectionBreak", out var breakKind) || breakKind == null)
            return false;
        {
            var sectProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["type"] = breakKind.ToString() ?? "nextPage"
            };
            foreach (var (k, v) in pNode.Format)
            {
                if (!k.StartsWith("sectionBreak.", StringComparison.OrdinalIgnoreCase)) continue;
                if (v == null) continue;
                var keyTail = k["sectionBreak.".Length..];
                var s = v switch { bool b => b ? "true" : "false", _ => v.ToString() ?? "" };
                if (s.Length > 0) sectProps[keyTail] = s;
            }
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = "/body",
                Type = "section",
                Props = sectProps
            });
            // BUG-DUMP4-04: a section-break paragraph can also carry visible
            // text runs (the carrier paragraph is just a regular paragraph
            // with sectPr in its pPr). AddSection appends a fresh paragraph
            // at /body/p[targetIndex]; emit each text-bearing run as
            // `add r` against that paragraph.
            var carrierRuns = (pNode.Children ?? new List<DocumentNode>())
                .Where(c =>
                {
                    // BUG-DUMP7-11: include inline w:sdt carrier children.
                    if (c.Type == "sdt") return true;
                    if (c.Type != "run" && c.Type != "r") return false;
                    if (!string.IsNullOrEmpty(c.Text)) return true;
                    // BUG-DUMP5-08 / BUG-R7B(BUG1): include empty footnote /
                    // endnote reference runs (their visible text comes via the
                    // typed emit branch below, not from c.Text). Detect via the
                    // actual reference child, not the arbitrary rStyle name.
                    var rsv = c.Format.TryGetValue("rStyle", out var rsraw) ? rsraw?.ToString() : null;
                    if (ClassifyNoteRefRun(word, c, rsv) != NoteRefKind.None)
                        return true;
                    return false;
                })
                .ToList();
            if (carrierRuns.Count > 0)
            {
                var carrierPath = $"/body/p[last()]";
                foreach (var run in carrierRuns)
                {
                    // BUG-DUMP7-11: inline SDT carrier — same prop whitelist
                    // as the body-paragraph inline-SDT branch.
                    if (run.Type == "sdt")
                    {
                        var sdtCarrierProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var key in new[] { "type", "alias", "tag", "items", "format" })
                        {
                            if (run.Format.TryGetValue(key, out var v) && v != null)
                            {
                                var s = v.ToString() ?? "";
                                if (s.Length > 0) sdtCarrierProps[key] = s;
                            }
                        }
                        if (!string.IsNullOrEmpty(run.Text))
                            sdtCarrierProps["text"] = run.Text!;
                        items.Add(new BatchItem
                        {
                            Command = "add",
                            Parent = carrierPath,
                            Type = "sdt",
                            Props = sdtCarrierProps
                        });
                        continue;
                    }
                    var rStyle = run.Format.TryGetValue("rStyle", out var rs) ? rs?.ToString() : null;
                    // BUG-R7B(BUG1): detect via the actual reference child, not
                    // the (arbitrary) rStyle name.
                    var carrierNoteKind = ctx != null
                        ? ClassifyNoteRefRun(word, run, rStyle) : NoteRefKind.None;
                    if (carrierNoteKind == NoteRefKind.Footnote)
                    {
                        var noteText = ctx!.FootnoteCursor.Index < ctx.FootnoteTexts.Count
                            ? ctx.FootnoteTexts[ctx.FootnoteCursor.Index] : "";
                        ctx.FootnoteCursor.Index++;
                        items.Add(new BatchItem
                        {
                            Command = "add",
                            Parent = carrierPath,
                            Type = "footnote",
                            Props = new() { ["text"] = noteText }
                        });
                        continue;
                    }
                    if (carrierNoteKind == NoteRefKind.Endnote)
                    {
                        var noteText = ctx!.EndnoteCursor.Index < ctx.EndnoteTexts.Count
                            ? ctx.EndnoteTexts[ctx.EndnoteCursor.Index] : "";
                        ctx.EndnoteCursor.Index++;
                        items.Add(new BatchItem
                        {
                            Command = "add",
                            Parent = carrierPath,
                            Type = "endnote",
                            Props = new() { ["text"] = noteText }
                        });
                        continue;
                    }
                    var rProps = FilterEmittableProps(run.Format);
                    if (!string.IsNullOrEmpty(run.Text))
                        rProps["text"] = run.Text!;
                    items.Add(new BatchItem
                    {
                        Command = "add",
                        Parent = carrierPath,
                        Type = "r",
                        Props = rProps
                    });
                }
            }
            return true;
        }
    }

    /// <summary>
    /// BUG-DUMP-TXBX-WRAPPER: a body paragraph whose only meaningful child is
    /// a textbox-bearing Drawing run (textboxes ship inside
    /// <c>&lt;mc:AlternateContent&gt;</c>, so Get reports the run as
    /// type=&quot;run&quot; with no Format hints) used to emit BOTH an empty
    /// <c>add p</c> wrapper AND a typed <c>add textbox</c> row. On replay
    /// AddTextbox creates its own host paragraph, leaving the target with
    /// one extra empty paragraph per textbox. Detect the
    /// textbox-only-paragraph shape here and emit only the textbox row.
    /// </summary>
    private static bool TryEmitTextboxOnlyParagraph(
        WordHandler word, DocumentNode pNode, string parentPath, bool autoPresent,
        List<BatchItem> items, BodyEmitContext? ctx)
    {
        // Wrapper coalescing only makes sense at /body — header/footer/cell
        // hosts of a textbox have their own pattern and we don't want to
        // skip wrapping paragraphs that carry visible run formatting.
        if (parentPath != "/body" || autoPresent) return false;
        if (!string.IsNullOrEmpty(pNode.Text)) return false;
        var children = pNode.Children ?? new List<DocumentNode>();
        // Need exactly one drawing-bearing child (run / picture) and nothing
        // else. Bookmarks / sdts / breaks need the paragraph wrapper to
        // anchor against and must not coalesce.
        if (children.Count != 1) return false;
        var run = children[0];
        // Source-side: AlternateContent wraps the drawing so Get reports the
        // run as plain "run"/"r" with no Format hints.
        // Target-side (after AddTextbox replay): Drawing sits directly under
        // Run with no AlternateContent, so Get reports it as "picture".
        // Both shapes must collapse here — otherwise source and target dumps
        // disagree on whether to emit the `add p` wrapper and the round-trip
        // drift grows on every textbox.
        if (run.Type != "run" && run.Type != "r" && run.Type != "picture") return false;
        // Don't gate on run.Text here: picture/textbox runs surface their
        // docPr name in DocumentNode.Text (e.g. "文本框 1") which is not
        // visible body text — it doesn't disqualify the wrapper-coalesce.
        var rawXml = word.GetElementXml(run.Path);
        if (string.IsNullOrEmpty(rawXml) || !IsTextboxDrawing(rawXml)) return false;
        // Delegate to the same emit path TryEmitPictureRun uses so geometry
        // props + inner-paragraph recursion stay identical.
        return TryEmitTextbox(word, run, rawXml, parentPath, items, ctx);
    }

    private static bool TryEmitTocParagraph(DocumentNode pNode, string parentPath, List<BatchItem> items)
    {
        // TOC field-bearing paragraph: a fldChar(begin) + instrText("TOC ...")
        // + fldChar(separate) + placeholder run + fldChar(end) chain. Get
        // exposes only the placeholder text on the parent paragraph, so
        // emitting a regular `add p text=...` would drop the field structure
        // entirely and Word would no longer auto-update the TOC on open.
        if (parentPath != "/body" || pNode.Children == null) return false;
        var instrChild = pNode.Children
            .FirstOrDefault(c => c.Type == "instrText"
                && (c.Format.TryGetValue("instruction", out var iv)
                    && iv?.ToString()?.TrimStart().StartsWith("TOC", StringComparison.OrdinalIgnoreCase) == true));
        if (instrChild == null) return false;
        var instr = instrChild.Format["instruction"]!.ToString()!;
        var tocProps = ParseTocInstruction(instr);
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = "/body",
            Type = "toc",
            Props = tocProps
        });
        return true;
    }

    private static bool ShouldCollapseSingleRun(WordHandler word, List<DocumentNode> runs, int breaksCount, int bookmarksCount, int inlineSdtsCount)
    {
        // Single-run / no-run paragraph: collapse run formatting into the
        // paragraph's prop bag (the schema-reflection layer accepts run-level
        // keys on a paragraph and routes them through ApplyRunFormatting).
        if (runs.Count > 1) return false;
        if (breaksCount > 0 || bookmarksCount > 0 || inlineSdtsCount > 0) return false;
        if (runs.Count == 0) return true;
        var r = runs[0];
        // Picture / ptab runs need their own typed `add` rows.
        if (r.Type == "picture" || r.Type == "ptab") return false;
        // OLE / embedded-object runs must stay on the explicit-run path so
        // TryEmitOleRun fires. Collapsing folds the ole metadata (progId,
        // fileSize, drawAspect, …) into `add p` props that AddParagraph
        // silently ignores — the embedded object vanishes with no warning.
        // The explicit path surfaces the documented "ole run dropped" warning
        // (full binary round-trip is a separate, deferred feature) and keeps
        // the host paragraph intact.
        if (r.Type == "ole") return false;
        // A single run carrying a drawing/shape payload (textbox, connector
        // line, autoshape — surfaced as a plain "run" when AlternateContent-
        // wrapped) must NOT collapse into `add p`: collapsing flattens the
        // paragraph and silently drops the shape. Keep it on the explicit-run
        // path so TryEmitPictureRun preserves it (typed textbox or raw-set).
        if ((r.Type == "run" || r.Type == "r"))
        {
            var probe = word.GetElementXml(r.Path);
            if (!string.IsNullOrEmpty(probe) && IsDrawingBearingRun(probe)) return false;
        }
        // R14-bug1+2: legacy form field synth — needs its own typed
        // `add formfield` row; collapsing into `add p` flattens the
        // ffData wrapper into paragraph props that AddParagraph ignores.
        if (r.Type == "formfield") return false;
        // BUG-DUMP6-05 / BUG-DUMP10-05: hyperlink-wrapped run (url, anchor,
        // or tooltip-only via isHyperlink sentinel) must re-emit as
        // `add hyperlink` — `add p` does not consume url/anchor.
        if (r.Format.ContainsKey("url") || r.Format.ContainsKey("anchor")
            || r.Format.ContainsKey("isHyperlink")) return false;
        // BUG-FUZZ-2 / BUG-R7B(BUG1): footnote/endnote reference runs need the
        // typed `add footnote/endnote` branch; AddParagraph doesn't consume the
        // reference. Detect via the actual reference child, not the arbitrary
        // rStyle name (real docs use ids like "a5", not "FootnoteReference").
        {
            var srStyle = r.Format.TryGetValue("rStyle", out var srraw) ? srraw?.ToString() : null;
            if (ClassifyNoteRefRun(word, r, srStyle) != NoteRefKind.None)
                return false;
        }
        // BUG-W14-EFFECTS / BUG-DUMP5-09 / 7-01 / 5-10: run-level w14 effects /
        // OpenType properties / sym / trackChange — AddParagraph's
        // ApplyRunFormatting fallback has no cases for these; they'd
        // surface as UNSUPPORTED on replay.
        if (r.Format.ContainsKey("w14shadow")
            || r.Format.ContainsKey("textOutline")
            || r.Format.ContainsKey("textFill")
            || r.Format.ContainsKey("w14glow")
            || r.Format.ContainsKey("w14reflection")
            || r.Format.ContainsKey("ligatures")
            || r.Format.ContainsKey("numForm")
            || r.Format.ContainsKey("numSpacing")
            || r.Format.ContainsKey("revision.type")
            || r.Format.ContainsKey("sym")) return false;
        // BUG-FIELD-COLLAPSE: a synthetic field run carries `instruction=…` —
        // collapse would lose the field chain on replay.
        if (r.Type == "field") return false;
        // BUG-DUMP7-03: inline equation must emit `add equation` explicitly.
        if (r.Type == "equation") return false;
        return true;
    }

    private static bool TryEmitBookmarkRun(DocumentNode run, string paraTargetPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // BUG-DUMP25-01: bookmark child emitted in DOM order so a
        // BookmarkStart between runs survives round-trip at its original
        // intra-paragraph offset. Deferred bookmarks (endPara=true) are
        // pushed onto ctx.DeferredBookmarks so the End sibling can land in
        // a downstream paragraph.
        if (run.Type != "bookmark") return false;
        var bmProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (run.Format.TryGetValue("name", out var bmName) && bmName != null)
        {
            var s = bmName.ToString();
            if (!string.IsNullOrEmpty(s)) bmProps["name"] = s;
        }
        if (bmProps.Count == 0) return true; // skip unnamed/anonymous bookmarks
        bool deferred = false;
        if (run.Format.TryGetValue("endPara", out var bmEnd) && bmEnd != null)
        {
            var s = bmEnd.ToString();
            if (!string.IsNullOrEmpty(s) && s != "0")
            {
                bmProps["endPara"] = s;
                deferred = true;
            }
        }
        var bmItem = new BatchItem
        {
            Command = "add",
            Parent = paraTargetPath,
            Type = "bookmark",
            Props = bmProps
        };
        if (deferred && ctx != null)
            ctx.DeferredBookmarks.Add(bmItem);
        else
            items.Add(bmItem);
        return true;
    }

    private static bool TryEmitBreakRun(DocumentNode run, string paraTargetPath, List<BatchItem> items)
    {
        // BUG-DUMP5-01/02: a soft <w:br/> with NO type attribute is a line
        // break, not a page break — fall back to type=line. Emitted inline
        // from the unified runs loop so each break stays at its source
        // position instead of being hoisted to the front of the paragraph.
        if (run.Type != "break") return false;
        var breakType = run.Format.TryGetValue("breakType", out var bt) ? bt?.ToString() : null;
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = paraTargetPath,
            Type = "pagebreak",
            Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["type"] = string.IsNullOrEmpty(breakType) ? "line" : breakType!
            }
        });
        return true;
    }

    private static bool TryEmitTabRun(DocumentNode run, string paraTargetPath, List<BatchItem> items)
    {
        // BUG-DUMP14-02: tab-only run (<w:r><w:tab/></w:r>) surfaces as
        // type="tab" with empty Text. AddText splits "\t" into TabChar, so
        // emit `add r text="\t"` to round-trip the tab character.
        if (run.Type != "tab") return false;
        var tabParent = ResolveHyperlinkParent(run, paraTargetPath, items);
        var tabProps = FilterEmittableProps(run.Format);
        tabProps["text"] = "\t";
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = tabParent,
            Type = "r",
            Props = tabProps
        });
        return true;
    }

    private static bool TryEmitPtabRun(DocumentNode run, string paraTargetPath, List<BatchItem> items)
    {
        // BUG-PTAB: ptab (positional tab) — Navigation surfaces its own run
        // type with align/relativeTo/leader on Format. Without an explicit
        // emit branch the runs filter would drop it and round-trip would
        // silently lose right-align/header-style tabs.
        if (run.Type != "ptab") return false;
        // BUG-DUMP-TABRPR: carry the ptab run's own rPr (font / size / szCs /
        // …) alongside its align/relativeTo/leader. Like a tab, a positional
        // tab paints a leader in the run's font and contributes to line
        // height, so its typography is meaningful — RunToNode keeps it on
        // run.Format and AddPtab applies it on replay.
        var ptabProps = FilterEmittableProps(run.Format);
        if (run.Format.TryGetValue("align", out var pAlign) && pAlign != null)
            ptabProps["alignment"] = pAlign.ToString() ?? "";
        if (run.Format.TryGetValue("relativeTo", out var pRel) && pRel != null)
            ptabProps["relativeTo"] = pRel.ToString() ?? "";
        if (run.Format.TryGetValue("leader", out var pLead) && pLead != null)
            ptabProps["leader"] = pLead.ToString() ?? "";
        var ptabParent = ResolveHyperlinkParent(run, paraTargetPath, items);
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = ptabParent,
            Type = "ptab",
            Props = ptabProps.Count > 0 ? ptabProps : null
        });
        return true;
    }

    // Build `add hyperlink` props from a source hyperlink node so a hyperlink
    // that carries no emittable runs (e.g. one wrapping only an <m:oMath>) can
    // still be materialized. Returns null when the hyperlink has no addressable
    // destination (neither url nor anchor) — the caller then routes the child
    // under the paragraph so its content still survives.
    private static Dictionary<string, string>? TryBuildHyperlinkAddProps(WordHandler word, string srcHlPath)
    {
        DocumentNode hl;
        try { hl = word.Get(srcHlPath); }
        catch { return null; }
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var key in new[] { "url", "anchor", "tooltip", "tgtFrame", "history" })
        {
            if (hl.Format.TryGetValue(key, out var v) && v != null)
            {
                var s = v.ToString();
                if (!string.IsNullOrEmpty(s)) props[key] = s!;
            }
        }
        return (props.ContainsKey("url") || props.ContainsKey("anchor")) ? props : null;
    }

    private static bool TryEmitEquationRun(WordHandler word, DocumentNode run, string paraTargetPath, List<BatchItem> items)
    {
        // BUG-DUMP7-03: inline <m:oMath> as paragraph child. Get surfaces it
        // as type="equation" with mode=inline and the LaTeX-ish formula in
        // Text. AddEquation accepts a paragraph parent for inline mode.
        // BUG-DUMP15-04: m:oMath inside w:hyperlink surfaces with a
        // hyperlink-scoped path (.../p[N]/hyperlink[K]/equation[M]). Strip
        // the trailing /equation[M] segment so the emitted Parent places the
        // equation INSIDE the hyperlink on replay.
        if (run.Type != "equation") return false;
        var eqMode = run.Format.TryGetValue("mode", out var emv) ? emv?.ToString() : "inline";
        var eqProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["mode"] = string.IsNullOrEmpty(eqMode) ? "inline" : eqMode!
        };
        // Always emit `formula` (even when empty); ToLatex may legitimately
        // return "" for minimal m:oMath.
        eqProps["formula"] = run.Text ?? "";
        var eqParent = paraTargetPath;
        if (!string.IsNullOrEmpty(run.Path))
        {
            var idxEq = run.Path.LastIndexOf("/equation[", StringComparison.Ordinal);
            if (idxEq > 0)
            {
                var derived = run.Path.Substring(0, idxEq);
                var hlIdx = derived.LastIndexOf("/hyperlink[", StringComparison.Ordinal);
                if (hlIdx > 0)
                {
                    // BUG-DBF-R3-02: the equation lives inside a <w:hyperlink>.
                    // A hyperlink is normally materialized via its child runs, but
                    // a hyperlink whose ONLY content is an <m:oMath> has no runs —
                    // so no `add hyperlink` row is emitted and replaying
                    // `add equation parent=…/hyperlink[K]` fails ("path not
                    // found"), dropping the math. Emit the missing `add hyperlink`
                    // (fetched from the source) before the equation so its parent
                    // exists; if the hyperlink can't be resolved, fall back to the
                    // paragraph so the math survives (the rare link wrapper is lost
                    // rather than the content).
                    var srcHlPath = run.Path.Substring(0, idxEq); // …/hyperlink[K]
                    var hlSeg = derived.Substring(hlIdx); // /hyperlink[K]
                    int alreadyEmitted = items.Count(it => it.Type == "hyperlink"
                        && string.Equals(it.Parent, paraTargetPath, StringComparison.Ordinal));
                    var rebasedHl = paraTargetPath + hlSeg;
                    int wantK = 0;
                    var kStr = hlSeg.Length > 11 ? hlSeg[11..^1] : "";
                    int.TryParse(kStr, out wantK);
                    if (alreadyEmitted < wantK)
                    {
                        var hlProps = TryBuildHyperlinkAddProps(word, srcHlPath);
                        if (hlProps != null)
                        {
                            items.Add(new BatchItem
                            {
                                Command = "add",
                                Parent = paraTargetPath,
                                Type = "hyperlink",
                                Props = hlProps
                            });
                            eqParent = rebasedHl;
                        }
                        // else: leave eqParent = paraTargetPath (math survives).
                    }
                    else
                    {
                        eqParent = rebasedHl;
                    }
                }
            }
        }
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = eqParent,
            Type = "equation",
            Props = eqProps
        });
        return true;
    }

    // R14-bug1+2: legacy form field synth from CollapseFieldChains. The
    // begin run's <w:ffData> was unpacked onto ff*-prefixed Format keys
    // (ffName, ffType, ffDefault, ffMaxLength, ffChecked, ffItems,
    // ffHelpText, ffStatusText, ffEntryMacro, ffExitMacro, ffCalcOnExit,
    // ffTextType, ffTextFormat, ffCheckBoxSize, ffEnabled). Map back to
    // AddFormField's accepted keys (drop the `ff` prefix, lowercase the
    // first letter) so replay rebuilds the FieldChar + FormFieldData +
    // Bookmark chain with every original wrapper intact.
    private static bool TryEmitFormFieldRun(DocumentNode run, string paraTargetPath, List<BatchItem> items)
    {
        if (run.Type != "formfield") return false;
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        // CONSISTENCY(formfield-keys): AddFormField accepts type/formfieldtype,
        // name, default, maxLength, checked, items, helpText, statusText,
        // entryMacro, exitMacro, calcOnExit, textType, textFormat,
        // checkBoxSize, enabled. Mirror the AddFormField accepted vocabulary.
        foreach (var (k, v) in run.Format)
        {
            if (v == null) continue;
            if (!k.StartsWith("ff", StringComparison.OrdinalIgnoreCase)) continue;
            // Strip the "ff" prefix and lowercase the first remaining char so
            // ffName→name, ffDefault→default, etc.
            if (k.Length <= 2) continue;
            var bare = char.ToLowerInvariant(k[2]) + k.Substring(3);
            var sv = v.ToString();
            if (string.IsNullOrEmpty(sv)) continue;
            props[bare] = sv!;
        }
        // Preserve cached display text where AddFormField would otherwise
        // emit the placeholder symbol (text input) or "false" (checkbox).
        if (!string.IsNullOrEmpty(run.Text) && !props.ContainsKey("text"))
            props["text"] = run.Text!;
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = paraTargetPath,
            Type = "formfield",
            Props = props,
        });
        return true;
    }

    private static bool TryEmitFieldRun(DocumentNode run, string paraTargetPath, List<BatchItem> items, BodyEmitContext? ctx = null)
    {
        // Synthetic field entry from CollapseFieldChains. Format carries
        // `instruction` (raw fldSimple/instrText) and Text holds the cached
        // display value. AddField parses the instruction and rebuilds the
        // fldChar chain on replay.
        // BUG-DUMP18-02: w:fldSimple / fldChar-chain field inside w:hyperlink
        // should replay INSIDE the hyperlink — but only when a prior
        // `add hyperlink` row actually landed at the target paragraph
        // (BUG-DUMP9-03 fldSimple-only hyperlinks never surface a hyperlink
        // row, and routing the field there would fail path lookup on replay).
        if (run.Type != "field") return false;
        // R10-bug7: CollapseFieldChains flagged a nested field (IF/REF with
        // an inner DATE/PAGE/MERGEFIELD branch). AddField rebuilds a flat
        // begin/instr/separate/display/end chain and cannot model the
        // nested branches — emitting an `add field` row here would either
        // throw (parser sees garbage), drop the inner branches, OR merge
        // the inner instruction into the outer expression. Backlog item:
        // teach AddField to accept a tree representation. Until then, the
        // cheapest correct behavior is to flag the loss in envelope
        // warnings — same model as the OLE warning above — so callers
        // don't ship a doc with the IF false-branch silently stripped.
        // R10-bug8: malformed field (begin without matching end) surfaces as
        // a synth from CollapseFieldChains with _unmatchedFieldBegin=true.
        // Same warning model as _nestedField — preserve cached display,
        // flag the partial instruction in envelope.warnings.
        if (run.Format.TryGetValue("_unmatchedFieldBegin", out var ufbObj) && ufbObj is bool ufbB && ufbB)
        {
            if (ctx != null)
            {
                var partialInstr = run.Format.TryGetValue("instruction", out var pIv)
                    ? pIv?.ToString() ?? "" : "";
                ctx.Warnings.Add(new DocxUnsupportedWarning(
                    Element: "field.unmatched_begin",
                    Path: run.Path,
                    Reason: $"fldChar(begin) without matching end; partial instruction='{partialInstr}' dropped"));
            }
            if (!string.IsNullOrEmpty(run.Text))
            {
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = paraTargetPath,
                    Type = "r",
                    Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["text"] = run.Text!
                    }
                });
            }
            return true;
        }
        if (run.Format.TryGetValue("_nestedField", out var nfObj) && nfObj is bool nfB && nfB)
        {
            if (ctx != null)
            {
                ctx.Warnings.Add(new DocxUnsupportedWarning(
                    Element: "field.nested",
                    Path: run.Path,
                    Reason: "nested field (begin inside a field's branch) cannot round-trip through add field; cached display preserved but inner field codes dropped"));
            }
            // Still emit the cached display so the paragraph isn't empty.
            if (!string.IsNullOrEmpty(run.Text))
            {
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = paraTargetPath,
                    Type = "r",
                    Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["text"] = run.Text!
                    }
                });
            }
            return true;
        }
        var instr = run.Format.TryGetValue("instruction", out var iv)
            ? iv?.ToString() ?? "" : "";
        var fieldProps = BuildFieldAddProps(instr, run.Text ?? "");
        if (fieldProps != null
            && run.Format.TryGetValue("_noFieldSeparator", out var nfs)
            && nfs is bool nfsB && nfsB)
        {
            fieldProps["noSeparator"] = "true";
        }
        // BUG-R12A(BUG1): forward the captured cached-result-run formatting
        // (stashed under `_resultFmt.` by CollapseFieldChains) onto the `add
        // field` prop bag. AddField applies font/size/bold/color uniformly to
        // every rebuilt field run, so a bold/red/20pt PAGE field round-trips
        // styled instead of as a plain "1". Keys AddField can't express
        // (italic/underline/strike) are surfaced as a warning — the field still
        // round-trips (instruction + cached value preserved), only the extra
        // emphasis is lost; raw-set field passthrough is a backlog item.
        if (fieldProps != null)
        {
            var unsupportedFmt = new List<string>();
            foreach (var (k, v) in run.Format)
            {
                if (v == null) continue;
                if (!k.StartsWith("_resultFmt.", StringComparison.OrdinalIgnoreCase)) continue;
                var bare = k["_resultFmt.".Length..];
                if (!FieldAddSupportedFormatKeys.Contains(bare))
                {
                    unsupportedFmt.Add(bare);
                    continue;
                }
                // Map font.latin/font.ascii/font.hAnsi onto AddField's `font`.
                var target = bare.StartsWith("font.", StringComparison.OrdinalIgnoreCase) ? "font" : bare;
                if (!fieldProps.ContainsKey(target))
                {
                    var s = v switch { bool b => b ? "true" : "false", _ => v.ToString() ?? "" };
                    if (s.Length > 0) fieldProps[target] = s;
                }
            }
            if (unsupportedFmt.Count > 0 && ctx != null)
            {
                ctx.Warnings.Add(new DocxUnsupportedWarning(
                    Element: "field.resultFormat",
                    Path: run.Path,
                    Reason: $"cached field-result run formatting ({string.Join(", ", unsupportedFmt)}) cannot be expressed via add field; field instruction + bold/color/size/font preserved, extra emphasis dropped"));
            }
        }
        var fldParent = paraTargetPath;
        string? candidateHlParent = null;
        if (!string.IsNullOrEmpty(run.Path))
        {
            var idxFld = run.Path.LastIndexOf("/field[", StringComparison.Ordinal);
            if (idxFld > 0)
            {
                var derived = run.Path.Substring(0, idxFld);
                if (derived.Contains("/hyperlink["))
                    candidateHlParent = derived;
            }
        }
        // fldChar-chain fields surface with a flat /…/r[N] path; the
        // hyperlink hint is in Format._hyperlinkParent.
        if (candidateHlParent == null
            && run.Format.TryGetValue("_hyperlinkParent", out var fhlpObj)
            && fhlpObj != null)
        {
            var hint = fhlpObj.ToString();
            if (!string.IsNullOrEmpty(hint)) candidateHlParent = hint;
        }
        if (candidateHlParent != null)
        {
            // Re-base the candidate path onto paraTargetPath and verify a
            // prior `add hyperlink` row landed under that same paragraph.
            const string hlMarker = "/hyperlink[";
            var hlIdxStart = candidateHlParent.LastIndexOf(hlMarker, StringComparison.Ordinal);
            if (hlIdxStart > 0)
            {
                var hlEnd = candidateHlParent.IndexOf(']', hlIdxStart);
                if (hlEnd > hlIdxStart)
                {
                    var kStr = candidateHlParent.Substring(hlIdxStart + hlMarker.Length,
                        hlEnd - hlIdxStart - hlMarker.Length);
                    if (int.TryParse(kStr, out var kIdx))
                    {
                        var rebased = paraTargetPath
                            + candidateHlParent.Substring(hlIdxStart);
                        int emittedHls = items.Count(it => it.Type == "hyperlink"
                            && string.Equals(it.Parent, paraTargetPath, StringComparison.Ordinal));
                        if (emittedHls >= kIdx)
                            fldParent = rebased;
                    }
                }
            }
        }
        if (fieldProps != null)
        {
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = fldParent,
                Type = "field",
                Props = fieldProps
            });
        }
        else if (!string.IsNullOrEmpty(run.Text))
        {
            // Unparseable instruction — fall back to plain text so the
            // paragraph still renders the cached value rather than going empty.
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = fldParent,
                Type = "r",
                Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["text"] = run.Text! }
            });
        }
        return true;
    }

    private static bool TryEmitPictureRun(WordHandler word, DocumentNode run, string paraTargetPath, string parentPath, int targetIndex, List<BatchItem> items, BodyEmitContext? ctx, string? sharedAttachPara = null)
    {
        // Drawing-bearing runs surface as type="picture" regardless of
        // whether the Drawing wraps an image (Blip) or a chart (c:chart).
        // Try the image path first; if no embedded image part the run is a
        // chart anchor — pull the next pre-resolved ChartSpec and emit a
        // typed `add chart` row.
        // Drawings wrapped in <mc:AlternateContent>/<mc:Choice> surface as a
        // plain "run" node (Run.GetFirstChild<Drawing>() returns null because
        // the Drawing lives inside the AlternateContent wrapper), so we also
        // accept "run" / "r" when the raw XML carries an obvious textbox
        // marker. Non-drawing runs without those markers short-circuit out
        // of the textbox/picture path immediately.
        if (run.Type != "picture")
        {
            if (run.Type != "run" && run.Type != "r") return false;
            var probeXml = word.GetElementXml(run.Path);
            if (string.IsNullOrEmpty(probeXml)) return false;
            bool isTextbox = IsTextboxDrawing(probeXml);
            // A genuine text/hyperlink run (no drawing payload at all) belongs
            // to the plain-run path — short-circuit out so EmitPlainOrHyperlink
            // run handles it. Only drawing-bearing runs continue here.
            if (!isTextbox && !IsDrawingBearingRun(probeXml)) return false;
            // BUG-R3 (dump emits `add textbox` into a table cell): see the
            // sibling site below — a cell-hosted textbox must attach to its
            // containing paragraph (paraTargetPath inside the cell), never as a
            // direct <w:tc> child. LibreOffice exports textboxes wrapped in
            // <mc:AlternateContent>, so they reach this "run"-typed branch.
            string? attachParaAc = sharedAttachPara
                ?? (parentPath.Contains("/tc[", StringComparison.Ordinal) ? paraTargetPath : null);
            if (isTextbox && TryEmitTextbox(word, run, probeXml, parentPath, items, ctx, attachParaAc))
                return true;
            // AlternateContent-wrapped non-textbox shapes — connector lines
            // (<wps:cNvCnPr>, e.g. a letterhead separator), autoshapes, groups
            // — and textboxes whose typed emit failed fall back to a raw-set
            // append so the shape survives round-trip instead of vanishing.
            if (parentPath == "/body" && !probeXml.Contains("r:embed") && !probeXml.Contains("r:id")
                && !DrawingHasUnreconstructableRel(probeXml))
            {
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = "/document",
                    // BUG-R1-01: attach to the most-recently-emitted paragraph
                    // via last(), NOT the navigation pIndex (targetIndex). pIndex
                    // deliberately does not advance for an oMathPara-bearing
                    // paragraph (EmitBody: display equations resolve as
                    // /body/oMathPara[N], not /body/p[N]), but in the LITERAL XML
                    // the equation is a real <w:p> wrapping <m:oMathPara>. So a
                    // numeric w:p[{pIndex}] under-counts by one per preceding
                    // equation and the shape lands on the wrong paragraph. The
                    // host paragraph was just added by this same EmitParagraph
                    // call, so it is always the last w:p at replay time — the same
                    // last()-relative attach the typed picture/textbox path uses.
                    Xpath = "/w:document/w:body/w:p[last()]",
                    Action = "append",
                    Xml = probeXml
                });
                return true;
            }
            // Drawing-bearing but not safely raw-set-able inline (lives in a
            // header/footer/cell, or carries an external relationship we can't
            // re-anchor). Flag the loss rather than silently dropping it.
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                "drawing", run.Path,
                "non-textbox drawing/shape could not be serialized for round-trip; it will be missing from the replayed document"));
            return true;
        }
        var binary = word.GetImageBinary(run.Path);
        if (binary.HasValue)
        {
            var (bytes, contentType) = binary.Value;
            var dataUri = $"data:{contentType};base64,{Convert.ToBase64String(bytes)}";
            var picProps = FilterEmittableProps(run.Format);
            picProps.Remove("id");
            picProps.Remove("contentType");
            picProps.Remove("fileSize");
            picProps["src"] = dataUri;
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "picture",
                Props = picProps
            });
            return true;
        }

        // Only consume a ChartSpec if the run is genuinely a chart. Picture-
        // typed runs that aren't images can also be background images, OLE
        // objects, SmartArt, watermark anchors etc — falling through
        // unconditionally would misalign chart positions.
        if (ctx != null && word.IsChartRun(run.Path)
            && ctx.ChartCursor.Index < ctx.ChartSpecs.Count)
        {
            var spec = ctx.ChartSpecs[ctx.ChartCursor.Index];
            ctx.ChartCursor.Index++;
            var chartProps = BuildChartProps(spec);
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "chart",
                Props = chartProps
            });
            return true;
        }
        // Drawing without image part and not a chart — most likely a wps
        // shape. BUG-DUMP-TXBX: textbox-bearing drawings get a typed
        // `add textbox` row plus recursive inner-paragraph/run emits so
        // round-trip preserves structure (raw-set fallback was emitting
        // BOTH the full <w:drawing> XML AND flattening the textbox's
        // inner runs back onto the host paragraph). Non-textbox shapes
        // still fall through to the raw-set append.
        var rawXml = word.GetElementXml(run.Path);
        if (!string.IsNullOrEmpty(rawXml) && IsTextboxDrawing(rawXml))
        {
            // BUG-R3 (dump emits `add textbox` into a table cell): a textbox is
            // a drawing carried by a run inside a paragraph, never a direct
            // <w:tc> child — `add textbox` with a bare cell parent is rejected
            // ("table cells only accept paragraphs, tables, or SDTs"). When the
            // host is a cell, attach the textbox to the run's containing
            // paragraph (paraTargetPath, which lives inside the cell) so
            // AddTextbox's paragraph-attach path (ResolveDrawingHostFromParagraph)
            // anchors it correctly — mirrors the body single-drawing path. Use
            // any already-computed sharedAttachPara (side-by-side multi-drawing
            // host) first; otherwise fall back to paraTargetPath for cells.
            string? attachPara = sharedAttachPara
                ?? (parentPath.Contains("/tc[", StringComparison.Ordinal) ? paraTargetPath : null);
            if (TryEmitTextbox(word, run, rawXml, parentPath, items, ctx, attachPara))
                return true;
        }
        // BUG-R3 (linked external image): a <w:drawing> carrying an
        // unreconstructable relationship (linked-image r:link, SmartArt
        // r:dm/r:lo/r:qs/r:cs) must NOT be raw-set verbatim — its relationship
        // target isn't recreated, so the replayed drawing would dangle and fail
        // [Semantic] validation. Drop it cleanly and surface the loss (mirrors
        // the SmartArt/non-textbox warn-and-skip on the AlternateContent path
        // above) instead of falling through to a silent raw-set or a silent
        // return.
        if (!string.IsNullOrEmpty(rawXml) && DrawingHasUnreconstructableRel(rawXml))
        {
            string detail = rawXml.Contains("r:link")
                ? "linked external image (<a:blip r:link>) references an external relationship that is not carried through dump→batch; the image will be missing from the replayed document"
                : "drawing references a relationship (SmartArt/diagram) that cannot be reconstructed; it will be missing from the replayed document";
            ctx?.Warnings.Add(new DocxUnsupportedWarning("drawing", run.Path, detail));
            return true;
        }
        if (!string.IsNullOrEmpty(rawXml) &&
            parentPath == "/body" &&
            !rawXml.Contains("r:embed") && !rawXml.Contains("r:id"))
        {
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = "/document",
                // BUG-R1-01 (second site, plain <w:drawing> wps shape): same
                // last()-relative attach as above — pIndex under-counts literal
                // w:p when a display equation precedes this shape's paragraph.
                Xpath = "/w:document/w:body/w:p[last()]",
                Action = "append",
                Xml = rawXml
            });
        }
        return true;
    }

    // SmartArt (diagram) drawings reference their data / layout / colors /
    // quickStyle parts via r:dm / r:lo / r:qs / r:cs — relationships, but NOT
    // r:embed/r:id. The dump never reconstructs those parts, so raw-setting the
    // drawing verbatim leaves dangling relationships: the SDK validator NREs in
    // RelationshipTypeConstraint and real Word refuses to open the file ("may be
    // corrupt"). Treat such drawings as unreconstructable so the caller flags a
    // loss warning instead of emitting a corrupt file. (Plain shapes/connectors
    // carry no relationships and still round-trip via raw-set.)
    private static bool DrawingHasUnreconstructableRel(string xml) =>
        xml.Contains("r:dm") || xml.Contains("r:lo") ||
        xml.Contains("r:qs") || xml.Contains("r:cs") ||
        // BUG-R3 (linked external image): <a:blip r:link="rIdN"> references an
        // EXTERNAL image relationship (TargetMode="External"), not an embedded
        // image part. The dump never recreates that relationship, so raw-setting
        // the drawing verbatim leaves a dangling r:link → [Semantic] "the
        // relationship referenced by r:link does not exist". Full linked-image
        // round-trip (carrying the external rel) is a separate feature; treat
        // the drawing as unreconstructable so the raw-set sites skip it cleanly
        // and the caller surfaces a loss warning (mirrors the SmartArt path).
        xml.Contains("r:link");

    private static bool IsTextboxDrawing(string rawXml)
    {
        // Mirrors WordHandler.CountTextboxesInHost / Navigation's textbox
        // selector — a textbox is a wps:wsp with txBox=1 cNvSpPr or a
        // wps:txbx child carrying w:txbxContent.
        return rawXml.Contains("txBox=\"1\"")
            || rawXml.Contains("<wps:txbx")
            || rawXml.Contains("txbxContent");
    }

    /// <summary>
    /// True when a run carries any drawing/shape payload (image, chart,
    /// textbox, connector line, autoshape, group) — i.e. a
    /// <c>&lt;w:drawing&gt;</c>, legacy VML <c>&lt;w:pict&gt;</c>, or an
    /// <c>&lt;mc:AlternateContent&gt;</c>/<c>&lt;wps:&gt;</c> wrapper around one.
    /// Used to decide whether a non-textbox run still needs preserving via a
    /// raw-set append rather than being treated as a plain text run.
    /// </summary>
    private static bool IsDrawingBearingRun(string rawXml)
    {
        return rawXml.Contains("<w:drawing")
            || rawXml.Contains("<w:pict")
            || rawXml.Contains("<mc:AlternateContent")
            || rawXml.Contains("<wps:");
    }

    /// <summary>
    /// BUG-DUMP-TXBX: emit a typed <c>add textbox</c> row for the host of
    /// the current drawing run, followed by recursive inner-paragraph/run
    /// emits under <c>/&lt;host&gt;/textbox[N]</c>. Geometry props
    /// (width/height/wrap/anchor.x/anchor.y/fill) are extracted from the
    /// raw drawing XML so the rebuilt textbox keeps its layout.
    /// </summary>
    private static bool TryEmitTextbox(WordHandler word, DocumentNode run, string rawXml,
                                       string parentPath, List<BatchItem> items, BodyEmitContext? ctx,
                                       string? attachParaPath = null)
    {
        if (ctx == null) return false;

        // Only emit a typed `add textbox` for hosts AddTextbox itself
        // supports: /body, /body/tbl[..]/tc[N], /header[N], /footer[N].
        // Other parents fall through to the raw-set append.
        string hostPath = parentPath;
        if (!IsTextboxHostPath(hostPath)) return false;

        // BUG-D1-MULTIDRAWING-HOST: when N textboxes share a source
        // paragraph (side-by-side card layout), attach each to the same
        // already-emitted host paragraph (attachParaPath = /body/p[last()])
        // instead of /body — otherwise AddTextbox creates a fresh host per
        // textbox and the side-by-side layout fans out into N stacked
        // paragraphs. The textbox INDEX still scopes to hostPath so
        // /body/textbox[K] addressing remains continuous across the doc.
        string emitParent = attachParaPath ?? hostPath;

        // Allocate next 1-based textbox index for this host.
        int n = ctx.TextboxCounters.TryGetValue(hostPath, out var prev) ? prev + 1 : 1;
        ctx.TextboxCounters[hostPath] = n;
        string textboxPath = hostPath == "/" ? "/textbox[" + n + "]" : $"{hostPath}/textbox[{n}]";

        // Extract geometry / wrap / fill / anchor from the drawing XML so the
        // rebuilt textbox keeps its layout. Conservative best-effort — any
        // attribute we can't parse falls back to AddTextbox's defaults.
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            var doc = System.Xml.Linq.XDocument.Parse(rawXml);
            System.Xml.Linq.XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
            System.Xml.Linq.XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";

            var anchor = doc.Descendants(wp + "anchor").FirstOrDefault()
                      ?? (System.Xml.Linq.XElement?)doc.Descendants(wp + "inline").FirstOrDefault();
            var extent = doc.Descendants(wp + "extent").FirstOrDefault();
            if (extent != null)
            {
                var cx = extent.Attribute("cx")?.Value;
                var cy = extent.Attribute("cy")?.Value;
                if (!string.IsNullOrEmpty(cx)) props["width"] = cx + "emu";
                if (!string.IsNullOrEmpty(cy)) props["height"] = cy + "emu";
            }
            if (anchor != null)
            {
                var posHEl = anchor.Element(wp + "positionH");
                var posVEl = anchor.Element(wp + "positionV");
                var posH = posHEl?.Element(wp + "posOffset")?.Value;
                var posV = posVEl?.Element(wp + "posOffset")?.Value;
                // A floating drawing positions each axis EITHER by an absolute
                // posOffset OR by a relative <wp:align> (left/center/right /
                // top/bottom). Capture the align form too — without it a
                // center-aligned textbox round-trips as posOffset=0 (i.e. jumps
                // hard left). posOffset wins when both somehow appear.
                var alignH = posHEl?.Element(wp + "align")?.Value;
                var alignV = posVEl?.Element(wp + "align")?.Value;
                if (!string.IsNullOrEmpty(posH)) props["anchor.x"] = posH + "emu";
                else if (!string.IsNullOrEmpty(alignH)) props["hAlign"] = alignH;
                if (!string.IsNullOrEmpty(posV)) props["anchor.y"] = posV + "emu";
                else if (!string.IsNullOrEmpty(alignV)) props["vAlign"] = alignV;
                // Anchor reference frames. AddTextbox hardcodes column/paragraph;
                // without forwarding these a textbox anchored relativeFrom="page"
                // round-trips as relativeFrom="paragraph" and floats off-position
                // (the source posOffset is measured from a different origin).
                var hRel = posHEl?.Attribute("relativeFrom")?.Value;
                var vRel = posVEl?.Attribute("relativeFrom")?.Value;
                if (!string.IsNullOrEmpty(hRel)) props["hRelative"] = hRel;
                if (!string.IsNullOrEmpty(vRel)) props["vRelative"] = vRel;
                // wrap token
                if (anchor.Element(wp + "wrapSquare") != null) props["wrap"] = "square";
                else if (anchor.Element(wp + "wrapTight") != null) props["wrap"] = "tight";
                else if (anchor.Element(wp + "wrapTopAndBottom") != null) props["wrap"] = "topAndBottom";
                else if (anchor.Element(wp + "wrapNone") != null) props["wrap"] = "none";
            }
            var spPr = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "spPr");
            // Geometry preset (rect default). roundRect etc. otherwise reverted
            // to a sharp rectangle on rebuild.
            var prst = spPr?.Element(a + "prstGeom")?.Attribute("prst")?.Value;
            if (!string.IsNullOrEmpty(prst) && prst != "rect") props["geometry"] = prst;
            // Rotation (<a:xfrm rot>, 60000ths of a degree). Round-trips raw.
            var rot = spPr?.Element(a + "xfrm")?.Attribute("rot")?.Value;
            if (!string.IsNullOrEmpty(rot) && rot != "0") props["rotation"] = rot;
            // Fill: solidFill (with optional alpha) or gradFill inside wps:spPr.
            var solidFill = spPr?.Element(a + "solidFill");
            var srgbClr = solidFill?.Element(a + "srgbClr");
            if (srgbClr?.Attribute("val")?.Value is string fillHex && !string.IsNullOrEmpty(fillHex))
            {
                props["fill"] = fillHex;
                var alpha = srgbClr.Element(a + "alpha")?.Attribute("val")?.Value;
                if (!string.IsNullOrEmpty(alpha)) props["fill.opacity"] = alpha;
            }
            else
            {
                var grad = spPr?.Element(a + "gradFill");
                if (grad != null)
                {
                    var stops = grad.Element(a + "gsLst")?.Elements(a + "gs")
                        .Select(gs => (gs.Element(a + "srgbClr")?.Attribute("val")?.Value, gs.Attribute("pos")?.Value))
                        .Where(t => !string.IsNullOrEmpty(t.Item1))
                        .Select(t => $"{t.Item1}@{t.Item2 ?? "0"}");
                    if (stops != null)
                    {
                        var joined = string.Join(";", stops);
                        if (!string.IsNullOrEmpty(joined)) props["fill.gradient"] = joined;
                    }
                }
            }
            // Shadow (<a:effectLst><a:outerShdw>): emit the compact tuple the
            // add path understands so the drop shadow round-trips faithfully.
            var shdw = spPr?.Element(a + "effectLst")?.Element(a + "outerShdw");
            if (shdw != null)
            {
                var sClr = shdw.Element(a + "srgbClr");
                props["shadow"] = string.Join(";",
                    shdw.Attribute("blurRad")?.Value ?? "50800",
                    shdw.Attribute("dist")?.Value ?? "38100",
                    shdw.Attribute("dir")?.Value ?? "5400000",
                    sClr?.Attribute("val")?.Value ?? "000000",
                    sClr?.Element(a + "alpha")?.Attribute("val")?.Value ?? "40000");
            }
            // Line / border (<a:ln>): width (EMU, round-trips through ParseEmu),
            // solidFill color, and dash style. Without this the textbox outline
            // was dropped entirely on dump→batch — borders vanished and content
            // reflowed. <a:noFill/> means the box explicitly has no border.
            var ln = spPr?.Element(a + "ln");
            if (ln != null)
            {
                if (ln.Element(a + "noFill") != null)
                {
                    props["line.style"] = "none";
                }
                else
                {
                    var lnW = ln.Attribute("w")?.Value;
                    if (!string.IsNullOrEmpty(lnW)) props["line.width"] = lnW;
                    var lnClr = ln.Element(a + "solidFill")?.Element(a + "srgbClr")?.Attribute("val")?.Value;
                    if (!string.IsNullOrEmpty(lnClr)) props["line.color"] = lnClr;
                    // a:prstDash@val is already an OOXML dash name; AddTextbox's
                    // MapDashStyle accepts it (lower-cased) and re-emits it.
                    var dash = ln.Element(a + "prstDash")?.Attribute("val")?.Value;
                    if (!string.IsNullOrEmpty(dash)) props["line.style"] = dash;
                }
            }
            // bodyPr text insets. AddTextbox hardcodes Word defaults
            // (91440/45720); a source with zero insets (common on tight
            // letterhead title boxes) otherwise loses ~0.2in of usable width
            // and its text rewraps. Forward each present inset verbatim (EMU).
            var bodyPr = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "bodyPr");
            if (bodyPr != null)
            {
                foreach (var (attr, key) in new[] { ("lIns", "inset.left"), ("tIns", "inset.top"), ("rIns", "inset.right"), ("bIns", "inset.bottom") })
                {
                    var v = bodyPr.Attribute(attr)?.Value;
                    if (!string.IsNullOrEmpty(v)) props[key] = v;
                }
                // Vertical text flow + vertical anchor. Without these a vertical
                // (eaVert) box renders as char-wrapped horizontal text and a
                // centered box anchors to the top.
                var vert = bodyPr.Attribute("vert")?.Value;
                if (!string.IsNullOrEmpty(vert) && vert != "horz") props["textDirection"] = vert;
                var bAnchor = bodyPr.Attribute("anchor")?.Value;
                if (!string.IsNullOrEmpty(bAnchor) && bAnchor != "t") props["textAnchor"] = bAnchor;
            }
            // docPr name → alt
            var docPr = doc.Descendants(wp + "docPr").FirstOrDefault();
            var altName = docPr?.Attribute("name")?.Value;
            if (!string.IsNullOrEmpty(altName) && altName != "Text Box") props["alt"] = altName;
        }
        catch
        {
            // Parsing failures: still emit the `add textbox` row with whatever
            // we managed to extract; defaults cover the rest.
        }

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = emitParent,
            Type = "textbox",
            Props = props.Count > 0 ? props : null
        });

        // Recurse over inner content. Get on /<host>/textbox[N] returns the
        // <w:txbxContent>; its children are the inner <w:p>. AddTextbox auto-
        // seeds one empty <w:p>, so the first source paragraph uses set-on-
        // existing (autoPresent: true) and the rest emit as fresh adds.
        try
        {
            var txbxNode = word.Get(textboxPath);
            var children = txbxNode.Children ?? new List<DocumentNode>();
            int innerPIdx = 0;
            int innerTblIdx = 0;
            bool firstParaSeen = false;
            foreach (var child in children)
            {
                if (child.Type == "paragraph" || child.Type == "p")
                {
                    innerPIdx++;
                    // The generic fallback fabricates child paths from the
                    // OOXML LocalName ("/body/txbxContent[N]/p[M]") which the
                    // Navigation layer can't re-resolve — the user-facing
                    // path segment is "textbox", not "txbxContent". Use the
                    // canonical /body/textbox[N]/p[M] form instead.
                    var sourceParaPath = $"{textboxPath}/p[{innerPIdx}]";
                    EmitParagraph(word, sourceParaPath, textboxPath, innerPIdx, items,
                                  autoPresent: !firstParaSeen, ctx);
                    firstParaSeen = true;
                }
                else if (child.Type == "table" || child.Type == "tbl")
                {
                    // BUG-D1-TXBX-TABLE: tables nested INSIDE a textbox were
                    // silently dropped on dump — the children loop only
                    // recognised paragraph types. Reuse EmitTable with the
                    // textbox path as containerPath so the resulting
                    // `add table` rows target /body/textbox[N]/tbl[K]
                    // (AddTable already accepts a TextBoxContent parent).
                    innerTblIdx++;
                    var sourceTblPath = $"{textboxPath}/tbl[{innerTblIdx}]";
                    EmitTable(word, sourceTblPath, innerTblIdx, items, ctx,
                              parentTablePath: null, containerPath: textboxPath);
                }
                else if (child.Type == "sdt")
                {
                    // BUG-R5B(BUG1): a block-level <w:sdt> nested inside a
                    // textbox (the canonical centered page-number footer:
                    // textbox → sdt → <w:p> with a PAGE field) was silently
                    // dropped — the inner walk only recognised paragraph and
                    // table types, so the SDT (and the field inside it) vanished
                    // on dump→batch and the footer lost its page number. There
                    // is no typed `add sdt` parent for a txbxContent host, and a
                    // PAGE-field SDT is "rich" content the typed text emit cannot
                    // reproduce anyway. Inject the SDT verbatim via raw-set into
                    // the just-created textbox's <w:txbxContent> (mirrors the
                    // rich-block-SDT raw-set in EmitSdt). The part is the host's
                    // own part (/document for body, /header[N] or /footer[N]
                    // otherwise); the (//w:txbxContent)[n] index selects the
                    // textbox we just emitted.
                    var sdtRaw = word.RawElementXml(child.Path);
                    if (!string.IsNullOrEmpty(sdtRaw))
                    {
                        var rawPart = hostPath == "/body" ? "/document" : hostPath;
                        // `add textbox` auto-seeds one empty <w:p> in the
                        // txbxContent. To keep document order, an SDT that
                        // precedes every source paragraph (the canonical
                        // page-number footer: sdt-then-empty-p) is prepended so
                        // it lands ahead of the seed paragraph; an SDT that
                        // follows already-emitted paragraphs is appended.
                        items.Add(new BatchItem
                        {
                            Command = "raw-set",
                            Part = rawPart,
                            Xpath = $"(//w:txbxContent)[{n}]",
                            Action = firstParaSeen ? "append" : "prepend",
                            Xml = sdtRaw
                        });
                    }
                    else
                    {
                        ctx.Warnings.Add(new DocxUnsupportedWarning(
                            Element: "textbox.sdt",
                            Path: child.Path,
                            Reason: "content control nested inside a textbox could not be serialized for round-trip; it will be missing from the replayed document"));
                    }
                }
            }
        }
        catch
        {
            // If the inner walk fails for any reason, the typed `add textbox`
            // still landed — round-trip recreates an empty textbox with the
            // right geometry, which beats the previous double-emit.
        }
        return true;
    }

    private static bool IsTextboxHostPath(string parentPath)
    {
        // Matches ResolveDrawingHost: /body, /body/tbl[..]/tr[..]/tc[N],
        // /header[N], /footer[N]. Reject anything else so non-supported
        // hosts fall through to the raw-set append.
        if (string.Equals(parentPath, "/body", StringComparison.Ordinal)) return true;
        if (parentPath.StartsWith("/header[", StringComparison.Ordinal)
            && parentPath.EndsWith("]", StringComparison.Ordinal)
            && !parentPath.Substring(8).Contains('/')) return true;
        if (parentPath.StartsWith("/footer[", StringComparison.Ordinal)
            && parentPath.EndsWith("]", StringComparison.Ordinal)
            && !parentPath.Substring(8).Contains('/')) return true;
        if (parentPath.Contains("/tc[", StringComparison.Ordinal)
            && parentPath.EndsWith("]", StringComparison.Ordinal)) return true;
        return false;
    }

    /// <summary>
    /// R10-bug1: OLE / embedded-object runs (Type=="ole"). Full round-trip:
    /// pull the embedded payload + VML icon binaries and the frame metadata
    /// (progId / DrawAspect / dimensions / alt-name) and emit a self-contained
    /// `add ole` carrying the bytes as data: URIs (mirrors picture-run base64
    /// inlining). AddOle rebuilds the embedded part, icon part and &lt;w:object&gt;
    /// wrapper from these — no external src file needed.
    ///
    /// If the payload can't be resolved (orphaned relationship / unreadable
    /// part), fall back to keeping the host paragraph and emitting a warning
    /// naming the path, rather than silently dropping the object.
    /// </summary>
    private static bool TryEmitOleRun(DocumentNode run, string paraTargetPath, List<BatchItem> items, BodyEmitContext? ctx, WordHandler word)
    {
        if (run.Type != "ole") return false;

        var data = word.GetOleEmitData(run.Path);
        if (data != null)
        {
            var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["src"] = $"data:{data.EmbeddedContentType};base64,{Convert.ToBase64String(data.EmbeddedBytes)}",
                ["oleKind"] = data.OleKind,
                ["contentType"] = data.EmbeddedContentType,
            };
            if (!string.IsNullOrEmpty(data.EmbeddedExt)) props["embedExt"] = data.EmbeddedExt;
            if (!string.IsNullOrEmpty(data.ProgId)) props["progId"] = data.ProgId!;
            if (!string.IsNullOrEmpty(data.Display)) props["display"] = data.Display!;
            if (!string.IsNullOrEmpty(data.Width)) props["width"] = data.Width!;
            if (!string.IsNullOrEmpty(data.Height)) props["height"] = data.Height!;
            if (!string.IsNullOrEmpty(data.Name)) props["name"] = data.Name!;
            if (data.IconBytes is { Length: > 0 })
                props["icon"] = $"data:{data.IconContentType ?? "image/png"};base64,{Convert.ToBase64String(data.IconBytes)}";
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "ole",
                Props = props,
            });
            return true;
        }

        if (ctx != null)
        {
            var progId = run.Format.TryGetValue("progId", out var pid) ? pid?.ToString() : null;
            var reason = progId != null
                ? $"ole run dropped (progId={progId}); embedded payload could not be resolved for round-trip"
                : "ole run dropped; embedded payload could not be resolved for round-trip";
            ctx.Warnings.Add(new DocxUnsupportedWarning(
                Element: "ole",
                Path: run.Path,
                Reason: reason));
        }
        return true;
    }

    private static bool TryEmitNoteRefRun(WordHandler word, DocumentNode run, string paraTargetPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // Footnote/endnote reference runs carry a <w:footnoteReference> /
        // <w:endnoteReference> child. Emit them as a typed footnote/endnote
        // add anchored on the host paragraph and pull the body text from the
        // pre-resolved ordered list — see BodyEmitContext for the
        // document-order assumption.
        //
        // BUG-R7B(BUG1): detection cannot rely on rStyle == "FootnoteReference":
        // that is only the default style name. Real documents reference the
        // note with an arbitrary style id (e.g. rStyle="a5"), so the run's
        // rStyle never matched and the reference (plus its note body) was
        // silently dropped — Get surfaces no Format key for the reference
        // element, so probe the raw run XML for the actual reference child.
        if (ctx == null) return false;
        var rStyle = run.Format.TryGetValue("rStyle", out var rs) ? rs?.ToString() : null;
        var noteKind = ClassifyNoteRefRun(word, run, rStyle);
        if (noteKind == NoteRefKind.Footnote)
        {
            var noteText = ctx.FootnoteCursor.Index < ctx.FootnoteTexts.Count
                ? ctx.FootnoteTexts[ctx.FootnoteCursor.Index]
                : "";
            ctx.FootnoteCursor.Index++;
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "footnote",
                Props = new() { ["text"] = noteText }
            });
            return true;
        }
        if (noteKind == NoteRefKind.Endnote)
        {
            var noteText = ctx.EndnoteCursor.Index < ctx.EndnoteTexts.Count
                ? ctx.EndnoteTexts[ctx.EndnoteCursor.Index]
                : "";
            ctx.EndnoteCursor.Index++;
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "endnote",
                Props = new() { ["text"] = noteText }
            });
            return true;
        }
        return false;
    }

    private enum NoteRefKind { None, Footnote, Endnote }

    // BUG-R7B(BUG1): a run is a footnote/endnote reference when it contains a
    // <w:footnoteReference>/<w:endnoteReference> child — the rStyle is only a
    // weak hint (defaults to FootnoteReference/EndnoteReference but is an
    // arbitrary style id in real documents). Probe the raw XML; fall back to
    // the rStyle name when raw XML is unavailable.
    private static NoteRefKind ClassifyNoteRefRun(WordHandler word, DocumentNode run, string? rStyle)
    {
        if (run.Type != "run" && run.Type != "r") return NoteRefKind.None;
        var raw = word.GetElementXml(run.Path);
        if (!string.IsNullOrEmpty(raw))
        {
            if (raw.Contains("footnoteReference", StringComparison.Ordinal)) return NoteRefKind.Footnote;
            if (raw.Contains("endnoteReference", StringComparison.Ordinal)) return NoteRefKind.Endnote;
            return NoteRefKind.None;
        }
        if (string.Equals(rStyle, "FootnoteReference", StringComparison.OrdinalIgnoreCase)) return NoteRefKind.Footnote;
        if (string.Equals(rStyle, "EndnoteReference", StringComparison.OrdinalIgnoreCase)) return NoteRefKind.Endnote;
        return NoteRefKind.None;
    }

    private static void EmitPlainOrHyperlinkRun(DocumentNode run, string paraTargetPath, List<BatchItem> items, BodyEmitContext? ctx = null)
    {
        // BUG-R12A(BUG1): a hyperlink wrapper with >1 run or any per-run rPr was
        // stashed by CoalesceHyperlinkRuns with its original runs in Children.
        // Emit the wrapper carrying the FIRST run's text + rPr, then one
        // structured `add run` per remaining run targeting the hyperlink path so
        // each run's bold/color/size/font survives round-trip (the flat
        // `add hyperlink text=` path would flatten them into one unformatted run).
        // Mirrors the R9 comment-body structured-add-run fix; raw-set is not an
        // option here because the <w:hyperlink> r:id relationship can't be
        // recreated by verbatim XML injection.
        if (run.Format.TryGetValue("_hlStructured", out var hlsObj) && hlsObj is bool hlsB && hlsB
            && run.Children is { Count: > 0 } hlRuns)
        {
            EmitStructuredHyperlink(hlRuns, paraTargetPath, items, ctx);
            return;
        }
        var rProps = FilterEmittableProps(run.Format);
        if (!string.IsNullOrEmpty(run.Text))
            rProps["text"] = run.Text!;
        // revision.beforeLost — see EmitParagraph counterpart. Strip from
        // the emitted props (AddRun would flag it UNSUPPORTED) and surface
        // a dump warning so the lost rPrChange snapshot is visible to the
        // caller.
        if (rProps.Remove("revision.beforeLost") && ctx != null)
        {
            ctx.Warnings.Add(new DocxUnsupportedWarning(
                Element: "revision.before",
                Path: run.Path,
                Reason: "rPrChange previous-properties snapshot dropped on dump→batch round-trip; Word's Reject-Change restoration state will be lost on replay"));
        }

        // Hyperlink-wrapped run: Get flattens a <w:hyperlink>'s child run
        // into a regular run-typed node but copies the resolved URL onto
        // Format["url"]. AddRun does not consume `url` — emitting type="r"
        // would silently drop the hyperlink wrapper. Re-emit as a typed
        // `add hyperlink` so the <w:hyperlink>+rel-relationship round-trip
        // rebuilds correctly.
        // CONSISTENCY(docx-hyperlink-canonical-url): canonical key is `url`
        // on both Get readback and Add input.
        if (rProps.ContainsKey("url") || rProps.ContainsKey("anchor")
            || rProps.ContainsKey("isHyperlink"))
        {
            // AddHyperlink writes its own color/underline defaults from theme;
            // drop the inferred `color: hyperlink` / `underline: single` Get
            // echoes back so we don't override those defaults.
            if (rProps.TryGetValue("color", out var hlColor)
                && string.Equals(hlColor, "hyperlink", StringComparison.OrdinalIgnoreCase))
                rProps.Remove("color");
            if (rProps.TryGetValue("underline", out var hlUl)
                && string.Equals(hlUl, "single", StringComparison.OrdinalIgnoreCase))
                rProps.Remove("underline");
            rProps.Remove("isHyperlink");
            // Bare <w:hyperlink> wrapper with no url/anchor/tooltip/tgtFrame
            // /history carries no round-trippable property — AddHyperlink
            // would reject it. Fall through and emit as a plain run.
            if (!rProps.ContainsKey("url") && !rProps.ContainsKey("anchor")
                && !rProps.ContainsKey("tooltip") && !rProps.ContainsKey("tgtFrame")
                && !rProps.ContainsKey("tgtframe") && !rProps.ContainsKey("history"))
            {
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = paraTargetPath,
                    Type = "r",
                    Props = rProps.Count > 0 ? rProps : null
                });
                return;
            }
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "hyperlink",
                Props = rProps,
            });
            return;
        }
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = paraTargetPath,
            Type = "r",
            Props = rProps.Count > 0 ? rProps : null
        });
    }

    // BUG-R12A(BUG1): emit a hyperlink wrapper + one structured `add run` per
    // wrapped run so per-run rPr survives. The first run materializes the
    // wrapper (AddHyperlink seeds the wrapper's first run from `text` + rPr
    // props); subsequent runs are appended via `add r` targeting the
    // hyperlink[K] path (verified working — AddRun accepts a hyperlink parent
    // and preserves bold/color/size/font). hlIndex is the 1-based ordinal of
    // this hyperlink among the rows already emitted under the host paragraph.
    private static void EmitStructuredHyperlink(List<DocumentNode> hlRuns, string paraTargetPath,
                                                List<BatchItem> items, BodyEmitContext? ctx)
    {
        // Build the wrapper add from the first run's props (url/anchor/tooltip/…
        // + the run's own rPr). Reuse the existing flat-emit logic by routing
        // the first run through EmitPlainOrHyperlinkRun with the structured flag
        // cleared, so the theme-default scrubbing + bare-wrapper fallback stay
        // identical.
        var first = hlRuns[0];
        var firstClone = new DocumentNode
        {
            Path = first.Path,
            Type = first.Type,
            Text = first.Text,
            Format = new Dictionary<string, object?>(first.Format, StringComparer.OrdinalIgnoreCase),
        };
        firstClone.Format.Remove("_hlStructured");
        int hlBefore = items.Count(it => it.Type == "hyperlink"
            && string.Equals(it.Parent, paraTargetPath, StringComparison.Ordinal));
        EmitPlainOrHyperlinkRun(firstClone, paraTargetPath, items, ctx);
        int hlAfter = items.Count(it => it.Type == "hyperlink"
            && string.Equals(it.Parent, paraTargetPath, StringComparison.Ordinal));
        // If the first run did not materialize a hyperlink row (bare wrapper with
        // no url/anchor/tooltip → AddHyperlink can't represent it, so it fell back
        // to a plain run), the wrapper is gone — emit the remaining runs as plain
        // runs too so their text/rPr still survive.
        if (hlAfter == hlBefore)
        {
            for (int k = 1; k < hlRuns.Count; k++)
            {
                var clone = new DocumentNode
                {
                    Path = hlRuns[k].Path,
                    Type = hlRuns[k].Type,
                    Text = hlRuns[k].Text,
                    Format = new Dictionary<string, object?>(hlRuns[k].Format, StringComparer.OrdinalIgnoreCase),
                };
                clone.Format.Remove("_hlStructured");
                clone.Format.Remove("url");
                clone.Format.Remove("anchor");
                EmitPlainOrHyperlinkRun(clone, paraTargetPath, items, ctx);
            }
            return;
        }
        var hlPath = $"{paraTargetPath}/hyperlink[{hlAfter}]";
        for (int k = 1; k < hlRuns.Count; k++)
        {
            var rProps = FilterEmittableProps(hlRuns[k].Format);
            // The hyperlink-wrapper keys belong to the <w:hyperlink>, not its
            // child runs — strip them so `add r` doesn't choke / re-wrap.
            foreach (var wk in HyperlinkWrapperOnlyKeys) rProps.Remove(wk);
            // Drop the theme-default echoes Get stamps on every hyperlink run;
            // the wrapper's own run already carries the real Hyperlink style.
            if (rProps.TryGetValue("color", out var c)
                && string.Equals(c, "hyperlink", StringComparison.OrdinalIgnoreCase))
                rProps.Remove("color");
            if (rProps.TryGetValue("underline", out var u)
                && string.Equals(u, "single", StringComparison.OrdinalIgnoreCase))
                rProps.Remove("underline");
            if (!string.IsNullOrEmpty(hlRuns[k].Text))
                rProps["text"] = hlRuns[k].Text!;
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = hlPath,
                Type = "r",
                Props = rProps.Count > 0 ? rProps : null
            });
        }
    }

    // BUG-R12A(BUG1): raw-set a rich inline (run-level) <w:sdt> into the host
    // paragraph so its per-run formatting survives. Returns true when it emitted
    // a raw-set (or a warning + fall-through is desired); false when the SDT is
    // simple enough for the flat `add sdt text=` fast path. The host paragraph
    // was just added by EmitParagraph, so it is the last <w:p> in the body at
    // replay time — the same last()-relative attach the inline-textbox raw-set
    // uses.
    private static bool TryEmitRichInlineSdt(WordHandler word, DocumentNode sdt,
                                             List<BatchItem> items, BodyEmitContext? ctx)
    {
        var rawXml = word.RawElementXml(sdt.Path);
        if (string.IsNullOrEmpty(rawXml) || !IsRichInlineSdt(rawXml!)) return false;
        // External relationship references (hyperlink r:id, image r:embed/r:link)
        // would dangle in the blank target — raw injection does not recreate the
        // matching rels. Fall back to the flat text emit and surface the loss.
        if (HasExternalRelRef(rawXml!))
        {
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "sdt.richContent",
                Path: sdt.Path,
                Reason: "inline content control with formatted runs AND external relationship references (hyperlinks/images) flattened to text on dump"));
            return false;
        }
        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/document",
            Xpath = "/w:document/w:body/w:p[last()]",
            Action = "append",
            Xml = rawXml
        });
        return true;
    }

    // A run-level <w:sdt> is "rich" when its sdtContent carries formatting the
    // text-only typed emit can't reproduce: more than one run, any run-level
    // <w:rPr>, a hyperlink/field, or an intra-run break/tab. Single plain run →
    // flat `add sdt text=` stays lossless. (Distinct from IsRichBlockSdt, which
    // keys on inner <w:p> count — a run-level SDT has no inner paragraph.)
    private static bool IsRichInlineSdt(string sdtXml)
    {
        // More than one content run.
        if (System.Text.RegularExpressions.Regex.Matches(sdtXml, "<w:r[ >]").Count > 1)
            return true;
        // Any run-level run properties (bold/color/size/font/…) inside content.
        if (sdtXml.Contains("<w:rPr", StringComparison.Ordinal)
            // exclude sdtPr/endPr-only rPr false positives: those live in
            // <w:sdtPr>/<w:sdtEndPr><w:rPr>; a content rPr sits under <w:r>.
            && System.Text.RegularExpressions.Regex.IsMatch(sdtXml, "<w:r[ >].*?<w:rPr"))
            return true;
        return sdtXml.Contains("<w:hyperlink", StringComparison.Ordinal)
            || sdtXml.Contains("<w:fldChar", StringComparison.Ordinal)
            || sdtXml.Contains("w:instrText", StringComparison.Ordinal)
            || sdtXml.Contains("<w:fldSimple", StringComparison.Ordinal)
            || sdtXml.Contains("<w:drawing", StringComparison.Ordinal)
            || sdtXml.Contains("<w:br", StringComparison.Ordinal)
            || sdtXml.Contains("<w:tab", StringComparison.Ordinal)
            || sdtXml.Contains("<w:cr", StringComparison.Ordinal);
    }

    // Collapse OOXML complex field chains (fldChar(begin) + instrText + …
    // + fldChar(end)) into a single synthetic "field" DocumentNode with
    // Format["instruction"] (raw code) and Text (cached display value).
    // Non-field children pass through untouched in original order. The TOC
    // chain is handled by the dedicated EmitParagraph branch above and never
    // reaches this collapsing step (early-return in that branch).
    // BUG-DUMP6-05: collapse consecutive runs sharing the same url/anchor
    // into a single synthetic node so dump emits ONE `add hyperlink` per
    // <w:hyperlink>, regardless of how many runs the source wrapped. The
    // synthesized node carries the merged Text (for AddHyperlink's `text`
    // prop) and the shared url/anchor/Hyperlink-style format keys.
    // Mirrors the field-emit hyperlink-parent rebase logic for tab/ptab runs.
    // Navigation marks tab-only runs that live inside w:hyperlink with a
    // Format["_hyperlinkParent"] hint (e.g. /body/p[1]/hyperlink[2]); without
    // re-routing on emit they would replay under the bare paragraph and lose
    // the hyperlink wrapper. The candidate-verify step (a prior `add hyperlink`
    // row must have landed under paraTargetPath) avoids dangling paths when
    // the hyperlink has no emittable runs and so was never added.
    private static string ResolveHyperlinkParent(DocumentNode run, string paraTargetPath, List<BatchItem> items)
    {
        string? candidateHlParent = null;
        if (run.Format.TryGetValue("_hyperlinkParent", out var hlpObj) && hlpObj != null)
        {
            var hint = hlpObj.ToString();
            if (!string.IsNullOrEmpty(hint)) candidateHlParent = hint;
        }
        if (candidateHlParent == null) return paraTargetPath;

        const string hlMarker = "/hyperlink[";
        var hlIdxStart = candidateHlParent.LastIndexOf(hlMarker, StringComparison.Ordinal);
        if (hlIdxStart <= 0) return paraTargetPath;
        var hlEnd = candidateHlParent.IndexOf(']', hlIdxStart);
        if (hlEnd <= hlIdxStart) return paraTargetPath;
        var kStr = candidateHlParent.Substring(hlIdxStart + hlMarker.Length,
            hlEnd - hlIdxStart - hlMarker.Length);
        if (!int.TryParse(kStr, out var kIdx)) return paraTargetPath;
        var rebased = paraTargetPath + candidateHlParent.Substring(hlIdxStart);
        int emittedHls = items.Count(it => it.Type == "hyperlink"
            && string.Equals(it.Parent, paraTargetPath, StringComparison.Ordinal));
        return emittedHls >= kIdx ? rebased : paraTargetPath;
    }

    // BUG-R12A(BUG1): hyperlink-wrapper keys + the theme defaults AddHyperlink
    // re-applies on its own. A run carrying ONLY these keys has no per-run
    // formatting worth preserving structurally, so the flat `add hyperlink
    // text=` fast path stays lossless. Any OTHER key (bold, color≠hyperlink,
    // size, font, …) means the run's rPr must survive — route to structured emit.
    private static readonly HashSet<string> HyperlinkWrapperOnlyKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "url", "anchor", "tooltip", "tgtFrame", "tgtframe", "history",
        "docLocation", "doclocation", "isHyperlink", "_hyperlinkParent",
        "rStyle", "rstyle",
    };

    // True when a hyperlink-wrapped run carries run-level formatting (rPr) that
    // AddHyperlink's flat `text=` path cannot reproduce. The theme-default
    // color=hyperlink / underline=single echoes that Get stamps back are NOT
    // real formatting (AddHyperlink re-applies them), so they don't count.
    private static bool HyperlinkRunHasRunFormatting(DocumentNode run)
    {
        var props = FilterEmittableProps(run.Format);
        foreach (var (k, v) in props)
        {
            if (HyperlinkWrapperOnlyKeys.Contains(k)) continue;
            if (string.Equals(k, "text", StringComparison.OrdinalIgnoreCase)) continue;
            if (string.Equals(k, "color", StringComparison.OrdinalIgnoreCase)
                && string.Equals(v, "hyperlink", StringComparison.OrdinalIgnoreCase)) continue;
            if (string.Equals(k, "underline", StringComparison.OrdinalIgnoreCase)
                && string.Equals(v, "single", StringComparison.OrdinalIgnoreCase)) continue;
            return true;
        }
        return false;
    }

    private static List<DocumentNode> CoalesceHyperlinkRuns(List<DocumentNode> runs)
    {
        var result = new List<DocumentNode>(runs.Count);
        int i = 0;
        while (i < runs.Count)
        {
            var run = runs[i];
            string? url = null, anchor = null, hlParent = null;
            if (run.Type == "run" || run.Type == "r")
            {
                if (run.Format.TryGetValue("url", out var u))
                    url = u?.ToString();
                if (run.Format.TryGetValue("anchor", out var a))
                    anchor = a?.ToString();
                // `_hyperlinkParent` (e.g. `/body/p[1]/hyperlink[2]`) is the
                // unique-per-wrapper marker NodeBuilder stamps on every run
                // hosted inside a `<w:hyperlink>`. Two sibling hyperlinks
                // sharing the same target URL (e.g. two "Read more →" links
                // pointing at /pricing) carry identical `url`/`anchor` but
                // distinct `_hyperlinkParent` paths, so the URL-only equality
                // pre-R12 collapsed them into one item whose text was
                // "Read more →Read more →" — silent loss of the second link.
                if (run.Format.TryGetValue("_hyperlinkParent", out var hp))
                    hlParent = hp?.ToString();
            }
            if (string.IsNullOrEmpty(url) && string.IsNullOrEmpty(anchor))
            {
                result.Add(run);
                i++;
                continue;
            }
            // Walk forward over consecutive runs with the same url/anchor AND
            // the same hyperlink-wrapper path. Differing `_hyperlinkParent`
            // values mark a hyperlink boundary even when URL/anchor match.
            int j = i + 1;
            var group = new List<DocumentNode> { run };
            var sb = new System.Text.StringBuilder(run.Text ?? "");
            while (j < runs.Count)
            {
                var next = runs[j];
                if (next.Type != "run" && next.Type != "r") break;
                next.Format.TryGetValue("url", out var nUrlObj);
                next.Format.TryGetValue("anchor", out var nAncObj);
                next.Format.TryGetValue("_hyperlinkParent", out var nHlpObj);
                var nUrl = nUrlObj?.ToString();
                var nAnchor = nAncObj?.ToString();
                var nHlParent = nHlpObj?.ToString();
                if (!string.Equals(nUrl, url, StringComparison.Ordinal)) break;
                if (!string.Equals(nAnchor, anchor, StringComparison.Ordinal)) break;
                if (!string.Equals(nHlParent, hlParent, StringComparison.Ordinal)) break;
                group.Add(next);
                sb.Append(next.Text ?? "");
                j++;
            }
            // BUG-R12A(BUG1): the flat `add hyperlink text=Part1Part2` fast path
            // drops every per-run rPr on the wrapped runs (a 2nd bold+red run
            // round-trips as plain text). Keep the fast path ONLY for the lossless
            // case: a single run with no run-level formatting. Otherwise stash the
            // original group runs in Children and flag the node so
            // EmitPlainOrHyperlinkRun emits the wrapper + structured `add run`
            // rows (mirrors the R9 comment-body fix), preserving each run's rPr.
            bool needsStructured = group.Count > 1 || group.Any(HyperlinkRunHasRunFormatting);
            var merged = new DocumentNode
            {
                Path = run.Path,
                Type = run.Type,
                Text = sb.ToString(),
                Format = new Dictionary<string, object?>(run.Format, StringComparer.OrdinalIgnoreCase),
            };
            if (needsStructured)
            {
                merged.Format["_hlStructured"] = true;
                merged.Children = group;
            }
            result.Add(merged);
            i = j;
        }
        return result;
    }

    // BUG-DUMP-HOIST: run-level character properties that WordHandler.Navigation
    // surfaces on the paragraph node (via the firstRun fallback) but which must
    // NOT ride on `add p` for multi-run paragraphs — every individual run gets
    // its own `add r` carrying its real props.
    private static readonly HashSet<string> RunCharacterPropsHoistedFromFirstRun = new(StringComparer.OrdinalIgnoreCase)
    {
        "bold", "italic", "size", "color", "underline", "underline.color",
        "strike", "highlight",
        "font.latin", "font.ea", "font.ascii", "font.hAnsi",
        // complex-script siblings populated by ReadComplexScriptRunFormatting
        "bold.cs", "italic.cs", "size.cs", "font.cs",
    };

    private static void StripRunCharacterPropsFromParagraph(Dictionary<string, string> props)
    {
        foreach (var k in RunCharacterPropsHoistedFromFirstRun)
            props.Remove(k);
    }

    // Layer per-stop `add tab` rows under a parent path that already has the
    // host paragraph/style created. tabs is the flat List<Dict> Get exposes.
    private static void EmitTabStops(string parentPath, object? tabsVal, List<BatchItem> items)
    {
        if (tabsVal is not IEnumerable<Dictionary<string, object?>> list) return;
        foreach (var t in list)
        {
            var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (t.TryGetValue("pos", out var p) && p != null) props["pos"] = p.ToString() ?? "";
            if (t.TryGetValue("val", out var v) && v != null) props["val"] = v.ToString() ?? "";
            if (t.TryGetValue("leader", out var l) && l != null) props["leader"] = l.ToString() ?? "";
            if (props.Count == 0 || !props.ContainsKey("pos")) continue;
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = parentPath,
                Type = "tab",
                Props = props
            });
        }
    }
}
