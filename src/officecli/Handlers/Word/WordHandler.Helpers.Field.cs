// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using Vml = DocumentFormat.OpenXml.Vml;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{

    /// <summary>
    /// Find body-level fields whose fldChar(begin)…fldChar(end) chain straddles
    /// more than one paragraph (e.g. a real, cached Table of Contents — the
    /// outer TOC field opens in the first entry paragraph and closes in the
    /// last). Returns inclusive 1-based <c>/body/p[N]</c> positional ranges
    /// (the same numbering EmitBody/Navigation use — oMathPara wrappers are not
    /// counted). WordBatchEmitter raw-passes each such span verbatim instead of
    /// collapsing the opener through AddToc/AddField, which can only model a
    /// self-contained single-paragraph field and would otherwise destroy the
    /// first entry's cached content and detach the rest of the result from the
    /// field. Self-contained fields (begin…end balanced within one paragraph)
    /// are NOT returned — they round-trip correctly through the typed path.
    /// </summary>
    internal List<(int Start, int End)> GetCrossParagraphFieldSpanRanges()
    {
        var spans = new List<(int, int)>();
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return spans;

        int pos = 0;        // /body/p[N] positional index of the current paragraph
        int depth = 0;      // open outer-field nesting carried across paragraphs
        int spanStart = -1; // /body/p[N] index where the open span began
        foreach (var el in body.ChildElements)
        {
            if (el is Paragraph p)
            {
                if (IsOMathParaWrapperParagraph(p)) continue;
                pos++;
                int begins = 0, ends = 0;
                foreach (var fc in p.Descendants<FieldChar>())
                {
                    if (fc.FieldCharType?.HasValue != true) continue;
                    var t = fc.FieldCharType.InnerText;
                    if (t == "begin") begins++;
                    else if (t == "end") ends++;
                }
                if (spanStart < 0)
                {
                    // Not currently inside a cross-paragraph field. A paragraph
                    // that opens more field chars than it closes starts one;
                    // a balanced paragraph (self-contained field) is left alone.
                    if (begins > ends) { spanStart = pos; depth = begins - ends; }
                }
                else
                {
                    depth += begins - ends;
                    if (depth <= 0) { spans.Add((spanStart, pos)); depth = 0; spanStart = -1; }
                }
            }
            else if (spanStart >= 0)
            {
                // A non-paragraph body child (table, sdt, …) interrupts an open
                // field span. Such a field can't be represented as a run of
                // consecutive paragraphs — abandon the span so its paragraphs
                // fall back to the normal per-paragraph emit (degraded but
                // safe) rather than producing a malformed raw slice.
                depth = 0; spanStart = -1;
            }
        }
        // An unterminated span (begin with no matching end before end-of-body)
        // is malformed; drop it so the opener falls back to the typed path.
        return spans;
    }

    // CONSISTENCY(field-cache-stale): true when <paramref name="run"/> sits
    // between an owning field's <w:fldChar w:fldCharType="separate"/> and
    // <w:fldChar w:fldCharType="end"/> — i.e. it is the cached result run
    // that Word will overwrite when it recomputes the field. Used by the
    // Set "text=" path to decide whether the caller needs the field marked
    // dirty so their manual edit is preserved on next Word open.
    private static bool IsFieldCachedRun(Run run)
    {
        // Walk backward; the most recent field-char we hit must be a
        // `separate` (with no closing `end` between us and it). Track depth
        // to ignore fully-closed nested fields.
        int closedDepth = 0;
        OpenXmlElement? sibling = run.PreviousSibling();
        while (sibling != null)
        {
            if (sibling is Run sibRun)
            {
                var fld = sibRun.GetFirstChild<FieldChar>();
                if (fld?.FieldCharType?.HasValue == true)
                {
                    var t = fld.FieldCharType.InnerText;
                    if (t == "end")
                        closedDepth++;
                    else if (t == "begin")
                    {
                        if (closedDepth == 0) return false; // begin without separate → not cached
                        closedDepth--;
                    }
                    else if (t == "separate" && closedDepth == 0)
                        return true;
                }
            }
            sibling = sibling.PreviousSibling();
        }
        return false;
    }

    // CONSISTENCY(field-cache-stale): walk back from a run carrying an
    // <w:instrText> to the OWNING field's <w:fldChar fldCharType="begin">
    // in the same paragraph and set its dirty="true" attribute so Word
    // recomputes the field on next open. Used by Set when the instruction
    // text is rewritten — without dirty, the cached result run keeps the
    // old display value (e.g. "PAGE → DATE" still shows the old page
    // number) until the user manually presses F9.
    private static void MarkOwningFieldDirty(Run run)
    {
        var para = run.Parent;
        if (para == null) return;
        // Walk siblings backward from this run looking for the OWNING
        // field's <w:fldChar w:fldCharType="begin">. Track depth so that
        // a fully-closed inner field does not get its begin mistaken for
        // the owner of an outer instr. Each `end` we pass while walking
        // means we entered a closed nested field (going backwards), so
        // its `begin` is below us — skip past it. Only the begin at
        // depth 0 is the owner. Use InnerText (not enum equality) since
        // SDK v3 enum equality on FieldCharValues is unreliable (same
        // trap as LineSpacingRuleValues — see WordHandler CLAUDE.md).
        int closedDepth = 0;
        OpenXmlElement? sibling = run.PreviousSibling();
        while (sibling != null)
        {
            if (sibling is Run sibRun)
            {
                var fld = sibRun.GetFirstChild<FieldChar>();
                if (fld?.FieldCharType?.HasValue == true)
                {
                    var t = fld.FieldCharType.InnerText;
                    if (t == "end")
                    {
                        closedDepth++;
                    }
                    else if (t == "begin")
                    {
                        if (closedDepth == 0)
                        {
                            fld.Dirty = true;
                            return;
                        }
                        closedDepth--;
                    }
                }
            }
            sibling = sibling.PreviousSibling();
        }
    }

    /// <summary>
    /// Generate a unique 8-character uppercase hex ID for w14:paraId / w14:textId.
    /// OOXML spec requires value &lt; 0x80000000 (MaxExclusive).
    /// Uses deterministic increment from _nextParaId, wraps around on overflow,
    /// skips IDs already in use.
    /// </summary>
    private string GenerateParaId()
    {
        const int maxExclusive = 0x7FFFFFFF; // OOXML spec limit
        const int minStartId = 0x100000;
        var startId = _nextParaId;
        while (true)
        {
            var id = _nextParaId.ToString("X8");
            _nextParaId++;
            if (_nextParaId > maxExclusive)
                _nextParaId = minStartId;
            if (_usedParaIds.Add(id))
                return id;
            // Safety: if we've wrapped all the way around, something is very wrong
            if (_nextParaId == startId)
                throw new InvalidOperationException("No available paraId slots");
        }
    }

    /// <summary>
    /// Generate a unique decimal revision id (w:id on w:ins/w:del/w:moveFrom/
    /// w:moveTo/w:rPrChange/w:pPrChange/w:sectPrChange/w:tblPrChange/...).
    /// Reuses the paraId allocator (counter + _usedParaIds HashSet) so
    /// revision ids are globally unique across the document and never collide
    /// with paraId/textId or other revision ids ever allocated in this handler
    /// instance. Replaces the old `(GenerateParaId().GetHashCode() &amp; 0x7FFFFFFF)`
    /// fallback whose hash output was not actually unique.
    /// </summary>
    private string GenerateRevisionId()
    {
        // Same allocator as paraId, formatted as decimal (OOXML w:id is xsd:integer).
        var hexId = GenerateParaId();
        return int.Parse(hexId, System.Globalization.NumberStyles.HexNumber)
                  .ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Enumerate every part root that can hold revision elements
    /// (body + headers + footers + footnotes + endnotes + comments).
    /// Mirrors the part fan-out in EnsureAllParaIds for the paraId scan.
    /// Used by both EnsureAllParaIds (revision id pre-registration) and any
    /// future revision-iteration logic.
    /// </summary>
    private static IEnumerable<OpenXmlElement> EnumerateRevisionRoots(MainDocumentPart mainPart)
    {
        if (mainPart.Document != null) yield return mainPart.Document;
        foreach (var hp in mainPart.HeaderParts)
            if (hp.Header != null) yield return hp.Header;
        foreach (var fp in mainPart.FooterParts)
            if (fp.Footer != null) yield return fp.Footer;
        if (mainPart.FootnotesPart?.Footnotes != null) yield return mainPart.FootnotesPart.Footnotes;
        if (mainPart.EndnotesPart?.Endnotes != null) yield return mainPart.EndnotesPart.Endnotes;
        if (mainPart.WordprocessingCommentsPart?.Comments != null)
            yield return mainPart.WordprocessingCommentsPart.Comments;
    }

    /// <summary>
    /// Assign paraId and textId to a paragraph if not already set.
    /// </summary>
    private void AssignParaId(Paragraph para)
    {
        if (string.IsNullOrEmpty(para.ParagraphId?.Value))
            para.ParagraphId = GenerateParaId();
        if (string.IsNullOrEmpty(para.TextId?.Value))
            para.TextId = GenerateParaId();
    }

    /// <summary>
    /// Ensure all paragraphs in the document have w14:paraId and w14:textId.
    /// Called on document open AND after every successful RawSet (raw XML can
    /// inject paragraphs with missing or colliding paraIds — see the
    /// CONSISTENCY(paraid-global-uniqueness) note in WordHandler.RawSet).
    /// </summary>
    private void EnsureAllParaIds()
    {
        var mainPart = _doc.MainDocumentPart;
        if (mainPart?.Document?.Body == null) return;

        _usedParaIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // CONSISTENCY(paraid-global-uniqueness): paraId is allocated from a
        // single _nextParaId counter shared across the entire handler, so
        // EVERY part that can hold paragraphs must contribute to the
        // collision set. Body + headers + footers were already covered;
        // footnotes/endnotes/comments were missed, letting newly generated
        // paraIds collide with paraIds Word had already written into those
        // parts (rare in practice but a real correctness gap).
        var allParagraphs = mainPart.Document.Body.Descendants<Paragraph>().AsEnumerable();
        foreach (var headerPart in mainPart.HeaderParts)
            if (headerPart.Header != null)
                allParagraphs = allParagraphs.Concat(headerPart.Header.Descendants<Paragraph>());
        foreach (var footerPart in mainPart.FooterParts)
            if (footerPart.Footer != null)
                allParagraphs = allParagraphs.Concat(footerPart.Footer.Descendants<Paragraph>());
        if (mainPart.FootnotesPart?.Footnotes != null)
            allParagraphs = allParagraphs.Concat(mainPart.FootnotesPart.Footnotes.Descendants<Paragraph>());
        if (mainPart.EndnotesPart?.Endnotes != null)
            allParagraphs = allParagraphs.Concat(mainPart.EndnotesPart.Endnotes.Descendants<Paragraph>());
        if (mainPart.WordprocessingCommentsPart?.Comments != null)
            allParagraphs = allParagraphs.Concat(mainPart.WordprocessingCommentsPart.Comments.Descendants<Paragraph>());

        var paragraphs = allParagraphs.ToList();

        // Collect existing IDs, detect duplicates, and track max for deterministic increment
        var paraIdSeen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int maxId = 0;

        foreach (var para in paragraphs)
        {
            // Fix duplicate paraId: if already seen, clear it so it gets reassigned below
            if (!string.IsNullOrEmpty(para.ParagraphId?.Value))
            {
                if (!paraIdSeen.Add(para.ParagraphId.Value))
                {
                    para.ParagraphId = null!; // duplicate — will be reassigned
                }
                else
                {
                    _usedParaIds.Add(para.ParagraphId.Value);
                    if (int.TryParse(para.ParagraphId.Value, System.Globalization.NumberStyles.HexNumber, null, out var numId) && numId > maxId)
                        maxId = numId;
                }
            }
            if (!string.IsNullOrEmpty(para.TextId?.Value))
            {
                _usedParaIds.Add(para.TextId.Value);
                if (int.TryParse(para.TextId.Value, System.Globalization.NumberStyles.HexNumber, null, out var numId) && numId > maxId)
                    maxId = numId;
            }
        }

        // Also collect existing revision w:id values (w:ins/w:del/w:moveFrom/
        // w:moveTo/w:rPrChange/w:pPrChange/w:sectPrChange/w:tblPrChange and
        // bare <w:ins>/<w:del> markers inside trPr/tcPr) into the same used-id
        // pool. Revision ids and paraIds live in different XML attribute slots
        // so there's no XML collision, but sharing one pool means
        // GenerateRevisionId() never picks a number that's already taken by
        // either paraId or another revision id. Decimal revision ids are
        // converted to the same 8-char hex form the pool uses.
        foreach (var rootElem in EnumerateRevisionRoots(mainPart))
        {
            foreach (var elem in rootElem.Descendants())
            {
                if (elem is InsertedRun or DeletedRun or MoveFromRun or MoveToRun
                    or RunPropertiesChange or ParagraphPropertiesChange
                    or SectionPropertiesChange or TablePropertiesChange
                    or TableCellPropertiesChange or TableRowPropertiesChange
                    or Inserted or Deleted or MoveFrom or MoveTo)
                {
                    var idAttr = elem.GetAttributes()
                        .FirstOrDefault(a => a.LocalName == "id");
                    if (idAttr.Value != null
                        && int.TryParse(idAttr.Value, System.Globalization.NumberStyles.Integer,
                            System.Globalization.CultureInfo.InvariantCulture, out var rid))
                    {
                        _usedParaIds.Add(rid.ToString("X8"));
                        if (rid > maxId) maxId = rid;
                    }
                }
            }
        }

        // Start deterministic increment from max+1, minimum 0x100000 to avoid conflicts with small IDs
        const int minStartId = 0x100000;
        _nextParaId = Math.Max(maxId + 1, minStartId);
        if (_nextParaId > 0x7FFFFFFF) _nextParaId = minStartId;

        // Assign IDs to paragraphs that don't have them (including cleared duplicates)
        foreach (var para in paragraphs)
        {
            if (string.IsNullOrEmpty(para.ParagraphId?.Value))
                para.ParagraphId = GenerateParaId();
            if (string.IsNullOrEmpty(para.TextId?.Value))
                para.TextId = GenerateParaId();
        }

        // Reassign accidentally-duplicated revision w:id values (runs after
        // _nextParaId is seeded so it draws from the same global allocator).
        DedupeRevisionIds(mainPart);

        // Ensure mc:Ignorable includes "w14" so Word 2007 skips w14:paraId/textId attributes
        var doc = mainPart.Document;
        const string mcNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        if (doc.LookupNamespace("mc") == null)
            doc.AddNamespaceDeclaration("mc", mcNs);
        if (doc.LookupNamespace("w14") == null)
            doc.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
        var ignorable = doc.MCAttributes?.Ignorable?.Value ?? "";
        if (!ignorable.Contains("w14"))
        {
            doc.MCAttributes ??= new DocumentFormat.OpenXml.MarkupCompatibilityAttributes();
            doc.MCAttributes.Ignorable = string.IsNullOrEmpty(ignorable) ? "w14" : $"{ignorable} w14";
        }
    }

    /// <summary>
    /// Reassign accidentally-duplicated revision <c>w:id</c> values. Word requires
    /// every revision marker's id to be unique EXCEPT a moveFrom/moveTo pair,
    /// which shares one id by design. On a dump→batch round-trip the content-run
    /// ins/del ids are regenerated and can collide with a paragraph-mark id
    /// preserved verbatim from the source (or with each other), producing
    /// "id should have unique value" schema errors. Pin every move-pair id first
    /// so a colliding non-move marker yields, then renumber the remaining
    /// duplicates from the shared paraId/revision allocator. Mirrors the
    /// duplicate sweep in <see cref="EnsureDocPropIds"/>; move ids are never
    /// touched so the pairing survives.
    /// </summary>
    private void DedupeRevisionIds(MainDocumentPart mainPart)
    {
        static bool IsMoveType(OpenXmlElement e) =>
            e is MoveFrom or MoveTo or MoveFromRun or MoveToRun;
        static bool IsRevisionMarker(OpenXmlElement e) =>
            e is InsertedRun or DeletedRun or MoveFromRun or MoveToRun
              or RunPropertiesChange or ParagraphPropertiesChange
              or SectionPropertiesChange or TablePropertiesChange
              or TableCellPropertiesChange or TableRowPropertiesChange
              or Inserted or Deleted or MoveFrom or MoveTo;

        var revElems = EnumerateRevisionRoots(mainPart)
            .SelectMany(r => r.Descendants())
            .Where(IsRevisionMarker)
            .ToList();

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        // Pass 1: pin move-pair ids (a moveFrom and its moveTo legitimately
        // share one id — never renumber these).
        foreach (var e in revElems)
        {
            if (!IsMoveType(e)) continue;
            var v = e.GetAttributes().FirstOrDefault(a => a.LocalName == "id").Value;
            if (!string.IsNullOrEmpty(v)) seen.Add(v);
        }
        // Pass 2: renumber non-move duplicates to a fresh allocator value.
        foreach (var e in revElems)
        {
            if (IsMoveType(e)) continue;
            var attr = e.GetAttributes().FirstOrDefault(a => a.LocalName == "id");
            if (attr.Value == null) continue;
            if (seen.Add(attr.Value)) continue; // first occurrence — keep
            var newId = GenerateRevisionId();
            e.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute(
                attr.Prefix, attr.LocalName, attr.NamespaceUri, newId));
            seen.Add(newId);
        }
    }

    /// <summary>
    /// Build a map from every <c>w:moveFrom</c>/<c>w:moveTo</c> run's own
    /// <c>w:id</c> to a single SHARED pairing id, grouping a moveFrom and its
    /// corresponding moveTo by the <c>w:name</c> their bracketing range markers
    /// share (<c>moveFromRangeStart</c>/<c>moveToRangeStart</c> carry the same
    /// <c>w:name</c> — ECMA-376 §17.13.5.20-23). In the source XML the moveFrom
    /// run and the moveTo run usually have DIFFERENT <c>w:id</c> values
    /// (e.g. 4 and 6); the pairing lives only on the range-marker name. The CLI
    /// representation of a move pair is instead a SHARED <c>revision.id</c> on
    /// both halves (so <see cref="WrapRunAsMoveFrom"/>'s <c>Move_{id}</c> marker
    /// name matches across the pair). The dump emitter consults this map to
    /// rewrite both halves to one id, restoring the four range markers + shared
    /// name on dump→batch replay. Returns an empty map when there are no
    /// bracketed move pairs.
    /// </summary>
    internal Dictionary<string, string> BuildMovePairIdMap()
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var mainPart = _doc.MainDocumentPart;
        if (mainPart == null) return result;

        // name → list of move-run ids bracketed by a range with that name.
        var byName = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

        foreach (var root in EnumerateRevisionRoots(mainPart))
        {
            // Active range-marker names keyed by their range id (a
            // moveFrom and moveTo range nest independently but share the
            // same w:name; track the set of names currently open).
            var openNames = new List<string>();
            foreach (var el in root.Descendants())
            {
                switch (el)
                {
                    case MoveFromRangeStart mfs when !string.IsNullOrEmpty(mfs.Name?.Value):
                        openNames.Add(mfs.Name!.Value!);
                        break;
                    case MoveToRangeStart mts when !string.IsNullOrEmpty(mts.Name?.Value):
                        openNames.Add(mts.Name!.Value!);
                        break;
                    case MoveFromRangeEnd:
                    case MoveToRangeEnd:
                        // Close the most recently opened range. The range
                        // marker End carries the range id (not the name), so
                        // pop the latest open name — move ranges don't
                        // interleave across the same name in practice.
                        if (openNames.Count > 0) openNames.RemoveAt(openNames.Count - 1);
                        break;
                    case MoveFromRun mf when openNames.Count > 0 && !string.IsNullOrEmpty(mf.Id?.Value):
                        AddMoveRunId(byName, openNames[^1], mf.Id!.Value!);
                        break;
                    case MoveToRun mt when openNames.Count > 0 && !string.IsNullOrEmpty(mt.Id?.Value):
                        AddMoveRunId(byName, openNames[^1], mt.Id!.Value!);
                        break;
                }
            }
        }

        foreach (var (_, ids) in byName)
        {
            if (ids.Count == 0) continue;
            // Pick the lexicographically-smallest numeric id as the shared
            // pairing id (deterministic; mirrors the "lowest free value"
            // convention used elsewhere for id reassignment).
            var shared = ids
                .OrderBy(v => long.TryParse(v, out var n) ? n : long.MaxValue)
                .ThenBy(v => v, StringComparer.Ordinal)
                .First();
            foreach (var id in ids) result[id] = shared;
        }
        return result;
    }

    private static void AddMoveRunId(Dictionary<string, List<string>> byName, string name, string id)
    {
        if (!byName.TryGetValue(name, out var list))
            byName[name] = list = new List<string>();
        if (!list.Contains(id, StringComparer.OrdinalIgnoreCase)) list.Add(id);
    }

    // ==================== SDT IDs (content controls) ====================

    /// <summary>
    /// Generate a deterministic unique SdtId by scanning max existing value + 1.
    /// </summary>
    private int NextSdtId()
    {
        const int overflowReset = 872011;
        int maxId = 0;
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body != null)
        {
            foreach (var sdtId in body.Descendants<SdtId>())
            {
                if (sdtId.Val?.HasValue == true && sdtId.Val.Value > maxId)
                    maxId = sdtId.Val.Value;
            }
        }
        var next = maxId + 1;
        return next > int.MaxValue - 1 ? overflowReset : next;
    }

    // ==================== DocPr IDs (pictures, charts) ====================

    /// <summary>
    /// Ensure unique ids for all drawing-object non-visual properties
    /// (<c>&lt;wp:docPr&gt;</c>, the SDK's DW.DocProperties) — the single id
    /// space shared by pictures, charts, textboxes and shapes. NOT file
    /// metadata (docProps/core.xml etc.) despite "DocProp" in the name, and
    /// NOT the nested <c>pic:cNvPr</c>/<c>wps:cNvPr</c> ids (those deliberately
    /// mirror their wrapping docPr and need only intra-group uniqueness — see
    /// CreateImageRun, which writes docPr.id == cNvPr.id).
    /// Scans body + headers + footers; reassigns duplicate/missing ids to the
    /// lowest free value. Called on document open AND after every successful
    /// RawSet — see CONSISTENCY(docpr-global-uniqueness) in WordHandler.RawSet.
    /// </summary>
    private void EnsureDocPropIds()
    {
        var mainPart = _doc.MainDocumentPart;
        if (mainPart?.Document?.Body == null) return;

        var allDocProps = mainPart.Document.Body.Descendants<DW.DocProperties>().ToList();

        foreach (var headerPart in mainPart.HeaderParts)
            if (headerPart.Header != null)
                allDocProps.AddRange(headerPart.Header.Descendants<DW.DocProperties>());
        foreach (var footerPart in mainPart.FooterParts)
            if (footerPart.Footer != null)
                allDocProps.AddRange(footerPart.Footer.Descendants<DW.DocProperties>());

        var usedIds = new HashSet<uint>();
        var duplicates = new List<DW.DocProperties>();

        foreach (var dp in allDocProps)
        {
            if (dp.Id?.HasValue == true && !usedIds.Add(dp.Id.Value))
                duplicates.Add(dp);
            else if (dp.Id?.HasValue != true)
                duplicates.Add(dp);
        }

        foreach (var dp in duplicates)
        {
            uint newId = 1;
            while (!usedIds.Add(newId)) newId++;
            dp.Id = newId;
        }
    }

    /// <summary>
    /// Ensure all structured-document-tag ids (<c>w:sdtPr/w:id</c>) are unique.
    /// Called on document open and after every raw mutation — the sibling of
    /// <see cref="EnsureDocPropIds"/> / EnsureAllParaIds. NextSdtId allocates
    /// max+1 for typed adds, but a raw-set can inject an sdt whose id collides
    /// with an existing one; without this dedup the duplicate would land on
    /// disk (an sdt id collision breaks content-control targeting in Word).
    /// </summary>
    private void EnsureSdtIds()
    {
        var mainPart = _doc.MainDocumentPart;
        if (mainPart?.Document?.Body == null) return;

        var allSdtIds = mainPart.Document.Body.Descendants<SdtId>().ToList();
        foreach (var headerPart in mainPart.HeaderParts)
            if (headerPart.Header != null)
                allSdtIds.AddRange(headerPart.Header.Descendants<SdtId>());
        foreach (var footerPart in mainPart.FooterParts)
            if (footerPart.Footer != null)
                allSdtIds.AddRange(footerPart.Footer.Descendants<SdtId>());

        var usedIds = new HashSet<int>();
        var duplicates = new List<SdtId>();

        foreach (var sid in allSdtIds)
        {
            if (sid.Val?.HasValue == true && !usedIds.Add(sid.Val.Value))
                duplicates.Add(sid);
            else if (sid.Val?.HasValue != true)
                duplicates.Add(sid);
        }

        foreach (var sid in duplicates)
        {
            int newId = 1;
            while (!usedIds.Add(newId)) newId++;
            sid.Val = newId;
        }
    }
}
