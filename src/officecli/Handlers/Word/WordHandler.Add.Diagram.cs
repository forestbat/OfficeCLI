// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core.Diagram;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // A 'diagram' is an ADD-only synthesizer (like 'equation'): it parses
    // mermaid text, lays out a graph via the shared format-agnostic engine
    // (Core/Diagram), and expands into native, editable drawing shapes +
    // connectors in the body. It is deliberately NOT a persistent element —
    // after Add it is a set of ordinary <w:drawing> shapes, so it has no
    // matching Set/Get/Query on a "diagram" node (documented exception to the
    // Add-and-Set feature checklist). The parse + layout are shared with the
    // pptx emitter; only this mapping onto docx DrawingML differs. The one
    // format-specific concern vs pptx: docx has no slide to resize, so the
    // diagram is scaled to fit the section's text-area width (never enlarged),
    // and all shapes are floating anchors positioned relative to the margin.
    private const double DiagramCmToEmu = 360000.0;

    private string AddDiagram(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        // Input mirrors `equation` / the pptx diagram: canonical `mermaid`
        // (+ aliases text/dsl) inline, or `src`/`path` to a .mmd file.
        var mermaidText = properties.GetValueOrDefault("mermaid")
                          ?? properties.GetValueOrDefault("text")
                          ?? properties.GetValueOrDefault("dsl");
        if (string.IsNullOrWhiteSpace(mermaidText)
            && (properties.TryGetValue("src", out var srcFile) || properties.TryGetValue("path", out srcFile))
            && !string.IsNullOrWhiteSpace(srcFile))
        {
            if (!System.IO.File.Exists(srcFile))
                throw new ArgumentException($"diagram source file not found: '{srcFile}'.");
            mermaidText = System.IO.File.ReadAllText(srcFile);
        }
        if (string.IsNullOrWhiteSpace(mermaidText))
            throw new ArgumentException("diagram requires inline 'mermaid' text (aliases: text, dsl) or a 'src' .mmd file path.");

        // render mode: native (built-in editable shapes) | image (real mermaid.js in
        // a headless browser → embedded PNG, covers EVERY mermaid type at full
        // fidelity) | auto (default: image when a browser is available, else native).
        var renderMode = (properties.GetValueOrDefault("render") ?? "auto").Trim().ToLowerInvariant();
        bool forceImage = renderMode is "image" or "svg" or "browser";
        if (forceImage && !MermaidImageRenderer.IsAvailable())
            throw new ArgumentException(
                "render=image needs mermaid-cli (mmdc) or a headless browser (Chrome/Chromium/Edge). "
                + "Install one, or use render=native for the built-in synthesizer.");
        bool wantImage = forceImage
            || (renderMode is not ("native" or "shapes") && MermaidImageRenderer.IsAvailable());
        if (wantImage)
            return AddDiagramAsImage(parent, parentPath, index, properties, mermaidText, allowNativeFallback: !forceImage);

        return AddDiagramNative(parent, parentPath, index, properties, mermaidText);
    }

    // Built-in synthesizer: mermaid → laid-out graph → native <w:drawing> shapes.
    private string AddDiagramNative(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties, string mermaidText)
    {
        var lo = DiagramCompiler.Compile(mermaidText);
        if (lo.Nodes.Count == 0)
            throw new ArgumentException("diagram parsed to zero nodes — check the mermaid syntax.");

        var (host, hostRoot) = ResolveDrawingHost(parent, parentPath);

        // Fit-to-box: docx has no slide to resize (no pptx-style poster), so the
        // diagram always scales into the available space. `width`/`height` give an
        // explicit box (may enlarge, mirroring picture/chart); with neither, fit the
        // section's text-area width and never enlarge (keeps small graphs readable).
        double natW = lo.SlideWidthCm, natH = lo.SlideHeightCm;
        double contentCm = SectionContentWidthCm();
        bool hasW = properties.TryGetValue("width", out var wStr);
        bool hasH = properties.TryGetValue("height", out var hStr);
        double scale;
        if (hasW || hasH)
        {
            double boxW = hasW ? ParseEmu(wStr) / DiagramCmToEmu : contentCm;
            double boxH = hasH ? ParseEmu(hStr) / DiagramCmToEmu : double.PositiveInfinity;
            scale = natW > 0.01 ? Math.Min(boxW / natW, boxH / natH) : 1.0;
        }
        else
        {
            scale = natW > 0.01 ? Math.Min(1.0, contentCm / natW) : 1.0;
        }
        long Emu(double cm) => (long)Math.Round(cm * scale * DiagramCmToEmu);
        // Font scales WITH the box (the layout sized every box to hold its text at
        // the base point size, so any uniform scale keeps text fitting). Floor at 1
        // only to avoid a 0pt run — a fixed higher floor (e.g. 6) forces the font
        // LARGER than the shrunken box on a heavily fit-scaled wide diagram →
        // overflow/mid-word wrap (the "text too big for the box" symptom). The node
        // bodyPr's normAutofit shrinks further if a rounding edge still overflows.
        int fontPt = Math.Max(1, (int)Math.Round(18 * lo.FontScale * scale));
        int labelPt = Math.Max(1, (int)Math.Round(10 * scale));

        var sb = new StringBuilder();
        int z = 0; // relativeHeight — nodes behind, edges above, labels on top
        // Drawings aren't in the document yet, so NextDocPropId() would return
        // the same value for each; allocate one base and increment locally so
        // every wp:docPr/@id is unique (else Word's "id must be unique" error).
        uint nextId = NextDocPropId();

        // nodes (lowest z)
        foreach (var n in lo.Nodes)
        {
            var (geom, fill, line) = DiagramStyles.ByShape[n.Shape];
            sb.Append(BuildDiagramNodeXml(nextId++, z++, geom, fill, line, n.Label, fontPt,
                Emu(n.X), Emu(n.Y), Emu(n.W), Emu(n.H)));
        }

        // edges: one custGeom polyline per edge (arrow on the last point)
        foreach (var e in lo.Edges)
        {
            if (e.Points.Count < 2) continue;
            sb.Append(BuildDiagramEdgeXml(nextId++, z++, e.Points, e.ArrowAtEnd, e.Dashed, Emu));
        }

        // edge labels (highest z — white masks sit on top of the lines)
        foreach (var lbl in lo.Labels)
        {
            double w = Math.Max(1.0, DiagramLabelWidthCm(lbl.Text));
            // Opaque (flowchart) labels mask the edge line; sequence labels sit in
            // empty space → no fill, so they don't break the lifeline they cross.
            sb.Append(BuildDiagramNodeXml(nextId++, z++, "rect", lbl.Opaque ? "FFFFFF" : null, null, lbl.Text, labelPt,
                Emu(lbl.Cx - w / 2), Emu(lbl.Cy - 0.26), Emu(w), Emu(0.52), label: true));
        }

        // All drawings live as runs in a single anchor paragraph (floating
        // anchors are positioned absolutely, so one host paragraph suffices).
        var para = new Paragraph();
        foreach (var drawingXml in SplitDrawings(sb.ToString()))
            para.AppendChild(new Run(ParseDrawingFromXml(drawingXml)));
        AssignParaId(para);
        InsertAtIndexOrAppend(host, para, index);

        // All drawings share one host paragraph; report it as a single shape
        // anchor (the diagram is add-only — users edit the resulting shapes).
        int shapeIdx = CountShapesInHost(host, para);
        return $"{hostRoot}/shape[{shapeIdx}]";
    }

    // High-fidelity path: render with the real mermaid.js (headless browser) to PNG
    // and embed it as a picture, stamping the source into alt-text so the diagram
    // travels in the file and is regenerable. In auto mode any render failure falls
    // back to the native synthesizer.
    private string AddDiagramAsImage(OpenXmlElement parent, string parentPath, int? index,
                                     Dictionary<string, string> properties, string mermaidText, bool allowNativeFallback)
    {
        string imgPath;
        try { imgPath = MermaidImageRenderer.RenderToPngFile(mermaidText); }
        catch when (allowNativeFallback) { return AddDiagramNative(parent, parentPath, index, properties, mermaidText); }
        try
        {
            var pic = new Dictionary<string, string>(properties);
            foreach (var k in new[] { "mermaid", "text", "dsl", "src", "path", "render", "poster" })
                pic.Remove(k);
            pic["src"] = imgPath;
            if (!(pic.TryGetValue("alt", out var a) && !string.IsNullOrEmpty(a)))
                pic["alt"] = MermaidImageRenderer.SourceTag + mermaidText;
            // Default sizing parity with the native path AND the pptx image path:
            // fit the aspect-correct PNG into the section's content BOX — both the
            // text-area width AND the available page height. Pinning width alone
            // (the old default) blew a tall/portrait diagram past the page bottom,
            // spilling one flowchart across several pages. Never enlarge; caller's
            // explicit width/height still wins.
            if (!pic.ContainsKey("width") && !pic.ContainsKey("height"))
            {
                double boxWCm = SectionContentWidthCm();
                double outWCm = boxWCm; // fallback: dims unreadable → old width-fit
                using (var s = System.IO.File.OpenRead(imgPath))
                {
                    var dims = OfficeCli.Core.ImageSource.TryGetDimensions(s);
                    if (dims is { Width: > 0, Height: > 0 } d)
                    {
                        double boxHCm = SectionContentHeightCm();
                        double fit = Math.Min(boxWCm / d.Width, boxHCm / d.Height);
                        outWCm = d.Width * fit; // aspect preserved by keeping height implicit
                    }
                }
                pic["width"] = outWCm.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) + "cm";
            }
            return AddPicture(parent, parentPath, index, pic);
        }
        finally { try { System.IO.File.Delete(imgPath); } catch { /* best effort */ } }
    }

    // Each drawing XML is a complete <w:drawing …>…</w:drawing>; split on the
    // closing tag so each becomes its own Run (ParseDrawingFromXml wants one).
    private static IEnumerable<string> SplitDrawings(string concatenated)
    {
        const string close = "</w:drawing>";
        int i = 0;
        while (true)
        {
            int end = concatenated.IndexOf(close, i, StringComparison.Ordinal);
            if (end < 0) yield break;
            yield return concatenated.Substring(i, end - i + close.Length);
            i = end + close.Length;
        }
    }

    private const string DiagramNs =
        "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
        "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
        "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"";

    private static string BuildDiagramNodeXml(uint docPropId, int z, string preset, string fill, string? line,
                                              string text, int fontPt, long x, long y, long cx, long cy,
                                              bool label = false)
    {
        string fillXml = string.IsNullOrEmpty(fill)
            ? "<a:noFill/>"
            : $"<a:solidFill><a:srgbClr val=\"{fill}\"/></a:solidFill>";
        string lnXml = string.IsNullOrEmpty(line)
            ? "<a:ln><a:noFill/></a:ln>"
            : $"<a:ln w=\"9525\"><a:solidFill><a:srgbClr val=\"{line}\"/></a:solidFill></a:ln>";
        int szHalfPt = fontPt * 2;
        // rFonts with an eastAsia slot so Word resolves CJK glyphs. Without it a
        // textbox run inherits the Latin default (Calibri) and East-Asian text can
        // render blank; PowerPoint auto-applies the theme's CJK font, Word doesn't.
        const string rFonts = "<w:rFonts w:eastAsia=\"SimSun\" w:hint=\"eastAsia\"/>";
        string txbx =
            "<wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr>" +
            $"<w:r><w:rPr>{rFonts}<w:sz w:val=\"{szHalfPt}\"/><w:szCs w:val=\"{szHalfPt}\"/></w:rPr>" +
            $"<w:t xml:space=\"preserve\">{SecurityElement.Escape(text)}</w:t></w:r></w:p></w:txbxContent></wps:txbx>";
        // Edge labels: single-line, no wrap, zero insets so the (fit-scaled,
        // short) box never clips the text. Nodes: wrap inside the box, centered.
        string bodyPr = label
            ? "<wps:bodyPr rot=\"0\" wrap=\"none\" lIns=\"0\" tIns=\"0\" rIns=\"0\" bIns=\"0\" anchor=\"ctr\" anchorCtr=\"1\"><a:noAutofit/></wps:bodyPr>"
            : "<wps:bodyPr rot=\"0\" anchor=\"ctr\" anchorCtr=\"0\"><a:normAutofit/></wps:bodyPr>";
        return
            $"<w:drawing {DiagramNs}><wp:anchor distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\" simplePos=\"0\" " +
            $"relativeHeight=\"{2510000 + z}\" behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\">" +
            "<wp:simplePos x=\"0\" y=\"0\"/>" +
            $"<wp:positionH relativeFrom=\"margin\"><wp:posOffset>{x}</wp:posOffset></wp:positionH>" +
            $"<wp:positionV relativeFrom=\"margin\"><wp:posOffset>{y}</wp:posOffset></wp:positionV>" +
            $"<wp:extent cx=\"{cx}\" cy=\"{cy}\"/><wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/><wp:wrapNone/>" +
            $"<wp:docPr id=\"{docPropId}\" name=\"DiagramShape {docPropId}\"/><wp:cNvGraphicFramePr/>" +
            "<a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">" +
            "<wps:wsp><wps:cNvSpPr/><wps:spPr>" +
            $"<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>" +
            $"<a:prstGeom prst=\"{preset}\"><a:avLst/></a:prstGeom>{fillXml}{lnXml}</wps:spPr>" +
            $"{txbx}{bodyPr}" +
            "</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>";
    }

    private static string BuildDiagramEdgeXml(uint docPropId, int z, IReadOnlyList<Core.Diagram.Pt> points,
                                              bool arrowAtEnd, bool dashed, Func<double, long> emu)
    {
        // Map points to EMU, then pad the thin axis so a purely axis-aligned
        // segment (cx==0 or cy==0) still has a non-degenerate bounding box that
        // Word will render.
        long[] px = points.Select(p => emu(p.X)).ToArray();
        long[] py = points.Select(p => emu(p.Y)).ToArray();
        long minX = px.Min(), maxX = px.Max(), minY = py.Min(), maxY = py.Max();
        long w = maxX - minX, h = maxY - minY;
        const long pad = 12700; // 1pt
        long shiftX = 0, shiftY = 0;
        if (w < pad) { shiftX = (pad - w) / 2; minX -= shiftX; w = pad; }
        if (h < pad) { shiftY = (pad - h) / 2; minY -= shiftY; h = pad; }

        var path = new StringBuilder();
        for (int i = 0; i < points.Count; i++)
        {
            long x = px[i] - minX, y = py[i] - minY;
            path.Append(i == 0
                ? $"<a:moveTo><a:pt x=\"{x}\" y=\"{y}\"/></a:moveTo>"
                : $"<a:lnTo><a:pt x=\"{x}\" y=\"{y}\"/></a:lnTo>");
        }

        string dash = dashed ? "<a:prstDash val=\"dash\"/>" : "";
        string arrow = arrowAtEnd ? "<a:tailEnd type=\"triangle\"/>" : "";
        string ln = $"<a:ln w=\"12700\" cap=\"flat\"><a:solidFill><a:srgbClr val=\"{DiagramStyles.EdgeColor}\"/></a:solidFill>{dash}<a:round/>{arrow}</a:ln>";
        string custGeom =
            $"<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l=\"0\" t=\"0\" r=\"{w}\" b=\"{h}\"/>" +
            $"<a:pathLst><a:path w=\"{w}\" h=\"{h}\">{path}</a:path></a:pathLst></a:custGeom>";
        return
            $"<w:drawing {DiagramNs}><wp:anchor distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\" simplePos=\"0\" " +
            $"relativeHeight=\"{2510000 + z}\" behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\">" +
            "<wp:simplePos x=\"0\" y=\"0\"/>" +
            $"<wp:positionH relativeFrom=\"margin\"><wp:posOffset>{minX}</wp:posOffset></wp:positionH>" +
            $"<wp:positionV relativeFrom=\"margin\"><wp:posOffset>{minY}</wp:posOffset></wp:positionV>" +
            $"<wp:extent cx=\"{w}\" cy=\"{h}\"/><wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/><wp:wrapNone/>" +
            $"<wp:docPr id=\"{docPropId}\" name=\"DiagramEdge {docPropId}\"/><wp:cNvGraphicFramePr/>" +
            "<a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">" +
            "<wps:wsp><wps:cNvSpPr/><wps:spPr>" +
            $"<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{w}\" cy=\"{h}\"/></a:xfrm>{custGeom}<a:noFill/>{ln}</wps:spPr>" +
            "<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>";
    }

    private static double DiagramLabelWidthCm(string text)
    {
        double w = 0;
        foreach (var c in text) w += c > 0x2E80 ? 0.58 : 0.30;
        return Math.Min(w, 5.0) + 0.4;
    }

    // Text-area width (page width − left/right margins) of the last section, in
    // cm. Falls back to US-Letter with 1in margins (~16.51cm) when unset.
    private double SectionContentWidthCm()
    {
        try
        {
            var body = _doc?.MainDocumentPart?.Document?.Body;
            var sect = body?.Elements<SectionProperties>().LastOrDefault()
                       ?? body?.Descendants<SectionProperties>().LastOrDefault();
            var pgSz = sect?.Elements<PageSize>().FirstOrDefault();
            var pgMar = sect?.Elements<PageMargin>().FirstOrDefault();
            long wTw = pgSz?.Width?.Value ?? 12240u;
            long lTw = pgMar?.Left?.Value ?? 1440u;
            long rTw = pgMar?.Right?.Value ?? 1440u;
            double contentCm = (wTw - lTw - rTw) / 1440.0 * 2.54;
            return contentCm > 1.0 ? contentCm : 16.51;
        }
        catch { return 16.51; }
    }

    // Text-area height (page height − top/bottom margins) of the last section, in
    // cm. Falls back to US-Letter with 1in margins (~24.13cm) when unset. Used to
    // cap a fit-to-page diagram image so a tall graph stays on one page.
    private double SectionContentHeightCm()
    {
        try
        {
            var body = _doc?.MainDocumentPart?.Document?.Body;
            var sect = body?.Elements<SectionProperties>().LastOrDefault()
                       ?? body?.Descendants<SectionProperties>().LastOrDefault();
            var pgSz = sect?.Elements<PageSize>().FirstOrDefault();
            var pgMar = sect?.Elements<PageMargin>().FirstOrDefault();
            long hTw = pgSz?.Height?.Value ?? 15840u;
            long tTw = (long)(pgMar?.Top?.Value ?? 1440);
            long bTw = (long)(pgMar?.Bottom?.Value ?? 1440);
            double contentCm = (hTw - tTw - bTw) / 1440.0 * 2.54;
            return contentCm > 1.0 ? contentCm : 24.13;
        }
        catch { return 24.13; }
    }
}
