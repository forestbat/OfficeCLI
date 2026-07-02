// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // ==================== Dump support ====================
    //
    // CONSISTENCY(emit-X-mirror): read-only enumeration surface consumed by
    // ExcelBatchEmitter, mirroring the public helper methods PowerPointHandler
    // grew for PptxBatchEmitter (GetSlideBulletImageParts, GetTimingAudioRels,
    // ...). The emitter lives outside the handler's partial-class family, so
    // everything it needs beyond Get/Query is exposed here.

    /// <summary>Sheet names in workbook (sldIdLst-equivalent) order.</summary>
    public List<string> GetDumpSheetNames() => GetWorksheets().Select(t => t.Name).ToList();

    /// <summary>
    /// Workbook-level settings node (date1904, calc.*, activeTab, protection).
    /// Same Format keys as PopulateWorkbookSettings emits on Get.
    /// </summary>
    public DocumentNode GetDumpWorkbookNode()
    {
        var node = new DocumentNode { Path = "/workbook", Type = "workbook" };
        PopulateWorkbookSettings(node);
        return node;
    }

    /// <summary>
    /// Enumerate every row of a sheet with ALL cells that carry content OR
    /// style. The bulk Get path (GetSheetChildNodes) intentionally omits
    /// styled-empty cells (&lt;c s="1"/&gt;, issue #149 bloat guard); a dump
    /// must include them because their xf holds user-visible formatting
    /// (filled header bands, bordered empty grids). Each cell node is built
    /// by the same CellToNode Get uses, so Format keys match Get exactly.
    /// A dump-only <c>__raw</c> Format key carries the raw stored
    /// &lt;x:v&gt; text so the emitter can reproduce numbers/dates without
    /// going through display formatting.
    /// </summary>
    public List<DocumentNode> GetDumpRowNodes(string sheetName)
    {
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        var rows = new List<DocumentNode>();
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>();
        if (sheetData == null) return rows;

        // One evaluator per sheet: CellToNode lazily creates a fresh
        // FormulaEvaluator per formula cell when none is passed, which is
        // O(cells × sheet-size) on formula-heavy sheets.
        var eval = new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart);
        var seenRowIndices = new HashSet<uint>();
        foreach (var row in sheetData.Elements<Row>())
        {
            var ridx = row.RowIndex?.Value ?? 0;
            if (ridx != 0 && !seenRowIndices.Add(ridx)) continue;

            var rowNode = new DocumentNode
            {
                Path = $"/{sheetName}/row[{ridx}]",
                Type = "row"
            };
            // CONSISTENCY(unit-qualified-readback): pt-suffix row height,
            // mirroring GetSheetChildNodes.
            if (row.Height?.Value != null)
                rowNode.Format["height"] = $"{row.Height.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)}pt";
            if (row.Hidden?.Value == true)
                rowNode.Format["hidden"] = true;
            if (row.OutlineLevel?.Value is { } rol && rol > 0)
                rowNode.Format["outlineLevel"] = (int)rol;
            if (row.Collapsed?.Value == true)
                rowNode.Format["collapsed"] = true;

            foreach (var cell in row.Elements<Cell>())
            {
                var hasContent = CellHasContent(cell);
                var hasStyle = cell.StyleIndex != null && cell.StyleIndex.Value != 0;
                if (!hasContent && !hasStyle) continue;
                var cellNode = CellToNode(sheetName, cell, worksheet, eval);
                var raw = cell.CellValue?.Text;
                if (!string.IsNullOrEmpty(raw))
                    cellNode.Format["__raw"] = raw;
                // Rich-text carrier: inline-string runs or a rich shared-string
                // entry can't ride the CSV baseline. Serialize the runs into
                // the `runs=<json>` shape ApplyRichTextToCell consumes so the
                // emitter can replay via `set type=richtext runs=...`.
                if (HasRichTextContent(cell))
                    cellNode.Format["__richruns"] = SerializeRichTextRuns(cell);
                rowNode.Children.Add(cellNode);
            }

            if (rowNode.Children.Count == 0 && rowNode.Format.Count == 0) continue;
            rowNode.ChildCount = rowNode.Children.Count;
            rows.Add(rowNode);
        }
        return rows;
    }

    private bool HasRichTextContent(Cell cell)
    {
        var inline = cell.GetFirstChild<InlineString>();
        if (inline != null && inline.Elements<Run>().Any()) return true;
        return GetRichTextHost(cell)?.Elements<Run>().Any() == true;
    }

    private OpenXmlElement? GetRichTextHost(Cell cell)
    {
        var inline = cell.GetFirstChild<InlineString>();
        if (inline != null && inline.Elements<Run>().Any()) return inline;
        if (cell.DataType?.Value == CellValues.SharedString
            && int.TryParse(cell.CellValue?.Text, out var ssIdx))
        {
            var ssItems = _doc.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
            return ssItems?.Elements<SharedStringItem>().ElementAtOrDefault(ssIdx);
        }
        return null;
    }

    /// <summary>
    /// Serialize a rich-text cell's runs into the `runs=<json>` array
    /// ApplyRichTextToCell consumes. Key vocabulary mirrors RunToNode's Get
    /// output (bold/italic/strike/underline/superscript/subscript/size/color/
    /// font) — the Add side accepts each of these verbatim.
    /// </summary>
    private string SerializeRichTextRuns(Cell cell)
    {
        var host = GetRichTextHost(cell);
        var arr = new System.Text.Json.Nodes.JsonArray();
        if (host == null) return "[]";
        // A rich shared-string item may open with a bare <t> (unformatted
        // leading segment) before the first <r>; carry it as a plain run.
        var leading = host.GetFirstChild<Text>();
        if (leading != null && !string.IsNullOrEmpty(leading.Text))
            arr.Add(new System.Text.Json.Nodes.JsonObject { ["text"] = leading.Text });
        foreach (var run in host.Elements<Run>())
        {
            var o = new System.Text.Json.Nodes.JsonObject { ["text"] = run.Text?.Text ?? "" };
            var rp = run.RunProperties;
            if (rp != null)
            {
                if (rp.GetFirstChild<Bold>() != null) o["bold"] = true;
                if (rp.GetFirstChild<Italic>() != null) o["italic"] = true;
                if (rp.GetFirstChild<Strike>() != null) o["strike"] = true;
                var ul = rp.GetFirstChild<Underline>();
                if (ul != null) o["underline"] = ul.Val?.InnerText == "double" ? "double" : "single";
                var va = rp.GetFirstChild<VerticalTextAlignment>();
                if (va?.Val?.Value == VerticalAlignmentRunValues.Superscript) o["superscript"] = true;
                if (va?.Val?.Value == VerticalAlignmentRunValues.Subscript) o["subscript"] = true;
                var sz = rp.GetFirstChild<FontSize>();
                if (sz?.Val?.Value != null) o["size"] = $"{sz.Val.Value:0.##}pt";
                var clr = rp.GetFirstChild<Color>();
                if (clr?.Rgb?.Value != null) o["color"] = ParseHelpers.FormatHexColor(clr.Rgb.Value!);
                var rf = rp.GetFirstChild<RunFont>();
                if (rf?.Val?.Value != null) o["font"] = rf.Val.Value!;
            }
            arr.Add(o);
        }
        return arr.ToJsonString();
    }

    /// <summary>Per-sheet element counts for the indexed Get paths the batch
    /// emitter transcribes (cf[N] / validation[N] / comment[N] / table[N] /
    /// chart[N] / sparkline[N]).</summary>
    public (int Tables, int Cfs, int Validations, int Comments, int Charts, int Sparklines, bool HasExtendedChart)
        GetDumpElementCounts(string sheetName)
    {
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        var ws = GetSheet(worksheet);
        var tables = worksheet.TableDefinitionParts.Count();
        var cfs = ws.Elements<ConditionalFormatting>().Count();
        var validations = ws.GetFirstChild<DataValidations>()?.Elements<DataValidation>().Count() ?? 0;
        var comments = worksheet.WorksheetCommentsPart?.Comments?
            .GetFirstChild<CommentList>()?.Elements<Comment>().Count() ?? 0;
        // chart[N] Get indexes over GetExcelCharts (standard + extended in
        // drawing order); extended (chartEx) charts have no Add vocabulary.
        int charts = 0; bool hasExtended = false;
        if (worksheet.DrawingsPart is { } dp)
        {
            var chartInfos = GetExcelCharts(dp);
            charts = chartInfos.Count;
            hasExtended = chartInfos.Any(c => c.IsExtended);
        }
        int sparklines = 0;
        var extLst = ws.GetFirstChild<WorksheetExtensionList>();
        if (extLst != null)
            sparklines = extLst.Descendants<DocumentFormat.OpenXml.Office2010.Excel.SparklineGroup>().Count();
        return (tables, cfs, validations, comments, charts, sparklines, hasExtended);
    }

    /// <summary>
    /// Column definitions of a sheet, one node per column LETTER (Column
    /// elements with min/max ranges are expanded). Format keys mirror the
    /// /Sheet/col[X] Get surface: width, hidden, customWidth, outlineLevel,
    /// collapsed. Range expansion is capped per definition so a stray
    /// A:XFD-wide entry cannot emit 16 000 rows — the overflow is reported
    /// through <paramref name="truncated"/>.
    /// </summary>
    public List<DocumentNode> GetDumpColumnNodes(string sheetName, out bool truncated)
    {
        truncated = false;
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        var nodes = new List<DocumentNode>();
        var cols = GetSheet(worksheet).GetFirstChild<Columns>();
        if (cols == null) return nodes;

        const int MaxExpandPerDef = 256;
        foreach (var col in cols.Elements<Column>())
        {
            var min = (int)(col.Min?.Value ?? 0);
            var max = (int)(col.Max?.Value ?? 0);
            if (min < 1 || max < min) continue;
            var span = max - min + 1;
            if (span > MaxExpandPerDef)
            {
                truncated = true;
                max = min + MaxExpandPerDef - 1;
            }
            for (int i = min; i <= max; i++)
            {
                var letter = IndexToColumnName(i);
                var node = new DocumentNode
                {
                    Path = $"/{sheetName}/col[{letter}]",
                    Type = "column",
                    Preview = letter
                };
                if (col.Width?.Value != null && col.CustomWidth?.Value == true)
                    node.Format["width"] = col.Width.Value;
                if (col.Hidden?.Value == true) node.Format["hidden"] = true;
                if (col.OutlineLevel?.Value is { } ol && ol > 0)
                    node.Format["outlineLevel"] = (int)ol;
                if (col.Collapsed?.Value == true) node.Format["collapsed"] = true;
                if (node.Format.Count > 0) nodes.Add(node);
            }
        }
        return nodes;
    }

    /// <summary>
    /// All merged ranges of a sheet, straight from the worksheet's
    /// MergeCells element. Per-cell Format["merge"] readback cannot drive
    /// this — a merge whose cells are all empty and unstyled has no cell
    /// node at all in the dump row enumeration.
    /// </summary>
    public List<string> GetDumpMergeRanges(string sheetName)
    {
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        var merges = GetSheet(worksheet).GetFirstChild<MergeCells>();
        if (merges == null) return new List<string>();
        return merges.Elements<MergeCell>()
            .Select(m => m.Reference?.Value)
            .Where(r => !string.IsNullOrEmpty(r))
            .Select(r => r!)
            .ToList();
    }

    /// <summary>
    /// Resolve a dump subtree token ("SheetName" or "sheet[N]") to the
    /// canonical sheet name; null when it does not resolve.
    /// </summary>
    public string? ResolveDumpSheetName(string token)
    {
        try
        {
            var resolved = ResolveSheetName(token);
            return FindWorksheet(resolved) != null ? resolved : null;
        }
        catch { return null; }
    }

    /// <summary>Pivot-table locations on a sheet ("E1:H10" ranges). The dump
    /// value baseline EXCLUDES cells inside these ranges — they are derived
    /// render output that `add pivottable` regenerates on replay; importing
    /// them as static values would collide with the rebuilt pivot.</summary>
    public List<string> GetDumpPivotLocations(string sheetName)
    {
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        var result = new List<string>();
        foreach (var pp in worksheet.PivotTableParts)
        {
            var loc = pp.PivotTableDefinition?.Location?.Reference?.Value;
            if (!string.IsNullOrEmpty(loc)) result.Add(loc!);
        }
        return result;
    }

    /// <summary>Number of slicers on a sheet (slicer[N] index space).</summary>
    public int GetDumpSlicerCount(string sheetName)
    {
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        var extLst = GetSheet(worksheet).GetFirstChild<WorksheetExtensionList>();
        if (extLst == null) return 0;
        // Each <x14:slicer r:id> under the slicerList ext references one
        // SlicerPart; the slicer[N] Get index space walks the same list.
        return extLst.Descendants()
            .Count(e => e.LocalName == "slicer" && e.NamespaceUri.Contains("2009/9/main"));
    }

    /// <summary>Number of pivot tables on a sheet (pivottable[N] index space).</summary>
    public int GetDumpPivotCount(string sheetName)
    {
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        return worksheet.PivotTableParts.Count();
    }

    /// <summary>Drawing-layer counts (pictures / leaf shapes) matching the
    /// picture[N] / shape[N] Get index spaces.</summary>
    public (int Pictures, int Shapes) GetDumpDrawingCounts(string sheetName)
    {
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        var wsDrawing = worksheet.DrawingsPart?.WorksheetDrawing;
        if (wsDrawing == null) return (0, 0);
        var pictures = wsDrawing.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>()
            .Count(a => a.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>().Any());
        var shapes = EnumerateLeafShapes(wsDrawing).Count();
        return (pictures, shapes);
    }

    /// <summary>
    /// Extract picture[index]'s image bytes as a data URI for the emit's
    /// `add picture src=` prop (ImageSource.Resolve round-trips data URIs).
    /// SVG dual-representation aware: when the blip carries an
    /// asvg:svgBlip extension, the TRUE source is the SVG part (the r:embed
    /// PNG is just the fallback AddPicture regenerates on replay).
    /// </summary>
    public string? GetDumpPictureDataUri(string sheetName, int index)
    {
        var worksheet = FindWorksheet(sheetName);
        var drawingsPart = worksheet?.DrawingsPart;
        var wsDrawing = drawingsPart?.WorksheetDrawing;
        if (worksheet == null || drawingsPart == null || wsDrawing == null) return null;
        var picAnchors = wsDrawing.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>()
            .Where(a => a.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>().Any())
            .ToList();
        if (index < 1 || index > picAnchors.Count) return null;
        var picture = picAnchors[index - 1]
            .Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>().First();
        var blip = picture.BlipFill?.Blip;
        if (blip == null) return null;

        // SVG extension takes precedence over the PNG fallback embed.
        var svgBlip = blip.Descendants()
            .FirstOrDefault(e => e.LocalName == "svgBlip");
        var relId = svgBlip?.GetAttributes()
                .FirstOrDefault(a => a.LocalName == "embed").Value
            ?? blip.Embed?.Value;
        if (string.IsNullOrEmpty(relId)) return null;

        if (drawingsPart.GetPartById(relId!) is not ImagePart imgPart) return null;
        using var s = imgPart.GetStream();
        using var ms = new MemoryStream();
        s.CopyTo(ms);
        return $"data:{imgPart.ContentType};base64,{Convert.ToBase64String(ms.ToArray())}";
    }

    /// <summary>
    /// Cheap package-part scan for per-sheet content the batch emit does not
    /// round-trip yet. Mirrors PptxBatchEmitter's EmitAuxiliaryPartsScan
    /// contract: silent data loss is worse than a noisy warning.
    /// </summary>
    public List<(string Element, string Reason)> GetDumpUnsupportedFeatures(string sheetName)
    {
        var result = new List<(string, string)>();
        var worksheet = FindWorksheet(sheetName);
        if (worksheet == null) return result;
        var ws = GetSheet(worksheet);

        void AddIf(bool present, string element, string reason)
        {
            if (present) result.Add((element, reason));
        }

        // PR2-4 round-trip tables/cf/validations/comments/charts/sparklines/
        // pictures/shapes/pivots semantically (ExcelBatchEmitter.Elements.cs);
        // this scan covers only what still has no emit channel.
        AddIf(worksheet.EmbeddedObjectParts.Any() || worksheet.EmbeddedPackageParts.Any(), "ole",
            "embedded OLE objects are not round-tripped by dump yet");
        var extLst = ws.GetFirstChild<WorksheetExtensionList>();
        if (extLst != null)
        {
            AddIf(extLst.OuterXml.Contains("slicerList", StringComparison.OrdinalIgnoreCase), "slicer",
                "slicers are not round-tripped by dump yet");
        }
        return result;
    }
}
