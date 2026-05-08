using IntelliCAD.ApplicationServices;
using IntelliCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Teigha.Colors;
using Teigha.DatabaseServices;
using Teigha.Geometry;
using Teigha.Runtime;
using Application = IntelliCAD.ApplicationServices.Application;

namespace MagnaSoft_Drafting_Assistant_Tool_CobbFendley
{
    public class AutoDimRectCommand_Pauley
    {
        public struct LayerConfig
        {
            public string Suffix { get; }
            public int SuffixCharCount { get; }
            public LayerConfig(string suffix, int suffixChars)
            {
                Suffix = suffix;
                SuffixCharCount = suffixChars;
            }
        }

        private static readonly Dictionary<string, List<LayerConfig>> LayerSuffixMap =
            new Dictionary<string, List<LayerConfig>>(StringComparer.OrdinalIgnoreCase)
            {
                { "WATER",              new List<LayerConfig> { new LayerConfig("' WTR",          5) } },
                { "SEWER",              new List<LayerConfig> { new LayerConfig("' SWR",          5) } },
                { "BOC",                new List<LayerConfig> { new LayerConfig("' BOC",          5),
                                                                new LayerConfig("' FOSW",         5) } },
                { "SIDEWALK",           new List<LayerConfig> { new LayerConfig("' BOSW",         5) } },
                { "PROP_BORE",          new List<LayerConfig> { new LayerConfig("' PROP.R/L",     9) } },
                { "PROPOSED TRENCH",    new List<LayerConfig> { new LayerConfig("' PROP. TRENCH",14) } },
                { "PROP_TRENCH",     new List<LayerConfig> { new LayerConfig("' PROP. TRENCH", 14) } },
                { "EOP",                new List<LayerConfig> { new LayerConfig("' EOP",          5) } },
                { "ROW",                new List<LayerConfig> { new LayerConfig("' ROW",          5) } },
                { "U_CenturyLink_Line", new List<LayerConfig> { new LayerConfig("' CLN",          5) } },
                { "TELCO",              new List<LayerConfig> { new LayerConfig("' T&E&C",        7) } },
                { "GAS LAT",            new List<LayerConfig> { new LayerConfig("' GAS",          5) } },
                { "PROPERTY LINE",      new List<LayerConfig>() },
                { "VNAE LINE",          new List<LayerConfig>() },
                { "EASEMENT",           new List<LayerConfig>() },
            };

        private static readonly List<(string Src, string Tgt, string Sfx, int SfxChars)>
            RelationalPairs = new List<(string, string, string, int)>
            {
                ("PROPERTY LINE", "VNAE LINE", "' VNAE", 6),
                ("ROW",           "EASEMENT",  "' PUE",  5),
            };

        private static readonly string[] ReferenceLayers =
            { "PROPERTY LINE", "VNAE LINE", "EASEMENT" };

        private const string DIM_LAYER = "DIMENSIONS";
        private const string DIM_STYLE_NAME = "BPGDIMS";
        private const double STACK_GAP = 8.0;
        private const double BEYOND_OFFSET = 15.0;
        private const double RAY_BUFFER = 50.0;   //20.0;
        private const double CHAR_ASPECT = 0.6;
        private const double CURVE_SAMPLE_STEP = 2.0;   //5.0;
        private const double VERTICAL_SPACING = 4.2;

        // ─────────────────────────────────────────────────────────────
        // CHANGE-6: ACCEPTED CURVE TYPE GUARD
        //
        // Single place that lists every entity type this tool treats as
        // a curve.  All four types inherit from Curve, so all downstream
        // Curve API calls (GetPointAtParameter, GetClosestPointTo,
        // IntersectWith, GeometricExtents) work uniformly.
        //
        // To support additional types (Arc, Spline, …) in the future,
        // just add them here — no other method needs changing.
        // ─────────────────────────────────────────────────────────────
        private static bool IsAcceptedCurve(Entity ent) =>
            ent is Polyline
            || ent is Polyline2d
            || ent is Polyline3d
            || ent is Line;

        // ─────────────────────────────────────────────────────────────
        // MAIN COMMAND
        // ─────────────────────────────────────────────────────────────
        [CommandMethod("AUTODIMRECT_PAULEY")]
        public void AutoDimRect()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            double liveDimtxt = Convert.ToDouble(Application.GetSystemVariable("DIMTXT"));
            double liveDimasz = Convert.ToDouble(Application.GetSystemVariable("DIMASZ"));
            if (liveDimtxt <= 0) liveDimtxt = 3.5;
            if (liveDimasz <= 0) liveDimasz = 2.75;

            int originalDimassoc = Convert.ToInt32(Application.GetSystemVariable("DIMASSOC"));
            ed.Command("_.DIMASSOC", 0);
            db.DimAssoc = 0;

            ObjectId visualRectId = ObjectId.Null;

            try
            {
                Point3d[] corners = PickFourPoints(ed);
                if (corners == null) { ed.WriteMessage("\nAutoDimRect: cancelled."); return; }

                double minX = corners.Min(p => p.X);
                double maxX = corners.Max(p => p.X);
                double minY = corners.Min(p => p.Y);
                double maxY = corners.Max(p => p.Y);
                              

                Extents2d rect = new Extents2d(new Point2d(minX, minY), new Point2d(maxX, maxY));
                visualRectId = DrawVisualRectangle(db, rect);

                ed.WriteMessage($"\n  Rect: MinX={rect.MinPoint.X:F2} MinY={rect.MinPoint.Y:F2} " +
               $"MaxX={rect.MaxPoint.X:F2} MaxY={rect.MaxPoint.Y:F2} " +
               $"Width={rect.MaxPoint.X - rect.MinPoint.X:F2} " +
               $"Height={rect.MaxPoint.Y - rect.MinPoint.Y:F2}");

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        // CHANGE-5: accepts Line-based centrelines too.
                        Polyline cl = FindCenterlineInRect(db, tr, ed, rect);
                        if (cl != null)
                        {
                            ed.WriteMessage($"\n  CL found. Length={cl.Length:F2}  " +
                                            $"Vertices={cl.NumberOfVertices}  " +
                                            $"Start={cl.GetPoint3dAt(0)}  " +
                                            $"End={cl.GetPoint3dAt(cl.NumberOfVertices - 1)}");
                        }

                        if (cl == null)
                        {
                            ed.WriteMessage("\nNo CENTERLINE polyline or line found inside the selected rectangle.");
                            tr.Abort();
                            if (!visualRectId.IsNull) EraseVisualRectangle(db, visualRectId);
                            return;
                        }

                        // CHANGE-3: collects both Polyline and Line entities per layer.
                        var layerEntities = CollectLayerEntitiesInRect(db, tr, rect);

                        ed.WriteMessage("\n--- Layer Entity Counts ---");
                        foreach (var kvp in layerEntities)
                            ed.WriteMessage($"\n  [{kvp.Key}] : {kvp.Value.Count} curves found");

                        // Also dump ALL layer names found inside the rect:
                        ed.WriteMessage("\n--- All layers inside rect ---");
                        using (Transaction dbgTr = db.TransactionManager.StartTransaction())
                        {
                            BlockTable bt2 = dbgTr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                            BlockTableRecord ms2 = dbgTr.GetObject(
                                bt2[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                            var foundLayers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                            foreach (ObjectId id in ms2)
                            {
                                Entity ent = dbgTr.GetObject(id, OpenMode.ForRead) as Entity;
                                if (ent == null) continue;
                                try
                                {
                                    Extents3d ext = ent.GeometricExtents;
                                    //if (ext.MaxPoint.X < rect.MinPoint.X || ext.MinPoint.X > rect.MaxPoint.X ||
                                    //    ext.MaxPoint.Y < rect.MinPoint.Y || ext.MinPoint.Y > rect.MaxPoint.Y)
                                    //    continue;
                                    double bboxTol = 1.0;
                                    if (ext.MaxPoint.X < rect.MinPoint.X - bboxTol ||
                                        ext.MinPoint.X > rect.MaxPoint.X + bboxTol ||
                                        ext.MaxPoint.Y < rect.MinPoint.Y - bboxTol ||
                                        ext.MinPoint.Y > rect.MaxPoint.Y + bboxTol)
                                        continue;
                                }
                                catch { continue; }
                                foundLayers.Add(ent.Layer);
                            }
                            foreach (string lyr in foundLayers.OrderBy(s => s))
                                ed.WriteMessage($"\n  LAYER: [{lyr}]");
                            dbgTr.Abort();
                        }
                        ed.WriteMessage("\n--- Centerline check ---");


                        if (!layerEntities.Any(kvp => kvp.Value.Count > 0))
                            ed.WriteMessage("\nWarning: no recognised utility layers found inside the rectangle.");

                        EnsureLayerExists(db, tr);
                        ObjectId dimStyleId = GetOrCreateDimStyle(db, tr, liveDimtxt, liveDimasz);

                        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord mSpace = tr.GetObject(
                            bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                        // PRE-PASS
                        var allLeftDims = new List<UtilityDimData>();
                        var allRightDims = new List<UtilityDimData>();
                        var preLeft = new HashSet<string>();
                        var preRight = new HashSet<string>();
                        CollectDimCandidates(cl, layerEntities, rect,allLeftDims, allRightDims, preLeft, preRight, ed);


                        double uniformLeft = ComputeUniformDimLineDist(allLeftDims, liveDimtxt, BEYOND_OFFSET, ed);
                        double uniformRight = ComputeUniformDimLineDist(allRightDims, liveDimtxt, BEYOND_OFFSET, ed);
                        ed.WriteMessage($"\nUniform dim-line dist  LEFT : {uniformLeft:F2}  RIGHT: {uniformRight:F2}");

                        // PLACEMENT PASS
                        var placeLeftDims = new List<UtilityDimData>();
                        var placeRightDims = new List<UtilityDimData>();
                        var placeLeftSeen = new HashSet<string>();
                        var placeRightSeen = new HashSet<string>();
                        CollectDimCandidates(cl, layerEntities, rect,placeLeftDims, placeRightDims, placeLeftSeen, placeRightSeen, ed);

                        var placedSuffixesLeft = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        var placedSuffixesRight = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                        var leftSideItems = placeLeftDims.OrderBy(d => SafeGetDistAtPoint(cl, d.StationPt)).ToList();
                        var rightSideItems = placeRightDims.OrderBy(d => SafeGetDistAtPoint(cl, d.StationPt)).ToList();

                        const double VERTICAL_STEP = 4.25;
                        PlaceUniformSideDims(tr, mSpace, cl, leftSideItems, dimStyleId, true, uniformLeft, liveDimtxt, placedSuffixesLeft, VERTICAL_STEP, layerEntities);
                        PlaceUniformSideDims(tr, mSpace, cl, rightSideItems, dimStyleId, false, uniformRight, liveDimtxt, placedSuffixesRight, VERTICAL_STEP, layerEntities);

                        tr.Commit();
                        ed.WriteMessage("\nAUTODIMRECT complete.");
                        EraseVisualRectangle(db, visualRectId);
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\nError: {ex.Message}\n{ex.StackTrace}");
                        tr.Abort();
                    }
                }
            }
            finally
            {
                ed.Command("_.DIMASSOC", originalDimassoc);
                db.DimAssoc = originalDimassoc;
            }
        }

        // ─────────────────────────────────────────────────────────────
        // PLACE UNIFORM SIDE DIMS
        // ─────────────────────────────────────────────────────────────
        private void PlaceUniformSideDims(
            Transaction tr, BlockTableRecord mSpace, Polyline cl,
            List<UtilityDimData> items, ObjectId dimStyleId, bool isLeft,
            double uniformDimLineDist, double dimtxt,
            HashSet<string> placedSuffixes, double gap,
            Dictionary<string, List<Curve>> layerEntities)
        {
            if (!items.Any()) return;
            double charW = dimtxt * CHAR_ASPECT;
            int globalIndex = 0;
            double startDist = SafeGetDistAtPoint(cl, items.First().StationPt);

            foreach (var dimData in items)
            {
                if (LayerSuffixMap.TryGetValue(dimData.LayerName, out var configs) && configs.Count > 0)
                {
                    foreach (var cfg in configs)
                    {
                        bool placed = CreateUniformDimension(
                                        tr, mSpace, cl, dimData.IntersectPt, dimData.Distance,
                                        cfg.Suffix, cfg.SuffixCharCount, startDist, globalIndex, gap, isLeft,
                                        uniformDimLineDist, dimStyleId, charW, dimData.LayerName, placedSuffixes,
                                        layerEntities);   // ADD THIS
                        if (placed) globalIndex++;
                    }
                }

                foreach (var pair in RelationalPairs)
                {
                    if (!dimData.LayerName.Equals(pair.Src, StringComparison.OrdinalIgnoreCase)) continue;
                    if (placedSuffixes.Contains(pair.Sfx)) continue;

                    var target = items.FirstOrDefault(d =>
                        d.LayerName.Equals(pair.Tgt, StringComparison.OrdinalIgnoreCase)
                        && d.StationPt.DistanceTo(dimData.StationPt) < CURVE_SAMPLE_STEP);
                    if (target == null) continue;

                    double pairDist = Math.Abs(dimData.Distance - target.Distance);
                    bool placed;

                    if (pair.Src.Equals("ROW", StringComparison.OrdinalIgnoreCase)
                        && pair.Tgt.Equals("EASEMENT", StringComparison.OrdinalIgnoreCase))
                    {
                        placed = CreatePueDimension(tr, mSpace, cl,
                                                    dimData.IntersectPt, target.IntersectPt, pairDist,
                                                    pair.Sfx, pair.SfxChars, startDist, globalIndex, gap, isLeft,
                                                    uniformDimLineDist, dimStyleId, charW, placedSuffixes,
                                                    layerEntities);   // ADD THIS
                    }
                    // AFTER — replace that else block with:
                    else if (pair.Src.Equals("PROPERTY LINE", StringComparison.OrdinalIgnoreCase)
                          && pair.Tgt.Equals("VNAE LINE", StringComparison.OrdinalIgnoreCase))
                    {
                        // dimData = PROPERTY LINE item, target = VNAE LINE item
                        placed = CreateVnaeDimension(
                                                    tr, mSpace, cl,
                                                    dimData.IntersectPt, target.IntersectPt, pairDist,
                                                    pair.Sfx, pair.SfxChars, startDist, globalIndex, gap, isLeft,
                                                    uniformDimLineDist, dimStyleId, charW, placedSuffixes,
                                                    layerEntities);   // ADD THIS
                    }
                    else
                    {
                        placed = CreateUniformDimension(
                                                        tr, mSpace, cl, target.IntersectPt, pairDist,
                                                        pair.Sfx, pair.SfxChars, startDist, globalIndex, gap, isLeft,
                                                        uniformDimLineDist, dimStyleId, charW, pair.Tgt, placedSuffixes,
                                                        layerEntities);   // ADD THIS
                    }
                    //else
                    //{
                    //    placed = CreateUniformDimension(
                    //        tr, mSpace, cl, target.IntersectPt, pairDist,
                    //        pair.Sfx, pair.SfxChars, startDist, globalIndex, gap, isLeft,
                    //        uniformDimLineDist, dimStyleId, charW, pair.Tgt, placedSuffixes);
                    //}
                    if (placed) globalIndex++;
                }
            }
        }

        // ─────────────────────────────────────────────────────────────
        // CREATE UNIFORM DIMENSION
        // ─────────────────────────────────────────────────────────────
        private bool CreateUniformDimension(
    Transaction tr, BlockTableRecord mSpace, Polyline cl,
    Point3d intersectPt, double distance,
    string suffix, int suffixChars,
    double startDist, int index, double gap, bool isLeft,
    double uniformDimLineDist, ObjectId dimStyleId,
    double charW, string layerName, HashSet<string> placedSuffixes,
    Dictionary<string, List<Curve>> layerEntities)  // ADD THIS PARAMETER
        {
            if (placedSuffixes.Contains(suffix)) return false;

            double rungDist = Math.Min(startDist + (index * gap), cl.Length);
            Point3d rungPt = cl.GetPointAtDist(rungDist);
            Vector3d tangent = cl.GetFirstDerivative(rungPt).GetNormal();
            Vector3d perpDir = isLeft
                ? new Vector3d(-tangent.Y, tangent.X, 0)
                : new Vector3d(tangent.Y, -tangent.X, 0);

            // xLine1: snapped to centerline at rung point (already exact)
            Point3d xLine1 = cl.GetClosestPointTo(rungPt, false);

            // xLine2: snapped to the actual utility polyline/line
            Point3d snappedIntersect = GetSnappedPointOnCurve(intersectPt, layerName, layerEntities);
            Point3d xLine2 = ProjectToPerp(rungPt, perpDir, snappedIntersect);
            Point3d dimLinePt = xLine1 + (perpDir * uniformDimLineDist);

            int roundedDist = (int)Math.Round(distance, 0);
            string dimText = $"{roundedDist}{suffix}";

            AlignedDimension ad = new AlignedDimension(xLine1, xLine2, dimLinePt, string.Empty, dimStyleId);
            ad.DimensionText = dimText;
            ad.Layer = DIM_LAYER;
            ad.Color = Color.FromColorIndex(ColorMethod.ByLayer, 256);
            if (suffix == "' PROP.R/L")
            {
                ad.Dimclrt = Color.FromColorIndex(ColorMethod.ByLayer, 6);
            }
            mSpace.AppendEntity(ad);
            tr.AddNewlyCreatedDBObject(ad, true);

            int totalChars = roundedDist.ToString().Length + suffixChars;
            double textOffset = uniformDimLineDist - ((totalChars * charW) * 0.5);
            ad.TextPosition = xLine1 + (perpDir * textOffset);
            ad.TextRotation = 0.0;

            placedSuffixes.Add(suffix);
            return true;
        }
        private Point3d GetSnappedPointOnCurve(
    Point3d approxPt,
    string layerName,
    Dictionary<string, List<Curve>> layerEntities)
        {
            if (!layerEntities.TryGetValue(layerName, out var curves) || curves.Count == 0)
                return approxPt;

            Point3d best = approxPt;
            double minDist = double.MaxValue;

            foreach (Curve c in curves)
            {
                try
                {
                    Point3d snapped = c.GetClosestPointTo(approxPt, false);
                    double d = approxPt.DistanceTo(snapped);
                    if (d < minDist) { minDist = d; best = snapped; }
                }
                catch { continue; }
            }

            return best;
        }

        // ─────────────────────────────────────────────────────────────
        // CREATE PUE DIMENSION
        // ─────────────────────────────────────────────────────────────
        private bool CreatePueDimension(
    Transaction tr, BlockTableRecord mSpace, Polyline cl,
    Point3d rowIntersectPt, Point3d easementIntersectPt, double distance,
    string suffix, int suffixChars,
    double startDist, int index, double gap, bool isLeft,
    double uniformDimLineDist, ObjectId dimStyleId,
    double charW, HashSet<string> placedSuffixes,
    Dictionary<string, List<Curve>> layerEntities)  // ADD THIS PARAMETER
        {
            if (placedSuffixes.Contains(suffix)) return false;

            double rungDist = Math.Min(startDist + (index * gap), cl.Length);
            Point3d rungPt = cl.GetPointAtDist(rungDist);
            Vector3d tangent = cl.GetFirstDerivative(rungPt).GetNormal();
            Vector3d perpDir = isLeft
                ? new Vector3d(-tangent.Y, tangent.X, 0)
                : new Vector3d(tangent.Y, -tangent.X, 0);

            // Snap each endpoint to its actual polyline
            Point3d snappedRow = GetSnappedPointOnCurve(rowIntersectPt, "ROW", layerEntities);
            Point3d snappedEasement = GetSnappedPointOnCurve(easementIntersectPt, "EASEMENT", layerEntities);

            Point3d xLine1 = ProjectToPerp(rungPt, perpDir, snappedRow);
            Point3d xLine2 = ProjectToPerp(rungPt, perpDir, snappedEasement);
            Point3d dimLinePt = xLine2 + (perpDir * BEYOND_OFFSET);

            int roundedDist = (int)Math.Round(distance, 0);
            string dimText = $"{roundedDist}{suffix}";

            AlignedDimension ad = new AlignedDimension(xLine1, xLine2, dimLinePt, string.Empty, dimStyleId);
            ad.DimensionText = dimText;
            ad.Layer = DIM_LAYER;
            ad.Color = Color.FromColorIndex(ColorMethod.ByLayer, 256);

            mSpace.AppendEntity(ad);
            tr.AddNewlyCreatedDBObject(ad, true);

            int totalChars = roundedDist.ToString().Length + suffixChars;
            double textOffset = uniformDimLineDist - ((totalChars * charW) * 0.5);
            ad.TextPosition = rungPt + (perpDir * textOffset);
            ad.TextRotation = 0.0;

            placedSuffixes.Add(suffix);
            return true;
        }

        private bool CreateVnaeDimension(
    Transaction tr, BlockTableRecord mSpace, Polyline cl,
    Point3d propLineIntersectPt, Point3d vnaeIntersectPt, double distance,
    string suffix, int suffixChars,
    double startDist, int index, double gap, bool isLeft,
    double uniformDimLineDist, ObjectId dimStyleId,
    double charW, HashSet<string> placedSuffixes,
    Dictionary<string, List<Curve>> layerEntities)  // ADD THIS PARAMETER
        {
            if (placedSuffixes.Contains(suffix)) return false;

            double rungDist = Math.Min(startDist + (index * gap), cl.Length);
            Point3d rungPt = cl.GetPointAtDist(rungDist);
            Vector3d tangent = cl.GetFirstDerivative(rungPt).GetNormal();
            Vector3d perpDir = isLeft
                ? new Vector3d(-tangent.Y, tangent.X, 0)
                : new Vector3d(tangent.Y, -tangent.X, 0);

            // Snap each endpoint to its actual polyline
            Point3d snappedPropLine = GetSnappedPointOnCurve(propLineIntersectPt, "PROPERTY LINE", layerEntities);
            Point3d snappedVnae = GetSnappedPointOnCurve(vnaeIntersectPt, "VNAE LINE", layerEntities);

            Point3d xLine1 = ProjectToPerp(rungPt, perpDir, snappedPropLine);
            Point3d xLine2 = ProjectToPerp(rungPt, perpDir, snappedVnae);
            Point3d dimLinePt = xLine2 + (perpDir * BEYOND_OFFSET);

            int roundedDist = (int)Math.Round(distance, 0);
            string dimText = $"{roundedDist}{suffix}";

            AlignedDimension ad = new AlignedDimension(xLine1, xLine2, dimLinePt, string.Empty, dimStyleId);
            ad.DimensionText = dimText;
            ad.Layer = DIM_LAYER;
            ad.Color = Color.FromColorIndex(ColorMethod.ByLayer, 256);
            mSpace.AppendEntity(ad);
            tr.AddNewlyCreatedDBObject(ad, true);

            int totalChars = roundedDist.ToString().Length + suffixChars;
            double textOffset = uniformDimLineDist - ((totalChars * charW) * 0.5);
            ad.TextPosition = rungPt + (perpDir * textOffset);
            ad.TextRotation = 0.0;

            placedSuffixes.Add(suffix);
            return true;
        }

        // ─────────────────────────────────────────────────────────────
        // PLACE SIDE DIMENSIONS  (legacy helper, kept for compat)
        // ─────────────────────────────────────────────────────────────
        private void PlaceSideDimensions(
            Transaction tr, BlockTableRecord mSpace,
            Vector3d tangent, Vector3d perpDir,
            List<UtilityDimData> dims, ObjectId dimStyleId,
            Extents2d rect, HashSet<string> placedKeys,
            double uniformDimLineDist, double dimtxt)
        {
            if (!dims.Any()) return;
            double charW = dimtxt * CHAR_ASPECT;
            var itemsToPlace = new List<LabelItem>();
            foreach (var dim in dims)
            {
                if (!LayerSuffixMap.TryGetValue(dim.LayerName, out var cfgs)) continue;
                foreach (var cfg in cfgs) itemsToPlace.Add(new LabelItem { Dim = dim, Config = cfg });
            }
            itemsToPlace = itemsToPlace.OrderBy(li => li.Dim.StationPt.DistanceTo(Point3d.Origin)).ToList();
            int globalIndex = 0;
            Point3d firstStation = itemsToPlace.First().Dim.StationPt;
            foreach (var li in itemsToPlace)
            {
                string labelKey = $"{li.Dim.LayerName}|{li.Config.Suffix}|{Math.Round(li.Dim.Distance, 0)}";
                if (placedKeys.Contains(labelKey)) continue;
                Point3d uniformStation = firstStation + (tangent * (globalIndex * VERTICAL_SPACING));
                Point3d xLine1 = uniformStation;
                Point3d xLine2 = ProjectToPerp(uniformStation, perpDir, li.Dim.IntersectPt);
                Point3d dimLinePt = xLine1 + (perpDir * uniformDimLineDist);
                string dimText = $"{(int)Math.Round(li.Dim.Distance, 0)}{li.Config.Suffix}";
                AlignedDimension ad = new AlignedDimension(xLine1, xLine2, dimLinePt, dimText, dimStyleId);
                ad.Layer = DIM_LAYER;
                if (li.Dim.LayerName.Equals("PROPOSED TRENCH", StringComparison.OrdinalIgnoreCase))
                    ad.Color = Color.FromColorIndex(ColorMethod.ByAci, 30);
                mSpace.AppendEntity(ad);
                tr.AddNewlyCreatedDBObject(ad, true);
                int totalChars = dimText.Length;
                double textOffset = uniformDimLineDist - ((totalChars * charW) * 0.5);
                ad.TextPosition = xLine1 + (perpDir * textOffset);
                ad.TextRotation = 0.0;
                placedKeys.Add(labelKey);
                globalIndex++;
            }
        }

        private bool IsLeftOfLine(Point3d start, Point3d end, Point3d pt) =>
            ((end.X - start.X) * (pt.Y - start.Y) - (end.Y - start.Y) * (pt.X - start.X)) > 0;

        private (double minDist, double maxDist) GetClDistRangeInRect(
    Polyline cl, Extents2d rect)
        {
            double minDist = double.MaxValue;
            double maxDist = double.MinValue;

            // Use fine step — rect may be small relative to total CL length
            double step = Math.Min(1.0, Math.Min(
                                rect.MaxPoint.X - rect.MinPoint.X,
                                rect.MaxPoint.Y - rect.MinPoint.Y) / 20.0);
            int samples = Math.Max(100, (int)Math.Ceiling(cl.Length / step));

            double tol = 2.0; // tight — we want points actually inside rect

            for (int i = 0; i <= samples; i++)
            {
                double d = cl.Length * i / samples;
                Point3d pt;
                try { pt = cl.GetPointAtDist(d); }
                catch { continue; }

                if (pt.X < rect.MinPoint.X - tol || pt.X > rect.MaxPoint.X + tol ||
                    pt.Y < rect.MinPoint.Y - tol || pt.Y > rect.MaxPoint.Y + tol)
                    continue;

                if (d < minDist) minDist = d;
                if (d > maxDist) maxDist = d;
            }

            if (minDist == double.MaxValue)
            {
                // CL passes through but no sampled point landed inside —
                // find the closest CL point to rect centre as fallback
                Point3d rectCentre = new Point3d(
                    (rect.MinPoint.X + rect.MaxPoint.X) / 2.0,
                    (rect.MinPoint.Y + rect.MaxPoint.Y) / 2.0, 0);
                try
                {
                    Point3d foot = cl.GetClosestPointTo(rectCentre, false);
                    double fd = cl.GetDistAtPoint(foot);
                    double half = Math.Max(
                                    rect.MaxPoint.X - rect.MinPoint.X,
                                    rect.MaxPoint.Y - rect.MinPoint.Y);
                    minDist = Math.Max(0, fd - half);
                    maxDist = Math.Min(cl.Length, fd + half);
                }
                catch
                {
                    minDist = 0;
                    maxDist = cl.Length;
                }
            }

            return (minDist, maxDist);
        }
        // ─────────────────────────────────────────────────────────────
        // COLLECT DIM CANDIDATES
        //
        // No structural change: GetPointAtParameter / GetClosestPointTo /
        // IntersectWith are Curve-API methods that work equally for
        // Polyline and Line.  Layer entities now include Line objects
        // (CHANGE-3) so they are sampled and hit-tested here automatically.
        // ─────────────────────────────────────────────────────────────
        private void CollectDimCandidates(
    Polyline cl,
    Dictionary<string, List<Curve>> layerEntities,
    Extents2d rect,
    List<UtilityDimData> leftDims,
    List<UtilityDimData> rightDims,
    HashSet<string> leftSeen,
    HashSet<string> rightSeen,
    Editor ed)
        {
            var (clMinDist, clMaxDist) = GetClDistRangeInRect(cl, rect);
            ed.WriteMessage($"\n  CL dist range in rect: {clMinDist:F2} → {clMaxDist:F2}");

            foreach (var kvp in layerEntities)
            {
                string lyrName = kvp.Key;
                List<Curve> curves = kvp.Value;
                if (curves.Count == 0) continue;

                foreach (Curve utilityCurve in curves)
                {
                    double curveLen = GetCurveLength(utilityCurve);
                    if (curveLen <= 0) continue;

                    int sampleCount = Math.Max(2,
                        (int)Math.Ceiling(curveLen / CURVE_SAMPLE_STEP));

                    for (int i = 0; i <= sampleCount; i++)
                    {
                        double t = (double)i / sampleCount;
                        Point3d samplePt;
                        try
                        {
                            samplePt = utilityCurve.GetPointAtParameter(
                                utilityCurve.StartParam + t *
                                (utilityCurve.EndParam - utilityCurve.StartParam));
                        }
                        catch { continue; }

                        // GATE 1: sample must be inside rect
                        if (!PointInRect(samplePt, rect)) continue;

                        // Get closest point on CL — may be outside rect, that's OK
                        Point3d clFoot;
                        try { clFoot = cl.GetClosestPointTo(samplePt, false); }
                        catch { continue; }

                        // ── REMOVED: clFoot rect check ──────────────────────────
                        // The CL may run parallel beside the rect. clFoot will
                        // legitimately land outside the rect in that case.
                        // The only guards we need are: samplePt in rect (above)
                        // and hitPt in rect (below). Distance is still valid.
                        // ────────────────────────────────────────────────────────

                        Vector3d tangent;
                        try { tangent = cl.GetFirstDerivative(clFoot).GetNormal(); }
                        catch { continue; }

                        Vector3d perpLeft = new Vector3d(-tangent.Y, tangent.X, 0);
                        Vector3d perpRight = new Vector3d(tangent.Y, -tangent.X, 0);
                        bool isLeft = (samplePt - clFoot).DotProduct(perpLeft) >= 0;
                        Vector3d perpDir = isLeft ? perpLeft : perpRight;
                        var seen = isLeft ? leftSeen : rightSeen;
                        var list = isLeft ? leftDims : rightDims;

                        // Fire ray from clFoot outward — use rect diagonal as max length
                        double rectDiag = Math.Sqrt(
                                                Math.Pow(rect.MaxPoint.X - rect.MinPoint.X, 2) +
                                                Math.Pow(rect.MaxPoint.Y - rect.MinPoint.Y, 2));
                        double rayLen = rectDiag + RAY_BUFFER;
                        Point3d rayEnd = clFoot + perpDir * rayLen;
                        Point3d? hitPt = FindNearestIntersection(
                                            clFoot, rayEnd, curves, rayLen);
                        if (!hitPt.HasValue) continue;

                        // GATE 2: hit point must be inside rect
                        if (!PointInRectTolerant(hitPt.Value, rect, 1.0)) continue;

                        double dist = clFoot.DistanceTo(hitPt.Value);
                        string key = $"{lyrName}|{Math.Round(dist, 0)}";
                        if (seen.Contains(key)) continue;
                        seen.Add(key);

                        list.Add(new UtilityDimData
                        {
                            LayerName = lyrName,
                            Distance = dist,
                            IntersectPt = hitPt.Value,
                            StationPt = clFoot,
                            DedupKey = key
                        });
                    }
                }
            }
        }

        private bool PointInRectTolerant(Point3d pt, Extents2d rect, double tol) =>
            pt.X >= rect.MinPoint.X - tol && pt.X <= rect.MaxPoint.X + tol &&
            pt.Y >= rect.MinPoint.Y - tol && pt.Y <= rect.MaxPoint.Y + tol;
        // ─────────────────────────────────────────────────────────────
        // GROUP BY STATION
        // ─────────────────────────────────────────────────────────────
        private Dictionary<Point3d, List<UtilityDimData>> GroupByStation(List<UtilityDimData> dims)
        {
            double tol = CURVE_SAMPLE_STEP * 0.5;
            var groups = new Dictionary<Point3d, List<UtilityDimData>>();
            foreach (UtilityDimData dim in dims)
            {
                Point3d? matchKey = null;
                foreach (Point3d existing in groups.Keys)
                    if (existing.DistanceTo(dim.StationPt) <= tol) { matchKey = existing; break; }
                if (matchKey.HasValue) groups[matchKey.Value].Add(dim);
                else groups[dim.StationPt] = new List<UtilityDimData> { dim };
            }
            return groups;
        }

        private double SafeGetDistAtPoint(Polyline cl, Point3d pt)
        {
            try { return cl.GetDistAtPoint(pt); }
            catch { return 0.0; }
        }

        // ─────────────────────────────────────────────────────────────
        // GET CURVE LENGTH  (CHANGE-4)
        //
        // Explicit Line branch returns Line.Length (true geometric length)
        // instead of the bounding-box diagonal that was used before.
        // Accurate length → correct CURVE_SAMPLE_STEP-based sample count.
        // ─────────────────────────────────────────────────────────────
        private double GetCurveLength(Curve c)
        {
            try
            {
                if (c is Polyline pl) return pl.Length;   // polyline: exact
                if (c is Line ln) return ln.Length;   // CHANGE-4: Line: exact

                // Polyline2d / Polyline3d: bounding-box diagonal as upper bound
                Extents3d ext = c.GeometricExtents;
                double dx = ext.MaxPoint.X - ext.MinPoint.X;
                double dy = ext.MaxPoint.Y - ext.MinPoint.Y;
                return Math.Sqrt(dx * dx + dy * dy);
            }
            catch { return 0.0; }
        }

        // ─────────────────────────────────────────────────────────────
        // RAY LENGTH TO RECT EDGE
        // ─────────────────────────────────────────────────────────────
        private double RayLengthToEdge(Point3d clPt, Vector3d perpDir, Extents2d rect)
        {
            double tMin = double.MaxValue;
            if (Math.Abs(perpDir.X) > 1e-9)
            {
                double t1 = (rect.MinPoint.X - clPt.X) / perpDir.X;
                double t2 = (rect.MaxPoint.X - clPt.X) / perpDir.X;
                if (t1 > 0) tMin = Math.Min(tMin, t1);
                if (t2 > 0) tMin = Math.Min(tMin, t2);
            }
            if (Math.Abs(perpDir.Y) > 1e-9)
            {
                double t3 = (rect.MinPoint.Y - clPt.Y) / perpDir.Y;
                double t4 = (rect.MaxPoint.Y - clPt.Y) / perpDir.Y;
                if (t3 > 0) tMin = Math.Min(tMin, t3);
                if (t4 > 0) tMin = Math.Min(tMin, t4);
            }
            if (tMin == double.MaxValue || tMin <= 0)
                tMin = Math.Max(rect.MaxPoint.X - rect.MinPoint.X, rect.MaxPoint.Y - rect.MinPoint.Y);
            return tMin;
        }

        // ─────────────────────────────────────────────────────────────
        // COMPUTE UNIFORM DIM-LINE DISTANCE
        // ─────────────────────────────────────────────────────────────
        private double ComputeUniformDimLineDist(
            List<UtilityDimData> allDims, double dimtxt, double beyondOffset, Editor ed)
        {
            if (!allDims.Any()) return 0.0;
            double charW = dimtxt * CHAR_ASPECT;
            double maxColumnRight = 0.0;
            foreach (UtilityDimData dim in allDims)
            {
                if (!LayerSuffixMap.TryGetValue(dim.LayerName, out var cfgs) || cfgs.Count == 0) continue;
                LayerConfig cfg = cfgs.OrderByDescending(c => c.SuffixCharCount).First();
                int rounded = (int)Math.Round(dim.Distance, 0);
                int totalChars = rounded.ToString().Length + cfg.SuffixCharCount;
                double colRight = dim.Distance + (totalChars * charW);
                if (colRight > maxColumnRight) maxColumnRight = colRight;
            }
            double uniformDist = maxColumnRight + beyondOffset;
            ed.WriteMessage($"\n  maxColumnRight={maxColumnRight:F2}  uniformDimLineDist={uniformDist:F2}");
            return uniformDist;
        }

        // ─────────────────────────────────────────────────────────────
        // PLACE STACKED ALIGNED DIMENSIONS  (legacy helper)
        // ─────────────────────────────────────────────────────────────
        private void PlaceStackedDims(
            Transaction tr, BlockTableRecord mSpace,
            Vector3d tangent, Vector3d perpDir,
            List<UtilityDimData> dims, ObjectId dimStyleId,
            double stackGap, double beyondOffset,
            Extents2d rect, HashSet<string> placedKeys,
            double uniformDimLineDist, double dimtxt)
        {
            if (!dims.Any()) return;
            double charW = dimtxt * CHAR_ASPECT;
            var labelItems = new List<LabelItem>();
            foreach (UtilityDimData dim in dims)
            {
                if (!LayerSuffixMap.TryGetValue(dim.LayerName, out var cfgs) || cfgs.Count == 0) continue;
                foreach (LayerConfig cfg in cfgs) labelItems.Add(new LabelItem { Dim = dim, Config = cfg });
            }
            labelItems = labelItems.OrderBy(li => li.Dim.StationPt.DistanceTo(Point3d.Origin)).ToList();
            int index = 0;
            Point3d anchorStation = labelItems.First().Dim.StationPt;
            foreach (LabelItem li in labelItems)
            {
                string labelKey = $"{li.Dim.LayerName}|{li.Config.Suffix}|{Math.Round(li.Dim.Distance, 0)}";
                if (placedKeys.Contains(labelKey)) continue;
                Point3d uniformStation = anchorStation + (tangent * (index * VERTICAL_SPACING));
                Point3d xLine1 = uniformStation;
                Point3d xLine2 = ProjectToPerp(uniformStation, perpDir, li.Dim.IntersectPt);
                Point3d dimLinePt = xLine1 + perpDir * uniformDimLineDist;
                int roundedDist = (int)Math.Round(li.Dim.Distance, 0);
                string dimText = roundedDist.ToString() + li.Config.Suffix;
                int totalChars = roundedDist.ToString().Length + li.Config.SuffixCharCount;
                double textOffset = uniformDimLineDist - ((totalChars * charW) * 0.5);
                AlignedDimension ad = new AlignedDimension(xLine1, xLine2, dimLinePt, string.Empty, dimStyleId);
                ad.DimensionText = dimText;
                ad.Layer = DIM_LAYER;
                ad.Color = li.Dim.LayerName.Equals("PROPOSED TRENCH", StringComparison.OrdinalIgnoreCase)
                           ? Color.FromColorIndex(ColorMethod.ByAci, 30)
                           : Color.FromColorIndex(ColorMethod.ByBlock, 0);
                mSpace.AppendEntity(ad);
                tr.AddNewlyCreatedDBObject(ad, true);
                ad.TextPosition = xLine1 + perpDir * textOffset;
                ad.TextRotation = 0.0;
                placedKeys.Add(labelKey);
                index++;
            }
        }

        // ─────────────────────────────────────────────────────────────
        // PICK FOUR POINTS
        // ─────────────────────────────────────────────────────────────
        private Point3d[] PickFourPoints(Editor ed)
        {
            var pts = new Point3d[4];
            string[] labels = { "1st", "2nd", "3rd", "4th" };
            for (int i = 0; i < 4; i++)
            {
                PromptPointOptions ppo = new PromptPointOptions($"\nPick {labels[i]} corner of work rectangle: ");
                ppo.AllowNone = false;
                if (i > 0) ppo.BasePoint = pts[i - 1];
                PromptPointResult ppr = ed.GetPoint(ppo);
                if (ppr.Status != PromptStatus.OK) return null;
                pts[i] = ppr.Value;
            }
            return pts;
        }

        // ─────────────────────────────────────────────────────────────
        // DRAW / ERASE VISUAL RECTANGLE
        // ─────────────────────────────────────────────────────────────
        private ObjectId DrawVisualRectangle(Database db, Extents2d rect)
        {
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord mSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                Polyline rp = new Polyline();
                rp.AddVertexAt(0, rect.MinPoint, 0, 0, 0);
                rp.AddVertexAt(1, new Point2d(rect.MaxPoint.X, rect.MinPoint.Y), 0, 0, 0);
                rp.AddVertexAt(2, rect.MaxPoint, 0, 0, 0);
                rp.AddVertexAt(3, new Point2d(rect.MinPoint.X, rect.MaxPoint.Y), 0, 0, 0);
                rp.Closed = true;
                rp.Color = Color.FromColorIndex(ColorMethod.ByAci, 1);
                rp.Layer = "0";
                mSpace.AppendEntity(rp);
                tr.AddNewlyCreatedDBObject(rp, true);
                ObjectId id = rp.ObjectId;
                tr.Commit();
                return id;
            }
        }

        private void EraseVisualRectangle(Database db, ObjectId rectId)
        {
            if (rectId == ObjectId.Null || !rectId.IsValid) return;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    DBObject obj = tr.GetObject(rectId, OpenMode.ForWrite);
                    if (obj != null && !obj.IsErased) obj.Erase();
                    tr.Commit();
                }
                catch { tr.Abort(); }
            }
        }

        // ─────────────────────────────────────────────────────────────
        // FIND NEAREST INTERSECTION
        // ─────────────────────────────────────────────────────────────
        private Point3d? FindNearestIntersection(
            Point3d rayStart, Point3d rayEnd,
            List<Curve> curves, double maxPlausibleDist)
        {
            Point3d? nearest = null;
            double minDist = double.MaxValue;
            using (Line ray = new Line(rayStart, rayEnd))
            {
                foreach (Curve curve in curves)
                {
                    var pts = new Point3dCollection();
                    try
                    {
                        curve.IntersectWith(ray, Intersect.ExtendThis,
                            pts, IntPtr.Zero, IntPtr.Zero);
                    }
                    catch { continue; }
                    foreach (Point3d pt in pts)
                    {
                        Vector3d toIntersect = pt - rayStart;
                        Vector3d rayDir = rayEnd - rayStart;
                        if (toIntersect.DotProduct(rayDir) < 0) continue;
                        double d = rayStart.DistanceTo(pt);
                        if (d > maxPlausibleDist) continue;
                        if (d < minDist) { minDist = d; nearest = pt; }
                    }
                }
            }
            return nearest;
        }

        // ─────────────────────────────────────────────────────────────
        // FIND CENTERLINE INSIDE RECTANGLE  (CHANGE-5)
        //
        // Accepts Polyline and Line entities on any CENTERLINE layer.
        //
        // A plain Line is wrapped in a synthetic single-segment Polyline
        // so all downstream callers (GetClosestPointTo, GetFirstDerivative,
        // GetDistAtPoint, GetPointAtDist) work without further changes.
        // The synthetic Polyline is NOT added to the database.
        //
        // Priority: real Polyline > Line.  A Polyline hit returns
        // immediately; a Line hit is remembered and returned only if no
        // Polyline is found after scanning all of mSpace.
        // ─────────────────────────────────────────────────────────────
        private Polyline FindCenterlineInRect(
    Database db, Transaction tr, Editor ed, Extents2d rect)
        {
            BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
            BlockTableRecord mSpace = tr.GetObject(
                bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

            Polyline rectPoly = BuildRectPolyline(rect);
            Point3d rectCentre = new Point3d(
                (rect.MinPoint.X + rect.MaxPoint.X) / 2.0,
                (rect.MinPoint.Y + rect.MaxPoint.Y) / 2.0, 0);

            // Collect ALL candidate centerlines with their score
            var candidates = new List<(Polyline pl, double score, bool isSynthetic)>();

            foreach (ObjectId id in mSpace)
            {
                Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent == null) continue;
                if (ent.Layer.IndexOf("CENTERLINE", StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                // ── Branch A: Polyline ────────────────────────────────────
                if (ent is Polyline pl)
                {
                    bool intersects = false;

                    // Check vertices inside rect
                    for (int v = 0; v < pl.NumberOfVertices; v++)
                    {
                        Point2d vpt = pl.GetPoint2dAt(v);
                        if (PointInRect(new Point3d(vpt.X, vpt.Y, 0), rect))
                        { intersects = true; break; }
                    }

                    // Check intersection with rect boundary
                    if (!intersects)
                    {
                        var intPts = new Point3dCollection();
                        try
                        {
                            pl.IntersectWith(rectPoly, Intersect.OnBothOperands,
                                intPts, IntPtr.Zero, IntPtr.Zero);
                        }
                        catch { }
                        if (intPts.Count > 0) intersects = true;
                    }

                    // Check if CL spans entirely over rect
                    if (!intersects)
                    {
                        try
                        {
                            Extents3d ext = pl.GeometricExtents;
                            if (ext.MinPoint.X <= rect.MinPoint.X &&
                                ext.MinPoint.Y <= rect.MinPoint.Y &&
                                ext.MaxPoint.X >= rect.MaxPoint.X &&
                                ext.MaxPoint.Y >= rect.MaxPoint.Y)
                                intersects = true;
                        }
                        catch { }
                    }

                    if (!intersects) continue;

                    // Score = distance from rect centre to closest point on this CL
                    // Lower score = better candidate
                    double score = double.MaxValue;
                    try
                    {
                        Point3d foot = pl.GetClosestPointTo(rectCentre, false);
                        score = rectCentre.DistanceTo(foot);
                    }
                    catch { continue; }

                    ed.WriteMessage(
                        $"\n  CL candidate: Layer=[{pl.Layer}] " +
                        $"Vertices={pl.NumberOfVertices} " +
                        $"Score(dist to rect centre)={score:F2}");

                    candidates.Add((pl, score, false));
                }

                // ── Branch B: Line → synthetic Polyline ──────────────────
                else if (ent is Line ln)
                {
                    bool lineInRect =
                        PointInRect(ln.StartPoint, rect) ||
                        PointInRect(ln.EndPoint, rect);

                    if (!lineInRect)
                    {
                        var intPts = new Point3dCollection();
                        try
                        {
                            ln.IntersectWith(rectPoly, Intersect.OnBothOperands,
                                intPts, IntPtr.Zero, IntPtr.Zero);
                        }
                        catch { }
                        lineInRect = intPts.Count > 0;
                    }

                    if (!lineInRect) continue;

                    double score = double.MaxValue;
                    try { score = rectCentre.DistanceTo(ln.GetClosestPointTo(rectCentre, false)); }
                    catch { continue; }

                    Polyline synth = new Polyline();
                    synth.AddVertexAt(0, new Point2d(ln.StartPoint.X, ln.StartPoint.Y), 0, 0, 0);
                    synth.AddVertexAt(1, new Point2d(ln.EndPoint.X, ln.EndPoint.Y), 0, 0, 0);
                    synth.Closed = false;

                    ed.WriteMessage(
                        $"\n  CL candidate (Line→synth): Layer=[{ln.Layer}] " +
                        $"Score={score:F2}");

                    candidates.Add((synth, score, true));
                }
            }

            rectPoly.Dispose();

            if (!candidates.Any())
            {
                ed.WriteMessage("\n  No CENTERLINE candidates found.");
                return null;
            }

            // Pick the candidate closest to the rect centre
            var best = candidates.OrderBy(c => c.score).First();

            // Dispose synthetic polylines that were not selected
            foreach (var c in candidates)
                if (c.isSynthetic && c.pl != best.pl)
                    c.pl.Dispose();

            ed.WriteMessage(
                $"\n  Selected CL: Score={best.score:F2} " +
                $"Vertices={best.pl.NumberOfVertices}");

            return best.pl;
        }

        // ─────────────────────────────────────────────────────────────
        // BUILD TEMP RECT POLYLINE
        // ─────────────────────────────────────────────────────────────
        private Polyline BuildRectPolyline(Extents2d rect)
        {
            Polyline rp = new Polyline();
            rp.AddVertexAt(0, rect.MinPoint, 0, 0, 0);
            rp.AddVertexAt(1, new Point2d(rect.MaxPoint.X, rect.MinPoint.Y), 0, 0, 0);
            rp.AddVertexAt(2, rect.MaxPoint, 0, 0, 0);
            rp.AddVertexAt(3, new Point2d(rect.MinPoint.X, rect.MaxPoint.Y), 0, 0, 0);
            rp.Closed = true;
            return rp;
        }

        // ─────────────────────────────────────────────────────────────
        // COLLECT LAYER ENTITIES IN RECTANGLE  (CHANGE-3)
        //
        // IsAcceptedCurve() now gates entry, making it explicit that
        // Polyline, Polyline2d, Polyline3d, and Line are all accepted.
        // Text, BlockReference, Hatch, Arc, Spline, etc. are excluded.
        // ─────────────────────────────────────────────────────────────
        private Dictionary<string, List<Curve>> CollectLayerEntitiesInRect(
            Database db, Transaction tr, Extents2d rect)
        {
            var result = new Dictionary<string, List<Curve>>(StringComparer.OrdinalIgnoreCase);
            foreach (string key in LayerSuffixMap.Keys) result[key] = new List<Curve>();

            BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
            BlockTableRecord mSpace = tr.GetObject(
                bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

            foreach (ObjectId id in mSpace)
            {
                Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent == null) continue;
                if (!result.ContainsKey(ent.Layer)) continue;

                // CHANGE-3/6: explicit type guard — Polyline, Polyline2d,
                // Polyline3d, and Line are accepted; everything else is skipped.
                if (!IsAcceptedCurve(ent)) continue;

                Curve c = ent as Curve;   // safe: IsAcceptedCurve guarantees this

                try
                {
                    Extents3d ext = c.GeometricExtents;
                    if (ext.MaxPoint.X < rect.MinPoint.X || ext.MinPoint.X > rect.MaxPoint.X ||
                        ext.MaxPoint.Y < rect.MinPoint.Y || ext.MinPoint.Y > rect.MaxPoint.Y)
                        continue;
                }
                catch { continue; }

                result[ent.Layer].Add(c);
            }

            return result;
        }

        // ─────────────────────────────────────────────────────────────
        // PROJECT INTERSECTION ONTO PERP AT STATION
        // ─────────────────────────────────────────────────────────────
        private Point3d ProjectToPerp(Point3d stationPt, Vector3d perpDir, Point3d originalIntersect)
        {
            double dist = (originalIntersect - stationPt).DotProduct(perpDir);
            return stationPt + perpDir * dist;
        }

        // ─────────────────────────────────────────────────────────────
        // POINT IN RECT HELPER
        // ─────────────────────────────────────────────────────────────
        private bool PointInRect(Point3d pt, Extents2d rect) =>
            pt.X >= rect.MinPoint.X && pt.X <= rect.MaxPoint.X &&
            pt.Y >= rect.MinPoint.Y && pt.Y <= rect.MaxPoint.Y;

        // ─────────────────────────────────────────────────────────────
        // GET OR CREATE DIM STYLE
        // ─────────────────────────────────────────────────────────────
        private ObjectId GetOrCreateDimStyle(Database db, Transaction tr, double dimtxt, double dimasz)
        {
            DimStyleTable dst = tr.GetObject(db.DimStyleTableId, OpenMode.ForRead) as DimStyleTable;
            DimStyleTableRecord dstr;
            if (dst.Has(DIM_STYLE_NAME))
                dstr = tr.GetObject(dst[DIM_STYLE_NAME], OpenMode.ForWrite) as DimStyleTableRecord;
            else
            {
                dst.UpgradeOpen();
                dstr = new DimStyleTableRecord { Name = DIM_STYLE_NAME };
                dst.Add(dstr);
                tr.AddNewlyCreatedDBObject(dstr, true);
            }
            dstr.Dimscale = 1.0; dstr.Dimasz = dimasz; dstr.Dimtsz = 0.0;
            dstr.Dimexo = dimtxt * 0.5; dstr.Dimexe = dimtxt * 0.25;
            dstr.Dimse1 = false; dstr.Dimse2 = false; dstr.Dimdle = 0.0; dstr.Dimdli = 0.0;
            dstr.Dimclrd = Color.FromColorIndex(ColorMethod.ByBlock, 0);
            dstr.Dimclre = Color.FromColorIndex(ColorMethod.ByBlock, 0);
            dstr.Dimclrt = Color.FromColorIndex(ColorMethod.ByBlock, 0);
            dstr.Dimtxt = dimtxt; dstr.Dimgap = dimtxt * 0.09;
            dstr.Dimatfit = 0; dstr.Dimtmove = 2; dstr.Dimtad = 0;
            dstr.Dimtoh = true; dstr.Dimtih = false; dstr.Dimtix = false;
            dstr.Dimsoxd = false; dstr.Dimdec = 0; dstr.Dimzin = 8;
            dstr.Dimrnd = 1.0; dstr.Dimlunit = 2; dstr.Dimdsep = '.';
            dstr.Dimlfac = 1.0; dstr.Dimtol = false; dstr.Dimlim = false;

            TextStyleTable tst = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
            if (tst.Has("ARIAL")) dstr.Dimtxsty = tst["ARIAL"];
            else if (tst.Has("Arial")) dstr.Dimtxsty = tst["Arial"];
            else
            {
                tst.UpgradeOpen();
                TextStyleTableRecord arialStyle = new TextStyleTableRecord
                { Name = "ARIAL", FileName = "arial.ttf", TextSize = 0.0 };
                ObjectId arialId = tst.Add(arialStyle);
                tr.AddNewlyCreatedDBObject(arialStyle, true);
                dstr.Dimtxsty = arialId;
            }
            return dstr.ObjectId;
        }

        // ─────────────────────────────────────────────────────────────
        // ENSURE DIMENSIONS LAYER EXISTS
        // ─────────────────────────────────────────────────────────────
        private void EnsureLayerExists(Database db, Transaction tr)
        {
            LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
            if (lt.Has(DIM_LAYER)) return;
            lt.UpgradeOpen();
            LayerTableRecord ltr = new LayerTableRecord
            { Name = DIM_LAYER, Color = Color.FromColorIndex(ColorMethod.ByAci, 7) };
            lt.Add(ltr);
            tr.AddNewlyCreatedDBObject(ltr, true);
        }

        // ─────────────────────────────────────────────────────────────
        // HELPER CLASSES
        // ─────────────────────────────────────────────────────────────
        public class UtilityDimData
        {
            public string LayerName { get; set; }
            public double Distance { get; set; }
            public Point3d IntersectPt { get; set; }
            public Point3d StationPt { get; set; }
            public string DedupKey { get; set; }
        }

        internal struct LabelItem
        {
            public UtilityDimData Dim { get; set; }
            public AutoDimRectCommand_Pauley.LayerConfig Config { get; set; }
        }
    }
}