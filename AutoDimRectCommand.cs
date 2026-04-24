using IntelliCAD.ApplicationServices;
using IntelliCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Teigha.Colors;
using Teigha.DatabaseServices;
using Teigha.Geometry;
using Teigha.Runtime;
using Application = IntelliCAD.ApplicationServices.Application;

namespace Rough_Works
{
    public class AutoDimRectCommand
    {
        // ─────────────────────────────────────────────
        // CONFIGURATION
        // ─────────────────────────────────────────────
        public struct LayerConfig
        {
            public string Suffix { get; }
            public double TextWidth { get; }   // pre-measured suffix pixel width

            public LayerConfig(string suffix, double textWidth)
            {
                Suffix = suffix;
                TextWidth = textWidth;
            }
        }
        private static readonly Dictionary<string, LayerConfig> LayerSuffixMap =
            new Dictionary<string, LayerConfig>(StringComparer.OrdinalIgnoreCase)
            {
                { "WATER",           new LayerConfig("' WATER",        20.5  ) },
                { "SS",              new LayerConfig("' SEWER",        20.5  ) },
                { "BOC",             new LayerConfig("' BOC",          11.66 ) },
                { "SIDEWALK",        new LayerConfig("' S/W",          10.41 ) },
                { "ROW",             new LayerConfig("' ROW",          13.39 ) },
                { "EASEMENTS",       new LayerConfig("' PUE",          11.01 ) },
                { "PROPOSED TRENCH", new LayerConfig("' PROP. TRENCH", 26.7  ) },
            };

        private const string DIM_LAYER = "DIMENSIONS";
        private const string DIM_STYLE_NAME = "BPGDIMS";
        private const double INTERVAL = 50.0;
        private const double STACK_GAP = 8.0;
        private const double BEYOND_OFFSET = 15.0;


        // ─────────────────────────────────────────────
        // MAIN COMMAND
        // ─────────────────────────────────────────────
        [CommandMethod("AUTODIMRECT")]
        public void AutoDimRect()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            int originalDimassoc = Convert.ToInt32(
                Application.GetSystemVariable("DIMASSOC"));

            ed.Command("_.DIMASSOC", 0);
            db.DimAssoc = 0;

            try
            {
                // ── Step 1: Pick 4 points manually ──
                Point3d[] corners = PickFourPoints(ed);
                if (corners == null)
                {
                    ed.WriteMessage("\n❌ Rectangle selection cancelled.");
                    return;
                }

                // ── Step 2: Build axis-aligned extents from 4 points ──
                double minX = corners.Min(p => p.X);
                double maxX = corners.Max(p => p.X);
                double minY = corners.Min(p => p.Y);
                double maxY = corners.Max(p => p.Y);

                Extents2d rect = new Extents2d(
                    new Point2d(minX, minY),
                    new Point2d(maxX, maxY));

                // ── Step 3: Draw a visual rectangle on screen ──
                ObjectId visualRectId = DrawVisualRectangle(db, rect);

                // ── Step 4: Compute rayLength from full model extents ──
                double rayLength = ComputeRayLength(db);

                ed.WriteMessage(
                    $"\nRectangle defined. Ray length: {rayLength:F1} units.");

                // ── Step 5: Main dimension transaction ──
                using (Transaction tr =
                           db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        // ── Find CENTERLINE ──
                        Polyline cl =
                            FindCenterlineInRect(db, tr, ed, rect);
                        if (cl == null)
                        {
                            ed.WriteMessage(
                                "\n❌ No CENTERLINE polyline found " +
                                "inside the selected rectangle.");
                            tr.Abort();
                            return;
                        }

                        // ── Collect utility layers ──
                        var layerEntities =
                            CollectLayerEntitiesInRect(db, tr, rect);

                        if (!layerEntities.Any(kvp => kvp.Value.Count > 0))
                            ed.WriteMessage(
                                "\n⚠️  No recognised utility layers " +
                                "found inside the rectangle.");

                        // ── Setup dim layer & style ──
                        EnsureLayerExists(db, tr);
                        ObjectId dimStyleId = GetOrCreateDimStyle(db, tr);

                        BlockTable bt = tr.GetObject(
                            db.BlockTableId,
                            OpenMode.ForRead) as BlockTable;
                        BlockTableRecord mSpace = tr.GetObject(
                            bt[BlockTableRecord.ModelSpace],
                            OpenMode.ForWrite) as BlockTableRecord;

                        // ── Dedup trackers — one per side ──
                        var placedLeft = new HashSet<string>();
                        var placedRight = new HashSet<string>();

                        // ─────────────────────────────────────────
                        // PRE-PASS: collect ALL dims on both sides
                        // across every station so we can find the
                        // true largest distance before placing anything.
                        // ─────────────────────────────────────────
                        var allLeftDims = new List<DimResult>();
                        var allRightDims = new List<DimResult>();

                        double total = cl.Length;
                        double currentDist = 0.0;

                        // Temporary dedup sets just for the pre-pass
                        var preLeft = new HashSet<string>();
                        var preRight = new HashSet<string>();

                        while (currentDist <= total)
                        {
                            Point3d clPt = cl.GetPointAtDist(currentDist);

                            if (!PointInRect(clPt, rect))
                            {
                                currentDist += INTERVAL;
                                continue;
                            }

                            Vector3d tangent =
                                cl.GetFirstDerivative(clPt).GetNormal();
                            Vector3d perpLeft =
                                new Vector3d(-tangent.Y, tangent.X, 0);
                            Vector3d perpRight =
                                new Vector3d(tangent.Y, -tangent.X, 0);

                            foreach (var kvp in layerEntities)
                            {
                                string lyrName = kvp.Key;
                                List<Curve> curves = kvp.Value;

                                // Left
                                Point3d? lPt = FindNearestIntersection(
                                    clPt,
                                    clPt + perpLeft * rayLength,
                                    curves);
                                if (lPt.HasValue)
                                {
                                    double dist = clPt.DistanceTo(lPt.Value);
                                    string key =
                                        $"{lyrName}|{Math.Round(dist, 0)}";
                                    if (!preLeft.Contains(key))
                                    {
                                        preLeft.Add(key);
                                        allLeftDims.Add(new DimResult
                                        {
                                            LayerName = lyrName,
                                            Distance = dist,
                                            IntersectPt = lPt.Value,
                                            DedupKey = key
                                        });
                                    }
                                }

                                // Right
                                Point3d? rPt = FindNearestIntersection(
                                    clPt,
                                    clPt + perpRight * rayLength,
                                    curves);
                                if (rPt.HasValue)
                                {
                                    double dist = clPt.DistanceTo(rPt.Value);
                                    string key =
                                        $"{lyrName}|{Math.Round(dist, 0)}";
                                    if (!preRight.Contains(key))
                                    {
                                        preRight.Add(key);
                                        allRightDims.Add(new DimResult
                                        {
                                            LayerName = lyrName,
                                            Distance = dist,
                                            IntersectPt = rPt.Value,
                                            DedupKey = key
                                        });
                                    }
                                }
                            }

                            currentDist += INTERVAL;
                        }

                        // ── Compute the largest distance on each side ──
                        // This is the value that replaces the hardcoded 43.
                        // It represents the measured length of the longest
                        // dimension string so text anchor is always offset
                        // correctly regardless of how large the dims get.
                        double largestLeftDist = allLeftDims.Any()
                            ? allLeftDims.Max(d => d.Distance)
                            : 0.0;
                        double largestRightDist = allRightDims.Any()
                            ? allRightDims.Max(d => d.Distance)
                            : 0.0;

                        //// Convert largest distance to a text-position
                        //// anchor offset using the same formula pattern
                        //// already in PlaceStackedDims:
                        ////   distValue = ((distValue + suffixLen) / 2) + anchor
                        //// We compute the raw "measured length" contribution
                        //// from the largest distance value the same way the
                        //// dimension number string length is computed.
                        //double largestLeftAnchor =
                        //    ComputeAnchorOffset(largestLeftDist, LayerSuffixMap);
                        //double largestRightAnchor =
                        //    ComputeAnchorOffset(largestRightDist, LayerSuffixMap);

                        //ed.WriteMessage(
                        //    $"\nLargest LEFT dist : {largestLeftDist:F1}  " +
                        //    $"anchor: {largestLeftAnchor:F2}");
                        //ed.WriteMessage(
                        //    $"\nLargest RIGHT dist: {largestRightDist:F1}  " +
                        //    $"anchor: {largestRightAnchor:F2}");

                        double uniformLeftOffset = ComputeUniformOffset(
                                allLeftDims, LayerSuffixMap);
                        double uniformRightOffset = ComputeUniformOffset(
                                                        allRightDims, LayerSuffixMap);

                        ed.WriteMessage(
                            $"\nUniform LEFT  offset: {uniformLeftOffset:F2}");
                        ed.WriteMessage(
                            $"\nUniform RIGHT offset: {uniformRightOffset:F2}");

                        // ── Main placement pass ──
                        currentDist = 0.0;

                        while (currentDist <= total)
                        {
                            Point3d clPt = cl.GetPointAtDist(currentDist);

                            if (!PointInRect(clPt, rect))
                            {
                                currentDist += INTERVAL;
                                continue;
                            }

                            Vector3d tangent =
                                cl.GetFirstDerivative(clPt).GetNormal();
                            Vector3d perpLeft =
                                new Vector3d(-tangent.Y, tangent.X, 0);
                            Vector3d perpRight =
                                new Vector3d(tangent.Y, -tangent.X, 0);

                            var leftDims = new List<DimResult>();
                            var rightDims = new List<DimResult>();

                            foreach (var kvp in layerEntities)
                            {
                                string lyrName = kvp.Key;
                                List<Curve> curves = kvp.Value;

                                // Left
                                Point3d? lPt = FindNearestIntersection(
                                    clPt,
                                    clPt + perpLeft * rayLength,
                                    curves);
                                if (lPt.HasValue)
                                {
                                    double dist = clPt.DistanceTo(lPt.Value);
                                    string key =
                                        $"{lyrName}|{Math.Round(dist, 0)}";
                                    if (!placedLeft.Contains(key))
                                        leftDims.Add(new DimResult
                                        {
                                            LayerName = lyrName,
                                            Distance = dist,
                                            IntersectPt = lPt.Value,
                                            DedupKey = key
                                        });
                                }

                                // Right
                                Point3d? rPt = FindNearestIntersection(
                                    clPt,
                                    clPt + perpRight * rayLength,
                                    curves);
                                if (rPt.HasValue)
                                {
                                    double dist = clPt.DistanceTo(rPt.Value);
                                    string key =
                                        $"{lyrName}|{Math.Round(dist, 0)}";
                                    if (!placedRight.Contains(key))
                                        rightDims.Add(new DimResult
                                        {
                                            LayerName = lyrName,
                                            Distance = dist,
                                            IntersectPt = rPt.Value,
                                            DedupKey = key
                                        });
                                }
                            }

                            leftDims = leftDims.OrderBy(d => d.Distance).ToList();
                            rightDims = rightDims.OrderBy(d => d.Distance).ToList();

                            //PlaceStackedDims(tr, mSpace,
                            //    clPt, tangent, perpLeft,
                            //    leftDims, dimStyleId,
                            //    STACK_GAP, BEYOND_OFFSET,
                            //    isLeft: true, rect,
                            //    placedLeft,
                            //    largestLeftAnchor);   // ← largest anchor

                            //PlaceStackedDims(tr, mSpace,
                            //    clPt, tangent, perpRight,
                            //    rightDims, dimStyleId,
                            //    STACK_GAP, BEYOND_OFFSET,
                            //    isLeft: false, rect,
                            //    placedRight,
                            //    largestRightAnchor);  // ← largest anchor

                            PlaceStackedDims(tr, mSpace,
                                clPt, tangent, perpLeft,
                                leftDims, dimStyleId,
                                STACK_GAP, BEYOND_OFFSET,
                                isLeft: true, rect,
                                placedLeft,
                                uniformLeftOffset);    // ← global uniform offset

                            PlaceStackedDims(tr, mSpace,
                                clPt, tangent, perpRight,
                                rightDims, dimStyleId,
                                STACK_GAP, BEYOND_OFFSET,
                                isLeft: false, rect,
                                placedRight,
                                uniformRightOffset);   // ← global uniform offset

                            currentDist += INTERVAL;
                        }

                        tr.Commit();
                        ed.WriteMessage("\n✅ AUTODIMRECT complete.");
                        // ── Erase the visual rectangle now that
                        //    dims are placed ──
                        EraseVisualRectangle(db, visualRectId);
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage(
                            $"\n❌ Error: {ex.Message}\n{ex.StackTrace}");
                        tr.Abort();
                    }
                }
            }
            finally
            {
                ed.Command("_.DIMASSOC", originalDimassoc);
                db.DimAssoc = originalDimassoc;
                ed.WriteMessage($"\nDIMASSC restored to: {originalDimassoc}");
            }
        }

        // ─────────────────────────────────────────────
        // ERASE VISUAL RECTANGLE
        // ─────────────────────────────────────────────
        private void EraseVisualRectangle(Database db, ObjectId rectId)
        {
            if (rectId == ObjectId.Null || !rectId.IsValid)
                return;

            using (Transaction tr =
                       db.TransactionManager.StartTransaction())
            {
                try
                {
                    DBObject obj = tr.GetObject(rectId, OpenMode.ForWrite);
                    if (obj != null && !obj.IsErased)
                        obj.Erase();

                    tr.Commit();
                }
                catch
                {
                    tr.Abort();
                }
            }
        }
        // ─────────────────────────────────────────────
        // COMPUTE ANCHOR OFFSET FROM LARGEST DISTANCE
        // ─────────────────────────────────────────────
        /// <summary>
        /// Mirrors the distValue formula in PlaceStackedDims but uses
        /// the largest distance on a side as the base, so every dim on
        /// that side uses a consistent anchor regardless of its own value.
        ///
        /// Formula (same as PlaceStackedDims):
        ///   rawDist    = rounded distance as string length × 3.5
        ///   suffixLen  = longest suffix in the map × 3.34
        ///   anchor     = ((rawDist + suffixLen) / 2) + largestMeasuredLen
        ///
        /// The "largestMeasuredLen" term is what replaces the old 43 —
        /// it is derived from the largest distance value itself so the
        /// text always sits proportionally beyond the longest dim line.
        /// </summary>
        private double ComputeAnchorOffset(
            double largestDist,
            Dictionary<string, LayerConfig> suffixMap)
        {
            if (largestDist <= 0) return 43.0; // safe fallback

            //// Reproduce how the dim number string length is estimated
            //string distStr = Math.Round(largestDist, 0).ToString();
            //double rawDistLen = distStr.Length * 3.5;

            //// Use the longest suffix in the map as the worst-case width
            //double maxTextWidth = suffixMap.Values.Max(c => c.TextWidth);
            //double measuredLenAnchor = rawDistLen;

            //return ((rawDistLen + maxTextWidth) / 2.0) + measuredLenAnchor;
            string distStr = Math.Round(largestDist, 0).ToString();
            double rawDistLen = distStr.Length * 3.5;

            return rawDistLen;
        }

        private double ComputeUniformOffset(
    List<DimResult> allDims,
    Dictionary<string, LayerConfig> suffixMap)
        {
            double maxOffset = 0.0;
            string maxLayerDbg = "";
            double maxDistDbg = 0.0;

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            foreach (DimResult dim in allDims)
            {
                LayerConfig cfg;
                if (!suffixMap.TryGetValue(dim.LayerName, out cfg))
                    continue;

                int roundedDist = (int)Math.Round(dim.Distance, 0);
                string distStr = roundedDist.ToString();
                double numWidth = distStr.Length * 3.5;
                double suffixWidth = cfg.TextWidth;
                //double totalTextWidth = numWidth + suffixWidth;
                //double offset = (totalTextWidth / 2.0) + totalTextWidth;
                double totalTextWidth = roundedDist;
                double offset = totalTextWidth;

                ed.WriteMessage(
                    $"\n   [{dim.LayerName}] dist={roundedDist} " +
                    $"numW={numWidth:F2} sufW={suffixWidth:F2} " +
                    $"total={totalTextWidth:F2} offset={offset:F2}");

                if (offset > maxOffset)
                {
                    maxOffset = offset;
                    maxLayerDbg = dim.LayerName;
                    maxDistDbg = dim.Distance;
                }
            }

            ed.WriteMessage(
                $"\n★ Winning layer: '{maxLayerDbg}' " +
                $"dist={maxDistDbg:F1} → uniformOffset={maxOffset:F2}");

            return maxOffset > 0 ? maxOffset : 43.0;
        }

        // ─────────────────────────────────────────────
        // PLACE STACKED ALIGNED DIMENSIONS
        // ─────────────────────────────────────────────
        private void PlaceStackedDims(
    Transaction tr,
    BlockTableRecord mSpace,
    Point3d clPt,
    Vector3d tangent,
    Vector3d perpDir,
    List<DimResult> dims,
    ObjectId dimStyleId,
    double stackGap,
    double beyondOffset,
    bool isLeft,
    Extents2d rect,
    HashSet<string> placedKeys,
    double uniformOffset)
        {
            if (!dims.Any()) return;

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            double maxDist = dims.Max(d => d.Distance);

            for (int i = 0; i < dims.Count; i++)
            {
                DimResult dim = dims[i];

                if (placedKeys.Contains(dim.DedupKey))
                    continue;

                LayerConfig config;
                if (!LayerSuffixMap.TryGetValue(dim.LayerName, out config))
                {
                    ed.WriteMessage(
                        $"\n⚠️ No config found for layer: '{dim.LayerName}'" +
                        $" — skipping.");
                    placedKeys.Add(dim.DedupKey);
                    continue;
                }

                double tangentShift = i * stackGap;
                Point3d stationPt = clPt + tangent * tangentShift;

                Point3d xLine1 = stationPt;
                Point3d xLine2 = ProjectToPerp(
                                        stationPt, perpDir, dim.IntersectPt);
                Point3d dimLinePt = stationPt
                                    + perpDir * (maxDist + beyondOffset);

                if (!PointInRect(xLine2, rect))
                    continue;

                double dist = Common_functions
                                      .GetDistanceBetweenPoints(xLine1, xLine2);
                double dimAngle = Common_functions
                                      .GetAngleBetweenPoints(xLine1, xLine2);

                int roundedDist = (int)Math.Round(dist, 0);
                string dimensionText = roundedDist.ToString() + config.Suffix;

                string distStr = roundedDist.ToString();
                double numWidth = distStr.Length * 2.5;
                double suffixWidth = config.TextWidth;
                double totalTextWidth = numWidth + suffixWidth;

                //MessageBox.Show("Suffix : " + config.Suffix + "\nSuffix length: " + config.TextWidth);
                AlignedDimension ad = new AlignedDimension(
                    xLine1, xLine2, dimLinePt,
                    string.Empty, dimStyleId);

                ad.DimensionText = dimensionText;

                ad.Layer = DIM_LAYER;
                ad.Dimasz = 2.75;
                ad.Dimtxt = 3.5;
                ad.Color = dim.LayerName.Equals(
                                "PROPOSED TRENCH",
                                StringComparison.OrdinalIgnoreCase)
                            ? Color.FromColorIndex(ColorMethod.ByAci, 30)
                            : Color.FromColorIndex(ColorMethod.ByBlock, 0);

                // ── All dims use the same globally computed offset ──
                // This guarantees text left-aligns across all stations.
                Point3d txtpnt = Common_functions
                                     .PolarPoint(xLine1, dimAngle, uniformOffset + totalTextWidth);

                //MessageBox.Show("Suffix : " + config.Suffix + "\nSuffix length: " + config.TextWidth + "Offset distance: " + uniformOffset);
                ad.TextPosition = txtpnt;
                ad.TextRotation = 0.0;

                mSpace.AppendEntity(ad);
                tr.AddNewlyCreatedDBObject(ad, true);
                ad.DimLinePoint = ad.DimLinePoint;

                placedKeys.Add(dim.DedupKey);
            }
        }


        // ─────────────────────────────────────────────
        // PICK FOUR POINTS
        // ─────────────────────────────────────────────
        private Point3d[] PickFourPoints(Editor ed)
        {
            var pts = new Point3d[4];
            string[] labels = { "1st", "2nd", "3rd", "4th" };

            for (int i = 0; i < 4; i++)
            {
                PromptPointOptions ppo = new PromptPointOptions(
                    $"\nPick {labels[i]} corner of work rectangle: ");
                ppo.AllowNone = false;

                if (i > 0)
                    ppo.BasePoint = pts[i - 1];

                PromptPointResult ppr = ed.GetPoint(ppo);
                if (ppr.Status != PromptStatus.OK) return null;

                pts[i] = ppr.Value;
            }

            return pts;
        }


        // ─────────────────────────────────────────────
        // DRAW VISUAL RECTANGLE ON SCREEN
        // ─────────────────────────────────────────────
        private ObjectId DrawVisualRectangle(Database db, Extents2d rect)
        {
            using (Transaction tr =
                       db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(
                    db.BlockTableId,
                    OpenMode.ForRead) as BlockTable;
                BlockTableRecord mSpace = tr.GetObject(
                    bt[BlockTableRecord.ModelSpace],
                    OpenMode.ForWrite) as BlockTableRecord;

                Polyline rp = new Polyline();
                rp.AddVertexAt(0,
                    new Point2d(rect.MinPoint.X, rect.MinPoint.Y), 0, 0, 0);
                rp.AddVertexAt(1,
                    new Point2d(rect.MaxPoint.X, rect.MinPoint.Y), 0, 0, 0);
                rp.AddVertexAt(2,
                    new Point2d(rect.MaxPoint.X, rect.MaxPoint.Y), 0, 0, 0);
                rp.AddVertexAt(3,
                    new Point2d(rect.MinPoint.X, rect.MaxPoint.Y), 0, 0, 0);
                rp.Closed = true;
                rp.Color = Color.FromColorIndex(ColorMethod.ByAci, 1);
                rp.Layer = "0";

                mSpace.AppendEntity(rp);
                tr.AddNewlyCreatedDBObject(rp, true);

                // ── Capture ObjectId before commit ──
                ObjectId rectId = rp.ObjectId;

                tr.Commit();

                return rectId;   // ← return so caller can erase it later
            }
        }


        // ─────────────────────────────────────────────
        // COMPUTE RAY LENGTH FROM MODEL EXTENTS
        // ─────────────────────────────────────────────
        private double ComputeRayLength(Database db)
        {
            try
            {
                Point3d eMin = db.Extmin;
                Point3d eMax = db.Extmax;

                double w = Math.Abs(eMax.X - eMin.X);
                double h = Math.Abs(eMax.Y - eMin.Y);
                double diagonal = Math.Sqrt(w * w + h * h);

                return Math.Max(diagonal * 2.0, 1000.0);
            }
            catch
            {
                return 100000.0;
            }
        }


        // ─────────────────────────────────────────────
        // FIND CENTERLINE INSIDE RECTANGLE
        // ─────────────────────────────────────────────
        private Polyline FindCenterlineInRect(
            Database db,
            Transaction tr,
            Editor ed,
            Extents2d rect)
        {
            BlockTable bt = tr.GetObject(
                db.BlockTableId, OpenMode.ForRead) as BlockTable;
            BlockTableRecord mSpace = tr.GetObject(
                bt[BlockTableRecord.ModelSpace],
                OpenMode.ForRead) as BlockTableRecord;

            Polyline rectPoly = BuildRectPolyline(rect);

            foreach (ObjectId id in mSpace)
            {
                Entity ent =
                    tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent == null) continue;

                if (ent.Layer.IndexOf(
                        "CENTERLINE",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                Polyline pl = ent as Polyline;
                if (pl == null) continue;

                // ── Test 1: Any vertex inside rectangle ──
                bool found = false;
                for (int v = 0; v < pl.NumberOfVertices; v++)
                {
                    Point2d vpt = pl.GetPoint2dAt(v);
                    if (PointInRect(new Point3d(vpt.X, vpt.Y, 0), rect))
                    {
                        found = true;
                        break;
                    }
                }
                if (found)
                {
                    rectPoly.Dispose();
                    return pl;
                }

                // ── Test 2: Segment crosses rectangle boundary ──
                var intPts = new Point3dCollection();
                try
                {
                    pl.IntersectWith(
                        rectPoly,
                        Intersect.OnBothOperands,
                        intPts,
                        IntPtr.Zero,
                        IntPtr.Zero);
                }
                catch { }

                if (intPts.Count > 0)
                {
                    rectPoly.Dispose();
                    return pl;
                }

                // ── Test 3: Rectangle fully inside centerline extents ──
                try
                {
                    Extents3d ext = pl.GeometricExtents;
                    if (ext.MinPoint.X <= rect.MinPoint.X &&
                        ext.MinPoint.Y <= rect.MinPoint.Y &&
                        ext.MaxPoint.X >= rect.MaxPoint.X &&
                        ext.MaxPoint.Y >= rect.MaxPoint.Y)
                    {
                        rectPoly.Dispose();
                        return pl;
                    }
                }
                catch { }
            }

            rectPoly.Dispose();
            return null;
        }


        // ─────────────────────────────────────────────
        // BUILD TEMPORARY RECTANGLE POLYLINE
        // ─────────────────────────────────────────────
        private Polyline BuildRectPolyline(Extents2d rect)
        {
            Polyline rp = new Polyline();
            rp.AddVertexAt(0,
                new Point2d(rect.MinPoint.X, rect.MinPoint.Y), 0, 0, 0);
            rp.AddVertexAt(1,
                new Point2d(rect.MaxPoint.X, rect.MinPoint.Y), 0, 0, 0);
            rp.AddVertexAt(2,
                new Point2d(rect.MaxPoint.X, rect.MaxPoint.Y), 0, 0, 0);
            rp.AddVertexAt(3,
                new Point2d(rect.MinPoint.X, rect.MaxPoint.Y), 0, 0, 0);
            rp.Closed = true;
            return rp;
        }


        // ─────────────────────────────────────────────
        // COLLECT LAYER ENTITIES IN RECTANGLE
        // ─────────────────────────────────────────────
        private Dictionary<string, List<Curve>> CollectLayerEntitiesInRect(
            Database db,
            Transaction tr,
            Extents2d rect)
        {
            var result = new Dictionary<string, List<Curve>>(
                StringComparer.OrdinalIgnoreCase);

            foreach (var key in LayerSuffixMap.Keys)
                result[key] = new List<Curve>();

            BlockTable bt = tr.GetObject(
                db.BlockTableId, OpenMode.ForRead) as BlockTable;
            BlockTableRecord mSpace = tr.GetObject(
                bt[BlockTableRecord.ModelSpace],
                OpenMode.ForRead) as BlockTableRecord;

            foreach (ObjectId id in mSpace)
            {
                Entity ent =
                    tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent == null) continue;
                if (!result.ContainsKey(ent.Layer)) continue;

                Curve c = ent as Curve;
                if (c == null) continue;

                try
                {
                    Extents3d ext = c.GeometricExtents;
                    if (ext.MaxPoint.X < rect.MinPoint.X ||
                        ext.MinPoint.X > rect.MaxPoint.X ||
                        ext.MaxPoint.Y < rect.MinPoint.Y ||
                        ext.MinPoint.Y > rect.MaxPoint.Y)
                        continue;
                }
                catch { continue; }

                result[ent.Layer].Add(c);
            }

            return result;
        }


        // ─────────────────────────────────────────────
        // FIND NEAREST INTERSECTION
        // ─────────────────────────────────────────────
        private Point3d? FindNearestIntersection(
            Point3d rayStart,
            Point3d rayEnd,
            List<Curve> curves)
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
                        curve.IntersectWith(
                            ray,
                            Intersect.ExtendThis,
                            pts,
                            IntPtr.Zero,
                            IntPtr.Zero);
                    }
                    catch { continue; }

                    foreach (Point3d pt in pts)
                    {
                        Vector3d toIntersect = pt - rayStart;
                        Vector3d rayDir = rayEnd - rayStart;

                        if (toIntersect.DotProduct(rayDir) < 0)
                            continue;

                        double d = rayStart.DistanceTo(pt);
                        if (d < minDist)
                        {
                            minDist = d;
                            nearest = pt;
                        }
                    }
                }
            }
            return nearest;
        }


        // ─────────────────────────────────────────────
        // PROJECT INTERSECTION TO PERP AT STATION
        // ─────────────────────────────────────────────
        private Point3d ProjectToPerp(
            Point3d stationPt,
            Vector3d perpDir,
            Point3d originalIntersect)
        {
            Vector3d toIntersect = originalIntersect - stationPt;
            double dist = toIntersect.DotProduct(perpDir);
            return stationPt + perpDir * dist;
        }


        // ─────────────────────────────────────────────
        // POINT IN RECT HELPER
        // ─────────────────────────────────────────────
        private bool PointInRect(Point3d pt, Extents2d rect)
        {
            return pt.X >= rect.MinPoint.X && pt.X <= rect.MaxPoint.X &&
                   pt.Y >= rect.MinPoint.Y && pt.Y <= rect.MaxPoint.Y;
        }


        // ─────────────────────────────────────────────
        // GET OR CREATE DIM STYLE
        // ─────────────────────────────────────────────
        private ObjectId GetOrCreateDimStyle(Database db, Transaction tr)
        {
            DimStyleTable dst = tr.GetObject(
                db.DimStyleTableId,
                OpenMode.ForRead) as DimStyleTable;

            DimStyleTableRecord dstr;

            if (dst.Has(DIM_STYLE_NAME))
            {
                dstr = tr.GetObject(
                           dst[DIM_STYLE_NAME],
                           OpenMode.ForWrite) as DimStyleTableRecord;
            }
            else
            {
                dst.UpgradeOpen();
                dstr = new DimStyleTableRecord { Name = DIM_STYLE_NAME };
                dst.Add(dstr);
                tr.AddNewlyCreatedDBObject(dstr, true);
            }

            // --- Lines & Arrows ---
            dstr.Dimscale = 100.0;
            dstr.Dimasz = 0.03;
            dstr.Dimexo = 0.625;
            dstr.Dimexe = 0.125;
            dstr.Dimse1 = false;
            dstr.Dimse2 = false;
            dstr.Dimdle = 0.0;
            dstr.Dimdli = 0.0;
            dstr.Dimclrd =
                Color.FromColorIndex(ColorMethod.ByBlock, 0);
            dstr.Dimclre =
                Color.FromColorIndex(ColorMethod.ByBlock, 0);

            // --- Text ---
            dstr.Dimtxt = 0.04;
            dstr.Dimgap = 0.09;
            dstr.Dimtad = 1;
            dstr.Dimtoh = true;
            dstr.Dimtih = false;
            dstr.Dimtvp = 0.0;
            dstr.Dimclrt =
                Color.FromColorIndex(ColorMethod.ByBlock, 0);

            // --- Text Style: ARIAL ---
            TextStyleTable tst = tr.GetObject(
                db.TextStyleTableId,
                OpenMode.ForRead) as TextStyleTable;

            if (tst.Has("ARIAL"))
                dstr.Dimtxsty = tst["ARIAL"];
            else if (tst.Has("Arial"))
                dstr.Dimtxsty = tst["Arial"];
            else
            {
                tst.UpgradeOpen();
                TextStyleTableRecord arialStyle =
                    new TextStyleTableRecord
                    {
                        Name = "ARIAL",
                        FileName = "arial.ttf",
                        TextSize = 0.0
                    };
                ObjectId arialId = tst.Add(arialStyle);
                tr.AddNewlyCreatedDBObject(arialStyle, true);
                dstr.Dimtxsty = arialId;
            }

            // --- Fit ---
            dstr.Dimatfit = 3;
            dstr.Dimtmove = 0;
            dstr.Dimtix = false;
            dstr.Dimsoxd = false;

            // --- Primary Units ---
            dstr.Dimdec = 0;
            dstr.Dimzin = 8;
            dstr.Dimrnd = 1.0;
            dstr.Dimlunit = 2;
            dstr.Dimdsep = '.';
            dstr.Dimlfac = 1.0;

            // --- Tolerances ---
            dstr.Dimtol = false;
            dstr.Dimlim = false;

            return dstr.ObjectId;
        }


        // ─────────────────────────────────────────────
        // ENSURE DIMENSIONS LAYER EXISTS
        // ─────────────────────────────────────────────
        private void EnsureLayerExists(Database db, Transaction tr)
        {
            LayerTable lt = tr.GetObject(
                db.LayerTableId,
                OpenMode.ForRead) as LayerTable;
            if (lt.Has(DIM_LAYER)) return;

            lt.UpgradeOpen();
            LayerTableRecord ltr = new LayerTableRecord
            {
                Name = DIM_LAYER,
                Color = Color.FromColorIndex(ColorMethod.ByAci, 7)
            };
            lt.Add(ltr);
            tr.AddNewlyCreatedDBObject(ltr, true);
        }
    }


    // ─────────────────────────────────────────────
    // HELPER CLASS
    // ─────────────────────────────────────────────
    public class DimResults
    {
        public string LayerName { get; set; }
        public double Distance { get; set; }
        public Point3d IntersectPt { get; set; }
        public string DedupKey { get; set; }
    }
}