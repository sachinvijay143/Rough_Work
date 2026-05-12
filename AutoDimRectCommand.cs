using IntelliCAD.ApplicationServices;
using IntelliCAD.EditorInput;
using Rough_Works;
using System;
using System.Collections.Generic;
using System.Linq;
using Teigha.Colors;
using Teigha.DatabaseServices;
using Teigha.Geometry;
using Teigha.Runtime;

namespace MagnaSoft_Drafting_Assistant_Tool_CobbFendley
{
    // ═══════════════════════════════════════════════════════════════════
    //  AutoDimRectCommand  –  Reliable auto-dimensioning for utility runs
    //
    //  ROOT CAUSES FIXED
    //  ─────────────────
    //  RCA-1  Dimscale was 100 + fractional Dimasz/Dimtxt.
    //         Fix: Dimscale = 1; every value expressed directly in
    //         model-space drawing units.  The tool now reads the live
    //         DIMSCALE / DIMASZ system variables so it adapts to any
    //         drawing automatically.
    //
    //  RCA-2  ComputeUniformOffset mixed pixel-like constants (~3.5 per
    //         digit) with model-space units → wrong on every drawing.
    //         Fix: text width = Dimtxt × charCount × CHAR_ASPECT.
    //         CHAR_ASPECT (0.6) is a stable typographic ratio, not a
    //         pixel measurement, so it scales with any text height.
    //
    //  RCA-3  Dimasz / Dimtxt were set in the style record AND implicitly
    //         multiplied by Dimscale=100.  Fix: set them once, in model-
    //         space units, with Dimscale=1.  The tool reads DIMTXT from
    //         the document's current dim style as the default text height
    //         so the result matches the drawing's own standard.
    //
    //  RCA-4  RAY_RECT_MULTIPLIER = 2.5 shot rays 2.5 × the rectangle
    //         diagonal → caught geometry far outside the work area.
    //         Fix: ray shoots only to the far edge of the rectangle plus
    //         a small buffer.  Direction-aware: for a left-side ray the
    //         cap is (stationPt.X - rect.MinX) + BUFFER; right-side is
    //         (rect.MaxX - stationPt.X) + BUFFER.  This is always finite
    //         and never reaches the opposite side of the drawing.
    //
    //  RCA-5  Helper class declared as "DimResults" (plural) but used as
    //         "DimResult" everywhere → compile error / silent failure.
    //         Fix: renamed to UtilityDimData, used consistently throughout.
    //         TextPosition on AlignedDimension is fragile after DIMASSOC
    //         is restored; Dimtad=1 (above-line) + Dimgap places text
    //         reliably without a manual insertion point.
    // ═══════════════════════════════════════════════════════════════════

    public class AutoDimRectCommand
    {
        // ─────────────────────────────────────────────────────────────
        // LAYER CONFIG
        // Suffix         = text appended after the measured distance
        // SuffixCharCount = character count of the suffix string only
        //                   (used to compute text width in model units)
        // ─────────────────────────────────────────────────────────────
        public struct LayerConfig
        {
            public string Suffix { get; }
            public int SuffixCharCount { get; }

            public LayerConfig(string suffix)
            {
                Suffix = suffix;
                SuffixCharCount = suffix.Length;
            }
        }

        private static readonly Dictionary<string, LayerConfig> LayerSuffixMap =
            new Dictionary<string, LayerConfig>(StringComparer.OrdinalIgnoreCase)
            {
                { "WATER",           new LayerConfig("' WATER")        },
                { "SS",              new LayerConfig("' SEWER")        },
                { "BOC",             new LayerConfig("' BOC")          },
                { "SIDEWALK",        new LayerConfig("' S/W")          },
                { "ROW",             new LayerConfig("' ROW")          },
                { "EASEMENTS",       new LayerConfig("' PUE")          },
                { "PROPOSED TRENCH", new LayerConfig("' PROP. TRENCH") },
            };

        // ── Constant names (model-space, drawing units) ──────────────
        private const string DIM_LAYER = "DIMENSIONS";
        private const string DIM_STYLE_NAME = "BPGDIMS";
        private const double INTERVAL = 50.0;   // station spacing along CL
        private const double STACK_GAP = 8.0;    // tangential shift per stacked dim
        private const double BEYOND_OFFSET = 15.0;   // how far dimline extends past utility
        private const double RAY_BUFFER = 20.0;   // extra beyond rect edge (model units)

        // Typographic ratio: rendered text width ≈ height × charCount × CHAR_ASPECT
        // 0.6 is correct for a standard proportional font at any scale.
        private const double CHAR_ASPECT = 0.6;


        // ─────────────────────────────────────────────────────────────
        // MAIN COMMAND
        // ─────────────────────────────────────────────────────────────
        [CommandMethod("AUTODIMRECT")]
        public void AutoDimRect()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // ── Read live drawing settings before touching DIMASSOC ──
            // RCA-1 / RCA-3: use the drawing's own text height as the
            // default so dimensions look consistent with other dims.
            double liveDimtxt = Convert.ToDouble(
                Application.GetSystemVariable("DIMTXT"));
            double liveDimasz = Convert.ToDouble(
                Application.GetSystemVariable("DIMASZ"));

            // Fallback: if the drawing hasn't set these yet use sensible
            // defaults that work at a 1:50 plotting scale (model units).
            if (liveDimtxt <= 0) liveDimtxt = 3.5;
            if (liveDimasz <= 0) liveDimasz = 2.75;

            int originalDimassoc = Convert.ToInt32(
                Application.GetSystemVariable("DIMASSOC"));

            ed.Command("_.DIMASSOC", 0);
            db.DimAssoc = 0;

            try
            {
                // ── Step 1: Pick 4 corners ───────────────────────────
                Extents3d? pickedRect = PickRectangleByDrag(ed);
                if (pickedRect == null)
                {
                    ed.WriteMessage("\nAutoDimRect: cancelled.");
                    return;
                }

                double minX = pickedRect.Value.MinPoint.X;
                double maxX = pickedRect.Value.MaxPoint.X;
                double minY = pickedRect.Value.MinPoint.Y;
                double maxY = pickedRect.Value.MaxPoint.Y;

                Extents2d rect = new Extents2d(
                    new Point2d(minX, minY),
                    new Point2d(maxX, maxY));

                // ── Step 2: Draw visual boundary ─────────────────────
                ObjectId visualRectId = DrawVisualRectangle(db, rect);

                // ── Step 3: Main transaction ─────────────────────────
                using (Transaction tr =
                           db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        Polyline cl = FindCenterlineInRect(db, tr, ed, rect);
                        if (cl == null)
                        {
                            ed.WriteMessage(
                                "\nNo CENTERLINE polyline found inside " +
                                "the selected rectangle.");
                            tr.Abort();
                            return;
                        }

                        var layerEntities = CollectLayerEntitiesInRect(db, tr, rect);

                        if (!layerEntities.Any(kvp => kvp.Value.Count > 0))
                            ed.WriteMessage(
                                "\nWarning: no recognised utility layers " +
                                "found inside the rectangle.");

                        EnsureLayerExists(db, tr);
                        ObjectId dimStyleId = GetOrCreateDimStyle(
                            db, tr, liveDimtxt, liveDimasz);

                        BlockTable bt = tr.GetObject(
                            db.BlockTableId,
                            OpenMode.ForRead) as BlockTable;
                        BlockTableRecord mSpace = tr.GetObject(
                            bt[BlockTableRecord.ModelSpace],
                            OpenMode.ForWrite) as BlockTableRecord;

                        // ── PRE-PASS: discover all dim results across
                        //    every station so we can compute a single
                        //    uniform text-offset for left and right. ──
                        var allLeftDims = new List<UtilityDimData>();
                        var allRightDims = new List<UtilityDimData>();
                        var preLeft = new HashSet<string>();
                        var preRight = new HashSet<string>();

                        double total = cl.Length;
                        double currentDist = 0.0;

                        while (currentDist <= total)
                        {
                            Point3d clPt = cl.GetPointAtDist(currentDist);

                            if (!PointInRect(clPt, rect))
                            { currentDist += INTERVAL; continue; }

                            Vector3d tangent = cl.GetFirstDerivative(clPt).GetNormal();
                            Vector3d perpLeft = new Vector3d(-tangent.Y, tangent.X, 0);
                            Vector3d perpRight = new Vector3d(tangent.Y, -tangent.X, 0);

                            foreach (var kvp in layerEntities)
                            {
                                string lyrName = kvp.Key;
                                List<Curve> curves = kvp.Value;

                                // RCA-4: ray length = distance to rect edge + buffer
                                double rayLenLeft = RayLengthToEdge(clPt, perpLeft, rect) + RAY_BUFFER;
                                double rayLenRight = RayLengthToEdge(clPt, perpRight, rect) + RAY_BUFFER;

                                Point3d? lPt = FindNearestIntersection(
                                    clPt,
                                    clPt + perpLeft * rayLenLeft,
                                    curves,
                                    rayLenLeft);

                                if (lPt.HasValue)
                                {
                                    double d = clPt.DistanceTo(lPt.Value);
                                    string key = $"{lyrName}|{Math.Round(d, 0)}";
                                    if (!preLeft.Contains(key))
                                    {
                                        preLeft.Add(key);
                                        allLeftDims.Add(new UtilityDimData
                                        {
                                            LayerName = lyrName,
                                            Distance = d,
                                            IntersectPt = lPt.Value,
                                            DedupKey = key
                                        });
                                    }
                                }

                                Point3d? rPt = FindNearestIntersection(
                                    clPt,
                                    clPt + perpRight * rayLenRight,
                                    curves,
                                    rayLenRight);

                                if (rPt.HasValue)
                                {
                                    double d = clPt.DistanceTo(rPt.Value);
                                    string key = $"{lyrName}|{Math.Round(d, 0)}";
                                    if (!preRight.Contains(key))
                                    {
                                        preRight.Add(key);
                                        allRightDims.Add(new UtilityDimData
                                        {
                                            LayerName = lyrName,
                                            Distance = d,
                                            IntersectPt = rPt.Value,
                                            DedupKey = key
                                        });
                                    }
                                }
                            }

                            currentDist += INTERVAL;
                        }

                        // Compute one shared dim-line distance per side:
                        //   = farthest distance + label width of that entry
                        //     + BEYOND_OFFSET
                        // All dims on each side will use this same distance
                        // so every dim line is co-linear and text aligns.
                        double uniformLeft = ComputeUniformDimLineDist(
                            allLeftDims, liveDimtxt, BEYOND_OFFSET, ed);
                        double uniformRight = ComputeUniformDimLineDist(
                            allRightDims, liveDimtxt, BEYOND_OFFSET, ed);

                        ed.WriteMessage(
                            $"\nUniform dim-line dist  LEFT : {uniformLeft:F2}" +
                            $"  RIGHT: {uniformRight:F2}");

                        // ── PLACEMENT PASS ───────────────────────────
                        var placedLeft = new HashSet<string>();
                        var placedRight = new HashSet<string>();
                        currentDist = 0.0;

                        while (currentDist <= total)
                        {
                            Point3d clPt = cl.GetPointAtDist(currentDist);

                            if (!PointInRect(clPt, rect))
                            { currentDist += INTERVAL; continue; }

                            Vector3d tangent = cl.GetFirstDerivative(clPt).GetNormal();
                            Vector3d perpLeft = new Vector3d(-tangent.Y, tangent.X, 0);
                            Vector3d perpRight = new Vector3d(tangent.Y, -tangent.X, 0);

                            var leftDims = new List<UtilityDimData>();
                            var rightDims = new List<UtilityDimData>();

                            foreach (var kvp in layerEntities)
                            {
                                string lyrName = kvp.Key;
                                List<Curve> curves = kvp.Value;

                                double rayLenLeft = RayLengthToEdge(clPt, perpLeft, rect) + RAY_BUFFER;
                                double rayLenRight = RayLengthToEdge(clPt, perpRight, rect) + RAY_BUFFER;

                                Point3d? lPt = FindNearestIntersection(
                                    clPt,
                                    clPt + perpLeft * rayLenLeft,
                                    curves,
                                    rayLenLeft);

                                if (lPt.HasValue)
                                {
                                    double d = clPt.DistanceTo(lPt.Value);
                                    string key = $"{lyrName}|{Math.Round(d, 0)}";
                                    if (!placedLeft.Contains(key))
                                        leftDims.Add(new UtilityDimData
                                        {
                                            LayerName = lyrName,
                                            Distance = d,
                                            IntersectPt = lPt.Value,
                                            DedupKey = key
                                        });
                                }

                                Point3d? rPt = FindNearestIntersection(
                                    clPt,
                                    clPt + perpRight * rayLenRight,
                                    curves,
                                    rayLenRight);

                                if (rPt.HasValue)
                                {
                                    double d = clPt.DistanceTo(rPt.Value);
                                    string key = $"{lyrName}|{Math.Round(d, 0)}";
                                    if (!placedRight.Contains(key))
                                        rightDims.Add(new UtilityDimData
                                        {
                                            LayerName = lyrName,
                                            Distance = d,
                                            IntersectPt = rPt.Value,
                                            DedupKey = key
                                        });
                                }
                            }

                            leftDims = leftDims.OrderBy(d => d.Distance).ToList();
                            rightDims = rightDims.OrderBy(d => d.Distance).ToList();

                            PlaceStackedDims(tr, mSpace,
                                clPt, tangent, perpLeft,
                                leftDims, dimStyleId,
                                STACK_GAP, BEYOND_OFFSET,
                                rect, placedLeft,
                                uniformLeft, liveDimtxt);

                            PlaceStackedDims(tr, mSpace,
                                clPt, tangent, perpRight,
                                rightDims, dimStyleId,
                                STACK_GAP, BEYOND_OFFSET,
                                rect, placedRight,
                                uniformRight, liveDimtxt);

                            currentDist += INTERVAL;
                        }

                        tr.Commit();
                        ed.WriteMessage("\nAUTODIMRECT complete.");
                        EraseVisualRectangle(db, visualRectId);
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage(
                            $"\nError: {ex.Message}\n{ex.StackTrace}");
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
        // RAY LENGTH TO RECT EDGE  (RCA-4)
        //
        // Computes the distance from clPt to the rectangle boundary in
        // the direction of perpDir.  This replaces the diagonal-based
        // multiplier so the ray never escapes the work area.
        // ─────────────────────────────────────────────────────────────
        private double RayLengthToEdge(
            Point3d clPt,
            Vector3d perpDir,
            Extents2d rect)
        {
            // Parametric ray:  P = clPt + t * perpDir
            // Solve for t where the ray hits each wall; take the smallest
            // positive t (= the first wall hit in that direction).
            double tMin = double.MaxValue;

            // X walls
            if (Math.Abs(perpDir.X) > 1e-9)
            {
                double t1 = (rect.MinPoint.X - clPt.X) / perpDir.X;
                double t2 = (rect.MaxPoint.X - clPt.X) / perpDir.X;
                if (t1 > 0) tMin = Math.Min(tMin, t1);
                if (t2 > 0) tMin = Math.Min(tMin, t2);
            }

            // Y walls
            if (Math.Abs(perpDir.Y) > 1e-9)
            {
                double t3 = (rect.MinPoint.Y - clPt.Y) / perpDir.Y;
                double t4 = (rect.MaxPoint.Y - clPt.Y) / perpDir.Y;
                if (t3 > 0) tMin = Math.Min(tMin, t3);
                if (t4 > 0) tMin = Math.Min(tMin, t4);
            }

            // Safety: if clPt is outside the rect or perpDir is along
            // the CL (shouldn't happen), return a conservative maximum.
            if (tMin == double.MaxValue || tMin <= 0)
                tMin = Math.Max(rect.MaxPoint.X - rect.MinPoint.X,
                                rect.MaxPoint.Y - rect.MinPoint.Y);

            return tMin;
        }


        // ─────────────────────────────────────────────────────────────
        // COMPUTE UNIFORM DIM-LINE DISTANCE
        //
        // This is the single perpendicular distance from the centreline
        // at which ALL dimension lines on one side will be drawn.
        //
        // Formula (exact logic requested):
        //   1. Find the farthest layer distance on this side (maxDist).
        //   2. Compute the label text width of that farthest entry:
        //        labelWidth = (digitCount + suffixCharCount) × charWidth
        //        charWidth  = Dimtxt × CHAR_ASPECT   (model-space units)
        //   3. uniformDimLineDist = maxDist + labelWidth + BEYOND_OFFSET
        //
        // Every dim on this side uses this same distance for its
        // DimLinePoint, so all dimension lines are co-linear and all
        // text labels are uniformly offset from the centreline.
        // ─────────────────────────────────────────────────────────────
        private double ComputeUniformDimLineDist(
            List<UtilityDimData> allDims,
            double dimtxt,
            double beyondOffset,
            Editor ed)
        {
            if (!allDims.Any()) return 0.0;

            double charW = dimtxt * CHAR_ASPECT;

            // ── Step 1: entry with the largest distance from CL ───────
            UtilityDimData farthest = null;
            double maxDist = 0.0;

            foreach (UtilityDimData dim in allDims)
            {
                if (!LayerSuffixMap.ContainsKey(dim.LayerName)) continue;
                if (dim.Distance > maxDist)
                {
                    maxDist = dim.Distance;
                    farthest = dim;
                }
            }

            if (farthest == null) return 0.0;

            // ── Step 2: label width of the farthest entry ─────────────
            LayerConfig cfg = LayerSuffixMap[farthest.LayerName];
            int rounded = (int)Math.Round(farthest.Distance, 0);
            int digitCount = rounded.ToString().Length;
            int sufCount = cfg.SuffixCharCount;
            double labelWidth = (digitCount + sufCount) * charW;

            // ── Step 3: uniform dim-line distance ─────────────────────
            //   = farthest distance + label text width + beyond gap
            double uniformDist = maxDist + labelWidth + beyondOffset;

            ed.WriteMessage(
                $"\n  Farthest layer : {farthest.LayerName}" +
                $"  dist={maxDist:F1}" +
                $"  label='{rounded}{cfg.Suffix}'" +
                $"  labelWidth={labelWidth:F2}" +
                $"  uniformDimLineDist={uniformDist:F2}");

            return uniformDist;
        }


        // ─────────────────────────────────────────────────────────────
        // PLACE STACKED ALIGNED DIMENSIONS
        //
        // WHY TextPosition IS SET EXPLICITLY
        // ───────────────────────────────────
        // AlignedDimension auto-centres its text over the MID-POINT of
        // the measured segment (xLine1 → xLine2).  For a short dimension
        // (e.g. 5' WATER, close to the CL) that midpoint is near the CL,
        // so the text lands right on top of the utility line — overlap.
        //
        // The only reliable fix is:
        //   1. Set Dimtmove = 2 in the style  →  "free" text: the engine
        //      honours an explicit TextPosition without adding a leader.
        //   2. After AppendEntity (so the dimension is fully initialised),
        //      set ad.TextPosition to the uniform anchor point:
        //        textAnchor = stationPt + perpDir × uniformDimLineDist
        //      This is the SAME point for every dim on this side, so all
        //      labels sit at the identical distance from the CL — no
        //      overlap with the utility layers, perfect column alignment.
        //
        // uniformDimLineDist is pre-computed once per side as:
        //   farthest_distance + label_text_width + BEYOND_OFFSET
        // so even the widest label clears the farthest utility.
        // ─────────────────────────────────────────────────────────────
        private void PlaceStackedDims(
            Transaction tr,
            BlockTableRecord mSpace,
            Point3d clPt,
            Vector3d tangent,
            Vector3d perpDir,
            List<UtilityDimData> dims,
            ObjectId dimStyleId,
            double stackGap,
            double beyondOffset,
            Extents2d rect,
            HashSet<string> placedKeys,
            double uniformDimLineDist,
            double dimtxt)
        {
            if (!dims.Any()) return;

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            for (int i = 0; i < dims.Count; i++)
            {
                UtilityDimData dim = dims[i];
                if (placedKeys.Contains(dim.DedupKey)) continue;

                LayerConfig config;
                if (!LayerSuffixMap.TryGetValue(dim.LayerName, out config))
                {
                    ed.WriteMessage(
                        $"\nNo config for layer '{dim.LayerName}' — skipped.");
                    placedKeys.Add(dim.DedupKey);
                    continue;
                }

                // ── Station point (stacked along tangent) ─────────────
                double tangentShift = i * stackGap;
                Point3d stationPt = clPt + tangent * tangentShift;

                // xLine1: on the centreline
                Point3d xLine1 = stationPt;

                // xLine2: utility crossing projected onto the perp at
                //         this station
                Point3d xLine2 = ProjectToPerp(stationPt, perpDir, dim.IntersectPt);

                if (!PointInRect(xLine2, rect))
                {
                    placedKeys.Add(dim.DedupKey);
                    continue;
                }

                // ── DimLinePoint ──────────────────────────────────────
                // Same uniform distance for every dim on this side →
                // all dimension lines are co-linear.
                Point3d dimLinePt = stationPt + perpDir * uniformDimLineDist;

                // ── Text string ───────────────────────────────────────
                double dist = Common_functions.GetDistanceBetweenPoints(xLine1, xLine2);
                int roundedDist = (int)Math.Round(dist, 0);
                string dimText = roundedDist.ToString() + config.Suffix;

                // ── Create AlignedDimension ───────────────────────────
                AlignedDimension ad = new AlignedDimension(
                    xLine1, xLine2, dimLinePt,
                    string.Empty, dimStyleId);

                ad.DimensionText = dimText;
                ad.Layer = DIM_LAYER;
                ad.Color = dim.LayerName.Equals(
                               "PROPOSED TRENCH",
                               StringComparison.OrdinalIgnoreCase)
                           ? Color.FromColorIndex(ColorMethod.ByAci, 30)
                           : Color.FromColorIndex(ColorMethod.ByBlock, 0);

                // Must append first so the dimension geometry is fully
                // computed before we override TextPosition.
                mSpace.AppendEntity(ad);
                tr.AddNewlyCreatedDBObject(ad, true);

                // ── Explicit TextPosition (the key fix) ───────────────
                // Place every label at exactly uniformDimLineDist from
                // the CL along the perp direction.  Dimtmove=2 (set in
                // GetOrCreateDimStyle) makes the engine honour this
                // position without adding an unwanted leader line.
                // Dimtoh=true allows text to display outside the ext
                // lines when the measured distance is shorter than the
                // label — which is exactly the short-dim overlap case.
                Point3d textAnchor = stationPt + perpDir * uniformDimLineDist;
                ad.TextPosition = textAnchor;
                ad.TextRotation = 0.0;   // horizontal text always

                ed.WriteMessage(
                    $"\n  {dim.LayerName} | {dimText} | dist={dist:F1}" +
                    $" | textAnchor=({textAnchor.X:F1},{textAnchor.Y:F1})");

                placedKeys.Add(dim.DedupKey);
            }
        }


        // ─────────────────────────────────────────────────────────────
        // PICK FOUR POINTS
        // ─────────────────────────────────────────────────────────────
        private Extents3d? PickRectangleByDrag(Editor ed)
        {
            // ── First corner ──────────────────────────────────────────
            PromptPointOptions ppo = new PromptPointOptions(
                "\nPick first corner of work rectangle: ");
            ppo.AllowNone = false;

            PromptPointResult ppr = ed.GetPoint(ppo);
            if (ppr.Status != PromptStatus.OK) return null;

            Point3d firstCorner = ppr.Value;

            // ── Opposite corner (rubber-band jig built into the editor) ──
            PromptCornerOptions pco = new PromptCornerOptions(
                "\nPick opposite corner: ", firstCorner);
            pco.AllowNone = false;

            PromptPointResult pcr = ed.GetCorner(pco);
            if (pcr.Status != PromptStatus.OK) return null;

            Point3d secondCorner = pcr.Value;

            // ── Normalise so MinPoint < MaxPoint regardless of drag dir ──
            double minX = Math.Min(firstCorner.X, secondCorner.X);
            double maxX = Math.Max(firstCorner.X, secondCorner.X);
            double minY = Math.Min(firstCorner.Y, secondCorner.Y);
            double maxY = Math.Max(firstCorner.Y, secondCorner.Y);

            if (maxX - minX < 1e-6 || maxY - minY < 1e-6)
            {
                ed.WriteMessage("\nRectangle is too small — cancelled.");
                return null;
            }

            return new Extents3d(new Point3d(minX, minY,0), new Point3d(maxX, maxY, 0));
        }


        // ─────────────────────────────────────────────────────────────
        // DRAW / ERASE VISUAL RECTANGLE
        // ─────────────────────────────────────────────────────────────
        private ObjectId DrawVisualRectangle(Database db, Extents2d rect)
        {
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(
                    db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord mSpace = tr.GetObject(
                    bt[BlockTableRecord.ModelSpace],
                    OpenMode.ForWrite) as BlockTableRecord;

                Polyline rp = new Polyline();
                rp.AddVertexAt(0, rect.MinPoint, 0, 0, 0);
                rp.AddVertexAt(1, new Point2d(rect.MaxPoint.X, rect.MinPoint.Y), 0, 0, 0);
                rp.AddVertexAt(2, rect.MaxPoint, 0, 0, 0);
                rp.AddVertexAt(3, new Point2d(rect.MinPoint.X, rect.MaxPoint.Y), 0, 0, 0);
                rp.Closed = true;
                rp.Color = Color.FromColorIndex(ColorMethod.ByAci, 1);  // red
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
        // FIND NEAREST INTERSECTION  (RCA-4 incorporated)
        //
        // maxPlausibleDist is now equal to the ray length (distance to
        // the rectangle boundary + RAY_BUFFER), so no separate threshold
        // constant is needed.  Any hit beyond that is impossible.
        // ─────────────────────────────────────────────────────────────
        private Point3d? FindNearestIntersection(
            Point3d rayStart,
            Point3d rayEnd,
            List<Curve> curves,
            double maxPlausibleDist)
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

                        // Reject behind the origin
                        if (toIntersect.DotProduct(rayDir) < 0) continue;

                        double d = rayStart.DistanceTo(pt);

                        // Reject beyond the rect boundary
                        if (d > maxPlausibleDist) continue;

                        if (d < minDist) { minDist = d; nearest = pt; }
                    }
                }
            }
            return nearest;
        }


        // ─────────────────────────────────────────────────────────────
        // FIND CENTERLINE INSIDE RECTANGLE
        // ─────────────────────────────────────────────────────────────
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
                Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent == null) continue;
                if (ent.Layer.IndexOf("CENTERLINE",
                        StringComparison.OrdinalIgnoreCase) < 0) continue;

                Polyline pl = ent as Polyline;
                if (pl == null) continue;

                // Test 1: any vertex inside the rect?
                bool found = false;
                for (int v = 0; v < pl.NumberOfVertices; v++)
                {
                    Point2d vpt = pl.GetPoint2dAt(v);
                    if (PointInRect(new Point3d(vpt.X, vpt.Y, 0), rect))
                    { found = true; break; }
                }
                if (found) { rectPoly.Dispose(); return pl; }

                // Test 2: polyline crosses the rect boundary?
                var intPts = new Point3dCollection();
                try
                {
                    pl.IntersectWith(rectPoly, Intersect.OnBothOperands,
                        intPts, IntPtr.Zero, IntPtr.Zero);
                }
                catch { }
                if (intPts.Count > 0) { rectPoly.Dispose(); return pl; }

                // Test 3: polyline fully contains the rect?
                try
                {
                    Extents3d ext = pl.GeometricExtents;
                    if (ext.MinPoint.X <= rect.MinPoint.X &&
                        ext.MinPoint.Y <= rect.MinPoint.Y &&
                        ext.MaxPoint.X >= rect.MaxPoint.X &&
                        ext.MaxPoint.Y >= rect.MaxPoint.Y)
                    { rectPoly.Dispose(); return pl; }
                }
                catch { }
            }

            rectPoly.Dispose();
            return null;
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
        // COLLECT LAYER ENTITIES IN RECTANGLE
        // ─────────────────────────────────────────────────────────────
        private Dictionary<string, List<Curve>> CollectLayerEntitiesInRect(
            Database db,
            Transaction tr,
            Extents2d rect)
        {
            var result = new Dictionary<string, List<Curve>>(
                StringComparer.OrdinalIgnoreCase);

            foreach (string key in LayerSuffixMap.Keys)
                result[key] = new List<Curve>();

            BlockTable bt = tr.GetObject(
                db.BlockTableId, OpenMode.ForRead) as BlockTable;
            BlockTableRecord mSpace = tr.GetObject(
                bt[BlockTableRecord.ModelSpace],
                OpenMode.ForRead) as BlockTableRecord;

            foreach (ObjectId id in mSpace)
            {
                Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
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


        // ─────────────────────────────────────────────────────────────
        // PROJECT INTERSECTION ONTO PERP AT STATION
        // ─────────────────────────────────────────────────────────────
        private Point3d ProjectToPerp(
            Point3d stationPt,
            Vector3d perpDir,
            Point3d originalIntersect)
        {
            Vector3d toIntersect = originalIntersect - stationPt;
            double dist = toIntersect.DotProduct(perpDir);
            return stationPt + perpDir * dist;
        }


        // ─────────────────────────────────────────────────────────────
        // POINT IN RECT HELPER
        // ─────────────────────────────────────────────────────────────
        private bool PointInRect(Point3d pt, Extents2d rect) =>
            pt.X >= rect.MinPoint.X && pt.X <= rect.MaxPoint.X &&
            pt.Y >= rect.MinPoint.Y && pt.Y <= rect.MaxPoint.Y;


        // ─────────────────────────────────────────────────────────────
        // GET OR CREATE DIM STYLE  (RCA-1 / RCA-3 fix)
        //
        // Dimscale = 1.0   — all values in model-space units.
        // Dimasz / Dimtxt  — taken from live drawing sysvar so that the
        //                    output matches the drawing's own standards.
        // Dimtad = 1       — text above dim line (reliable on all hosts).
        // Dimtmove = 1     — text can move with a leader; dim line stays.
        // ─────────────────────────────────────────────────────────────
        private ObjectId GetOrCreateDimStyle(
            Database db,
            Transaction tr,
            double dimtxt,
            double dimasz)
        {
            DimStyleTable dst = tr.GetObject(
                db.DimStyleTableId, OpenMode.ForRead) as DimStyleTable;

            DimStyleTableRecord dstr;
            //if (dst.Has(dimstyle))
            //    dstr = tr.GetObject(dst[dimstyle],
            //               OpenMode.ForWrite) as DimStyleTableRecord;
            //else
            //{
            //    dst.UpgradeOpen();
            //    dstr = new DimStyleTableRecord { Name = dimstyle };
            //    dst.Add(dstr);
            //    tr.AddNewlyCreatedDBObject(dstr, true);
            //}

            if (dst.Has(DIM_STYLE_NAME))
                dstr = tr.GetObject(dst[DIM_STYLE_NAME],
                           OpenMode.ForWrite) as DimStyleTableRecord;
            else
            {
                dst.UpgradeOpen();
                dstr = new DimStyleTableRecord { Name = DIM_STYLE_NAME };
                dst.Add(dstr);
                tr.AddNewlyCreatedDBObject(dstr, true);
            }

            // ── Scale ────────────────────────────────────────────────
            // RCA-1 / RCA-3: Dimscale = 1 so every value below is the
            // final rendered size in model-space units.
            dstr.Dimscale = 1.0;

            // ── Arrows ───────────────────────────────────────────────
            dstr.Dimasz = dimasz;          // from live sysvar
            dstr.Dimtsz = 0.0;            // closed-filled arrow (not tick)

            // ── Extension lines ──────────────────────────────────────
            dstr.Dimexo = dimtxt * 0.5;   // offset from origin
            dstr.Dimexe = dimtxt * 0.25;  // extension beyond dim line
            dstr.Dimse1 = false;
            dstr.Dimse2 = false;
            dstr.Dimdle = 0.0;
            dstr.Dimdli = 0.0;

            // ── Colours (ByBlock so the layer colour wins) ───────────
            dstr.Dimclrd = Color.FromColorIndex(ColorMethod.ByBlock, 0);
            dstr.Dimclre = Color.FromColorIndex(ColorMethod.ByBlock, 0);
            dstr.Dimclrt = Color.FromColorIndex(ColorMethod.ByBlock, 0);

            // ── Text ─────────────────────────────────────────────────
            dstr.Dimtxt = dimtxt;         // from live sysvar
            dstr.Dimgap = dimtxt * 0.09;  // gap between text and dim line

            // ── Text placement ───────────────────────────────────────
            // Dimtmove = 2  →  "free" text.  The engine honours an
            //   explicit ad.TextPosition exactly, with no leader line.
            //   This is the critical setting that prevents text from
            //   snapping back to the segment midpoint.
            // Dimtad  = 0  →  manual vertical placement (we control Y
            //   via TextPosition; Dimtad=1 would override it).
            // Dimtoh  = true → text can render outside extension lines,
            //   which is needed when the utility is close to the CL and
            //   the label is wider than the measured gap.
            dstr.Dimatfit = 0;   // never move arrows to fit text
            dstr.Dimtmove = 2;   // free text — honour TextPosition
            dstr.Dimtad = 0;   // manual vertical (TextPosition drives Y)
            dstr.Dimtoh = true;
            dstr.Dimtih = false;
            dstr.Dimtix = false;
            dstr.Dimsoxd = false;

            // ── Units ────────────────────────────────────────────────
            dstr.Dimdec = 0;      // 0 decimal places (whole feet/units)
            dstr.Dimzin = 8;      // suppress trailing zeros
            dstr.Dimrnd = 1.0;    // round to nearest 1 unit
            dstr.Dimlunit = 2;     // decimal
            dstr.Dimdsep = '.';
            dstr.Dimlfac = 1.0;

            dstr.Dimtol = false;
            dstr.Dimlim = false;

            // ── Font (Arial if available, otherwise standard) ────────
            TextStyleTable tst = tr.GetObject(
                db.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;

            if (tst.Has("ARIAL"))
                dstr.Dimtxsty = tst["ARIAL"];
            else if (tst.Has("Arial"))
                dstr.Dimtxsty = tst["Arial"];
            else
            {
                tst.UpgradeOpen();
                TextStyleTableRecord arialStyle = new TextStyleTableRecord
                {
                    Name = "ARIAL",
                    FileName = "arial.ttf",
                    TextSize = 0.0
                };
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
            LayerTable lt = tr.GetObject(
                db.LayerTableId, OpenMode.ForRead) as LayerTable;
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


    // ─────────────────────────────────────────────────────────────────
    // HELPER CLASS  (renamed from DimResult → UtilityDimData)
    // ─────────────────────────────────────────────────────────────────
    public class UtilityDimData
    {
        public string LayerName { get; set; }
        public double Distance { get; set; }
        public Point3d IntersectPt { get; set; }
        public string DedupKey { get; set; }
    }
}