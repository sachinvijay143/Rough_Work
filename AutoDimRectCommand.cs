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
        private static readonly Dictionary<string, string> LayerSuffixMap =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
            { "WATER",           "' WATER"        },
            { "SS",              "' SEWER"        },
            { "BOC",             "' BOC"          },
            { "SIDEWALK",        "' S/W"          },
            { "ROW",             "' ROW"          },
            { "EASEMENTS",       "' PUE"          },
            { "PROPOSED TRENCH", "' PROP. TRENCH" },
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
                DrawVisualRectangle(db, rect);

                // ── Step 4: Compute rayLength from full model extents
                //            so it works at any zoom level ──
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

                        // ── Walk Centerline ──
                        double total = cl.Length;
                        double currentDist = 0.0;

                        while (currentDist <= total)
                        {
                            Point3d clPt =
                                cl.GetPointAtDist(currentDist);

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

                                // ── Left side ──
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

                                // ── Right side ──
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

                            leftDims =
                                leftDims.OrderBy(d => d.Distance).ToList();
                            rightDims =
                                rightDims.OrderBy(d => d.Distance).ToList();

                            PlaceStackedDims(tr, mSpace,
                                clPt, tangent, perpLeft,
                                leftDims, dimStyleId,
                                STACK_GAP, BEYOND_OFFSET,
                                isLeft: true, rect, placedLeft);

                            PlaceStackedDims(tr, mSpace,
                                clPt, tangent, perpRight,
                                rightDims, dimStyleId,
                                STACK_GAP, BEYOND_OFFSET,
                                isLeft: false, rect, placedRight);

                            currentDist += INTERVAL;
                        }

                        tr.Commit();
                        ed.WriteMessage("\n✅ AUTODIMRECT complete.");
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
        private void DrawVisualRectangle(Database db, Extents2d rect)
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
                tr.Commit();
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
            HashSet<string> placedKeys)
        {
            if (!dims.Any()) return;

            double maxDist = dims.Max(d => d.Distance);

            for (int i = 0; i < dims.Count; i++)
            {
                DimResult dim = dims[i];

                // ── Skip already placed layer+distance combos ──
                if (placedKeys.Contains(dim.DedupKey))
                    continue;

                double tangentShift = i * stackGap;
                Point3d stationPt = clPt + tangent * tangentShift;

                Point3d xLine1 = stationPt;
                Point3d xLine2 = ProjectToPerp(
                                        stationPt, perpDir, dim.IntersectPt);
                Point3d dimLinePt = stationPt
                                    + perpDir * (maxDist + beyondOffset);

                // ── Only check xLine2 against rect ──
                if (!PointInRect(xLine2, rect))
                    continue;

                AlignedDimension ad = new AlignedDimension(
                    xLine1, xLine2, dimLinePt,
                    string.Empty, dimStyleId);

                double dimAngle = Common_functions
                                       .GetAngleBetweenPoints(xLine1, xLine2);
                double dist = Common_functions.GetDistanceBetweenPoints(xLine1, xLine2);
                string dst=dist.ToString();
                double distValue = (dst.Length*3.5);
                

                ad.Layer = DIM_LAYER;
                ad.Dimasz = 2.75;
                ad.Dimtxt = 3.5;
                ad.Color = dim.LayerName.Equals(
                                "PROPOSED TRENCH",
                                StringComparison.OrdinalIgnoreCase)
                            ? Color.FromColorIndex(ColorMethod.ByAci, 30)
                            : Color.FromColorIndex(ColorMethod.ByBlock, 0);

                if (LayerSuffixMap.TryGetValue(
                        dim.LayerName, out string suffix))
                    ad.Suffix = suffix;
                distValue = ((distValue + (ad.Suffix.Length * 3.34)) /2) + 43;
                //MessageBox.Show("Suffix: " + suffix + "\\t distValue: " + distValue);
                Point3d txtpnt = Common_functions.PolarPoint(xLine1, dimAngle, distValue);
                ad.TextPosition = txtpnt;
                ad.TextRotation = 0.0;

                mSpace.AppendEntity(ad);
                tr.AddNewlyCreatedDBObject(ad, true);
                ad.DimLinePoint = ad.DimLinePoint;

                // ── Register as placed ──
                placedKeys.Add(dim.DedupKey);
            }
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
                        // Only accept points in the correct ray direction
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
