using IntelliCAD.ApplicationServices;
using IntelliCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Teigha.Colors;
using Teigha.DatabaseServices;
using Teigha.Geometry;
using Teigha.Runtime;

namespace Rough_Works
{
    public class AutoDimCommand
    {
        // ─────────────────────────────────────────────
        // CONFIGURATION
        // ─────────────────────────────────────────────
        private static readonly Dictionary<string, string> LayerSuffixMap =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
            { "WATER",            "' WATER"        },
            { "SS",               "' SEWER"        },
            { "BOC",              "' BOC"          },
            { "SIDEWALK",         "' S/W"          },
            { "ROW",              "' ROW"          },
            { "EASEMENTS",        "' PUE"          },
            { "PROPOSED TRENCH",  "' PROP. TRENCH" },
            };

        private const string DIM_LAYER = "DIMENSIONS";
        private const string DIM_STYLE_NAME = "BPGDIMS";
        private const double RAY_LENGTH = 500.0;

        // ─────────────────────────────────────────────
        // MAIN COMMAND
        // ─────────────────────────────────────────────
        [CommandMethod("AUTODIM")]
        public void AutoDim()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // ── Save original DIMASSOC ──
            int originalDimassoc = Convert.ToInt32(
                Application.GetSystemVariable("DIMASSOC"));

            // ✅ Step 1: Set system variable immediately (affects UI/prompts)
            Application.SetSystemVariable("DIMASSOC", 0);

            // ✅ Step 2: Commit db.DimAssoc in its own pre-transaction
            //    so it is fully flushed before the main transaction opens.
            //    This is the reliable way to silence associativity warnings.
            using (Transaction preTr = db.TransactionManager.StartTransaction())
            {
                db.DimAssoc = 0;
                preTr.Commit();
            }

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        // ── Select Centerline ──
                        PromptEntityOptions peo =
                            new PromptEntityOptions(
                                "\nSelect CENTERLINE Polyline: ");
                        peo.SetRejectMessage("Must be a Polyline.");
                        peo.AddAllowedClass(typeof(Polyline), true);
                        PromptEntityResult per = ed.GetEntity(peo);
                        if (per.Status != PromptStatus.OK) return;

                        Polyline cl = tr.GetObject(
                                        per.ObjectId,
                                        OpenMode.ForRead) as Polyline;
                        if (cl == null)
                        {
                            ed.WriteMessage("\nInvalid selection.");
                            return;
                        }

                        // ── Station Interval ──
                        PromptDoubleOptions pdoInterval =
                            new PromptDoubleOptions(
                                "\nEnter station interval (feet): ");
                        pdoInterval.DefaultValue = 50.0;
                        pdoInterval.AllowNegative = false;
                        pdoInterval.AllowZero = false;
                        PromptDoubleResult pdrInterval = ed.GetDouble(pdoInterval);
                        if (pdrInterval.Status != PromptStatus.OK) return;
                        double interval = pdrInterval.Value;

                        // ── Stack Gap ──
                        PromptDoubleOptions pdoGap =
                            new PromptDoubleOptions(
                                "\nEnter stack gap between dims (feet): ");
                        pdoGap.DefaultValue = 8.0;
                        pdoGap.AllowNegative = false;
                        pdoGap.AllowZero = false;
                        PromptDoubleResult pdrGap = ed.GetDouble(pdoGap);
                        if (pdrGap.Status != PromptStatus.OK) return;
                        double stackGap = pdrGap.Value;

                        // ── Beyond Offset ──
                        PromptDoubleOptions pdoBeyond =
                            new PromptDoubleOptions(
                                "\nEnter offset beyond farthest layer (feet): ");
                        pdoBeyond.DefaultValue = 15.0;
                        pdoBeyond.AllowNegative = false;
                        pdoBeyond.AllowZero = false;
                        PromptDoubleResult pdrBeyond = ed.GetDouble(pdoBeyond);
                        if (pdrBeyond.Status != PromptStatus.OK) return;
                        double beyondOffset = pdrBeyond.Value;

                        // ── Setup ──
                        EnsureLayerExists(db, tr);
                        ObjectId dimStyleId = GetOrCreateDimStyle(db, tr);

                        // ── DEBUG ──
                        DimStyleTableRecord debugStyle =
                            tr.GetObject(dimStyleId, OpenMode.ForRead)
                              as DimStyleTableRecord;
                        ed.WriteMessage(
                            "\n========== DIM STYLE DEBUG ==========");
                        ed.WriteMessage(
                            $"\nDimscale in use   : {debugStyle.Dimscale}");
                        ed.WriteMessage(
                            $"\nDimasz in use     : {debugStyle.Dimasz}");
                        ed.WriteMessage(
                            $"\nDimtxt in use     : {debugStyle.Dimtxt}");
                        ed.WriteMessage(
                            $"\nActual arrow size : " +
                            $"{debugStyle.Dimasz * debugStyle.Dimscale}");
                        ed.WriteMessage(
                            $"\nActual text height: " +
                            $"{debugStyle.Dimtxt * debugStyle.Dimscale}");
                        ed.WriteMessage(
                            "\n======================================\n");

                        BlockTable bt = tr.GetObject(
                            db.BlockTableId,
                            OpenMode.ForRead) as BlockTable;
                        BlockTableRecord mSpace = tr.GetObject(
                            bt[BlockTableRecord.ModelSpace],
                            OpenMode.ForWrite) as BlockTableRecord;

                        var layerEntities = CollectLayerEntities(db, tr);

                        // ── Walk Centerline ──
                        double total = cl.Length;
                        double currentDist = 0.0;

                        while (currentDist <= total)
                        {
                            Point3d clPt = cl.GetPointAtDist(currentDist);
                            Vector3d tangent = cl.GetFirstDerivative(clPt)
                                                 .GetNormal();

                            Vector3d perpLeft = new Vector3d(
                                                    -tangent.Y, tangent.X, 0);
                            Vector3d perpRight = new Vector3d(
                                                     tangent.Y, -tangent.X, 0);

                            var leftDims = new List<DimResult>();
                            var rightDims = new List<DimResult>();

                            foreach (var kvp in layerEntities)
                            {
                                string lyrName = kvp.Key;
                                List<Curve> curves = kvp.Value;

                                Point3d? lPt = FindNearestIntersection(
                                    clPt,
                                    clPt + perpLeft * RAY_LENGTH,
                                    curves);
                                if (lPt.HasValue)
                                    leftDims.Add(new DimResult
                                    {
                                        LayerName = lyrName,
                                        Distance = clPt.DistanceTo(lPt.Value),
                                        IntersectPt = lPt.Value
                                    });

                                Point3d? rPt = FindNearestIntersection(
                                    clPt,
                                    clPt + perpRight * RAY_LENGTH,
                                    curves);
                                if (rPt.HasValue)
                                    rightDims.Add(new DimResult
                                    {
                                        LayerName = lyrName,
                                        Distance = clPt.DistanceTo(rPt.Value),
                                        IntersectPt = rPt.Value
                                    });
                            }

                            leftDims = leftDims.OrderBy(d => d.Distance).ToList();
                            rightDims = rightDims.OrderBy(d => d.Distance).ToList();

                            PlaceStackedDims(tr, mSpace,
                                clPt, tangent, perpLeft,
                                leftDims, dimStyleId,
                                stackGap, beyondOffset, isLeft: true);

                            PlaceStackedDims(tr, mSpace,
                                clPt, tangent, perpRight,
                                rightDims, dimStyleId,
                                stackGap, beyondOffset, isLeft: false);

                            currentDist += interval;
                        }

                        tr.Commit();
                        ed.WriteMessage("\n✅ AUTODIM complete.");
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
                // ✅ Restore DIMASSOC in its own transaction (same pattern as setup)
                using (Transaction restoreTr =
                           db.TransactionManager.StartTransaction())
                {
                    db.DimAssoc = originalDimassoc;
                    restoreTr.Commit();
                }
                Application.SetSystemVariable("DIMASSOC", originalDimassoc);
                ed.WriteMessage($"\nDIMASSC restored to: {originalDimassoc}");
            }
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
            bool isLeft)
        {
            if (!dims.Any()) return;

            double maxDist = dims.Max(d => d.Distance);

            for (int i = 0; i < dims.Count; i++)
            {
                DimResult dim = dims[i];

                double tangentShift = i * stackGap;
                Point3d stationPt = clPt + tangent * tangentShift;

                Point3d xLine1 = stationPt;
                Point3d xLine2 = ProjectToPerp(
                                        stationPt, perpDir, dim.IntersectPt);
                Point3d dimLinePt = stationPt
                                    + perpDir * (maxDist + beyondOffset);

                AlignedDimension ad = new AlignedDimension(
                    xLine1,
                    xLine2,
                    dimLinePt,
                    string.Empty,
                    dimStyleId);

                double dimAngle = Common_functions.GetAngleBetweenPoints(xLine1, xLine2);
                double dist = Common_functions.GetDistanceBetweenPoints(xLine1, xLine2);
                dist = dist + 0;
                Point3d txtpnt = Common_functions.PolarPoint(xLine1, dimAngle, 53);

                ad.Layer = DIM_LAYER;
                ad.Dimasz = 2.75;
                ad.Dimtxt = 3.5;
                ad.Color = dim.LayerName.Equals(
                               "PROPOSED TRENCH",
                               StringComparison.OrdinalIgnoreCase)
                           ? Color.FromColorIndex(ColorMethod.ByAci, 30)
                           : Color.FromColorIndex(ColorMethod.ByBlock, 0);

                if (LayerSuffixMap.TryGetValue(dim.LayerName, out string suffix))
                    ad.Suffix = suffix;

                ad.TextPosition = txtpnt;
                ad.TextRotation = 0.0;

                mSpace.AppendEntity(ad);
                tr.AddNewlyCreatedDBObject(ad, true);

                // ✅ Flush any pending geometry-attachment AutoCAD queued
                //    during construction. No version-specific API needed —
                //    reassigning a property forces an internal state flush.
                ad.DimLinePoint = ad.DimLinePoint;
            }
        }


        // ─────────────────────────────────────────────
        // GET OR CREATE DIM STYLE
        // ─────────────────────────────────────────────
        private ObjectId GetOrCreateDimStyle(Database db, Transaction tr)
        {
            DimStyleTable dst = tr.GetObject(
                db.DimStyleTableId, OpenMode.ForRead) as DimStyleTable;

            DimStyleTableRecord dstr;

            if (dst.Has(DIM_STYLE_NAME))
            {
                dstr = tr.GetObject(dst[DIM_STYLE_NAME], OpenMode.ForWrite)
                         as DimStyleTableRecord;
            }
            else
            {
                dst.UpgradeOpen();
                dstr = new DimStyleTableRecord { Name = DIM_STYLE_NAME };
                dst.Add(dstr);
                tr.AddNewlyCreatedDBObject(dstr, true);
            }

            // --- Lines & Arrows ---
            dstr.Dimscale = 100.0;   // ✅ drawing scale factor
            dstr.Dimasz = 0.03;  // × 100 = 3.0 model units arrow
            dstr.Dimexo = 0.625;
            dstr.Dimexe = 0.125;
            dstr.Dimse1 = false;
            dstr.Dimse2 = false;
            dstr.Dimdle = 0.0;
            dstr.Dimdli = 0.0;
            dstr.Dimclrd = Color.FromColorIndex(ColorMethod.ByBlock, 0);
            dstr.Dimclre = Color.FromColorIndex(ColorMethod.ByBlock, 0);

            // --- Text ---
            dstr.Dimtxt = 0.04;   // × 100 = 4.0 model units text height
            dstr.Dimgap = 0.09;
            dstr.Dimtad = 1;
            dstr.Dimtoh = true;
            dstr.Dimtih = false;
            dstr.Dimtvp = 0.0;
            dstr.Dimclrt = Color.FromColorIndex(ColorMethod.ByBlock, 0);

            // --- Text Style: ARIAL ---
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
                        curve.IntersectWith(ray,
                            Intersect.OnBothOperands,
                            pts, IntPtr.Zero, IntPtr.Zero);
                    }
                    catch { continue; }

                    foreach (Point3d pt in pts)
                    {
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
        // COLLECT LAYER ENTITIES
        // ─────────────────────────────────────────────
        private Dictionary<string, List<Curve>> CollectLayerEntities(
            Database db, Transaction tr)
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
                Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent == null) continue;
                if (result.ContainsKey(ent.Layer) && ent is Curve c)
                    result[ent.Layer].Add(c);
            }
            return result;
        }


        // ─────────────────────────────────────────────
        // ENSURE DIMENSIONS LAYER EXISTS
        // ─────────────────────────────────────────────
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


    // ─────────────────────────────────────────────
    // HELPER
    // ─────────────────────────────────────────────
    public class DimResult
    {
        public string LayerName { get; set; }
        public double Distance { get; set; }
        public Point3d IntersectPt { get; set; }
        public string DedupKey { get; set; }
    }
}
