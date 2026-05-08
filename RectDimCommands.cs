using IntelliCAD.ApplicationServices;
using IntelliCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Teigha.DatabaseServices;
using Teigha.Geometry;
using Teigha.Runtime;

namespace Rough_Works
{
    /// <summary>
    /// Holds configuration for each target layer: display suffix and dimension offset value.
    /// </summary>
    public class LayerConfig
    {
        public string Suffix { get; }
        public double Offset { get; }

        public LayerConfig(string suffix, double offset)
        {
            Suffix = suffix;
            Offset = offset;
        }
    }

    public class RectDimCommands
    {
        // ─────────────────────────────────────────────────────────────
        //  Layer suffix / offset map (case-insensitive key matching)
        // ─────────────────────────────────────────────────────────────
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

        private const string CenterlineLayerKeyword = "CENTERLINE";
        private const string TempRectangleLayerName = "TEMP_SELECTION_RECT";

        // ─────────────────────────────────────────────────────────────
        //  Main command: RECTDIM
        // ─────────────────────────────────────────────────────────────
        [CommandMethod("RECTDIM")]
        public void RectDimCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            ed.WriteMessage("\n[RECTDIM] Pick 4 corner points for the selection rectangle.\n");

            // ── 1. Collect 4 points ──────────────────────────────────
            Point3d[] corners = PickFourPoints(ed);
            if (corners == null)
            {
                ed.WriteMessage("\nCommand cancelled.\n");
                return;
            }

            // ── 2. Build the selection polygon (convex hull order) ───
            Point3dCollection poly3d = BuildConvexQuad(corners);

            // ── 3. Draw a temporary rectangle on model space ─────────
            ObjectId tempRectId = DrawTemporaryRectangle(db, corners);

            // ── 4. Select objects inside the rectangle ───────────────
            SelectionFilter filter = new SelectionFilter(new[]
            {
                new TypedValue((int)DxfCode.Start, "LINE,LWPOLYLINE,POLYLINE,SPLINE")
            });

            PromptSelectionResult psr = ed.SelectCrossingPolygon(poly3d, filter);

            if (psr.Status != PromptStatus.OK || psr.Value.Count == 0)
            {
                ed.WriteMessage("\nNo lines or polylines found inside the rectangle.\n");
                RemoveTemporaryRectangle(db, tempRectId);
                return;
            }

            ed.WriteMessage($"\nFound {psr.Value.Count} object(s) inside the rectangle.\n");

            // ── 5. Analyse found objects ─────────────────────────────
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var centerlineSegments = new List<LineSegment3d>();  // from CENTERLINE entities
                var targetEntities = new List<(Entity ent, LayerConfig cfg)>();

                BlockTableRecord mspace = (BlockTableRecord)tr.GetObject(
                    db.CurrentSpaceId, OpenMode.ForWrite);

                foreach (SelectedObject so in psr.Value)
                {
                    Entity ent = (Entity)tr.GetObject(so.ObjectId, OpenMode.ForRead);
                    string layerUpper = ent.Layer.ToUpper();

                    // ── Check for CENTERLINE ─────────────────────────
                    if (layerUpper.Contains(CenterlineLayerKeyword))
                    {
                        centerlineSegments.AddRange(ExtractSegments(ent));
                        continue;
                    }

                    // ── Check for mapped layers ──────────────────────
                    foreach (var kv in LayerSuffixMap)
                    {
                        if (layerUpper.Contains(kv.Key.ToUpper()))
                        {
                            targetEntities.Add((ent, kv.Value));
                            break;
                        }
                    }
                }

                if (centerlineSegments.Count == 0)
                {
                    ed.WriteMessage("\nNo CENTERLINE entities found inside the rectangle. " +
                                    "Dimensions will not be created.\n");
                    tr.Commit();
                    RemoveTemporaryRectangle(db, tempRectId);
                    return;
                }

                ed.WriteMessage($"\nFound {centerlineSegments.Count} centerline segment(s) " +
                                $"and {targetEntities.Count} target entity/entities.\n");

                // ── 6. Create dimensions ─────────────────────────────
                int dimCount = 0;
                foreach (var (ent, cfg) in targetEntities)
                {
                    foreach (LineSegment3d clSeg in centerlineSegments)
                    {
                        // Pass the 'corners' array to the method
                        dimCount += CreateDimensionsBothSides(
                            db, tr, mspace, clSeg, ent, cfg, corners);
                    }
                }

                ed.WriteMessage($"\nCreated {dimCount} dimension object(s).\n");
                tr.Commit();
            }

            // ── 7. Remove the temporary rectangle ────────────────────
            RemoveTemporaryRectangle(db, tempRectId);
            ed.WriteMessage("\nRECTDIM complete.\n");
        }

        // ─────────────────────────────────────────────────────────────
        //  Step 1 – Ask the user for 4 points
        // ─────────────────────────────────────────────────────────────
        private Point3d[] PickFourPoints(Editor ed)
        {
            var pts = new Point3d[4];
            string[] prompts =
            {
                "\nPick Point 1 (e.g. top-left)    : ",
                "\nPick Point 2 (e.g. top-right)   : ",
                "\nPick Point 3 (e.g. bottom-right): ",
                "\nPick Point 4 (e.g. bottom-left) : ",
            };

            for (int i = 0; i < 4; i++)
            {
                PromptPointOptions ppo = new PromptPointOptions(prompts[i]);
                ppo.AllowNone = false;
                if (i > 0) ppo.UseBasePoint = false;

                PromptPointResult ppr = ed.GetPoint(ppo);
                if (ppr.Status != PromptStatus.OK) return null;
                pts[i] = ppr.Value;
            }
            return pts;
        }

        // ─────────────────────────────────────────────────────────────
        //  Build a 2-D polygon from the 4 corners (simple bounding box)
        // ─────────────────────────────────────────────────────────────
        private Point3dCollection BuildConvexQuad(Point3d[] corners)
        {
            // Use bounding box so the crossing selection always works
            double minX = corners.Min(p => p.X);
            double maxX = corners.Max(p => p.X);
            double minY = corners.Min(p => p.Y);
            double maxY = corners.Max(p => p.Y);

            // Change the return type to Point3dCollection
            Point3dCollection col = new Point3dCollection
            {
                new Point3d(minX, minY, 0),
                new Point3d(maxX, minY, 0),
                new Point3d(maxX, maxY, 0),
                new Point3d(minX, maxY, 0),
            };
            return col;
        }

        // ─────────────────────────────────────────────────────────────
        //  Draw a temporary lightweight polyline rectangle
        // ─────────────────────────────────────────────────────────────
        private ObjectId DrawTemporaryRectangle(Database db, Point3d[] corners)
        {
            double minX = corners.Min(p => p.X);
            double maxX = corners.Max(p => p.X);
            double minY = corners.Min(p => p.Y);
            double maxY = corners.Max(p => p.Y);

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Ensure the temp layer exists
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                if (!lt.Has(TempRectangleLayerName))
                {
                    lt.UpgradeOpen();
                    var ltr = new LayerTableRecord
                    {
                        Name = TempRectangleLayerName,
                        Color = Teigha.Colors.Color.FromColorIndex(
                                    Teigha.Colors.ColorMethod.ByAci, 1) // red
                    };
                    lt.Add(ltr);
                    tr.AddNewlyCreatedDBObject(ltr, true);
                }

                Polyline rect = new Polyline();
                rect.AddVertexAt(0, new Point2d(minX, minY), 0, 0, 0);
                rect.AddVertexAt(1, new Point2d(maxX, minY), 0, 0, 0);
                rect.AddVertexAt(2, new Point2d(maxX, maxY), 0, 0, 0);
                rect.AddVertexAt(3, new Point2d(minX, maxY), 0, 0, 0);
                rect.Closed = true;
                rect.Layer = TempRectangleLayerName;

                BlockTableRecord mspace = (BlockTableRecord)tr.GetObject(
                    db.CurrentSpaceId, OpenMode.ForWrite);
                ObjectId id = mspace.AppendEntity(rect);
                tr.AddNewlyCreatedDBObject(rect, true);

                tr.Commit();
                return id;
            }
        }

        // ─────────────────────────────────────────────────────────────
        //  Remove the temporary rectangle
        // ─────────────────────────────────────────────────────────────
        private void RemoveTemporaryRectangle(Database db, ObjectId id)
        {
            if (id.IsNull) return;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                Entity ent = (Entity)tr.GetObject(id, OpenMode.ForWrite);
                ent.Erase();
                tr.Commit();
            }
        }

        // ─────────────────────────────────────────────────────────────
        //  Extract line segments from LINE / LWPOLYLINE / POLYLINE
        // ─────────────────────────────────────────────────────────────
        private List<LineSegment3d> ExtractSegments(Entity ent)
        {
            var segs = new List<LineSegment3d>();

            if (ent is Line line)
            {
                segs.Add(new LineSegment3d(line.StartPoint, line.EndPoint));
            }
            else if (ent is Polyline pl)
            {
                for (int i = 0; i < pl.NumberOfVertices - 1; i++)
                    segs.Add(new LineSegment3d(
                        pl.GetPoint3dAt(i), pl.GetPoint3dAt(i + 1)));
                if (pl.Closed)
                    segs.Add(new LineSegment3d(
                        pl.GetPoint3dAt(pl.NumberOfVertices - 1),
                        pl.GetPoint3dAt(0)));
            }
            else if (ent is Polyline2d pl2)
            {
                var vertices = new List<Point3d>();
                foreach (ObjectId vid in pl2)
                {
                    using (Transaction tr2 =
                        ent.Database.TransactionManager.StartTransaction())
                    {
                        Vertex2d v = (Vertex2d)tr2.GetObject(vid, OpenMode.ForRead);
                        vertices.Add(v.Position);
                        tr2.Commit();
                    }
                }
                for (int i = 0; i < vertices.Count - 1; i++)
                    segs.Add(new LineSegment3d(vertices[i], vertices[i + 1]));
            }

            return segs;
        }

        // ─────────────────────────────────────────────────────────────
        //  Core: create aligned dimensions on both sides of the CL
        // ─────────────────────────────────────────────────────────────
        /// <summary>
        /// Projects the target entity onto the centerline direction and places
        /// one AlignedDimension on each side of the centerline.
        /// Returns the number of dimensions actually added (0, 1, or 2).
        /// </summary>
        private int CreateDimensionsBothSides(
    Database db,
    Transaction tr,
    BlockTableRecord mspace,
    LineSegment3d clSeg,
    Entity targetEnt,
    LayerConfig cfg,
    Point3d[] selectionCorners) // New parameter
        {
            Point3d clMid = clSeg.MidPoint;
            Vector3d clDir = (clSeg.EndPoint - clSeg.StartPoint).GetNormal();
            double segLength = clSeg.Length;
            Vector3d clNormal = new Vector3d(-clDir.Y, clDir.X, 0).GetNormal();

            Point3d? closestPt = GetClosestPointOnEntity(targetEnt, clMid);
            if (closestPt == null) return 0;
            Point3d tgt = closestPt.Value;

            double t = (tgt - clSeg.StartPoint).DotProduct(clDir);
            if (t < -1e-6 || t > segLength + 1e-6) return 0;

            Point3d clFoot = clSeg.StartPoint + clDir * t;

            // --- NEW: Clipping Check ---
            // Only create dimension if clFoot is inside the user's picked rectangle
            if (!IsPointInRectangle(clFoot, selectionCorners)) return 0;

            double signedDist = (tgt - clFoot).DotProduct(clNormal);
            if (Math.Abs(signedDist) < 1e-4) return 0;

            int count = 0;
            // Use the offset from your Dictionary: e.g., 20.5, 11.66, etc.
            // If you want a uniform 4.25' gap between layers, you can replace 
            // cfg.Offset with a calculated value.
            double dimLineOffset = cfg.Offset;

            for (int side = -1; side <= 1; side += 2)
            {
                Point3d witness2Point = clFoot + clNormal * (signedDist * side);

                // This ensures the dimension line is placed at the specific 
                // offset defined for that layer (e.g. WATER at 20.5')
                Point3d dimLinePoint = clFoot + (clNormal * (side * dimLineOffset));

                AlignedDimension dim = new AlignedDimension
                {
                    XLine1Point = clFoot,
                    XLine2Point = witness2Point,
                    DimLinePoint = dimLinePoint,
                    DimensionStyle = db.Dimstyle,
                    DimensionText = cfg.Suffix,
                };

                mspace.AppendEntity(dim);
                tr.AddNewlyCreatedDBObject(dim, true);
                count++;
            }
            return count;
        }

        private bool IsPointInRectangle(Point3d pt, Point3d[] corners)
        {
            double minX = corners.Min(p => p.X);
            double maxX = corners.Max(p => p.X);
            double minY = corners.Min(p => p.Y);
            double maxY = corners.Max(p => p.Y);

            return (pt.X >= minX && pt.X <= maxX && pt.Y >= minY && pt.Y <= maxY);
        }
        // ─────────────────────────────────────────────────────────────
        //  Closest point on any supported entity to a reference point
        // ─────────────────────────────────────────────────────────────
        private Point3d? GetClosestPointOnEntity(Entity ent, Point3d refPt)
        {
            try
            {
                // AutoCAD's Curve.GetClosestPointTo works for Line, Polyline, etc.
                if (ent is Curve curve)
                    return curve.GetClosestPointTo(refPt, false);
            }
            catch { /* fall through */ }

            // Fallback: midpoint of bounding-box
            Extents3d ext = ent.GeometricExtents;
            return new Point3d(
                (ext.MinPoint.X + ext.MaxPoint.X) / 2,
                (ext.MinPoint.Y + ext.MaxPoint.Y) / 2,
                0);
        }
    }
}
