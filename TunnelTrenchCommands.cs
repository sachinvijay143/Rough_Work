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
    public class TunnelTrenchCommands
    {
        private const string TUNNEL_LAYER = "TANNEL";
        private const string TRENCH_LAYER = "PROPOSED TRENCH";
        private const string MARK_LAYER = "Intersection_Marks";
        private const double CIRCLE_RADIUS = 1.0;
        private const double TOLERANCE = 1e-6;

        [CommandMethod("CheckTunnelTrench")]
        public void CheckTunnelTrench()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

                    // Ensure mark layer exists
                    EnsureLayerExists(db, tr, lt, MARK_LAYER);

                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord modelSpace = tr.GetObject(
                        bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    // Collect all polylines in "Proposed_Trench" layer
                    List<Polyline> trenchPolylines = new List<Polyline>();
                    List<Polyline2d> trenchPolylines2d = new List<Polyline2d>();

                    // Collect tunnel polylines
                    List<ObjectId> tunnelPolylineIds = new List<ObjectId>();

                    // Iterate over all entities in model space
                    foreach (ObjectId objId in modelSpace)
                    {
                        Entity ent = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                        if (ent == null) continue;

                        if (ent.Layer == TRENCH_LAYER)
                        {
                            if (ent is Polyline pl) trenchPolylines.Add(pl);
                            else if (ent is Polyline2d pl2d) trenchPolylines2d.Add(pl2d);
                        }
                        else if (ent.Layer == TUNNEL_LAYER)
                        {
                            if (ent is Polyline || ent is Polyline2d)
                                tunnelPolylineIds.Add(objId);
                        }
                    }

                    if (tunnelPolylineIds.Count == 0)
                    {
                        ed.WriteMessage($"\nNo polylines found in layer '{TUNNEL_LAYER}'.");
                        tr.Commit();
                        return;
                    }

                    int circlesAdded = 0, marksPlaced = 0, circlesRemoved = 0;

                    foreach (ObjectId tunnelId in tunnelPolylineIds)
                    {
                        Entity tunnelEnt = tr.GetObject(tunnelId, OpenMode.ForRead) as Entity;

                        Point3d startPt = Point3d.Origin;
                        Point3d endPt = Point3d.Origin;

                        if (tunnelEnt is Polyline tunnelPl)
                        {
                            startPt = tunnelPl.GetPoint3dAt(0);
                            endPt = tunnelPl.GetPoint3dAt(tunnelPl.NumberOfVertices - 1);
                        }
                        else if (tunnelEnt is Polyline2d tunnelPl2d)
                        {
                            var vertices = GetPolyline2dVertices(tunnelPl2d, tr);
                            if (vertices.Count < 2) continue;
                            startPt = vertices[0];
                            endPt = vertices[vertices.Count - 1];
                        }

                        // Process Start and End points
                        ProcessPoint(db, tr, modelSpace, startPt,
                            trenchPolylines, trenchPolylines2d, tr,
                            ref circlesAdded, ref marksPlaced, ref circlesRemoved);

                        ProcessPoint(db, tr, modelSpace, endPt,
                            trenchPolylines, trenchPolylines2d, tr,
                            ref circlesAdded, ref marksPlaced, ref circlesRemoved);
                    }

                    tr.Commit();

                    ed.WriteMessage($"\nDone! Circles added: {circlesAdded}, " +
                                   $"Marks placed: {marksPlaced}, " +
                                   $"Circles removed (no intersection): {circlesRemoved}");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError: {ex.Message}");
                    tr.Abort();
                }
            }
        }

        private void ProcessPoint(
            Database db,
            Transaction tr,
            BlockTableRecord modelSpace,
            Point3d center,
            List<Polyline> trenchPolylines,
            List<Polyline2d> trenchPolylines2d,
            Transaction outerTr,
            ref int circlesAdded,
            ref int marksPlaced,
            ref int circlesRemoved)
        {
            // Create circle at the point
            Circle circle = new Circle(center, Vector3d.ZAxis, CIRCLE_RADIUS);
            circle.Layer = TUNNEL_LAYER;
            modelSpace.AppendEntity(circle);
            tr.AddNewlyCreatedDBObject(circle, true);
            circlesAdded++;

            // Check intersection with all Proposed_Trench polylines
            bool intersects = false;

            foreach (Polyline trenchPl in trenchPolylines)
            {
                if (CircleIntersectsOrTouchesPolyline(circle, trenchPl))
                {
                    intersects = true;
                    break;
                }
            }

            if (!intersects)
            {
                foreach (Polyline2d trenchPl2d in trenchPolylines2d)
                {
                    if (CircleIntersectsOrTouchesPolyline2d(circle, trenchPl2d, tr))
                    {
                        intersects = true;
                        break;
                    }
                }
            }

            if (intersects)
            {
                //// Place a mark (Point entity) at the circle center
                //DBPoint mark = new DBPoint(center);
                //mark.Layer = MARK_LAYER;
                //modelSpace.AppendEntity(mark);
                //tr.AddNewlyCreatedDBObject(mark, true);

                //// Set PDMODE so points are visible (cross style)
                //db.Pdmode = 34;  // Cross inside circle
                //db.Pdsize = CIRCLE_RADIUS * 0.5;

                marksPlaced++;
            }
            else
            {
                // Remove the circle — no intersection found
                circle.UpgradeOpen();
                circle.Erase();
                circlesRemoved++;
            }
        }

        /// <summary>
        /// Checks if a circle intersects or touches a Polyline (LW Polyline).
        /// Strategy: for each segment of the polyline, find the minimum distance
        /// from the circle center to that segment. If <= radius, it intersects/touches.
        /// </summary>
        private bool CircleIntersectsOrTouchesPolyline(Circle circle, Polyline pl)
        {
            Point3d center = circle.Center;
            double radius = circle.Radius;

            int numSegments = pl.NumberOfVertices - 1;
            if (pl.Closed) numSegments = pl.NumberOfVertices;

            for (int i = 0; i < numSegments; i++)
            {
                SegmentType segType = pl.GetSegmentType(i);

                if (segType == SegmentType.Line)
                {
                    LineSegment3d seg = pl.GetLineSegmentAt(i);
                    double dist = seg.GetDistanceTo(center);
                    if (dist <= radius + TOLERANCE)
                        return true;
                }
                else if (segType == SegmentType.Arc)
                {
                    // For arc segments, use closest point on the arc
                    CircularArc3d arc = pl.GetArcSegmentAt(i);
                    Point3d closest = arc.GetClosestPointTo(center).Point;
                    double dist = center.DistanceTo(closest);
                    if (dist <= radius + TOLERANCE)
                        return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Checks if a circle intersects or touches a Polyline2d.
        /// </summary>
        private bool CircleIntersectsOrTouchesPolyline2d(
            Circle circle, Polyline2d pl2d, Transaction tr)
        {
            // Convert Polyline2d to a list of line segments via its vertices
            List<Point3d> pts = GetPolyline2dVertices(pl2d, tr);
            if (pts.Count < 2) return false;

            Point3d center = circle.Center;
            double radius = circle.Radius;

            int count = pl2d.Closed ? pts.Count : pts.Count - 1;

            for (int i = 0; i < count; i++)
            {
                Point3d p1 = pts[i];
                Point3d p2 = pts[(i + 1) % pts.Count];
                LineSegment3d seg = new LineSegment3d(p1, p2);
                double dist = seg.GetDistanceTo(center);
                if (dist <= radius + TOLERANCE)
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Extracts vertex positions from a Polyline2d.
        /// </summary>
        private List<Point3d> GetPolyline2dVertices(Polyline2d pl2d, Transaction tr)
        {
            var pts = new List<Point3d>();
            foreach (ObjectId vId in pl2d)
            {
                Vertex2d v = tr.GetObject(vId, OpenMode.ForRead) as Vertex2d;
                if (v != null)
                    pts.Add(v.Position);
            }
            return pts;
        }

        /// <summary>
        /// Ensures a layer exists in the drawing; creates it if not.
        /// </summary>
        private void EnsureLayerExists(Database db, Transaction tr, LayerTable lt, string layerName)
        {
            if (!lt.Has(layerName))
            {
                lt.UpgradeOpen();
                LayerTableRecord ltr = new LayerTableRecord();
                ltr.Name = layerName;
                lt.Add(ltr);
                tr.AddNewlyCreatedDBObject(ltr, true);
            }
        }
    }
}
