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
    public class UtilityIntersectionData
    {
        public Point3d PaperSpacePoint { get; set; }
        public string UtilityType { get; set; }
        public string LayoutName { get; set; }
    }

    public class IntersectionTracker
    {
        private readonly Dictionary<string, string> layerToUtility = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "E-ELECTRIC-UG", "ELEC" },
            { "E-GAS", "GAS" },
            { "E_GAS", "GAS" },
            { "E-SEWER", "SEWER" },
            { "E_WATER", "WATER" },
            { "E-WATER", "WATER" },
            { "E-STORM", "STORM" }
        };

        [CommandMethod("GetPaperSpaceIntersections")]
        public void GetPaperSpaceIntersections()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            List<UtilityIntersectionData> allResults = new List<UtilityIntersectionData>();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                DBDictionary layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    Layout lay = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                    if (lay.LayoutName.Equals("Model", StringComparison.OrdinalIgnoreCase)) continue;

                    ObjectIdCollection vportIds = lay.GetViewports();
                    if (vportIds.Count < 2) continue;

                    Viewport vp = (Viewport)tr.GetObject(vportIds[1], OpenMode.ForRead);

                    // --- MANUAL TRANSFORMATION CALCULATION ---
                    // 1. Move from Model Space WCS to Viewport Center
                    // 2. Scale by CustomScale
                    // 3. Move to Paper Space CenterPoint
                    // Convert Vector2d ViewCenter to Vector3d
                    Vector3d viewCenterVect = new Vector3d(vp.ViewCenter.X, vp.ViewCenter.Y, 0.0);

                    // Build the matrix correctly
                    Matrix3d modelToPaper =
                        Matrix3d.Displacement(vp.CenterPoint.GetAsVector()) *
                        Matrix3d.Scaling(vp.CustomScale, vp.CenterPoint) *
                        Matrix3d.Displacement(viewCenterVect.Negate());

                    BlockTableRecord ms = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

                    List<Polyline> fibers = new List<Polyline>();
                    List<Polyline> utilities = new List<Polyline>();

                    foreach (ObjectId id in ms)
                    {
                        Polyline pl = tr.GetObject(id, OpenMode.ForRead) as Polyline;
                        if (pl == null) continue;

                        if (pl.Layer.Equals("P-FIBER-UG", StringComparison.OrdinalIgnoreCase))
                            fibers.Add(pl);
                        else if (layerToUtility.ContainsKey(pl.Layer))
                            utilities.Add(pl);
                    }

                    foreach (Polyline fiber in fibers)
                    {
                        foreach (Polyline other in utilities)
                        {
                            Point3dCollection pts = new Point3dCollection();
                            // Intersecting on a plane handles different Z-elevations
                            fiber.IntersectWith(other, Intersect.OnBothOperands, new Plane(), pts, IntPtr.Zero, IntPtr.Zero);

                            foreach (Point3d msPt in pts)
                            {
                                // Apply the manual matrix
                                Point3d psPt = msPt.TransformBy(modelToPaper);

                                // Boundary Check: Is the intersection visible in this viewport?
                                if (IsPointInViewport(psPt, vp))
                                {
                                    allResults.Add(new UtilityIntersectionData
                                    {
                                        PaperSpacePoint = psPt,
                                        UtilityType = layerToUtility[other.Layer],
                                        LayoutName = lay.LayoutName
                                    });
                                }
                            }
                        }
                    }
                }
                tr.Commit();
            }

            // Output the final collected list
            ed.WriteMessage($"\nTotal Intersections Found: {allResults.Count}");
            foreach (var res in allResults)
            {
                ed.WriteMessage($"\nLayout: {res.LayoutName} | Utility: {res.UtilityType} | Paper Pt: {res.PaperSpacePoint}");
            }
        }

        private bool IsPointInViewport(Point3d psPt, Viewport vp)
        {
            double halfW = vp.Width / 2.0;
            double halfH = vp.Height / 2.0;
            return (psPt.X >= vp.CenterPoint.X - halfW && psPt.X <= vp.CenterPoint.X + halfW &&
                    psPt.Y >= vp.CenterPoint.Y - halfH && psPt.Y <= vp.CenterPoint.Y + halfH);
        }
    }
}
