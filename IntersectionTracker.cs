using IntelliCAD.ApplicationServices;
using IntelliCAD.EditorInput;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Teigha.DatabaseServices;
using Teigha.Geometry;
using Teigha.Runtime;

namespace Rough_Works
{
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

        [CommandMethod("ScanSortedIntersections")]
        public void ScanSortedIntersections()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                Plane worldXY = new Plane(Point3d.Origin, Vector3d.ZAxis);
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

                // 1. Map and Sort Grids by SHTNUM (Natural Alphanumeric Sort)
                List<GridData> allGrids = MapGrids(tr, ms);
                if (allGrids.Count == 0)
                {
                    ed.WriteMessage("\nError: No 'Grids_New' polylines or 'DWG_NO' blocks found.");
                    return;
                }

                // Use custom Natural Sort for the Sheet Numbers
                allGrids = allGrids.OrderBy(g => g.SheetNumber, new NaturalStringComparer()).ToList();

                // 2. Filter Curves in Model Space
                List<Curve> fibers = new List<Curve>();
                List<Curve> utilities = new List<Curve>();
                foreach (ObjectId id in ms)
                {
                    Curve ent = tr.GetObject(id, OpenMode.ForRead) as Curve;
                    if (ent == null) continue;
                    if (ent.Layer.Equals("P-FIBER-UG", StringComparison.OrdinalIgnoreCase)) fibers.Add(ent);
                    else if (layerToUtility.ContainsKey(ent.Layer)) utilities.Add(ent);
                }

                // 3. Find Unique Intersections and assign Grid Data
                List<UtilityIntersectionData> rawPoints = new List<UtilityIntersectionData>();
                HashSet<Point3d> uniquePts = new HashSet<Point3d>(new Point3dComparer(0.001));

                foreach (Curve fiber in fibers)
                {
                    foreach (Curve other in utilities)
                    {
                        Point3dCollection pts = new Point3dCollection();
                        fiber.IntersectWith(other, Intersect.OnBothOperands, worldXY, pts, IntPtr.Zero, IntPtr.Zero);

                        foreach (Point3d pt in pts)
                        {
                            if (uniquePts.Add(pt))
                            {
                                string sNum = "N/A";
                                int gridOrder = 9999;

                                for (int i = 0; i < allGrids.Count; i++)
                                {
                                    if (IsPointInPolyline(pt, allGrids[i].Boundary))
                                    {
                                        sNum = allGrids[i].SheetNumber;
                                        gridOrder = i;
                                        break;
                                    }
                                }

                                rawPoints.Add(new UtilityIntersectionData
                                {
                                    Point = pt,
                                    UtilityType = layerToUtility[other.Layer],
                                    SheetNum = sNum,
                                    GridSortIndex = gridOrder
                                });
                            }
                        }
                    }
                }

                // 4. Sort Results: First by Sheet Order, then by Location (Y then X)
                var sortedResults = rawPoints
                    .OrderBy(p => p.GridSortIndex)
                    .ThenByDescending(p => p.Point.Y)
                    .ThenBy(p => p.Point.X)
                    .ToList();

                // 5. Apply Sequence and Create MLeaders
                int globalCounter = 1;
                foreach (var item in sortedResults)
                {
                    item.Sequence = globalCounter++;
                    CreatePHMLeader(tr, db, item.Point, $"P{item.Sequence}");
                }

                if (sortedResults.Count > 0)
                    ExportToCSV(db, sortedResults);

                tr.Commit();
                ed.WriteMessage($"\nDone! Processed {sortedResults.Count} unique intersections.");
            }
        }

        private List<GridData> MapGrids(Transaction tr, BlockTableRecord ms)
        {
            List<Polyline> gridPolys = new List<Polyline>();
            List<BlockReference> dwgBlocks = new List<BlockReference>();
            List<GridData> results = new List<GridData>();

            foreach (ObjectId id in ms)
            {
                Entity ent = (Entity)tr.GetObject(id, OpenMode.ForRead);
                if (ent is Polyline pl && pl.Layer.Equals("Grids_New", StringComparison.OrdinalIgnoreCase))
                    gridPolys.Add(pl);
                else if (ent is BlockReference br && (br.Name.Equals("DWG_NO", StringComparison.OrdinalIgnoreCase) || br.Name.Contains("*U")))
                    dwgBlocks.Add(br);
            }

            foreach (Polyline poly in gridPolys)
            {
                foreach (BlockReference br in dwgBlocks)
                {
                    if (IsPointInPolyline(br.Position, poly))
                    {
                        string val = GetAttributeValue(tr, br, "SHTNUM");
                        if (!string.IsNullOrEmpty(val))
                        {
                            results.Add(new GridData { Boundary = poly, SheetNumber = val });
                            break;
                        }
                    }
                }
            }
            return results;
        }

        private bool IsPointInPolyline(Point3d pt, Polyline pl)
        {
            int intersections = 0;
            int n = pl.NumberOfVertices;
            for (int i = 0; i < n; i++)
            {
                Point3d p1 = pl.GetPoint3dAt(i);
                Point3d p2 = pl.GetPoint3dAt((i + 1) % n);
                if (((p1.Y > pt.Y) != (p2.Y > pt.Y)) && (pt.X < (p2.X - p1.X) * (pt.Y - p1.Y) / (p2.Y - p1.Y) + p1.X))
                    intersections++;
            }
            return (intersections % 2 != 0);
        }

        private string GetAttributeValue(Transaction tr, BlockReference br, string tag)
        {
            foreach (ObjectId id in br.AttributeCollection)
            {
                AttributeReference att = (AttributeReference)tr.GetObject(id, OpenMode.ForRead);
                if (att.Tag.Equals(tag, StringComparison.OrdinalIgnoreCase)) return att.TextString;
            }
            return "";
        }

        private void ExportToCSV(Database db, List<UtilityIntersectionData> results)
        {
            string csvPath = Path.ChangeExtension(db.Filename, ".csv");
            StringBuilder sb = new StringBuilder("Sequence,SHTNUM,Utility,X,Y,Z\n");
            foreach (var r in results)
                sb.AppendLine($"{r.Sequence},{r.SheetNum},{r.UtilityType},{r.Point.X},{r.Point.Y},{r.Point.Z}");
            File.WriteAllText(csvPath, sb.ToString());
        }

        public void CreatePHMLeader(Transaction tr, Database db, Point3d arrowheadPt, string phValue)
        {
            // 1. Define Fixed Offset
            double offsetX = 2.5;
            double offsetY = 10.5;
            Point3d bubblePt = new Point3d(arrowheadPt.X + offsetX, arrowheadPt.Y + offsetY, arrowheadPt.Z);

            // 2. Access Style and Block
            DBDictionary mlStyleDict = (DBDictionary)tr.GetObject(db.MLeaderStyleDictionaryId, OpenMode.ForRead);
            ObjectId styleId = db.MLeaderstyle;

            // Check for "BUBBLE" style specifically
            if (mlStyleDict.Contains("BUBBLE"))
                styleId = mlStyleDict.GetAt("BUBBLE");

            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            if (!bt.Has("CIRCLE FOR LEADER")) return;
            ObjectId blockId = bt["CIRCLE FOR LEADER"];

            // 3. Create MLeader Instance
            MLeader ml = new MLeader();
            ml.SetDatabaseDefaults();

            // --- FIX: ALWAYS TARGET MODEL SPACE ---
            // Instead of db.CurrentSpaceId, we use SymbolUtilityServices to get Model Space
            ObjectId modelSpaceId = SymbolUtilityServices.GetBlockModelSpaceId(db);
            BlockTableRecord ms = (BlockTableRecord)tr.GetObject(modelSpaceId, OpenMode.ForWrite);

            ms.AppendEntity(ml);
            tr.AddNewlyCreatedDBObject(ml, true);

            // 4. Set Content Properties
            ml.MLeaderStyle = styleId;
            ml.Layer = "0"; // Ensure Layer 0 is thawed and turned on
            ml.ContentType = ContentType.BlockContent;
            ml.BlockContentId = blockId;

            // Scale and Size
            //double s = 0.48425;
            double s = 19.37;
            ml.BlockScale = new Scale3d(s, s, s);
            ml.ArrowSize = 0.06053;
            ml.LandingGap = 0.012;
            ml.EnableAnnotationScale = false;

            // 5. Build Geometry
            // Note: The order of AddLeader -> AddLeaderLine -> AddVertex is strict
            int ldIdx = ml.AddLeader();
            int lnIdx = ml.AddLeaderLine(ldIdx);
            ml.AddFirstVertex(lnIdx, arrowheadPt);
            ml.AddLastVertex(lnIdx, bubblePt);

            // 6. Set Attribute Data (PH Value)
            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(blockId, OpenMode.ForRead);
            foreach (ObjectId id in btr)
            {
                AttributeDefinition attDef = tr.GetObject(id, OpenMode.ForRead) as AttributeDefinition;
                if (attDef != null && attDef.Tag.Equals("PH", StringComparison.OrdinalIgnoreCase))
                {
                    using (AttributeReference attRef = new AttributeReference())
                    {
                        attRef.SetAttributeFromBlock(attDef, Matrix3d.Identity);
                        attRef.TextString = phValue;
                        ml.SetBlockAttribute(id, attRef);
                    }
                    break;
                }
            }

            // 7. Force Geometry Recomputation
            // This is where we fix your error by passing 'true'
            ml.RecordGraphicsModified(true);

            // If your API supports it, this ensures the block content is drawn
            try { ml.DowngradeOpen(); } catch { }
        }

        [CommandMethod("BatchChSpaceMLeaders", CommandFlags.Modal)]
        public void BatchChSpaceMLeaders()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            LayoutManager lm = LayoutManager.Current;

            string originalLayout = lm.CurrentLayout;

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    DBDictionary layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                    List<string> layoutNames = new List<string>();
                    foreach (DBDictionaryEntry entry in layoutDict)
                    {
                        if (!entry.Key.Equals("Model", StringComparison.OrdinalIgnoreCase))
                            layoutNames.Add(entry.Key);
                    }

                    foreach (string layName in layoutNames)
                    {
                        lm.CurrentLayout = layName;

                        Layout lay = (Layout)tr.GetObject(layoutDict.GetAt(layName), OpenMode.ForRead);
                        ObjectIdCollection vportIds = lay.GetViewports();

                        if (vportIds.Count < 2) continue;

                        Viewport vp = (Viewport)tr.GetObject(vportIds[1], OpenMode.ForRead);

                        // Calculate MS boundaries
                        double scale = vp.CustomScale;
                        double halfMSWidth = (vp.Width / scale) / 2.0;
                        double halfMSHeight = (vp.Height / scale) / 2.0;
                        double minX = vp.ViewCenter.X - halfMSWidth;
                        double maxX = vp.ViewCenter.X + halfMSWidth;
                        double minY = vp.ViewCenter.Y - halfMSHeight;
                        double maxY = vp.ViewCenter.Y + halfMSHeight;

                        TypedValue[] filter = { new TypedValue(0, "MULTILEADER") };
                        SelectionFilter selFilter = new SelectionFilter(filter);

                        PromptSelectionResult selRes = ed.SelectAll(selFilter);

                        if (selRes.Status == PromptStatus.OK)
                        {
                            List<ObjectId> idsToMove = new List<ObjectId>();
                            foreach (ObjectId id in selRes.Value.GetObjectIds())
                            {
                                MLeader ml = (MLeader)tr.GetObject(id, OpenMode.ForRead);
                                Point3d arrowPt = GetMLeaderArrowhead(ml);

                                if (arrowPt.X >= minX && arrowPt.X <= maxX && arrowPt.Y >= minY && arrowPt.Y <= maxY)
                                {
                                    idsToMove.Add(id);
                                }
                            }

                            if (idsToMove.Count > 0)
                            {
                                ed.SwitchToModelSpace();
                                Application.SetSystemVariable("CVPORT", vp.Number);

                                ed.SetImpliedSelection(idsToMove.ToArray());
                                ed.Command("_.CHSPACE", "P", "", "");

                                ed.SwitchToPaperSpace();
                                ed.WriteMessage($"\nProcessed {idsToMove.Count} MLeaders in {layName}");
                            }
                        }
                    }
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nError: " + ex.Message);
            }
            finally
            {
                lm.CurrentLayout = originalLayout;
            }
        }

        // Updated helper to use ArrayList and explicit casting
        private Point3d GetMLeaderArrowhead(MLeader ml)
        {
            try
            {
                // 1. Get leader indexes as ArrayList
                ArrayList leaderIndexes = ml.GetLeaderIndexes();
                if (leaderIndexes != null && leaderIndexes.Count > 0)
                {
                    // 2. Get line indexes (Must cast leaderIndexes[0] from object to int)
                    ArrayList lineIndexes = ml.GetLeaderLineIndexes((int)leaderIndexes[0]);
                    if (lineIndexes != null && lineIndexes.Count > 0)
                    {
                        // 3. Get the vertex (Must cast lineIndexes[0] from object to int)
                        return ml.GetFirstVertex((int)lineIndexes[0]);
                    }
                }

                // Fallback to bounding box center
                Extents3d ext = ml.GeometricExtents;
                return ext.MinPoint + (ext.MaxPoint - ext.MinPoint) * 0.5;
            }
            catch
            {
                return Point3d.Origin;
            }
        }
    }

    // --- CUSTOM NATURAL SORT COMPARER ---
    public class NaturalStringComparer : IComparer<string>
    {
        public int Compare(string x, string y)
        {
            if (x == null || y == null) return 0;
            return NaturalSort(x, y);
        }

        private int NaturalSort(string s1, string s2)
        {
            return Regex.Replace(s1, "[0-9]+", match => match.Value.PadLeft(10, '0'))
                .CompareTo(Regex.Replace(s2, "[0-9]+", match => match.Value.PadLeft(10, '0')));
        }
    }

    public class GridData { public Polyline Boundary { get; set; } public string SheetNumber { get; set; } }
    // --- HELPER CLASS: Point Comparison with Tolerance ---
    public class Point3dComparer : IEqualityComparer<Point3d>
    {
        private readonly double _tolerance;
        public Point3dComparer(double tolerance) { _tolerance = tolerance; }

        public bool Equals(Point3d p1, Point3d p2)
        {
            return p1.DistanceTo(p2) < _tolerance;
        }

        public int GetHashCode(Point3d obj)
        {
            // Rounding to 3 decimals for the hash ensures nearby points group together
            return (Math.Round(obj.X, 3), Math.Round(obj.Y, 3), Math.Round(obj.Z, 3)).GetHashCode();
        }
    }

    public class UtilityIntersectionData
    {
        public int Sequence { get; set; }
        public Point3d Point { get; set; }
        public string UtilityType { get; set; }
        public string SheetNum { get; set; }
        public int GridSortIndex { get; set; }
    }
}
