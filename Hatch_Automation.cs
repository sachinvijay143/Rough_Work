using IntelliCAD.ApplicationServices;
using IntelliCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Teigha.DatabaseServices;
using Teigha.Runtime;

namespace Rough_Works
{
    internal class Hatch_Automation
    {
        [CommandMethod("HatchInsideWidth")]
        public void HatchInsideWidth()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptEntityOptions peo = new PromptEntityOptions("\nSelect Polyline with width: ");
            peo.SetRejectMessage("\nMust be a Polyline.");
            peo.AddAllowedClass(typeof(Polyline), true);

            PromptEntityResult per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                Polyline pline = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Polyline;
                if (pline == null) return;

                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                // Calculate offset distance (half of Global Width)
                // If Global Width is 1, offset is 0.5
                double offsetDist = pline.ConstantWidth / 2.0;

                // Create the inner boundary
                // Note: Positive or negative determines inside/outside depending on clockwise/counter-clockwise
                DBObjectCollection offsetCurves = pline.GetOffsetCurves(-offsetDist);

                if (offsetCurves.Count > 0)
                {
                    Entity boundaryEnt = offsetCurves[0] as Entity;
                    // Add temporary boundary to database to get an ObjectId
                    ObjectId boundaryId = btr.AppendEntity(boundaryEnt);
                    tr.AddNewlyCreatedDBObject(boundaryEnt, true);

                    using (Hatch hatch = new Hatch())
                    {
                        hatch.SetDatabaseDefaults();
                        hatch.SetHatchPattern(HatchPatternType.PreDefined, "ANSI31");
                        hatch.PatternScale = 1.0;

                        btr.AppendEntity(hatch);
                        tr.AddNewlyCreatedDBObject(hatch, true);

                        // Attach to the OFFSET boundary, not the original thick polyline
                        ObjectIdCollection ids = new ObjectIdCollection();
                        ids.Add(boundaryId);

                        hatch.AppendLoop(HatchLoopTypes.Default, ids);
                        hatch.EvaluateHatch(true);

                        // Optional: If you want the hatch to follow the ORIGINAL polyline 
                        // you usually keep it non-associative if deleting the boundary.
                        hatch.Associative = false;
                    }

                    // Delete the temporary offset curve so only the hatch remains
                    boundaryEnt.Erase();
                }

                tr.Commit();
                ed.WriteMessage("\nPrecision hatch created inside the polyline width.");
            }
        }
    }
}
