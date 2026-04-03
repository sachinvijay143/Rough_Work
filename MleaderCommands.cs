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
    public class MLeaderCommands
    {
        // ─────────────────────────────────────────────────────────────
        //  Helper DTO to carry viewport data between transactions
        // ─────────────────────────────────────────────────────────────
        private class ViewportInfo
        {
            public ObjectId ViewportId { get; set; }
            public int VpNumber { get; set; }   // CVPORT number
            public Point2d ViewCenter { get; set; }   // Model-space centre of the view
            public double ViewHeight { get; set; }   // Model-space height of the view
            public double PsWidth { get; set; }   // Paper-space width  of the viewport frame
            public double PsHeight { get; set; }   // Paper-space height of the viewport frame
        }



        // ─────────────────────────────────────────────────────────────
        //  Entry point
        // ─────────────────────────────────────────────────────────────

        [CommandMethod("MLEADER_CHSPACE_ALL")]
        public void MLeaderChspaceAll()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;



            // ── Save state we will restore at the end ──
            string savedLayout = LayoutManager.Current.CurrentLayout;
            object savedTileMode = Application.GetSystemVariable("TILEMODE");
            object savedCvport = Application.GetSystemVariable("CVPORT");
            object savedPickfirst = Application.GetSystemVariable("PICKFIRST");
            try
            {
                // Ensure PICKFIRST is ON so CHSPACE can consume the implied selection
                Application.SetSystemVariable("PICKFIRST", 1);
                // ── Step 1 : collect every paper-space layout name ──
                List<string> paperLayouts = CollectPaperLayouts(db);
                if (paperLayouts.Count == 0)
                {
                    ed.WriteMessage("\nNo paper-space layouts found. Aborting.");
                    return;
                }
                // ── Step 2 : loop through layouts ──
                foreach (string layoutName in paperLayouts)
                {
                    ed.WriteMessage($"\n\n=== Layout : {layoutName} ===");
                    // Switch to paper space for this layout
                    LayoutManager.Current.CurrentLayout = layoutName;
                    Application.SetSystemVariable("TILEMODE", 0); // ensure paper space
                    // Collect viewport data (own transaction – read-only)
                    List<ViewportInfo> viewports = CollectViewportsInLayout(db, layoutName);
                    if (viewports.Count == 0)
                    {
                        ed.WriteMessage("  No active viewports found.");
                        continue;
                    }

                    // ── Step 3 : loop through viewports ──

                    foreach (ViewportInfo vp in viewports)

                    {

                        ed.WriteMessage($"\n  ▶ Viewport #{vp.VpNumber} ...");

                        ProcessViewport(doc, ed, vp, layoutName);

                    }

                }

            }

            catch (Teigha.Runtime.Exception ex)

            {

                ed.WriteMessage($"\n** ERROR : {ex.Message}");

            }

            finally

            {

                // ── Restore original AutoCAD state ──

                try

                {

                    LayoutManager.Current.CurrentLayout = savedLayout;

                    Application.SetSystemVariable("TILEMODE", savedTileMode);

                    Application.SetSystemVariable("CVPORT", savedCvport);

                    Application.SetSystemVariable("PICKFIRST", savedPickfirst);

                }

                catch { /* best-effort restore */ }



                ed.WriteMessage("\n\n✔ MLEADER_CHSPACE_ALL finished.\n");

            }

        }



        // ─────────────────────────────────────────────────────────────

        //  Returns names of all non-model layouts

        // ─────────────────────────────────────────────────────────────

        private List<string> CollectPaperLayouts(Database db)

        {

            var result = new List<string>();

            using (Transaction tr = db.TransactionManager.StartTransaction())

            {

                var dict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                foreach (DBDictionaryEntry entry in dict)

                {

                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);

                    if (!layout.ModelType)

                        result.Add(layout.LayoutName);

                }

                tr.Commit();

            }

            return result;

        }



        // ─────────────────────────────────────────────────────────────

        //  Returns all "real" (non-overall) active viewports in a layout

        // ─────────────────────────────────────────────────────────────

        private List<ViewportInfo> CollectViewportsInLayout(Database db, string layoutName)

        {

            var result = new List<ViewportInfo>();

            using (Transaction tr = db.TransactionManager.StartTransaction())

            {

                var dict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                var layout = (Layout)tr.GetObject(dict.GetAt(layoutName), OpenMode.ForRead);

                var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);



                foreach (ObjectId id in btr)

                {

                    var ent = tr.GetObject(id, OpenMode.ForRead) as Viewport;



                    // VpNumber == 1 is the overall paper-space viewport – skip it

                    if (ent == null || ent.Number <= 1 || !ent.On)

                        continue;



                    result.Add(new ViewportInfo

                    {

                        ViewportId = id,

                        VpNumber = ent.Number,

                        ViewCenter = ent.ViewCenter,   // 2-D model-space centre of the view

                        ViewHeight = ent.ViewHeight,   // model-space height visible in the VP

                        PsWidth = ent.Width,        // paper-space frame width

                        PsHeight = ent.Height        // paper-space frame height

                    });

                }

                tr.Commit();

            }

            return result;

        }



        // ─────────────────────────────────────────────────────────────

        //  Activate MSPACE for one viewport, select its MLeaders,

        //  run CHSPACE synchronously via ed.Command()

        // ─────────────────────────────────────────────────────────────

        private void ProcessViewport(Document doc, Editor ed,

                                     ViewportInfo vp, string layoutName)

        {

            // ── Activate MSPACE inside this specific viewport ──

            // Setting CVPORT to the viewport's number switches to model space

            // inside that viewport (equivalent to double-clicking into it).

            Application.SetSystemVariable("CVPORT", vp.VpNumber);



            // ── Build the model-space bounding box of this viewport ──

            // ViewHeight  = model-space height the viewport shows

            // aspectRatio = PsWidth / PsHeight (paper-space frame shape)

            // modelWidth  = ViewHeight * aspectRatio

            double aspectRatio = vp.PsHeight > 0 ? vp.PsWidth / vp.PsHeight : 1.0;

            double halfW = (vp.ViewHeight * aspectRatio) / 2.0;

            double halfH = vp.ViewHeight / 2.0;



            // Z range is intentionally huge to catch MLeaders at any elevation

            Point3d minPt = new Point3d(vp.ViewCenter.X - halfW, vp.ViewCenter.Y - halfH, -1e15);

            Point3d maxPt = new Point3d(vp.ViewCenter.X + halfW, vp.ViewCenter.Y + halfH, 1e15);



            // ── Crossing-window selection of MULTILEADER objects only ──

            // Because we are now in MSPACE, SelectCrossingWindow operates

            // in model space – it will NOT pick up paper-space entities.

            var filter = new SelectionFilter(new[]

            {

            new TypedValue((int)DxfCode.Start, "MULTILEADER")

        });



            PromptSelectionResult selRes = ed.SelectCrossingWindow(minPt, maxPt, filter);



            if (selRes.Status != PromptStatus.OK || selRes.Value.Count == 0)

            {

                ed.WriteMessage("  No MLeaders found inside this viewport boundary.");

                return;

            }



            int count = selRes.Value.Count;

            ed.WriteMessage($"  Found {count} MLeader(s) – running CHSPACE …");



            // ── Load the selection into the pickfirst (implied) set ──

            ed.SetImpliedSelection(selRes.Value);



            // ── Run CHSPACE synchronously ──

            // ed.Command() is blocking; it waits until the AutoCAD command

            // completes before returning.  The trailing "" is the Enter

            // that ends the "Select objects:" prompt inside CHSPACE.

            // CHSPACE will consume the implied selection because PICKFIRST=1.

            ed.Command("_.CHSPACE", "");



            ed.WriteMessage($"  ✔ CHSPACE applied to {count} MLeader(s).");

        }

        [CommandMethod("ForceChspaceMLeader")]
        public void ForceChspaceMLeader()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // 1. Initial Check: Must be in Paper Space (Layout)
            if (db.TileMode)
            {
                ed.WriteMessage("\nSwitch to a Layout tab first.");
                return;
            }

            // 2. Select the MLeader while in Paper Space 
            // (Even if it's visible through a viewport, GetEntity works)
            PromptEntityOptions opt = new PromptEntityOptions("\nSelect the MLeader: ");
            opt.AddAllowedClass(typeof(MLeader), false);
            PromptEntityResult res = ed.GetEntity(opt);

            if (res.Status != PromptStatus.OK) return;
            ObjectId mlId = res.ObjectId;

            // 3. Switch to Model Space (Inside the Viewport)
            // We use the synchronous Command method to ensure it happens NOW
            ed.Command("._MSPACE");

            // 4. Force the MLeader into the "Active Selection Set" (PickFirst)
            // This is the C# equivalent of clicking the object manually
            ed.SetImpliedSelection(new ObjectId[] { mlId });

            // 5. Execute CHSPACE
            // Since the object is already selected (Implied), 
            // CHSPACE will immediately move it without asking for selection.
            doc.SendStringToExecute("._CHSPACE\n\n", true, false, false);

            // 6. Return to Paper Space
            doc.SendStringToExecute("._PSPACE\n", true, false, false);

            ed.WriteMessage("\nMLeader pushed to Paper Space.");
        }


        [CommandMethod("RUN_MLEADER_CHSPACE")]
        public void RunMLeaderChspace()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            string originalLayout = LayoutManager.Current.CurrentLayout;

            try
            {
                List<string> paperLayouts = GetPaperLayouts(db);

                if (paperLayouts.Count == 0)
                {
                    ed.WriteMessage("\nNo paper layouts found.");
                    return;
                }

                foreach (string layoutName in paperLayouts)
                {
                    ed.WriteMessage($"\n--- Processing Layout: {layoutName} ---");

                    // Switch to this paper layout (paper space)
                    LayoutManager.Current.CurrentLayout = layoutName;
                    Application.SetSystemVariable("TILEMODE", 0);

                    List<ViewportData> viewports = GetViewportsInLayout(db, layoutName);

                    if (viewports.Count == 0)
                    {
                        ed.WriteMessage("\n  No viewports found. Skipping.");
                        continue;
                    }

                    foreach (ViewportData vpData in viewports)
                    {
                        ed.WriteMessage($"\n  Entering MSPACE - Viewport #{vpData.Number}");

                        // Activate MSPACE inside this specific viewport
                        Application.SetSystemVariable("CVPORT", vpData.Number);

                        // Directly read MLeaders from model space and check against viewport bounds
                        ObjectId[] mleaderIds = GetMLeadersFromModelSpace(db, ed, vpData);

                        if (mleaderIds.Length == 0)
                        {
                            ed.WriteMessage($"\n  No MLeaders found in Viewport #{vpData.Number}. Skipping.");
                            continue;
                        }

                        ed.WriteMessage($"\n  Found {mleaderIds.Length} MLeader(s) in Viewport #{vpData.Number}.");

                        // Set the implied selection with found MLeaders
                        SelectionSet ss = SelectionSet.FromObjectIds(mleaderIds);
                        ed.SetImpliedSelection(ss);

                        // Execute CHSPACE - moves objects from model space to paper space
                        ed.WriteMessage($"\n  Executing CHSPACE...");
                        doc.SendStringToExecute("CHSPACE\n \n ", true, false, true);

                        ed.WriteMessage($"\n  CHSPACE executed for {mleaderIds.Length} MLeader(s).");
                    }
                }
            }
            finally
            {
                LayoutManager.Current.CurrentLayout = originalLayout;
                ed.WriteMessage("\n\nFinished. Original layout restored.");
            }
        }

        // ─────────────────────────────────────────────────────────────
        // Viewport data carrier
        // ─────────────────────────────────────────────────────────────
        private class ViewportData
        {
            public ObjectId Id { get; set; }
            public int Number { get; set; }
            public Point2d ViewCenter { get; set; }  // model-space centre of the view
            public double ViewHeight { get; set; }  // model-space height
            public double Width { get; set; }  // paper-space frame width
            public double Height { get; set; }  // paper-space frame height
            public Point3d CenterPS { get; set; }  // paper-space centre of the frame
        }

        // ─────────────────────────────────────────────────────────────
        // Get all paper-space layout names
        // ─────────────────────────────────────────────────────────────
        private List<string> GetPaperLayouts(Database db)
        {
            var layouts = new List<string>();
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var dict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                foreach (DBDictionaryEntry entry in dict)
                {
                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                    if (!layout.ModelType)
                        layouts.Add(layout.LayoutName);
                }
                tr.Commit();
            }
            return layouts;
        }

        // ─────────────────────────────────────────────────────────────
        // Get all active real viewports (skip pseudo-viewport #1)
        // ─────────────────────────────────────────────────────────────
        private List<ViewportData> GetViewportsInLayout(Database db, string layoutName)
        {
            var result = new List<ViewportData>();
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var dict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                var layout = (Layout)tr.GetObject(dict.GetAt(layoutName), OpenMode.ForRead);
                var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);

                foreach (ObjectId id in btr)
                {
                    var vp = tr.GetObject(id, OpenMode.ForRead) as Viewport;
                    if (vp == null || vp.Number <= 1 || !vp.On)
                        continue;

                    result.Add(new ViewportData
                    {
                        Id = id,
                        Number = vp.Number,
                        ViewCenter = vp.ViewCenter,
                        ViewHeight = vp.ViewHeight,
                        Width = vp.Width,
                        Height = vp.Height,
                        CenterPS = vp.CenterPoint
                    });
                }
                tr.Commit();
            }
            return result;
        }

        // ─────────────────────────────────────────────────────────────
        // Directly iterate model space block table record and
        // check each MULTILEADER against the viewport's model-space bounds.
        // This bypasses SelectCrossingWindow which is unreliable in MSPACE.
        // ─────────────────────────────────────────────────────────────
        private ObjectId[] GetMLeadersFromModelSpace(Database db, Editor ed, ViewportData vpData)
        {
            var found = new List<ObjectId>();

            // Compute model-space visible extents of this viewport
            double aspectRatio = (vpData.Height > 0) ? vpData.Width / vpData.Height : 1.0;
            double halfH = vpData.ViewHeight / 2.0;
            double halfW = halfH * aspectRatio;

            double minX = vpData.ViewCenter.X - halfW;
            double maxX = vpData.ViewCenter.X + halfW;
            double minY = vpData.ViewCenter.Y - halfH;
            double maxY = vpData.ViewCenter.Y + halfH;

            ed.WriteMessage($"\n    VP Bounds → X:[{minX:F2} to {maxX:F2}] Y:[{minY:F2} to {maxY:F2}]");

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Open model space block table record directly
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                foreach (ObjectId id in btr)
                {
                    // Only process MULTILEADER entities
                    var mleader = tr.GetObject(id, OpenMode.ForRead) as MLeader;
                    if (mleader == null)
                        continue;

                    // Use the MLeader's GeometricExtents to check if it falls within viewport bounds
                    try
                    {
                        Extents3d ext = mleader.GeometricExtents;

                        // Check if the MLeader's extents overlap or sit within the viewport bounds
                        bool withinX = ext.MinPoint.X <= maxX && ext.MaxPoint.X >= minX;
                        bool withinY = ext.MinPoint.Y <= maxY && ext.MaxPoint.Y >= minY;

                        if (withinX && withinY)
                        {
                            found.Add(id);
                            ed.WriteMessage($"\n    ✔ MLeader found at [{ext.MinPoint.X:F2}, {ext.MinPoint.Y:F2}]");
                        }
                    }
                    catch
                    {
                        // GeometricExtents can throw if entity has no geometry — skip safely
                    }
                }

                tr.Commit();
            }

            return found.ToArray();
        }
    }
}
