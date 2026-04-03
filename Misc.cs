
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
using Exception = System.Exception;

namespace Rough_Works
{
    internal class Misc
    {
        [CommandMethod("Place_PHLeader")]
        public void Place_PHLeader()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // 1. Get Point and String Inputs
            PromptPointOptions ppo = new PromptPointOptions("\nSpecify arrowhead point: ");
            PromptPointResult ppr1 = ed.GetPoint(ppo);
            if (ppr1.Status != PromptStatus.OK) return;

            PromptPointOptions ppo2 = new PromptPointOptions("\nSpecify landing point: ");
            ppo2.UseBasePoint = true;
            ppo2.BasePoint = ppr1.Value;
            PromptPointResult ppr2 = ed.GetPoint(ppo2);
            if (ppr2.Status != PromptStatus.OK) return;

            PromptStringOptions pso = new PromptStringOptions("\nEnter PH value: ");
            PromptResult psr = ed.GetString(pso);
            if (psr.Status != PromptStatus.OK) return;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                string blockName = "CIRCLE FOR LEADER";

                if (!bt.Has(blockName))
                {
                    ed.WriteMessage("\nError: Block 'CIRCLE FOR LEADER' not found.");
                    return;
                }

                ObjectId blockId = bt[blockName];

                // 2. Initialize the MLeader
                MLeader ml = new MLeader();
                ml.SetDatabaseDefaults();
                ml.ContentType = ContentType.BlockContent;
                ml.BlockContentId = blockId;

                // 3. Define the geometry (Leader Line)
                int ldIdx = ml.AddLeader();
                int lnIdx = ml.AddLeaderLine(ldIdx);
                ml.AddFirstVertex(lnIdx, ppr1.Value); // Arrowhead
                ml.AddLastVertex(lnIdx, ppr2.Value);  // Landing/Block location

                // 4. Set the Attribute Value
                // Note: Instead of ml.BlockTransform, we use a simple Matrix3d.Identity 
                // or the block's current scale/rotation.
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(blockId, OpenMode.ForRead);

                foreach (ObjectId id in btr)
                {
                    AttributeDefinition attDef = tr.GetObject(id, OpenMode.ForRead) as AttributeDefinition;
                    if (attDef != null && attDef.Tag.Equals("PH", StringComparison.OrdinalIgnoreCase))
                    {
                        // Create a new reference
                        using (AttributeReference attRef = new AttributeReference())
                        {
                            // Copy properties from the definition
                            attRef.SetAttributeFromBlock(attDef, Matrix3d.Identity);
                            attRef.TextString = psr.StringResult;

                            // Apply to the MLeader using the Attribute Definition's ObjectId
                            ml.SetBlockAttribute(id, attRef);
                        }
                        break;
                    }
                }

                // 5. Finalize
                BlockTableRecord curSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                curSpace.AppendEntity(ml);
                tr.AddNewlyCreatedDBObject(ml, true);

                tr.Commit();
                ed.WriteMessage("\nMLeader created successfully.");
            }
        }

        //[CommandMethod("PlacePHLeaderFinal")]
        //public void PlacePHLeaderFinal()
        //{
        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    Database db = doc.Database;
        //    Editor ed = doc.Editor;

        //    // 1. Get Inputs
        //    PromptPointOptions ppo1 = new PromptPointOptions("\nSpecify arrowhead point (Dot location): ");
        //    PromptPointResult ppr1 = ed.GetPoint(ppo1);
        //    if (ppr1.Status != PromptStatus.OK) return;

        //    PromptPointOptions ppo2 = new PromptPointOptions("\nSpecify bubble point: ");
        //    ppo2.UseBasePoint = true;
        //    ppo2.BasePoint = ppr1.Value;
        //    PromptPointResult ppr2 = ed.GetPoint(ppo2);
        //    if (ppr2.Status != PromptStatus.OK) return;

        //    PromptStringOptions pso = new PromptStringOptions("\nEnter PH value: ");
        //    PromptResult psr = ed.GetString(pso);
        //    if (psr.Status != PromptStatus.OK) return;

        //    using (Transaction tr = db.TransactionManager.StartTransaction())
        //    {
        //        // 2. Access the Style Dictionary correctly
        //        DBDictionary mlStyleDict = (DBDictionary)tr.GetObject(db.MLeaderStyleDictionaryId, OpenMode.ForRead);

        //        ObjectId styleId = db.MLeaderstyle; // Default style
        //        if (mlStyleDict.Contains("BUBBLE"))
        //        {
        //            styleId = mlStyleDict.GetAt("BUBBLE");
        //        }

        //        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        //        if (!bt.Has("CIRCLE FOR LEADER")) return;
        //        ObjectId blockId = bt["CIRCLE FOR LEADER"];

        //        // 3. Create MLeader
        //        MLeader ml = new MLeader();
        //        ml.SetDatabaseDefaults();
        //        ml.MLeaderStyle = styleId;
        //        ml.ContentType = ContentType.BlockContent;
        //        ml.BlockContentId = blockId;

        //        // 4. Setup Geometry
        //        int ldIdx = ml.AddLeader();
        //        int lnIdx = ml.AddLeaderLine(ldIdx);
        //        ml.AddFirstVertex(lnIdx, ppr1.Value);
        //        ml.AddLastVertex(lnIdx, ppr2.Value);

        //        // Properties from your screenshot
        //        ml.ArrowSize = 0.06053;
        //        // Setting the arrowhead to "Dot"
        //        // Note: This often requires the "Dot" block to be present in the BlockTable
        //        if (bt.Has("_Dot")) ml.SetArrowSymbolId(ldIdx, bt["_Dot"]);

        //        // 5. Set Attribute Value
        //        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(blockId, OpenMode.ForRead);
        //        foreach (ObjectId id in btr)
        //        {
        //            AttributeDefinition attDef = tr.GetObject(id, OpenMode.ForRead) as AttributeDefinition;
        //            if (attDef != null && attDef.Tag.Equals("PH", StringComparison.OrdinalIgnoreCase))
        //            {
        //                using (AttributeReference attRef = new AttributeReference())
        //                {
        //                    // Use Matrix3d.Identity since the MLeader manages the block transform internally
        //                    attRef.SetAttributeFromBlock(attDef, Matrix3d.Identity);
        //                    attRef.TextString = psr.StringResult;
        //                    ml.SetBlockAttribute(id, attRef);
        //                }
        //                break;
        //            }
        //        }

        //        // 6. Finalize
        //        BlockTableRecord curSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
        //        curSpace.AppendEntity(ml);
        //        tr.AddNewlyCreatedDBObject(ml, true);

        //        tr.Commit();
        //        ed.WriteMessage("\nMLeader placed successfully with Dot arrowhead.");
        //    }
        //}

        //[CommandMethod("PlacePHLeaderFinalFixed")]
        //public void PlacePHLeaderFinalFixed()
        //{
        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    Database db = doc.Database;
        //    Editor ed = doc.Editor;

        //    // 1. Get Inputs
        //    PromptPointOptions ppo1 = new PromptPointOptions("\nSpecify arrowhead point: ");
        //    PromptPointResult ppr1 = ed.GetPoint(ppo1);
        //    if (ppr1.Status != PromptStatus.OK) return;

        //    PromptPointOptions ppo2 = new PromptPointOptions("\nSpecify bubble point: ");
        //    ppo2.UseBasePoint = true;
        //    ppo2.BasePoint = ppr1.Value;
        //    PromptPointResult ppr2 = ed.GetPoint(ppo2);
        //    if (ppr2.Status != PromptStatus.OK) return;

        //    PromptStringOptions pso = new PromptStringOptions("\nEnter PH value: ");
        //    PromptResult psr = ed.GetString(pso);
        //    if (psr.Status != PromptStatus.OK) return;

        //    using (Transaction tr = db.TransactionManager.StartTransaction())
        //    {
        //        // 2. Access Dictionary and Block Table
        //        DBDictionary mlStyleDict = (DBDictionary)tr.GetObject(db.MLeaderStyleDictionaryId, OpenMode.ForRead);

        //        ObjectId styleId = db.MLeaderstyle;
        //        if (mlStyleDict.Contains("BUBBLE"))
        //            styleId = mlStyleDict.GetAt("BUBBLE");

        //        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        //        if (!bt.Has("CIRCLE FOR LEADER")) return;
        //        ObjectId blockId = bt["CIRCLE FOR LEADER"];

        //        // 3. Create MLeader
        //        MLeader ml = new MLeader();
        //        ml.SetDatabaseDefaults();
        //        ml.MLeaderStyle = styleId;
        //        ml.ContentType = ContentType.BlockContent;
        //        ml.BlockContentId = blockId;

        //        // --- FIX FOR "BIG BLOCK" ON SELECTION ---

        //        // 1. Set the block scale using Scale3d
        //        double s = 0.48425;
        //        ml.BlockScale = new Scale3d(s, s, s);

        //        // 2. Explicitly handle Annotative property
        //        // If your style is annotative, the Scale 1:1 might override your BlockScale 0.48
        //        // Setting this to False prevents the "jumping" size issue
        //        ml.Annotative = AnnotativeStates.False;

        //        // 3. Ensure the block connection is set to 'Center' 
        //        // (This matches your 'Center Extents' property)
        //        ml.BlockConnectionPoint = ppr2.Value;

        //        // 4. Setup Geometry (Leader Line)
        //        int ldIdx = ml.AddLeader();
        //        int lnIdx = ml.AddLeaderLine(ldIdx);
        //        ml.AddFirstVertex(lnIdx, ppr1.Value);
        //        ml.AddLastVertex(lnIdx, ppr2.Value);

        //        ml.ArrowSize = 0.06053;

        //        // 5. Set Attribute Value
        //        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(blockId, OpenMode.ForRead);
        //        foreach (ObjectId id in btr)
        //        {
        //            AttributeDefinition attDef = tr.GetObject(id, OpenMode.ForRead) as AttributeDefinition;
        //            if (attDef != null && attDef.Tag.Equals("PH", StringComparison.OrdinalIgnoreCase))
        //            {
        //                using (AttributeReference attRef = new AttributeReference())
        //                {
        //                    // --- TRANSFORM FIX: Use Matrix3d.Identity ---
        //                    // The MLeader manages the block's actual transform internally
        //                    attRef.SetAttributeFromBlock(attDef, Matrix3d.Identity);
        //                    attRef.TextString = psr.StringResult;

        //                    // Link the attribute reference to the MLeader instance
        //                    ml.SetBlockAttribute(id, attRef);
        //                }
        //                break;
        //            }
        //        }

        //        // 6. Append to current space
        //        BlockTableRecord curSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
        //        curSpace.AppendEntity(ml);
        //        tr.AddNewlyCreatedDBObject(ml, true);

        //        tr.Commit();
        //        ed.WriteMessage("\nMLeader created successfully.");
        //    }
        //}

        [CommandMethod("PlacePHLeader")]
        public void PlacePHLeader()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // 1. Get User Inputs
            PromptPointOptions ppo1 = new PromptPointOptions("\nSpecify arrowhead point: ");
            PromptPointResult ppr1 = ed.GetPoint(ppo1);
            if (ppr1.Status != PromptStatus.OK) return;

            PromptPointOptions ppo2 = new PromptPointOptions("\nSpecify bubble point: ");
            ppo2.UseBasePoint = true;
            ppo2.BasePoint = ppr1.Value;
            PromptPointResult ppr2 = ed.GetPoint(ppo2);
            if (ppr2.Status != PromptStatus.OK) return;

            PromptStringOptions pso = new PromptStringOptions("\nEnter PH value: ");
            PromptResult psr = ed.GetString(pso);
            if (psr.Status != PromptStatus.OK) return;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 2. Access Style and Block
                DBDictionary mlStyleDict = (DBDictionary)tr.GetObject(db.MLeaderStyleDictionaryId, OpenMode.ForRead);
                ObjectId styleId = db.MLeaderstyle;
                if (mlStyleDict.Contains("BUBBLE")) styleId = mlStyleDict.GetAt("BUBBLE");

                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (!bt.Has("CIRCLE FOR LEADER")) return;
                ObjectId blockId = bt["CIRCLE FOR LEADER"];

                // 3. Create MLeader and set requested properties
                MLeader ml = new MLeader();
                ml.SetDatabaseDefaults();
                ml.MLeaderStyle = styleId;

                // --- REQUESTED PROPERTIES ---
                ml.Layer = "0";                               // Layer: 0
                ml.Color = Teigha.Colors.Color.FromColorIndex(Teigha.Colors.ColorMethod.ByColor, 7); // Color: White/ByLayer
                ml.Scale = 0.0;                                // Overall Scale: 0
                ml.LandingGap = 0.0;                           // Landing Distance: 0
                ml.Annotative = AnnotativeStates.False;        // Ensure fixed scaling

                ml.ContentType = ContentType.BlockContent;
                ml.BlockContentId = blockId;

                // Block internal scale (as requested previously)
                double s = 0.48425;
                ml.BlockScale = new Scale3d(s, s, s);

                // 4. Setup Geometry
                int ldIdx = ml.AddLeader();
                int lnIdx = ml.AddLeaderLine(ldIdx);
                ml.AddFirstVertex(lnIdx, ppr1.Value);
                ml.AddLastVertex(lnIdx, ppr2.Value);

                ml.ArrowSize = 0.06053;

                // 5. Set the Attribute (PH Tag)
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(blockId, OpenMode.ForRead);
                foreach (ObjectId id in btr)
                {
                    AttributeDefinition attDef = tr.GetObject(id, OpenMode.ForRead) as AttributeDefinition;
                    if (attDef != null && attDef.Tag.Equals("PH", StringComparison.OrdinalIgnoreCase))
                    {
                        using (AttributeReference attRef = new AttributeReference())
                        {
                            attRef.SetAttributeFromBlock(attDef, Matrix3d.Identity);
                            attRef.TextString = psr.StringResult;
                            ml.SetBlockAttribute(id, attRef);
                        }
                        break;
                    }
                }

                // 6. Append to current space
                BlockTableRecord curSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                curSpace.AppendEntity(ml);
                tr.AddNewlyCreatedDBObject(ml, true);

                tr.Commit();
                ed.WriteMessage("\nMLeader placed on Layer '0' with Landing Distance 0.");
            }
        }

        [CommandMethod("PlacePHLeaderFinal")]
        public void PlacePHLeaderCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            // 1. Get Inputs (These would be passed automatically in a loop)
            PromptPointOptions ppo1 = new PromptPointOptions("\nSpecify arrowhead point: ");
            PromptPointResult ppr1 = ed.GetPoint(ppo1);
            if (ppr1.Status != PromptStatus.OK) return;

            PromptStringOptions pso = new PromptStringOptions("\nEnter PH value: ");
            PromptResult psr = ed.GetString(pso);
            if (psr.Status != PromptStatus.OK) return;

            // 2. Call the Reusable Method
            // We pass the transaction-ready logic inside
            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                CreatePHMLeader(tr, doc.Database, ppr1.Value, psr.StringResult);
                tr.Commit();
            }

            ed.WriteMessage("\nMLeader placed via reusable method.");
        }

        /// <summary>
        /// Reusable method to place a PH MLeader with a fixed offset bubble.
        /// </summary>
        public void CreatePHMLeader(Transaction tr, Database db, Point3d arrowheadPt, string phValue)
        {
            // 1. Define Fixed Offset (Adjust these values to change bubble distance)
            // Example: Place the bubble 2.0 units to the right and 2.0 units up from the arrowhead
            double offsetX = 2.0;
            double offsetY = 2.0;
            Point3d bubblePt = new Point3d(arrowheadPt.X + offsetX, arrowheadPt.Y + offsetY, arrowheadPt.Z);

            // 2. Access Style and Block
            DBDictionary mlStyleDict = (DBDictionary)tr.GetObject(db.MLeaderStyleDictionaryId, OpenMode.ForRead);
            ObjectId styleId = db.MLeaderstyle;
            if (mlStyleDict.Contains("BUBBLE")) styleId = mlStyleDict.GetAt("BUBBLE");

            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            if (!bt.Has("CIRCLE FOR LEADER")) return;
            ObjectId blockId = bt["CIRCLE FOR LEADER"];

            // 3. Create MLeader
            MLeader ml = new MLeader();
            ml.SetDatabaseDefaults();
            ml.MLeaderStyle = styleId;

            // Properties
            ml.Layer = "0";
            ml.Color = Teigha.Colors.Color.FromColorIndex(Teigha.Colors.ColorMethod.ByColor, 7);
            ml.Scale = 0.0;
            ml.LandingGap = 0.0;
            ml.Annotative = AnnotativeStates.False;
            ml.ContentType = ContentType.BlockContent;
            ml.BlockContentId = blockId;

            // Scale
            double s = 0.48425;
            ml.BlockScale = new Scale3d(s, s, s);

            // 4. Geometry
            int ldIdx = ml.AddLeader();
            int lnIdx = ml.AddLeaderLine(ldIdx);
            ml.AddFirstVertex(lnIdx, arrowheadPt); // Start at user pick
            ml.AddLastVertex(lnIdx, bubblePt);    // End at calculated offset

            ml.ArrowSize = 0.06053;

            // 5. Set Attribute
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

            // 6. Append
            BlockTableRecord curSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            curSpace.AppendEntity(ml);
            tr.AddNewlyCreatedDBObject(ml, true);
        }


        [CommandMethod("DEBUG_MLSTYLE_SCALE")]
        public void DebugMLeaderStyleScale()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                DBDictionary mlStyleDict = (DBDictionary)tr.GetObject(db.MLeaderStyleDictionaryId, OpenMode.ForRead);

                ed.WriteMessage("\n========== MLeader Style Diagnostics ==========");

                foreach (DBDictionaryEntry entry in mlStyleDict)
                {
                    MLeaderStyle mls = tr.GetObject(entry.Value, OpenMode.ForRead) as MLeaderStyle;
                    if (mls == null) continue;

                    Scale3d bs = mls.BlockScale;

                    ed.WriteMessage($"\nStyle Name    : {entry.Key}");
                    ed.WriteMessage($"  BlockScale  X: {bs.X}");
                    ed.WriteMessage($"  BlockScale  Y: {bs.Y}");
                    ed.WriteMessage($"  BlockScale  Z: {bs.Z}");
                    ed.WriteMessage($"  Scale        : {mls.Scale}");
                    ed.WriteMessage($"  ContentType  : {mls.ContentType}");
                }

                ed.WriteMessage("\n\n========== Live MLeader on screen (select one) ==========");

                PromptEntityOptions peo = new PromptEntityOptions("\nSelect an existing MLeader to inspect: ");
                peo.SetRejectMessage("\nMust be a MLeader.");
                peo.AddAllowedClass(typeof(MLeader), true);
                PromptEntityResult per = ed.GetEntity(peo);

                if (per.Status == PromptStatus.OK)
                {
                    MLeader ml = tr.GetObject(per.ObjectId, OpenMode.ForRead) as MLeader;
                    if (ml != null)
                    {
                        Scale3d bs2 = ml.BlockScale;
                        ed.WriteMessage($"\n  ml.BlockScale X  : {bs2.X}");
                        ed.WriteMessage($"\n  ml.BlockScale Y  : {bs2.Y}");
                        ed.WriteMessage($"\n  ml.BlockScale Z  : {bs2.Z}");
                        ed.WriteMessage($"\n  ml.ArrowSize     : {ml.ArrowSize}");
                        ed.WriteMessage($"\n  ml.LandingGap    : {ml.LandingGap}");
                        ed.WriteMessage($"\n  ml.Scale         : {ml.Scale}");

                        MLeaderStyle mls2 = tr.GetObject(ml.MLeaderStyle, OpenMode.ForRead) as MLeaderStyle;
                        if (mls2 != null)
                        {
                            Scale3d ss = mls2.BlockScale;
                            ed.WriteMessage($"\n  Style.BlockScale X : {ss.X}");
                            ed.WriteMessage($"\n  Style.Scale        : {mls2.Scale}");
                            ed.WriteMessage($"\n\n  ── Computed combinations ──");
                            ed.WriteMessage($"\n  ml.BlockScale.X × Style.BlockScale.X : {bs2.X * ss.X}");
                            ed.WriteMessage($"\n  ml.BlockScale.X × Style.Scale        : {bs2.X * mls2.Scale}");
                            ed.WriteMessage($"\n  ml.Scale        × Style.BlockScale.X : {ml.Scale * ss.X}");
                            ed.WriteMessage($"\n  ml.Scale        × Style.Scale        : {ml.Scale * mls2.Scale}");
                        }
                    }
                }

                tr.Commit();
            }
        }

        [CommandMethod("DEBUG_CREATED_MLEADER")]
        public void DebugCreatedMLeader()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                DBDictionary mlStyleDict = (DBDictionary)tr.GetObject(
                    db.MLeaderStyleDictionaryId, OpenMode.ForRead);

                ObjectId styleId = db.MLeaderstyle;
                if (mlStyleDict.Contains("BUBBLE"))
                    styleId = mlStyleDict.GetAt("BUBBLE");

                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (!bt.Has("CIRCLE FOR LEADER"))
                {
                    ed.WriteMessage("\nBlock 'CIRCLE FOR LEADER' not found.");
                    return;
                }
                ObjectId blockId = bt["CIRCLE FOR LEADER"];

                // ── Read block native size by iterating its entities directly ──
                BlockTableRecord btrCheck = (BlockTableRecord)tr.GetObject(blockId, OpenMode.ForRead);
                double nativeW = 0, nativeH = 0;
                try
                {
                    double minX = double.MaxValue, minY = double.MaxValue;
                    double maxX = double.MinValue, maxY = double.MinValue;

                    foreach (ObjectId entId in btrCheck)
                    {
                        Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                        if (ent == null) continue;
                        try
                        {
                            Extents3d e = ent.GeometricExtents;
                            if (e.MinPoint.X < minX) minX = e.MinPoint.X;
                            if (e.MinPoint.Y < minY) minY = e.MinPoint.Y;
                            if (e.MaxPoint.X > maxX) maxX = e.MaxPoint.X;
                            if (e.MaxPoint.Y > maxY) maxY = e.MaxPoint.Y;
                        }
                        catch { }
                    }

                    nativeW = maxX - minX;
                    nativeH = maxY - minY;
                    ed.WriteMessage($"\n  Block native width  : {nativeW}");
                    ed.WriteMessage($"\n  Block native height : {nativeH}");
                    ed.WriteMessage($"\n  Block native radius : {nativeW / 2.0}");
                }
                catch { ed.WriteMessage("\n  Could not read block extents."); }

                // ── Create a test MLeader exactly as your code does ──
                Point3d arrowPt = new Point3d(0, 0, 0);
                Point3d bubblePt = new Point3d(2.5, 10.5, 0);

                MLeader ml = new MLeader();
                ml.SetDatabaseDefaults();

                ObjectId msId = SymbolUtilityServices.GetBlockModelSpaceId(db);
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(msId, OpenMode.ForWrite);
                ms.AppendEntity(ml);
                tr.AddNewlyCreatedDBObject(ml, true);

                ml.MLeaderStyle = styleId;
                ml.Layer = "0";
                ml.ContentType = ContentType.BlockContent;
                ml.BlockContentId = blockId;
                ml.BlockScale = new Scale3d(1.0, 1.0, 1.0);
                ml.ArrowSize = 0.06053;
                ml.LandingGap = 0.012;
                ml.EnableAnnotationScale = false;

                int ldIdx = ml.AddLeader();
                int lnIdx = ml.AddLeaderLine(ldIdx);
                ml.AddFirstVertex(lnIdx, arrowPt);
                ml.AddLastVertex(lnIdx, bubblePt);

                // ── Report everything on the newly created MLeader ──
                ed.WriteMessage("\n\n========== Newly Created MLeader ==========");
                ed.WriteMessage($"\n  ml.BlockScale X  : {ml.BlockScale.X}");
                ed.WriteMessage($"\n  ml.BlockScale Y  : {ml.BlockScale.Y}");
                ed.WriteMessage($"\n  ml.Scale         : {ml.Scale}");

                // ── Read rendered size via GeometricExtents ──
                try
                {
                    Extents3d ext = ml.GeometricExtents;
                    double renderedW = ext.MaxPoint.X - ext.MinPoint.X;
                    double renderedH = ext.MaxPoint.Y - ext.MinPoint.Y;
                    ed.WriteMessage($"\n  Rendered width   : {renderedW}");
                    ed.WriteMessage($"\n  Rendered height  : {renderedH}");
                    ed.WriteMessage($"\n  Rendered radius  : {renderedW / 2.0}");
                    ed.WriteMessage($"\n\n  ── Scale factor needed ──");
                    ed.WriteMessage($"\n  Target diameter  : 0.48425");
                    ed.WriteMessage($"\n  Rendered diameter: {renderedW}");
                    ed.WriteMessage($"\n  Correction factor: {0.48425 / renderedW}");
                }
                catch (Exception ex)
                {
                    ed.WriteMessage($"\n  GeometricExtents failed: {ex.Message}");

                    // Fallback: compute correction from native block size directly
                    if (nativeW > 0)
                    {
                        ed.WriteMessage($"\n\n  ── Fallback correction from native block size ──");
                        ed.WriteMessage($"\n  Target diameter  : 0.48425");
                        ed.WriteMessage($"\n  Native diameter  : {nativeW}");
                        ed.WriteMessage($"\n  Correction factor: {0.48425 / nativeW}");
                    }
                }

                tr.Commit();
            }

            ed.WriteMessage("\n\nTest MLeader created at origin. UNDO after checking.");
        }

        [CommandMethod("DEBUG_BLOCK_REAL_SIZE")]
        public void DebugBlockRealSize()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptEntityOptions peo = new PromptEntityOptions("\nSelect existing correct MLeader (P2): ");
            peo.SetRejectMessage("\nMust select a MLeader.");          // ── must come BEFORE AddAllowedClass
            peo.AddAllowedClass(typeof(MLeader), true);
            PromptEntityResult per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                MLeader ml = tr.GetObject(per.ObjectId, OpenMode.ForRead) as MLeader;
                if (ml == null) { ed.WriteMessage("\nNot a MLeader."); return; }

                ed.WriteMessage("\n========== Existing MLeader (P2) ==========");
                ed.WriteMessage($"\n  ml.BlockScale X : {ml.BlockScale.X}");
                ed.WriteMessage($"\n  ml.BlockScale Y : {ml.BlockScale.Y}");
                ed.WriteMessage($"\n  ml.BlockScale Z : {ml.BlockScale.Z}");
                ed.WriteMessage($"\n  ml.Scale        : {ml.Scale}");
                ed.WriteMessage($"\n  ml.BlockPosition: {ml.BlockPosition}");

                ObjectId blockId = ml.BlockContentId;
                BlockTableRecord btr = tr.GetObject(blockId, OpenMode.ForRead) as BlockTableRecord;

                if (btr != null)
                {
                    // Measure native block size by iterating its entities
                    double minX = double.MaxValue, minY = double.MaxValue;
                    double maxX = double.MinValue, maxY = double.MinValue;

                    foreach (ObjectId entId in btr)
                    {
                        Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                        if (ent == null) continue;
                        try
                        {
                            Extents3d e = ent.GeometricExtents;
                            if (e.MinPoint.X < minX) minX = e.MinPoint.X;
                            if (e.MinPoint.Y < minY) minY = e.MinPoint.Y;
                            if (e.MaxPoint.X > maxX) maxX = e.MaxPoint.X;
                            if (e.MaxPoint.Y > maxY) maxY = e.MaxPoint.Y;
                        }
                        catch { }
                    }

                    double nativeW = maxX - minX;
                    ed.WriteMessage($"\n  Block native diameter : {nativeW}");
                    ed.WriteMessage($"\n  BlockScale.X applied  : {ml.BlockScale.X}");
                    ed.WriteMessage($"\n  native × BlockScale.X : {nativeW * ml.BlockScale.X}");

                    // ── The correct MLeader P2 visually shows 0.48425 diameter ──
                    // So whatever nativeW × BlockScale.X gives is the rendered size
                    // We need: what BlockScale produces 0.48425 from nativeW?
                    if (nativeW > 0)
                    {
                        double correctBlockScale = 0.48425 / nativeW;
                        ed.WriteMessage($"\n\n  ✔ BlockScale needed for 0.48425 diameter : {correctBlockScale}");
                        ed.WriteMessage($"\n  (i.e. 0.48425 / {nativeW} = {correctBlockScale})");
                    }
                }

                tr.Commit();
            }
        }
        [CommandMethod("DEBUG_COPY_MLEADER_PROPS")]
        public void DebugCopyMLeaderProps()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptEntityOptions peo = new PromptEntityOptions("\nSelect the CORRECT existing MLeader: ");
            peo.SetRejectMessage("\nMust select a MLeader.");
            peo.AddAllowedClass(typeof(MLeader), true);
            PromptEntityResult per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                MLeader src = tr.GetObject(per.ObjectId, OpenMode.ForRead) as MLeader;
                if (src == null) return;

                ed.WriteMessage("\n========== Source MLeader Properties ==========");
                ed.WriteMessage($"\n  BlockScale X          : {src.BlockScale.X}");
                ed.WriteMessage($"\n  BlockScale Y          : {src.BlockScale.Y}");
                ed.WriteMessage($"\n  BlockScale Z          : {src.BlockScale.Z}");
                ed.WriteMessage($"\n  Scale                 : {src.Scale}");
                ed.WriteMessage($"\n  ArrowSize             : {src.ArrowSize}");
                ed.WriteMessage($"\n  LandingGap            : {src.LandingGap}");
                ed.WriteMessage($"\n  EnableAnnotationScale : {src.EnableAnnotationScale}");
                ed.WriteMessage($"\n  BlockPosition         : {src.BlockPosition}");
                ed.WriteMessage($"\n  BlockRotation         : {src.BlockRotation}");
                //ed.WriteMessage($"\n  LeaderLineCount       : {src.LeaderLineCount}");

                // ── Create clone using CopyFrom ──
                ed.WriteMessage("\n\n========== Creating Clone ==========");

                MLeader clone = new MLeader();
                clone.SetDatabaseDefaults();

                ObjectId modelSpaceId = SymbolUtilityServices.GetBlockModelSpaceId(db);
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(modelSpaceId, OpenMode.ForWrite);
                ms.AppendEntity(clone);
                tr.AddNewlyCreatedDBObject(clone, true);

                // CopyFrom copies ALL internal state including hidden scale factors
                clone.CopyFrom(src);

                ed.WriteMessage($"\n  clone.Scale      : {clone.Scale}");
                ed.WriteMessage($"\n  clone.BlockScale : {clone.BlockScale.X}");

                // ── Rebuild leader geometry at offset position ──
                Point3d newArrow = new Point3d(src.BlockPosition.X + 5, src.BlockPosition.Y, 0);
                Point3d newBubble = new Point3d(newArrow.X + 2.5, newArrow.Y + 10.5, 0);

                // Remove all existing leaders from clone
                try
                {
                    System.Collections.ArrayList cloneLeaders = clone.GetLeaderIndexes();
                    // Iterate reverse to avoid index shifting during removal
                    for (int i = cloneLeaders.Count - 1; i >= 0; i--)
                        clone.RemoveLeader((int)cloneLeaders[i]);
                }
                catch (Exception ex)
                {
                    ed.WriteMessage($"\n  RemoveLeader error: {ex.Message}");
                }

                // Add fresh leader at new position
                int newLdIdx = clone.AddLeader();
                int newLnIdx = clone.AddLeaderLine(newLdIdx);
                clone.AddFirstVertex(newLnIdx, newArrow);
                clone.AddLastVertex(newLnIdx, newBubble);
                clone.BlockPosition = newBubble;

                // Set the PH attribute on the clone
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(src.BlockContentId, OpenMode.ForRead);
                foreach (ObjectId id in btr)
                {
                    AttributeDefinition attDef = tr.GetObject(id, OpenMode.ForRead) as AttributeDefinition;
                    if (attDef != null && attDef.Tag.Equals("PH", StringComparison.OrdinalIgnoreCase))
                    {
                        using (AttributeReference attRef = new AttributeReference())
                        {
                            attRef.SetAttributeFromBlock(attDef, Matrix3d.Identity);
                            attRef.TextString = "TEST";
                            clone.SetBlockAttribute(id, attRef);
                        }
                        break;
                    }
                }

                clone.RecordGraphicsModified(true);

                ed.WriteMessage($"\n  ✔ Clone placed at ({newBubble.X:F4}, {newBubble.Y:F4})");
                ed.WriteMessage("\n  >>> Check drawing — does clone MATCH source size? <<<");

                tr.Commit();
            }
        }
    }
}
