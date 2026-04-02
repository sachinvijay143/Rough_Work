
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
    }
}
