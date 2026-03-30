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
    internal class Mleader_Reading
    {  
        [CommandMethod("ReadMLeaderBlock")]
        public void ReadMLeaderBlock()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // 1. Select the MLeader
            PromptEntityOptions opt = new PromptEntityOptions("\nSelect an MLeader with a block: ");
            opt.SetRejectMessage("\nObject must be an MLeader.");
            opt.AddAllowedClass(typeof(MLeader), false);

            PromptEntityResult res = ed.GetEntity(opt);
            if (res.Status != PromptStatus.OK) return;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                MLeader ml = tr.GetObject(res.ObjectId, OpenMode.ForRead) as MLeader;

                if (ml != null && ml.ContentType == ContentType.BlockContent)
                {
                    ObjectId blockDefId = ml.BlockContentId;
                    BlockTableRecord btr = tr.GetObject(blockDefId, OpenMode.ForRead) as BlockTableRecord;

                    ed.WriteMessage($"\n--- MLeader Block: {btr.Name} ---");

                    foreach (ObjectId id in btr)
                    {
                        // We only care about Attribute Definitions inside the block
                        AttributeDefinition attDef = tr.GetObject(id, OpenMode.ForRead) as AttributeDefinition;

                        if (attDef != null)
                        {
                            // Instead of GetBlockAttributeValue, we use GetBlockAttribute
                            // This returns an AttributeReference object for that specific Tag
                            using (AttributeReference attRef = ml.GetBlockAttribute(id))
                            {
                                string val = (attRef != null) ? attRef.TextString : "N/A";
                                ed.WriteMessage($"\nTag: {attDef.Tag} | Value: {val}");
                            }
                        }
                    }
                }
                else
                {
                    ed.WriteMessage("\nThe selected MLeader does not contain block content.");
                }
                tr.Commit();
            }
        }

        [CommandMethod("ResequenceMLeaders")]
        public void ResequenceMLeaders()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            int totalProcessedCount = 0; // Final counter

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 1. Get Layouts excluding Model Space, sorted by TabOrder
                DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                var sortedLayouts = layoutDict
                    .Cast<DBDictionaryEntry>()
                    .Select(entry => tr.GetObject(entry.Value, OpenMode.ForRead) as Layout)
                    .Where(lay => !lay.ModelType)
                    .OrderBy(lay => lay.TabOrder);

                int sequenceCounter = 1;
                string targetBlockName = "CIRCLE FOR LEADER";
                string targetTagName = "PH";
                string prefix = "P";

                ed.WriteMessage("\n--- Starting Resequence Process ---");

                foreach (Layout lay in sortedLayouts)
                {
                    // 2. Access the Paper Space BlockTableRecord for this layout
                    BlockTableRecord btr = tr.GetObject(lay.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;

                    foreach (ObjectId entId in btr)
                    {
                        // Check if entity is an MLeader
                        if (entId.ObjectClass.Name == "AcDbMLeader")
                        {
                            MLeader ml = tr.GetObject(entId, OpenMode.ForRead) as MLeader;

                            if (ml != null && ml.ContentType == ContentType.BlockContent)
                            {
                                BlockTableRecord mLeaderBlkDef = tr.GetObject(ml.BlockContentId, OpenMode.ForRead) as BlockTableRecord;

                                // 3. Verify Block Name
                                if (mLeaderBlkDef.Name.Equals(targetBlockName, StringComparison.OrdinalIgnoreCase))
                                {
                                    // 4. Find the "PH" Attribute Definition in the block
                                    foreach (ObjectId attId in mLeaderBlkDef)
                                    {
                                        AttributeDefinition attDef = tr.GetObject(attId, OpenMode.ForRead) as AttributeDefinition;

                                        if (attDef != null && attDef.Tag.Equals(targetTagName, StringComparison.OrdinalIgnoreCase))
                                        {
                                            string expectedVal = prefix + sequenceCounter.ToString();

                                            using (AttributeReference attRef = ml.GetBlockAttribute(attId))
                                            {
                                                string currentVal = attRef.TextString;

                                                // 5. Update only if value is different
                                                if (currentVal != expectedVal)
                                                {
                                                    ml.UpgradeOpen();
                                                    // Note: We pass a new AttributeReference with the updated text
                                                    ml.SetBlockAttribute(attId, new AttributeReference { TextString = expectedVal });
                                                }
                                            }

                                            sequenceCounter++;
                                            totalProcessedCount++;
                                            break; // Found the tag, move to next MLeader
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                tr.Commit();

                // 6. Display final summary
                ed.WriteMessage("\n------------------------------------");
                ed.WriteMessage($"\nResequencing complete!");
                ed.WriteMessage($"\nTotal MLeader Blocks Found/Updated: {totalProcessedCount}");
                ed.WriteMessage("\n------------------------------------");
            }
        }

    }
}
