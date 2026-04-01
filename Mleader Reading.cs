using IntelliCAD.ApplicationServices;
using IntelliCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions; // Add this at the top
using System.Threading.Tasks;
using System.Windows.Forms;
using Teigha.DatabaseServices;
using Teigha.Runtime;
using Application = IntelliCAD.ApplicationServices.Application;

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

        [CommandMethod("DirectSyncTable")]
        public void DirectSyncTable()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // This map specifically tracks OldValue -> NewValue for every MLeader that changed
            Dictionary<string, string> migrationMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // --- STEP 1: RESEQUENCE AND CAPTURE CHANGES ---
                DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                var sortedLayouts = layoutDict.Cast<DBDictionaryEntry>()
                    .Select(entry => tr.GetObject(entry.Value, OpenMode.ForRead) as Layout)
                    .Where(lay => !lay.ModelType)
                    .OrderBy(lay => lay.TabOrder);

                int counter = 1;
                string targetBlock = "CIRCLE FOR LEADER";
                string targetTag = "PH";

                foreach (Layout lay in sortedLayouts)
                {
                    BlockTableRecord btr = tr.GetObject(lay.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                    foreach (ObjectId id in btr)
                    {
                        if (id.ObjectClass.Name == "AcDbMLeader")
                        {
                            MLeader ml = tr.GetObject(id, OpenMode.ForRead) as MLeader;
                            if (ml != null && ml.ContentType == ContentType.BlockContent)
                            {
                                BlockTableRecord blkDef = tr.GetObject(ml.BlockContentId, OpenMode.ForRead) as BlockTableRecord;
                                if (blkDef.Name.Equals(targetBlock, StringComparison.OrdinalIgnoreCase))
                                {
                                    foreach (ObjectId attId in blkDef)
                                    {
                                        AttributeDefinition attDef = tr.GetObject(attId, OpenMode.ForRead) as AttributeDefinition;
                                        if (attDef != null && attDef.Tag.Equals(targetTag, StringComparison.OrdinalIgnoreCase))
                                        {
                                            string newVal = "P" + counter;
                                            using (AttributeReference attRef = ml.GetBlockAttribute(attId))
                                            {
                                                string oldVal = attRef.TextString;

                                                // If it changed, record the exact link
                                                if (oldVal != newVal)
                                                {
                                                    // We use the last assigned value for that ID to update the table
                                                    migrationMap[oldVal] = newVal;

                                                    ml.UpgradeOpen();
                                                    ml.SetBlockAttribute(attId, new AttributeReference { TextString = newVal });
                                                }
                                            }
                                            counter++;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // --- STEP 2: UPDATE TABLE BASED ON PREVIOUS VALUE MATCH ---
                if (migrationMap.Count > 0)
                {
                    foreach (Layout lay in sortedLayouts)
                    {
                        BlockTableRecord btr = tr.GetObject(lay.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                        foreach (ObjectId id in btr)
                        {
                            if (id.ObjectClass.Name == "AcDbTable")
                            {
                                Table tbl = tr.GetObject(id, OpenMode.ForRead) as Table;
                                if (tbl.Layer.Equals("NOTES", StringComparison.OrdinalIgnoreCase))
                                {
                                    tbl.UpgradeOpen();
                                    for (int r = 0; r < tbl.Rows.Count; r++)
                                    {
                                        for (int c = 0; c < tbl.Columns.Count; c++)
                                        {
                                            string raw = tbl.Cells[r, c].TextString;

                                            // Extract the ID from the formatted string (e.g., P2)
                                            Match m = Regex.Match(raw, @"P\d+");
                                            if (m.Success)
                                            {
                                                string cellValue = m.Value;

                                                // If the existing cell value matches a "Previous Value" we recorded...
                                                if (migrationMap.ContainsKey(cellValue))
                                                {
                                                    string updatedVal = migrationMap[cellValue];

                                                    // Update the cell with the new linked value
                                                    tbl.Cells[r, c].TextString = raw.Replace(cellValue, updatedVal);
                                                    ed.WriteMessage($"\nTable Sync: Cell {cellValue} updated to {updatedVal}");
                                                }
                                            }
                                        }
                                    }
                                    tbl.RecomputeTableBlock(true);
                                }
                            }
                        }
                    }
                }
                tr.Commit();
            }
        }

        
        [CommandMethod("ResequenceAndSyncTable")]
        public void ResequenceAndSyncTable()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Dictionary to store the "Before and After" of every MLeader
            // Key = Previous Value (e.g., P2), Value = New Value (e.g., P3)
            Dictionary<string, string> migrationMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // To satisfy your "If already matching, no need to edit" rule:
            HashSet<string> alreadyCorrectValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 1. Get Sorted Layouts
                DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                var sortedLayouts = layoutDict.Cast<DBDictionaryEntry>()
                    .Select(entry => tr.GetObject(entry.Value, OpenMode.ForRead) as Layout)
                    .Where(lay => !lay.ModelType)
                    .OrderBy(lay => lay.TabOrder)
                    .ToList();

                int sequenceCounter = 1;
                string targetBlockName = "CIRCLE FOR LEADER";
                string targetTagName = "PH";

                ed.WriteMessage("\n--- Phase 1: Processing MLeaders ---");

                foreach (Layout lay in sortedLayouts)
                {
                    BlockTableRecord btr = tr.GetObject(lay.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                    foreach (ObjectId entId in btr)
                    {
                        if (entId.ObjectClass.Name == "AcDbMLeader")
                        {
                            MLeader ml = tr.GetObject(entId, OpenMode.ForRead) as MLeader;
                            if (ml != null && ml.ContentType == ContentType.BlockContent)
                            {
                                BlockTableRecord mLeaderBlkDef = tr.GetObject(ml.BlockContentId, OpenMode.ForRead) as BlockTableRecord;
                                if (mLeaderBlkDef.Name.Equals(targetBlockName, StringComparison.OrdinalIgnoreCase))
                                {
                                    foreach (ObjectId attId in mLeaderBlkDef)
                                    {
                                        AttributeDefinition attDef = tr.GetObject(attId, OpenMode.ForRead) as AttributeDefinition;
                                        if (attDef != null && attDef.Tag.Equals(targetTagName, StringComparison.OrdinalIgnoreCase))
                                        {
                                            string newValue = "P" + sequenceCounter;
                                            using (AttributeReference attRef = ml.GetBlockAttribute(attId))
                                            {
                                                string oldValue = attRef.TextString;

                                                // Record the "Old -> New" transition
                                                if (oldValue != newValue)
                                                {
                                                    if (!migrationMap.ContainsKey(oldValue))
                                                        migrationMap.Add(oldValue, newValue);

                                                    ml.UpgradeOpen();
                                                    ml.SetBlockAttribute(attId, new AttributeReference { TextString = newValue });
                                                }

                                                // Keep track of what the "Correct" value is now
                                                alreadyCorrectValues.Add(newValue);
                                            }
                                            sequenceCounter++;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // 2. Phase 2: Targeted Table Update
                ed.WriteMessage("\n--- Phase 2: Updating Tables (Layer: NOTES) ---");
                foreach (Layout lay in sortedLayouts)
                {
                    BlockTableRecord btr = tr.GetObject(lay.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                    foreach (ObjectId entId in btr)
                    {
                        if (entId.ObjectClass.Name == "AcDbTable")
                        {
                            Table tbl = tr.GetObject(entId, OpenMode.ForRead) as Table;
                            if (tbl.Layer.Equals("NOTES", StringComparison.OrdinalIgnoreCase))
                            {
                                tbl.UpgradeOpen();
                                for (int r = 0; r < tbl.Rows.Count; r++)
                                {
                                    for (int c = 0; c < tbl.Columns.Count; c++)
                                    {
                                        string rawText = tbl.Cells[r, c].TextString;
                                        Match match = Regex.Match(rawText, @"P\d+");

                                        if (match.Success)
                                        {
                                            string cellValue = match.Value;

                                            // RULE A: If cell matches the NEW value already, skip (No need to edit)
                                            if (alreadyCorrectValues.Contains(cellValue))
                                            {
                                                continue;
                                            }

                                            // RULE B: If cell matches an OLD value that was changed, update it
                                            if (migrationMap.ContainsKey(cellValue))
                                            {
                                                string updatedValue = migrationMap[cellValue];
                                                tbl.Cells[r, c].TextString = rawText.Replace(cellValue, updatedValue);
                                                ed.WriteMessage($"\nSyncing Cell: {cellValue} -> {updatedValue}");
                                            }
                                        }
                                    }
                                }
                                tbl.RecomputeTableBlock(true);
                            }
                        }
                    }
                }
                tr.Commit();
                ed.WriteMessage("\nDone.");
            }
        }


        [CommandMethod("ResequenceWithLog")]
        public void ResequenceWithLog()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            Dictionary<string, string> migrationMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> alreadyCorrectValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // To build the display list
            StringBuilder changeLog = new StringBuilder();
            changeLog.AppendLine("\n--- MLeader Value Change List ---");
            changeLog.AppendLine("OLD VALUE\t->\tNEW VALUE");
            changeLog.AppendLine("------------------------------------");

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var sortedLayouts = GetSortedLayouts(tr, db);

                int sequenceCounter = 1;
                string targetBlockName = "CIRCLE FOR LEADER";
                string targetTagName = "PH";

                // --- PHASE 1: PROCESS MLEADERS ---
                foreach (Layout lay in sortedLayouts)
                {
                    BlockTableRecord btr = tr.GetObject(lay.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                    foreach (ObjectId entId in btr)
                    {
                        if (entId.ObjectClass.Name == "AcDbMLeader")
                        {
                            MLeader ml = tr.GetObject(entId, OpenMode.ForRead) as MLeader;
                            if (ml != null && ml.ContentType == ContentType.BlockContent)
                            {
                                BlockTableRecord mLeaderBlkDef = tr.GetObject(ml.BlockContentId, OpenMode.ForRead) as BlockTableRecord;
                                if (mLeaderBlkDef.Name.Equals(targetBlockName, StringComparison.OrdinalIgnoreCase))
                                {
                                    foreach (ObjectId attId in mLeaderBlkDef)
                                    {
                                        AttributeDefinition attDef = tr.GetObject(attId, OpenMode.ForRead) as AttributeDefinition;
                                        if (attDef != null && attDef.Tag.Equals(targetTagName, StringComparison.OrdinalIgnoreCase))
                                        {
                                            string newValue = "P" + sequenceCounter;
                                            using (AttributeReference attRef = ml.GetBlockAttribute(attId))
                                            {
                                                string oldValue = attRef.TextString;

                                                if (oldValue != newValue)
                                                {
                                                    if (!migrationMap.ContainsKey(oldValue))
                                                    {
                                                        migrationMap.Add(oldValue, newValue);
                                                        // Add to our display list
                                                        changeLog.AppendLine($"{oldValue}\t\t->\t{newValue}");
                                                    }

                                                    ml.UpgradeOpen();
                                                    ml.SetBlockAttribute(attId, new AttributeReference { TextString = newValue });
                                                }
                                                alreadyCorrectValues.Add(newValue);
                                            }
                                            sequenceCounter++;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // --- PHASE 2: UPDATE TABLES ---
                foreach (Layout lay in sortedLayouts)
                {
                    BlockTableRecord btr = tr.GetObject(lay.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                    foreach (ObjectId entId in btr)
                    {
                        if (entId.ObjectClass.Name == "AcDbTable")
                        {
                            Table tbl = tr.GetObject(entId, OpenMode.ForRead) as Table;
                            if (tbl.Layer.Equals("NOTES", StringComparison.OrdinalIgnoreCase))
                            {
                                tbl.UpgradeOpen();
                                for (int r = 0; r < tbl.Rows.Count; r++)
                                {
                                    for (int c = 0; c < tbl.Columns.Count; c++)
                                    {
                                        string rawText = tbl.Cells[r, c].TextString;
                                        Match match = Regex.Match(rawText, @"P\d+");

                                        if (match.Success)
                                        {
                                            string cellValue = match.Value;

                                            // DEBUG 1: See every P# found in the table
                                            ed.WriteMessage($"\nChecking Table Cell: {cellValue}");

                                            // CHECK 1: Is it already correct?
                                            if (alreadyCorrectValues.Contains(cellValue))
                                            {
                                                ed.WriteMessage($" -> Skipping {cellValue} (Already exists in drawing)");
                                                continue;
                                            }

                                            // CHECK 2: Do we have a migration for it?
                                            if (migrationMap.ContainsKey(cellValue))
                                            {
                                                string updatedValue = migrationMap[cellValue];

                                                // SUCCESS DEBUG: This should now pop up
                                                System.Windows.Forms.MessageBox.Show(
                                                    $"Matching Migration Found!\n" +
                                                    $"Old Value in Cell: {cellValue}\n" +
                                                    $"New Value to Apply: {updatedValue}",
                                                    "Sync Action");

                                                tbl.Cells[r, c].TextString = rawText.Replace(cellValue, updatedValue);
                                            }
                                            else
                                            {
                                                // DEBUG 2: If it's not correct AND not in the map
                                                ed.WriteMessage($" -> No migration found for {cellValue}");
                                            }
                                        }
                                    }
                                }
                                tbl.RecomputeTableBlock(true);
                            }
                        }
                    }
                }

                tr.Commit();
            }

            // --- DISPLAY THE LIST ---
            if (migrationMap.Count > 0)
            {
                ed.WriteMessage(changeLog.ToString());
                // Optional: Show in a popup window
                // System.Windows.Forms.MessageBox.Show(changeLog.ToString(), "Sequence Update Summary");
            }
            else
            {
                ed.WriteMessage("\nNo values required updating.");
            }
        }

        private IEnumerable<Layout> GetSortedLayouts(Transaction tr, Database db)
        {
            DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
            return layoutDict.Cast<DBDictionaryEntry>()
                .Select(entry => tr.GetObject(entry.Value, OpenMode.ForRead) as Layout)
                .Where(lay => !lay.ModelType)
                .OrderBy(lay => lay.TabOrder);
        }

        [CommandMethod("SyncByLayout")]
        public void SyncByLayout()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var sortedLayouts = GetSortedLayouts(tr, db);
                int globalSequenceCounter = 1; // Keep this global if you want P1, P2, P3 across all pages

                string targetBlockName = "CIRCLE FOR LEADER";
                string targetTagName = "PH";

                foreach (Layout lay in sortedLayouts)
                {
                    // --- NEW: These lists are RESET for every single page ---
                    HashSet<string> validIdsOnThisPage = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    Dictionary<string, string> migrationOnThisPage = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                    BlockTableRecord btr = tr.GetObject(lay.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;

                    // 1. SCAN MLEADERS ON THIS PAGE ONLY
                    foreach (ObjectId entId in btr)
                    {
                        if (entId.ObjectClass.Name == "AcDbMLeader")
                        {
                            MLeader ml = tr.GetObject(entId, OpenMode.ForRead) as MLeader;
                            if (ml != null && ml.ContentType == ContentType.BlockContent)
                            {
                                BlockTableRecord blkDef = tr.GetObject(ml.BlockContentId, OpenMode.ForRead) as BlockTableRecord;
                                if (blkDef.Name.Equals(targetBlockName, StringComparison.OrdinalIgnoreCase))
                                {
                                    foreach (ObjectId attId in blkDef)
                                    {
                                        AttributeDefinition attDef = tr.GetObject(attId, OpenMode.ForRead) as AttributeDefinition;
                                        if (attDef != null && attDef.Tag.Equals(targetTagName, StringComparison.OrdinalIgnoreCase))
                                        {
                                            string newValue = "P" + globalSequenceCounter;
                                            using (AttributeReference attRef = ml.GetBlockAttribute(attId))
                                            {
                                                string oldValue = attRef.TextString;
                                                if (oldValue != newValue)
                                                {
                                                    migrationOnThisPage[oldValue] = newValue;
                                                    ml.UpgradeOpen();
                                                    ml.SetBlockAttribute(attId, new AttributeReference { TextString = newValue });
                                                }
                                            }
                                            validIdsOnThisPage.Add(newValue);
                                            globalSequenceCounter++;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // 2. SCAN TABLES ON THIS PAGE ONLY
                    foreach (ObjectId entId in btr)
                    {
                        if (entId.ObjectClass.Name == "AcDbTable")
                        {
                            Table tbl = tr.GetObject(entId, OpenMode.ForRead) as Table;
                            if (tbl.Layer.Equals("NOTES", StringComparison.OrdinalIgnoreCase))
                            {
                                tbl.UpgradeOpen();
                                for (int r = 0; r < tbl.Rows.Count; r++)
                                {
                                    for (int c = 0; c < tbl.Columns.Count; c++)
                                    {
                                        string rawText = tbl.Cells[r, c].TextString;
                                        Match match = Regex.Match(rawText, @"P\d+");

                                        if (match.Success)
                                        {
                                            string cellValue = match.Value;

                                            // Skip if the ID already exists as a valid MLeader on this specific page
                                            if (validIdsOnThisPage.Contains(cellValue))
                                            {
                                                ed.WriteMessage($"\nLayout {lay.LayoutName}: Keeping {cellValue} (exists on page)");
                                                continue;
                                            }

                                            // If the cell value is an 'old' value we have a replacement for
                                            if (migrationOnThisPage.ContainsKey(cellValue))
                                            {
                                                string newValue = migrationOnThisPage[cellValue];

                                                // UPDATED MESSAGE: Showing Old vs New
                                                ed.WriteMessage($"\n[Layout: {lay.LayoutName}] Table Sync: Original ID '{cellValue}' replaced with New ID '{newValue}'");

                                                // Perform the replacement in the cell
                                                tbl.Cells[r, c].TextString = rawText.Replace(cellValue, newValue);
                                            }
                                        }
                                    }
                                }
                                tbl.RecomputeTableBlock(true);
                            }
                        }
                    }
                }
                tr.Commit();
            }
        }
        [CommandMethod("AdvancedSync")]
        public void AdvancedSync()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var sortedLayouts = GetSortedLayouts(tr, db);
                int globalCounter = 1; // Change to 1 inside the loop if you want reset per page

                foreach (Layout lay in sortedLayouts)
                {
                    BlockTableRecord btr = tr.GetObject(lay.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;

                    // 1. COLLECT AND SORT MLEADERS BY CURRENT ATTRIBUTE VALUE
                    var mleadersOnPage = GetSortedMLeadersOnPage(tr, btr);

                    if (mleadersOnPage.Count == 0) continue;

                    // 2. ASSIGN NEW VALUES AND STORE MAPPING
                    // List of (OldValue, NewValue)
                    List<Tuple<string, string>> pageMapping = new List<Tuple<string, string>>();
                    List<string> finalValuesOnPage = new List<string>();

                    foreach (var item in mleadersOnPage)
                    {
                        string newValue = "P" + globalCounter;
                        if (item.CurrentValue != newValue)
                        {
                            item.MLeaderObj.UpgradeOpen();
                            item.MLeaderObj.SetBlockAttribute(item.AttDefId, new AttributeReference { TextString = newValue });
                        }
                        pageMapping.Add(new Tuple<string, string>(item.CurrentValue, newValue));
                        finalValuesOnPage.Add(newValue);
                        globalCounter++;
                    }

                    // 3. SYNC TABLE
                    SyncTableOnPage(tr, btr, pageMapping, finalValuesOnPage, ed, lay.LayoutName);
                }
                tr.Commit();
            }
        }

        private void SyncTableOnPage(Transaction tr, BlockTableRecord btr, List<Tuple<string, string>> mapping, List<string> finalValues, Editor ed, string layoutName)
        {
            foreach (ObjectId id in btr)
            {
                if (id.ObjectClass.Name == "AcDbTable")
                {
                    Table tbl = tr.GetObject(id, OpenMode.ForRead) as Table;
                    if (tbl.Layer.Equals("NOTES", StringComparison.OrdinalIgnoreCase))
                    {
                        tbl.UpgradeOpen();

                        // Strategy: We are going to collect all P-values currently in the table
                        // and replace them with the new sequence in order.
                        int matchIndex = 0;

                        for (int r = 0; r < tbl.Rows.Count; r++)
                        {
                            for (int c = 0; c < tbl.Columns.Count; c++)
                            {
                                string raw = tbl.Cells[r, c].TextString;
                                if (Regex.IsMatch(raw, @"P\d+"))
                                {
                                    // If we still have MLeaders to put in the table
                                    if (matchIndex < finalValues.Count)
                                    {
                                        string newValue = finalValues[matchIndex];
                                        // Update cell with the new sequence value
                                        tbl.Cells[r, c].TextString = Regex.Replace(raw, @"P\d+", newValue);
                                        matchIndex++;
                                    }
                                    else
                                    {
                                        // If there are more table cells than MLeaders, clear the cell (Item was deleted)
                                        tbl.Cells[r, c].TextString = "";
                                        ed.WriteMessage($"\n[{layoutName}] Removed orphaned ID from table cell.");
                                    }
                                }
                            }
                        }

                        // If there were more MLeaders than Table Cells (New items added)
                        if (matchIndex < finalValues.Count)
                        {
                            ed.WriteMessage($"\n[{layoutName}] WARNING: {finalValues.Count - matchIndex} MLeaders added but table has no more rows.");
                        }

                        tbl.RecomputeTableBlock(true);
                    }
                }
            }
        }

        private List<MLeaderData> GetSortedMLeadersOnPage(Transaction tr, BlockTableRecord btr)
        {
            List<MLeaderData> list = new List<MLeaderData>();
            foreach (ObjectId id in btr)
            {
                if (id.ObjectClass.Name == "AcDbMLeader")
                {
                    MLeader ml = tr.GetObject(id, OpenMode.ForRead) as MLeader;
                    if (ml != null && ml.ContentType == ContentType.BlockContent)
                    {
                        BlockTableRecord blkDef = tr.GetObject(ml.BlockContentId, OpenMode.ForRead) as BlockTableRecord;
                        if (blkDef.Name.Equals("CIRCLE FOR LEADER", StringComparison.OrdinalIgnoreCase))
                        {
                            foreach (ObjectId attId in blkDef)
                            {
                                AttributeDefinition ad = tr.GetObject(attId, OpenMode.ForRead) as AttributeDefinition;
                                if (ad.Tag.Equals("PH", StringComparison.OrdinalIgnoreCase))
                                {
                                    using (AttributeReference ar = ml.GetBlockAttribute(attId))
                                    {
                                        list.Add(new MLeaderData
                                        {
                                            MLeaderObj = ml,
                                            CurrentValue = ar.TextString,
                                            AttDefId = attId,
                                            // Extract number for sorting (P12 -> 12)
                                            NumericValue = ExtractNumber(ar.TextString)
                                        });
                                    }
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            // Sort by the numerical value of the existing attribute
            return list.OrderBy(x => x.NumericValue).ToList();
        }

        private int ExtractNumber(string val)
        {
            Match m = Regex.Match(val, @"\d+");
            return m.Success ? int.Parse(m.Value) : 0;
        }

        [CommandMethod("FinalResequence")]
        public void FinalResequence()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            int globalTotalCount = 0;
            List<string> fileLogLines = new List<string>();
            fileLogLines.Add("LAYOUT\t\tOLD VALUE\t->\tNEW VALUE");
            fileLogLines.Add("---------------------------------------------------");

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var sortedLayouts = GetSortedLayouts(tr, db);
                int globalCounter = 1;

                foreach (Layout lay in sortedLayouts)
                {
                    BlockTableRecord btr = tr.GetObject(lay.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;

                    // 1. Get Sorted MLeaders for this specific page
                    var mleadersOnPage = GetSortedMLeadersOnPage(tr, btr);
                    if (mleadersOnPage.Count == 0) continue;

                    List<string> finalValuesOnPage = new List<string>();

                    foreach (var item in mleadersOnPage)
                    {
                        string newValue = "P" + globalCounter;

                        // Add to File Log
                        fileLogLines.Add($"{lay.LayoutName}\t\t{item.CurrentValue}\t\t->\t{newValue}");

                        if (item.CurrentValue != newValue)
                        {
                            item.MLeaderObj.UpgradeOpen();
                            item.MLeaderObj.SetBlockAttribute(item.AttDefId, new AttributeReference { TextString = newValue });
                        }

                        finalValuesOnPage.Add(newValue);
                        globalCounter++;
                        globalTotalCount++;
                    }

                    // 2. Sync Tables on this page (NOTES layer logic)
                    SyncTableOnPage(tr, btr, finalValuesOnPage);
                }

                // 3. Update Summary Table on Layout 1
                UpdateSummaryTable(tr, sortedLayouts, globalTotalCount, ed);

                tr.Commit();
            }

            // 4. Export to Text File
            ExportLogToFile(fileLogLines, globalTotalCount);

            ed.WriteMessage($"\nProcess Complete. Total MLeaders: {globalTotalCount}. Log saved to Drawing Folder.");
        }

        private void UpdateSummaryTable(Transaction tr, IEnumerable<Layout> layouts, int totalCount, Editor ed)
        {
            // Find Layout 1 (usually TabOrder 1)
            Layout lay1 = layouts.FirstOrDefault(l => l.TabOrder == 1);
            if (lay1 == null) return;

            BlockTableRecord btr = tr.GetObject(lay1.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
            foreach (ObjectId id in btr)
            {
                if (id.ObjectClass.Name == "AcDbTable")
                {
                    Table tbl = tr.GetObject(id, OpenMode.ForRead) as Table;
                    tbl.UpgradeOpen();

                    for (int r = 0; r < tbl.Rows.Count; r++)
                    {
                        // Check first column for the specific utility text
                        string firstColText = tbl.Cells[r, 0].TextString;
                        if (firstColText.Contains("POT-HOLE EXISTING UTILITY"))
                        {
                            // Update second column (Index 1) with the total count
                            tbl.Cells[r, 1].TextString = totalCount.ToString();
                            tbl.RecomputeTableBlock(true);
                            ed.WriteMessage($"\nSummary Table updated on {lay1.LayoutName} with count: {totalCount}");
                            return;
                        }
                    }
                }
            }
        }

        private void ExportLogToFile(List<string> lines, int total)
        {
            string desktop = System.IO.Path.GetDirectoryName(
                                Application.DocumentManager.MdiActiveDocument.Name);
            string drawingName =
                System.IO.Path.GetFileNameWithoutExtension(
                    Application.DocumentManager.MdiActiveDocument.Name);
            string newFolderPath = System.IO.Path.Combine(
                desktop, "BOM Output");
            if (!System.IO.Directory.Exists(newFolderPath))
                System.IO.Directory.CreateDirectory(newFolderPath);
            string filePath = Path.Combine(newFolderPath,drawingName + "POT-HOLE_Resequence_Log.txt");

            using (StreamWriter sw = new StreamWriter(filePath))
            {
                sw.WriteLine($"TOTAL POT-HOLE COUNT: {total}");
                sw.WriteLine($"DATE: {DateTime.Now}");
                sw.WriteLine();
                foreach (string line in lines) sw.WriteLine(line);
            }
        }

        // --- Reuse existing helper methods from previous steps ---
        private void SyncTableOnPage(Transaction tr, BlockTableRecord btr, List<string> finalValues)
        {
            foreach (ObjectId id in btr)
            {
                if (id.ObjectClass.Name == "AcDbTable")
                {
                    Table tbl = tr.GetObject(id, OpenMode.ForRead) as Table;
                    if (tbl.Layer.Equals("NOTES", StringComparison.OrdinalIgnoreCase))
                    {
                        tbl.UpgradeOpen();
                        int matchIndex = 0;
                        for (int r = 0; r < tbl.Rows.Count; r++)
                        {
                            for (int c = 0; c < tbl.Columns.Count; c++)
                            {
                                string raw = tbl.Cells[r, c].TextString;
                                if (Regex.IsMatch(raw, @"P\d+"))
                                {
                                    if (matchIndex < finalValues.Count)
                                    {
                                        tbl.Cells[r, c].TextString = Regex.Replace(raw, @"P\d+", finalValues[matchIndex]);
                                        matchIndex++;
                                    }
                                    else { tbl.Cells[r, c].TextString = ""; }
                                }
                            }
                        }
                        tbl.RecomputeTableBlock(true);
                    }
                }
            }
        }      

        
        
        
        
    }
    public class MLeaderData
    {
        public MLeader MLeaderObj;
        public string CurrentValue;
        public ObjectId AttDefId;
        public int NumericValue;
    }
}

   

