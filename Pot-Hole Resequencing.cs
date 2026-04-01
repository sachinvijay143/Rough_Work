using IntelliCAD.ApplicationServices;
using IntelliCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Teigha.DatabaseServices;
using Teigha.Runtime;
using Application = IntelliCAD.ApplicationServices.Application;

namespace Rough_Works
{
    internal class Pot_Hole_Resequencing
    {
        [CommandMethod("FinalResequenceWithTableLog")]
        public void FinalResequenceWithTableLog()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            int globalTotalCount = 0;
            List<string> fileLogLines = new List<string>();
            fileLogLines.Add($"RESEQUENCE LOG - {DateTime.Now}");
            fileLogLines.Add("------------------------------------------------------------");

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var sortedLayouts = GetSortedLayouts(tr, db);
                int globalCounter = 1;

                foreach (Layout lay in sortedLayouts)
                {
                    fileLogLines.Add($"\nLAYOUT: {lay.LayoutName}");
                    fileLogLines.Add("  MLeader Changes:");

                    BlockTableRecord btr = tr.GetObject(lay.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;

                    var mleadersOnPage = GetSortedMLeadersOnPage(tr, btr);

                    if (mleadersOnPage.Count == 0)
                    {
                        fileLogLines.Add("    -> No MLeaders found on this layout.");
                    }
                    else
                    {
                        List<string> finalValuesOnPage = new List<string>();

                        foreach (var item in mleadersOnPage)
                        {
                            string newValue = "P" + globalCounter;
                            fileLogLines.Add($"    {item.CurrentValue.PadRight(10)} -> {newValue}");

                            if (item.CurrentValue != newValue)
                            {
                                item.MLeaderObj.UpgradeOpen();
                                item.MLeaderObj.SetBlockAttribute(item.AttDefId, new AttributeReference { TextString = newValue });
                            }

                            finalValuesOnPage.Add(newValue);
                            globalCounter++;
                            globalTotalCount++;
                        }

                        // Sync Tables and get status for log
                        string tableStatus = SyncTableOnPage(tr, btr, finalValuesOnPage);
                        fileLogLines.Add($"  Table Status: {tableStatus}");
                    }
                    fileLogLines.Add("------------------------------------------------------------");
                }

                // Update Summary Table on Layout 1 (POT-HOLE EXISTING UTILITY)
                UpdateSummaryTable(tr, sortedLayouts, globalTotalCount, ed);

                tr.Commit();
            }

            ExportLogToFile(fileLogLines, globalTotalCount);
            ed.WriteMessage($"\nProcess Complete. Total MLeaders: {globalTotalCount}. Log saved to Desktop.");
        }

        private string SyncTableOnPage(Transaction tr, BlockTableRecord btr, List<string> finalValues)
        {
            bool tableFound = false;
            int cellsUpdated = 0;
            int totalPCellsFound = 0;
            string mismatchAlert = "";

            foreach (ObjectId id in btr)
            {
                if (id.ObjectClass.Name == "AcDbTable")
                {
                    Table tbl = tr.GetObject(id, OpenMode.ForRead) as Table;
                    if (tbl.Layer.Equals("NOTES", StringComparison.OrdinalIgnoreCase))
                    {
                        tableFound = true;
                        tbl.UpgradeOpen();
                        int matchIndex = 0;

                        // 1. First pass: Count how many cells currently have "P#"
                        for (int r = 0; r < tbl.Rows.Count; r++)
                        {
                            for (int c = 0; c < tbl.Columns.Count; c++)
                            {
                                if (Regex.IsMatch(tbl.Cells[r, c].TextString, @"P\d+"))
                                    totalPCellsFound++;
                            }
                        }

                        // 2. Second pass: Update or Clear based on MLeader count
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
                                        cellsUpdated++;
                                    }
                                    else
                                    {
                                        // MLeader was deleted - Clear current and subsequent columns in this row
                                        // We use a loop with a check to ensure we don't exceed tbl.Columns.Count
                                        for (int colToClear = c; colToClear < tbl.Columns.Count; colToClear++)
                                        {
                                            // Clear up to 4 columns (c, c+1, c+2, c+3) if they exist
                                            if (colToClear < c + 4)
                                            {
                                                tbl.Cells[r, colToClear].TextString = "";
                                            }
                                        }

                                        // Optional: Only show message once per row to avoid annoying the user
                                        // MessageBox.Show($"Row {r + 1} cleared: MLeader reference no longer exists.");

                                        // Break the inner loop for this row since we've cleared the data for this entry
                                        break;
                                    }
                                }
                            }
                        }

                        // 3. Logic for Mismatch reporting
                        if (finalValues.Count > totalPCellsFound)
                        {
                            int diff = finalValues.Count - totalPCellsFound;
                            mismatchAlert = $" [MISMATCH: {diff} MLeader(s) have no table rows!]";
                        }
                        else if (totalPCellsFound > finalValues.Count)
                        {
                            int diff = totalPCellsFound - finalValues.Count;
                            mismatchAlert = $" [MISMATCH: {diff} Table row(s) were cleared (MLeader deleted)]";
                        }
                        else
                        {
                            mismatchAlert = " [Count Matches]";
                        }

                        tbl.RecomputeTableBlock(true);
                    }
                }
            }

            if (!tableFound) return "No 'NOTES' layer table found.";

            return $"Success ({cellsUpdated} cells updated/verified out of {totalPCellsFound}){mismatchAlert}";
        }

        private void UpdateSummaryTable(Transaction tr, IEnumerable<Layout> layouts, int totalCount, Editor ed)
        {
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
                        if (tbl.Cells[r, 0].TextString.Contains("POT-HOLE EXISTING UTILITY"))
                        {
                            tbl.Cells[r, 1].TextString = totalCount.ToString();
                            tbl.RecomputeTableBlock(true);
                            return;
                        }
                    }
                }
            }
        }

        private void ExportLogToFile(List<string> lines, int total)
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filePath = Path.Combine(desktopPath, "MLeader_Sync_Report.txt");

            using (StreamWriter sw = new StreamWriter(filePath))
            {
                sw.WriteLine($"FINAL TOTAL COUNT: {total}");
                foreach (string line in lines) sw.WriteLine(line);
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
            return list.OrderBy(x => x.NumericValue).ToList();
        }

        private int ExtractNumber(string val)
        {
            Match m = Regex.Match(val, @"\d+");
            return m.Success ? int.Parse(m.Value) : 0;
        }

        private IEnumerable<Layout> GetSortedLayouts(Transaction tr, Database db)
        {
            DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
            return layoutDict.Cast<DBDictionaryEntry>()
                .Select(entry => tr.GetObject(entry.Value, OpenMode.ForRead) as Layout)
                .Where(lay => !lay.ModelType)
                .OrderBy(lay => lay.TabOrder);
        }
    }    
}
