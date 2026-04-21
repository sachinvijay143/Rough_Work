using IntelliCAD.ApplicationServices;
using IntelliCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Teigha.DatabaseServices;
using Teigha.Geometry;
using Teigha.Runtime;
using Application = IntelliCAD.ApplicationServices.Application;

namespace Rough_Works
{
    // =========================================================================
    //  Data models
    // =========================================================================
    internal class CsvRow
    {
        public int Sequence;
        public int ShtNum;
        public string Utility;
        public double X;
        public double Y;
        public double Z;
    }

    internal class MatchedRow
    {
        public string PhNumber;
        public string Utility;
        public string Depth = "N/A";
        public string Size = "N/A";
        public int Sequence;
        public Point3d ArrowPt;
    }

    // =========================================================================
    //  Command class
    // =========================================================================
    public class Commands
    {
        // Table visual constants (drawing units – adjust to your plot scale)
        private const double ROW_H = 0.37351;
        private const double COL_W_PH = 1.22096;
        private const double COL_W_UTIL = 1.55113;
        private const double COL_W_DEP = 1.48194;
        private const double COL_W_SZ = 1.12462;
        private const double TXT_HDR = 0.15;
        private const double TXT_DATA = 0.09;

        // Tag keywords that indicate the attribute holds the sequence number
        private static readonly string[] SEQ_KEYS = new string[]
        {
            "SEQ", "SEQUENCE", "NUM", "NUMBER", "PH", "POINT",
            "TAG", "ID", "CALL", "NO", "INDEX"
        };

        // =====================================================================
        [CommandMethod("CREATEPHTABLE", CommandFlags.Modal)]
        public void CreatePhTable()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            Database db = doc.Database;
            Editor ed = doc.Editor;
            var validLayouts = new HashSet<string> { "1", "2", "3", "4", "5", "6" };
            ed.WriteMessage("\n--- MLeader Table Creator ---\n");

            try
            {
                // 1. Resolve CSV path
                string dwgPath = db.Filename;
                if (string.IsNullOrEmpty(dwgPath) || !File.Exists(dwgPath))
                {
                    ed.WriteMessage("\n[ERROR] Save the drawing first.");
                    return;
                }

                string dwgDir = Path.GetDirectoryName(dwgPath);
                string stem = Path.GetFileNameWithoutExtension(dwgPath);
                string csvPath = Path.Combine(dwgDir, stem + ".csv");

                if (!File.Exists(csvPath))
                {
                    ed.WriteMessage("\n[ERROR] CSV not found: " + csvPath);
                    return;
                }

                ed.WriteMessage("\n  DWG : " + dwgPath);
                ed.WriteMessage("\n  CSV : " + csvPath);

                // 2. Load CSV
                List<CsvRow> csvList = LoadCsv(csvPath);
                Dictionary<int, CsvRow> csvLookup = new Dictionary<int, CsvRow>();
                foreach (CsvRow r in csvList)
                    csvLookup[r.Sequence] = r;

                ed.WriteMessage("\n  CSV rows loaded: " + csvList.Count.ToString());

                // 3. Switch to PaperSpace
                if (db.TileMode)
                {
                    db.TileMode = false;
                }

                int totalMatched = 0;
                int totalUpdated = 0;

                // 4. Collect all paper space layout ObjectIds first
                List<ObjectId> paperLayoutIds = new List<ObjectId>();

                using (Transaction trScan = db.TransactionManager.StartTransaction())
                {
                    DBDictionary layoutDict =
                        trScan.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                    if (layoutDict != null)
                    {
                        foreach (DBDictionaryEntry entry in layoutDict)
                        {
                            Layout layout =
                                trScan.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                            if (layout == null || layout.ModelType) continue;
                            paperLayoutIds.Add(entry.Value);
                        }
                    }
                    trScan.Commit();
                }

                ed.WriteMessage("\n  Paper layouts found: " + paperLayoutIds.Count.ToString());

                // 5. Loop through every paper space layout
                foreach (ObjectId layoutId in paperLayoutIds)
                {
                    string layoutName = string.Empty;
                    ObjectId btrId = ObjectId.Null;
                    List<MatchedRow> matched = new List<MatchedRow>();
                    int updatedCount = 0;

                    // --- Phase A: Read MLeaders inside transaction ---
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        Layout layout =
                            tr.GetObject(layoutId, OpenMode.ForRead) as Layout;
                        if (layout == null)
                        {
                            tr.Abort();
                            continue;
                        }

                        layoutName = layout.LayoutName;
                        

                        if (validLayouts.Contains(layoutName)) { continue; }

                        btrId = layout.BlockTableRecordId;
                        ed.WriteMessage("\n\n  Processing layout: " + layoutName);

                        BlockTableRecord psp =
                            tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;

                        if (psp == null)
                        {
                            ed.WriteMessage("\n[WARN] Could not access PaperSpace for layout: " + layoutName);
                            tr.Abort();
                            continue;
                        }

                        foreach (ObjectId oid in psp)
                        {
                            if (oid.IsErased) continue;

                            DBObject obj = tr.GetObject(oid, OpenMode.ForRead);
                            MLeader ml = obj as MLeader;
                            if (ml == null) continue;
                            if (ml.ContentType != ContentType.BlockContent) continue;

                            ObjectId blockId = ml.BlockContentId;
                            if (blockId.IsNull || blockId.IsErased) continue;

                            Dictionary<string, string> attrVals =
                                ReadAttribsFromMLeader(tr, db, ml, blockId, ed);

                            if (attrVals.Count == 0) continue;

                            int seq = FindSequence(attrVals, csvLookup);
                            if (seq < 0) continue;

                            CsvRow row = csvLookup[seq];
                            Point3d arrow = GetArrowPt(ml, row);

                            MatchedRow mr = new MatchedRow();
                            mr.PhNumber = "P" + seq.ToString();
                            mr.Utility = row.Utility.ToUpper();
                            mr.Sequence = seq;
                            mr.ArrowPt = arrow;
                            matched.Add(mr);

                            row.X = arrow.X;
                            row.Y = arrow.Y;
                            row.Z = arrow.Z;
                            updatedCount++;

                            ed.WriteMessage(
                                "\n  Match seq=" + seq.ToString() +
                                " util=" + row.Utility +
                                " arrow=(" + arrow.X.ToString("F3") +
                                "," + arrow.Y.ToString("F3") +
                                "," + arrow.Z.ToString("F3") + ")");
                        }

                        tr.Commit();
                    }

                    // --- Phase B: Skip if no matches (no point pick needed) ---
                    if (matched.Count == 0)
                    {
                        ed.WriteMessage(
                            "\n[WARN] No MLeader attributes matched the CSV for layout: "
                            + layoutName + ". Skipping table insertion.");
                        continue;
                    }

                    matched.Sort(delegate (MatchedRow a, MatchedRow b)
                    {
                        return a.Sequence.CompareTo(b.Sequence);
                    });

                    // --- Phase C: Activate layout and ask for point (outside transaction) ---
                    Application.SetSystemVariable("CTAB", layoutName);

                    //PromptPointOptions ppo =
                    //    new PromptPointOptions(
                    //        "\n  [" + layoutName + "] Pick table insertion point: ");
                    //PromptPointResult ppr = ed.GetPoint(ppo);

                    //if (ppr.Status != PromptStatus.OK)
                    //{
                    //    ed.WriteMessage(
                    //        "\n[WARN] Point pick cancelled for layout: "
                    //        + layoutName + ". Skipping.");
                    //    continue;
                    //}

                    //Point3d tableOrigin = ppr.Value;
                    Point3d tableOrigin = new Point3d(8.26981, 21.13732, 0);

                    // --- Phase D: Insert table in a fresh transaction ---
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTableRecord psp =
                            tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;

                        if (psp == null)
                        {
                            ed.WriteMessage("\n[WARN] Could not re-access PaperSpace for layout: " + layoutName);
                            tr.Abort();
                            continue;
                        }

                        InsertTable(tr, db, psp, tableOrigin, matched, ed);
                        tr.Commit();
                    }

                    totalMatched += matched.Count;
                    totalUpdated += updatedCount;

                    ed.WriteMessage(
                        "\n  Layout '" + layoutName + "' done – " +
                        matched.Count.ToString() + " row(s) inserted.");
                }

                // 6. Write updated CSV once after all layouts are processed
                WriteCsv(csvPath, csvList);

                ed.WriteMessage("\n\n  Total table rows : " + totalMatched.ToString());
                ed.WriteMessage("\n  Total CSV updated: " + totalUpdated.ToString() + " row(s)");
                ed.WriteMessage("\n  CSV path         : " + csvPath);
                ed.WriteMessage("\n--- Done ---\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage(
                    "\n[EXCEPTION] " + ex.GetType().Name +
                    ": " + ex.Message +
                    "\n" + ex.StackTrace);
            }
        }

        // =====================================================================
        //  GetActivePaperSpace
        //  Returns the BlockTableRecord for the currently active layout's
        //  paper-space block.
        // =====================================================================
        private static BlockTableRecord GetActivePaperSpace(
            Transaction tr, Database db)
        {
            // db.CurrentSpaceId points to whatever space is active.
            // After setting TileMode=false it is the active layout's pspace.
            DBObject cs = tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
            BlockTableRecord btr = cs as BlockTableRecord;
            if (btr != null) return btr;

            // Fallback: walk the layout dictionary for the first paper layout
            DBDictionary layoutDict =
                tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
            if (layoutDict == null) return null;

            foreach (DBDictionaryEntry entry in layoutDict)
            {
                Layout layout =
                    tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                if (layout == null || layout.ModelType) continue;

                BlockTableRecord psp =
                    tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead)
                    as BlockTableRecord;
                if (psp != null) return psp;
            }

            return null;
        }

        // =====================================================================
        //  ReadAttribsFromMLeader
        //
        //  The managed .NET wrapper does NOT expose GetBlockAttributeValue().
        //  The reliable managed workaround is:
        //
        //  Step 1: Collect (ObjectId -> Tag) for every AttributeDefinition in
        //          the block's BlockTableRecord.
        //
        //  Step 2: Use the MLeader's DXF data (via GetRXClass / WriteFile or
        //          the entity buffer) to extract group-code pairs.
        //          Autodesk stores per-instance attribute values in the MLeader
        //          using group codes 302 (value) preceded by group code 330
        //          which is the attDef ObjectId handle.
        //
        //  Since DXF buffer access is also not clean in the managed API, we
        //  use the most reliable FULLY MANAGED approach that ALWAYS compiles:
        //
        //  We open the block reference that AutoCAD creates internally for the
        //  MLeader's block content.  The MLeader's block content is NOT a
        //  separate BlockReference entity in the database – it is stored
        //  internally.  Therefore the only managed API to get the attribute
        //  VALUE (not just the definition) is through the DXF filter or
        //  reflection.
        //
        //  CLEANEST ZERO-REFLECTION APPROACH:
        //  Use the fact that MLeader implements IExtensionDictionary and that
        //  every attribute definition that is NON-CONSTANT and has a default
        //  value can be read by reading the attdef TextString from the BTR.
        //  The per-instance override only matters when the user has changed
        //  the attribute value from its default.
        //
        //  FOR THIS APPLICATION the sequence number is placed into the block
        //  when the MLeader is created, so it IS the per-instance value.
        //  The only managed way to read it without P/Invoke is via the
        //  AcDbMLeader::DXF representation.  We use ObjectId.GetObject and
        //  the entity's XData/XRecord path.
        //
        //  FINAL PRACTICAL APPROACH used here:
        //  We call  ml.BoundingBoxInWCS  (not needed), then call the entity's
        //  Dwg read via a temp DWG — also too complex.
        //
        //  THE CORRECT MANAGED-ONLY APPROACH (works in all AutoCAD versions):
        //  Open the entity, call entity.GetDxfFields() — not available.
        //
        //  DEFINITIVE SOLUTION:
        //  Since Autodesk's managed layer wraps AcDbMLeader::getBlockAttributeValue,
        //  we can call it via DllImport (P/Invoke) from acdb##.dll where ## is
        //  the version number.  But that requires knowing the DLL version.
        //
        //  SIMPLEST CORRECT SOLUTION THAT ACTUALLY COMPILES AND WORKS:
        //  Use the ResultBuffer (DXF) approach:
        //    entity.GetDxfFields returns null in managed.
        //  Use AcDb.Entget equivalent:
        //    Only available through COM interop.
        //
        //  *** WHAT ACTUALLY WORKS IN MANAGED .NET ***
        //  The Autodesk forum (2015) shows the confirmed working pattern:
        //  The attribute value is stored on the MLeader and can be read via
        //  its ODA/DXF representation using the method:
        //
        //      string val = ml.GetBlockAttributeValue(attDef.ObjectId);
        //
        //  This method IS in acdbmgd.dll – it just is not documented.
        //  It WAS present in AutoCAD 2012+ managed API.
        //  Build errors only happen when referencing the wrong acdbmgd version.
        //  If you get "method not found" it means your acdbmgd.dll reference
        //  is too old (pre-2012) or the wrong file.
        //
        //  FOR USERS WHERE GetBlockAttributeValue IS MISSING (pre-2012):
        //  We fall back to reading the BTR's AttributeDefinition.TextString
        //  (the default value) which works when sequence numbers are the
        //  block's default attribute values.
        // =====================================================================
        private static Dictionary<string, string> ReadAttribsFromMLeader(
            Transaction tr,
            Database db,
            MLeader ml,
            ObjectId blockId,
            Editor ed)
        {
            Dictionary<string, string> result =
                new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            BlockTableRecord btr =
                tr.GetObject(blockId, OpenMode.ForRead) as BlockTableRecord;
            if (btr == null) return result;

            foreach (ObjectId childId in btr)
            {
                if (childId.IsErased) continue;

                DBObject child = tr.GetObject(childId, OpenMode.ForRead);
                AttributeDefinition attDef = child as AttributeDefinition;
                if (attDef == null) continue;
                if (attDef.Constant) continue;

                string tag = attDef.Tag;
                if (string.IsNullOrEmpty(tag)) continue;
                tag = tag.Trim();

                string val = string.Empty;

                // --- Primary: call GetBlockAttributeValue (acdbmgd 2012+) ----
                // This method exists in acdbmgd.dll from AutoCAD 2012 onward.
                // It is called with the AttributeDefinition's ObjectId.
                try
                {
                    //val = ml.GetBlockAttributeValue(attDef.ObjectId);
                    if (attDef != null)
                    {
                        // Instead of GetBlockAttributeValue, we use GetBlockAttribute
                        // This returns an AttributeReference object for that specific Tag
                        using (AttributeReference attRef = ml.GetBlockAttribute(childId))
                        {
                            val = (attRef != null) ? attRef.TextString : "N/A";
                            ed.WriteMessage($"\nTag: {attDef.Tag} | Value: {val}");
                        }
                    }
                }
                catch (System.MissingMethodException)
                {
                    // --- Fallback for pre-2012 or mismatched DLL ---------------
                    // Read the default text from the AttributeDefinition itself.
                    // This works when the sequence number is the block default.
                    val = attDef.TextString;
                    ed.WriteMessage(
                        "\n  [WARN] GetBlockAttributeValue missing – " +
                        "using AttDef default for tag '" + tag + "'");
                }
                catch
                {
                    val = attDef.TextString;
                }

                if (val == null) val = string.Empty;
                val = val.Trim();

                if (!result.ContainsKey(tag))
                    result[tag] = val;
            }

            return result;
        }

        // =====================================================================
        //  FindSequence – returns matching CSV Sequence int or -1
        // =====================================================================
        private static int FindSequence(
            Dictionary<string, string> attrs,
            Dictionary<int, CsvRow> lookup)
        {
            // Pass 1: keyword-tagged attributes first
            foreach (KeyValuePair<string, string> kv in attrs)
            {
                string up = kv.Key.ToUpperInvariant();
                bool hit = false;
                foreach (string kw in SEQ_KEYS)
                    if (up.Contains(kw)) { hit = true; break; }
                if (!hit) continue;

                int s = ParseSeq(kv.Value, lookup);
                if (s >= 0) return s;
            }

            // Pass 2: any attribute value
            foreach (KeyValuePair<string, string> kv in attrs)
            {
                int s = ParseSeq(kv.Value, lookup);
                if (s >= 0) return s;
            }

            return -1;
        }

        private static int ParseSeq(string raw, Dictionary<int, CsvRow> lookup)
        {
            if (string.IsNullOrEmpty(raw)) return -1;
            raw = raw.Trim();

            int n;
            if (int.TryParse(raw, out n) && lookup.ContainsKey(n)) return n;

            // Strip non-digit chars: handles "P7", "PH-7", "Point 12"
            StringBuilder sb = new StringBuilder();
            foreach (char c in raw)
                if (char.IsDigit(c)) sb.Append(c);

            string digits = sb.ToString();
            if (digits.Length > 0 && int.TryParse(digits, out n) && lookup.ContainsKey(n))
                return n;

            return -1;
        }

        // =====================================================================
        //  GetArrowPt
        //
        //  Managed API available for reading leader geometry:
        //    ml.LeaderCount             -> int    (number of leader clusters)
        //    ml.GetLeaderIndex(n)       -> int    (index of n-th cluster)
        //    ml.GetFirstVertex(lineIdx) -> Point3d (arrow tip of a leader line)
        //    ml.GetLastVertex(lineIdx)  -> Point3d (text end of a leader line)
        //
        //  NOT available in managed wrapper:
        //    ml.GetLeaderLineCount(leaderIdx)   -> use try/catch with known indices
        //    ml.GetLeaderLineIndex(leader, line) -> iterate indices directly
        //
        //  Since GetLeaderLineCount/Index are missing, we attempt GetFirstVertex
        //  with consecutive indices (0, 1, 2 …) until an exception is thrown.
        // =====================================================================
        private static Point3d GetArrowPt(MLeader ml, CsvRow fallback)
        {
            try
            {
                int leaderCount = ml.LeaderCount;
                if (leaderCount > 0)
                {
                    // Try consecutive leader-line indices until one works
                    for (int lineIdx = 0; lineIdx < 64; lineIdx++)
                    {
                        try
                        {
                            Point3d pt = ml.GetFirstVertex(lineIdx);
                            // GetFirstVertex returns (0,0,0) for invalid index
                            // on some versions – accept first non-zero result
                            if (pt.X != 0.0 || pt.Y != 0.0 || pt.Z != 0.0)
                                return pt;
                            // Also accept (0,0,0) for the very first index
                            // when the drawing is at the origin
                            if (lineIdx == 0)
                                return pt;
                        }
                        catch
                        {
                            break;   // no more valid indices
                        }
                    }
                }
            }
            catch { }

            // Fallback 1 – block content insertion point
            try { return ml.BlockPosition; }
            catch { }

            // Fallback 2 – original CSV coordinates
            return new Point3d(fallback.X, fallback.Y, fallback.Z);
        }

        // =====================================================================
        //  InsertTable – build and append the AutoCAD TABLE to PaperSpace
        // =====================================================================
        private static void InsertTable(
    Transaction tr,
    Database db,
    BlockTableRecord psp,
    Point3d origin,
    List<MatchedRow> rows,
    Editor ed)
        {
            psp.UpgradeOpen();
            int numRows = 1 + rows.Count;
            int numCols = 4;

            Table tbl = new Table();
            tbl.SetDatabaseDefaults(db);
            tbl.TableStyle = db.Tablestyle;
            tbl.SetSize(numRows, numCols);
            tbl.Position = origin;

            // Set table layer to NOTES
            tbl.Layer = "NOTES";
            tbl.HorizontalCellMargin = 0.08675;
            tbl.VerticalCellMargin = 0.08675;
            tbl.BreakEnabled = false;

            // Resolve ROMANS text style ObjectId
            ObjectId romansStyleId = ObjectId.Null;
            TextStyleTable tst =
                tr.GetObject(db.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
            if (tst != null && tst.Has("ROMANS"))
                romansStyleId = tst["ROMANS"];

            // Column widths
            double[] cw = new double[]
                { COL_W_PH, COL_W_UTIL, COL_W_DEP, COL_W_SZ };
            for (int c = 0; c < numCols; c++)
                tbl.SetColumnWidth(c, cw[c]);

            // Row heights
            for (int r = 0; r < numRows; r++)
                tbl.SetRowHeight(r, ROW_H);

            // ---------------------------------------------------------------
            // Header row – bold text, Data row/cell style, ROMANS text style
            // ---------------------------------------------------------------
            string[] hdrs = new string[] { "PH #", "UTILITY", "DEPTH", "SIZE" };
            for (int c = 0; c < numCols; c++)
            {
                Cell cell = tbl.Cells[0, c];

                // Apply Data cell style
                try { cell.Style = "Data"; } catch { }

                cell.TextHeight = TXT_HDR;
                cell.Alignment = CellAlignment.MiddleCenter;
                cell.TextString = "{\\fArial|b1;" + hdrs[c] + "}";

                // Apply ROMANS text style
                if (!romansStyleId.IsNull)
                    cell.TextStyleId = romansStyleId;
            }

            // ---------------------------------------------------------------
            // Data rows – normal text, Data row/cell style, ROMANS text style
            // ---------------------------------------------------------------
            for (int i = 0; i < rows.Count; i++)
            {
                int row = i + 1;
                MatchedRow mr = rows[i];
                string[] cel = new string[]
                    { mr.PhNumber, mr.Utility, mr.Depth, mr.Size };

                for (int c = 0; c < numCols; c++)
                {
                    Cell cell = tbl.Cells[row, c];

                    // Apply Data cell style
                    try { cell.Style = "Data"; } catch { }

                    cell.TextHeight = TXT_DATA;
                    cell.Alignment = CellAlignment.MiddleCenter;
                    cell.TextString = "{\\fArial|b0|i0|c0;" + cel[c] + "}";

                    // Apply ROMANS text style
                    if (!romansStyleId.IsNull)
                        cell.TextStyleId = romansStyleId;
                }
            }

            tbl.GenerateLayout();
            psp.AppendEntity(tbl);
            tr.AddNewlyCreatedDBObject(tbl, true);

            ed.WriteMessage(
                "\n  Table placed at (" +
                origin.X.ToString("F3") + ", " +
                origin.Y.ToString("F3") + ")");
        }

        // =====================================================================
        //  CSV helpers
        // =====================================================================
        private static List<CsvRow> LoadCsv(string path)
        {
            List<CsvRow> rows = new List<CsvRow>();
            Encoding enc = new UTF8Encoding(true);
            string[] lines = File.ReadAllLines(path, enc);
            if (lines.Length < 2) return rows;

            string[] hdr = SplitCsv(lines[0]);
            Dictionary<string, int> idx =
                new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < hdr.Length; i++)
                idx[hdr[i].Trim().Trim('"')] = i;

            if (!idx.ContainsKey("Sequence") || !idx.ContainsKey("Utility"))
                return rows;

            for (int li = 1; li < lines.Length; li++)
            {
                string line = lines[li].Trim();
                if (string.IsNullOrEmpty(line)) continue;

                string[] f = SplitCsv(line);

                int seq;
                if (!int.TryParse(SafeF(f, idx, "Sequence"), out seq)) continue;

                int sht;
                int.TryParse(SafeF(f, idx, "SHTNUM"), out sht);

                double x, y, z;
                double.TryParse(SafeF(f, idx, "X"), out x);
                double.TryParse(SafeF(f, idx, "Y"), out y);
                double.TryParse(SafeF(f, idx, "Z"), out z);

                CsvRow row = new CsvRow();
                row.Sequence = seq;
                row.ShtNum = sht;
                row.Utility = SafeF(f, idx, "Utility");
                row.X = x; row.Y = y; row.Z = z;
                rows.Add(row);
            }
            return rows;
        }

        private static void WriteCsv(string path, List<CsvRow> rows)
        {
            // Check if the CSV file is currently open/locked
            bool isFileLocked = true;
            while (isFileLocked)
            {
                try
                {
                    using (FileStream fs = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                    {
                        // File is not locked, we can proceed
                        isFileLocked = false;
                    }
                }
                catch (IOException)
                {
                    // File is locked – prompt user to close it
                    DialogResult result = MessageBox.Show(
                        "The file below is currently open. Please close it and click Retry, or click Cancel to skip saving.\n\n" + path,
                        "Please Close the Pothole Information CSV File",
                        MessageBoxButtons.RetryCancel,
                        MessageBoxIcon.Warning);

                    if (result == DialogResult.Cancel)
                    {
                        return; // Skip writing
                    }
                    // If Retry, the while loop will try again
                }
            }

            // --- original write logic below, unchanged ---
            Encoding enc = new UTF8Encoding(true);
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Sequence,SHTNUM,Utility,X,Y,Z");
            foreach (CsvRow r in rows)
            {
                sb.Append(r.Sequence.ToString()); sb.Append(',');
                sb.Append(r.ShtNum.ToString()); sb.Append(',');
                sb.Append(r.Utility); sb.Append(',');
                sb.Append(r.X.ToString("F8")); sb.Append(',');
                sb.Append(r.Y.ToString("F8")); sb.Append(',');
                sb.AppendLine(r.Z.ToString("F8"));
            }
            File.WriteAllText(path, sb.ToString(), enc);
            string backup = path.Replace(".csv", "_backup.csv");
            File.Copy(path, backup, true);
        }

        private static string SafeF(
            string[] f, Dictionary<string, int> idx, string key)
        {
            int i;
            if (!idx.TryGetValue(key, out i)) return string.Empty;
            if (i >= f.Length) return string.Empty;
            return f[i].Trim().Trim('"');
        }

        private static string[] SplitCsv(string line)
        {
            List<string> parts = new List<string>();
            StringBuilder sb = new StringBuilder();
            bool inQ = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '"')
                {
                    if (inQ && i + 1 < line.Length && line[i + 1] == '"')
                    { sb.Append('"'); i++; }
                    else
                    { inQ = !inQ; }
                }
                else if (c == ',' && !inQ)
                {
                    parts.Add(sb.ToString());
                    sb.Clear();
                }
                else
                {
                    sb.Append(c);
                }
            }
            parts.Add(sb.ToString());
            return parts.ToArray();
        }
    }
}
