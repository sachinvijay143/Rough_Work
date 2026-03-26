using IntelliCAD.ApplicationServices;
using IntelliCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Teigha.DatabaseServices;
using Teigha.Geometry;
using Teigha.Runtime;
using Exception = System.Exception;

namespace Rough_Works
{
    // Simple data holder for each MText found
    public class MTextInfo
    {
        public string RawText { get; set; }
        public string CleanText { get; set; }
        public Point3d Location { get; set; }
        public int Level { get; set; }   // which explode depth it came from
        public string SourceType { get; set; }   // MText or DBText
    }

    public class StationLabelCommands
    {
        [CommandMethod("READALLSTATIONMTEXT")]
        public void ReadAllStationMText()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
                // ─────────────────────────────────────────
                // STEP 1: Select the AlignmentStationLabeling object
                // ─────────────────────────────────────────
                PromptEntityOptions peo = new PromptEntityOptions(
                    "\nSelect AlignmentStationLabeling object: ");
                peo.AllowNone = false;

                PromptEntityResult per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nSelection cancelled.");
                    return;
                }

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    DBObject selectedObj = tr.GetObject(per.ObjectId, OpenMode.ForRead);
                    ed.WriteMessage($"\nSelected Type: {selectedObj.GetType().Name}");

                    // ─────────────────────────────────────────
                    // STEP 2: Collect ALL MText with locations
                    // ─────────────────────────────────────────
                    List<MTextInfo> allTexts = new List<MTextInfo>();
                    CollectAllMTexts(selectedObj as Entity, allTexts, ed, level: 1);

                    // ─────────────────────────────────────────
                    // STEP 3: Display results
                    // ─────────────────────────────────────────
                    if (allTexts.Count == 0)
                    {
                        ed.WriteMessage("\n✘ No MText found in the selected label.");
                        return;
                    }

                    ed.WriteMessage($"\n\n{'─',60}");
                    ed.WriteMessage($"\n Total MText objects found: {allTexts.Count}");
                    ed.WriteMessage($"\n{'─',60}");

                    System.Text.StringBuilder sb = new System.Text.StringBuilder();
                    sb.AppendLine($"Station Label — All MText Values ({allTexts.Count} found)");
                    sb.AppendLine(new string('─', 50));

                    for (int i = 0; i < allTexts.Count; i++)
                    {
                        MTextInfo info = allTexts[i];

                        // Command line output
                        ed.WriteMessage($"\n[{i + 1}] Type     : {info.SourceType}");
                        ed.WriteMessage($"\n    Level    : {info.Level}");
                        ed.WriteMessage($"\n    Location : X={info.Location.X:F4}  Y={info.Location.Y:F4}  Z={info.Location.Z:F4}");
                        ed.WriteMessage($"\n    Raw      : {info.RawText}");
                        ed.WriteMessage($"\n    Clean    : {info.CleanText}");
                        ed.WriteMessage($"\n{'─',60}");

                        // Dialog output
                        sb.AppendLine($"\n[{i + 1}] {info.SourceType} (Level {info.Level})");
                        sb.AppendLine($"  Location : X={info.Location.X:F4}  Y={info.Location.Y:F4}  Z={info.Location.Z:F4}");
                        sb.AppendLine($"  Raw Text : {info.RawText}");
                        sb.AppendLine($"  Clean    : {info.CleanText}");
                        sb.AppendLine(new string('─', 50));
                    }

                    Application.ShowAlertDialog(sb.ToString());
                    tr.Commit();
                }
            }
            catch (Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}\n{ex.StackTrace}");
            }
        }


        /// <summary>
        /// Recursively explodes any entity and collects ALL MText/DBText
        /// objects found at any depth, along with their locations.
        /// </summary>
        private void CollectAllMTexts(
            Entity entity,
            List<MTextInfo> results,
            Editor ed,
            int level)
        {
            if (entity == null) return;

            // Safety: limit recursion depth to avoid infinite loops
            if (level > 5) return;

            DBObjectCollection exploded = new DBObjectCollection();

            try
            {
                entity.Explode(exploded);
                ed.WriteMessage(
                    $"\n[Level {level}] Exploded '{entity.GetType().Name}'" +
                    $" → {exploded.Count} object(s)");
            }
            catch (Exception ex)
            {
                ed.WriteMessage(
                    $"\n[Level {level}] Explode failed on '{entity.GetType().Name}'" +
                    $": {ex.Message}");
                return;
            }

            foreach (DBObject obj in exploded)
            {
                try
                {
                    // ── MText found ──────────────────────────────────────────
                    if (obj is MText mText)
                    {
                        string raw = mText.Contents;
                        string clean = StripMTextFormatting(raw);

                        // Capture even empty strings so caller sees everything
                        results.Add(new MTextInfo
                        {
                            RawText = raw,
                            CleanText = clean,
                            Location = mText.Location,   // insertion point of MText
                            Level = level,
                            SourceType = "MText"
                        });

                        ed.WriteMessage(
                            $"\n  ✔ MText @ X={mText.Location.X:F2}" +
                            $" Y={mText.Location.Y:F2} → \"{clean}\"");
                    }

                    // ── DBText found ─────────────────────────────────────────
                    else if (obj is DBText dbText)
                    {
                        string raw = dbText.TextString;
                        string clean = raw?.Trim() ?? string.Empty;

                        results.Add(new MTextInfo
                        {
                            RawText = raw,
                            CleanText = clean,
                            Location = dbText.Position,  // insertion point of DBText
                            Level = level,
                            SourceType = "DBText"
                        });

                        ed.WriteMessage(
                            $"\n  ✔ DBText @ X={dbText.Position.X:F2}" +
                            $" Y={dbText.Position.Y:F2} → \"{clean}\"");
                    }

                    // ── BlockReference → recurse one level deeper ────────────
                    else if (obj is BlockReference blkRef)
                    {
                        ed.WriteMessage(
                            $"\n  → BlockRef '{blkRef.Name}' found, recursing...");
                        CollectAllMTexts(blkRef, results, ed, level + 1);
                    }

                    // ── AttributeReference (block attribute text) ────────────
                    else if (obj is AttributeReference attRef)
                    {
                        string raw = attRef.TextString;
                        string clean = raw?.Trim() ?? string.Empty;

                        if (!string.IsNullOrWhiteSpace(clean))
                        {
                            results.Add(new MTextInfo
                            {
                                RawText = raw,
                                CleanText = clean,
                                Location = attRef.Position,
                                Level = level,
                                SourceType = "AttributeRef"
                            });

                            ed.WriteMessage(
                                $"\n  ✔ AttribRef @ X={attRef.Position.X:F2}" +
                                $" Y={attRef.Position.Y:F2} → \"{clean}\"");
                        }
                    }
                }
                catch (Exception ex)
                {
                    ed.WriteMessage(
                        $"\n  ✘ Error reading object '{obj?.GetType().Name}': {ex.Message}");
                }
                finally
                {
                    // Always dispose in-memory exploded objects
                    try { if (!obj.IsDisposed) obj.Dispose(); } catch { }
                }
            }
        }


        /// <summary>
        /// Strips MText formatting codes to return plain readable text.
        /// e.g.  "\H2.5;\C1;11+98 W"  →  "11+98 W"
        /// </summary>
        private string StripMTextFormatting(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return string.Empty;

            string result = System.Text.RegularExpressions.Regex.Replace(
                raw,
                @"\\[A-Za-z][^;]*;|[{}]|\\[pPqQlLrRoO]",
                string.Empty);

            result = result.Replace(@"\P", "\n").Replace(@"\n", "\n");
            return result.Trim();
        }
    }
}
