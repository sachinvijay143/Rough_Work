using IntelliCAD.ApplicationServices;
using IntelliCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Teigha.DatabaseServices;
using Teigha.Geometry;
using Teigha.Runtime;
using Application = IntelliCAD.ApplicationServices.Application;
using Exception = System.Exception;

namespace Rough_Works
{
    public class MTextInfo
    {
        public string RawText { get; set; }
        public string CleanText { get; set; }
        public string Value { get; set; }
        public string DirectionSuffix { get; set; }
        public int WholeNumber { get; set; }
        public Point3d Location { get; set; }
        public string SourceType { get; set; }
        public double DistanceToBlock { get; set; }
    }

    public class LabelCandidate
    {
        public Entity Entity { get; set; }
        public ObjectId ObjectId { get; set; }
        public double Distance { get; set; }
        public List<MTextInfo> Texts { get; set; } = new List<MTextInfo>();
    }

    public class StationLabelCommands
    {
        // ───────────────────────────────────────────────────────────────
        // COMMAND 1: READALLSTATIONMTEXT
        // ───────────────────────────────────────────────────────────────
        [CommandMethod("READALLSTATIONMTEXT")]
        public void ReadAllStationMText()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
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
                    DBObject selectedObj = tr.GetObject(
                        per.ObjectId, OpenMode.ForRead);

                    ed.WriteMessage(
                        $"\nSelected Type: {selectedObj.GetType().Name}");

                    List<MTextInfo> allTexts = new List<MTextInfo>();
                    CollectAllMTexts(
                        selectedObj as Entity, allTexts, ed,
                        level: 1, blockPos: null);

                    if (allTexts.Count == 0)
                    {
                        ed.WriteMessage("\n✘ No MText found.");
                        tr.Commit();
                        return;
                    }

                    ed.WriteMessage($"\n\n{'─',60}");
                    ed.WriteMessage($"\n  Total MText found: {allTexts.Count}");
                    ed.WriteMessage($"\n{'─',60}");

                    System.Text.StringBuilder sb = new System.Text.StringBuilder();
                    sb.AppendLine(
                        $"Station Label — All MText ({allTexts.Count} found)");
                    sb.AppendLine(new string('─', 50));

                    for (int i = 0; i < allTexts.Count; i++)
                    {
                        MTextInfo info = allTexts[i];

                        ed.WriteMessage($"\n[{i + 1}] Type       : {info.SourceType}");
                        ed.WriteMessage(
                            $"\n    Location   : X={info.Location.X:F4}" +
                            $"  Y={info.Location.Y:F4}");
                        ed.WriteMessage($"\n    Raw        : {info.RawText}");
                        ed.WriteMessage($"\n    Clean      : {info.CleanText}");
                        ed.WriteMessage($"\n    Value      : {info.Value}");
                        ed.WriteMessage($"\n    WholeNumber: {info.WholeNumber}");
                        ed.WriteMessage($"\n    Suffix     : {info.DirectionSuffix}");
                        ed.WriteMessage($"\n{'─',60}");

                        sb.AppendLine($"\n[{i + 1}] {info.SourceType}");
                        sb.AppendLine(
                            $"  Location   : X={info.Location.X:F4}" +
                            $"  Y={info.Location.Y:F4}");
                        sb.AppendLine($"  Raw        : {info.RawText}");
                        sb.AppendLine($"  Clean      : {info.CleanText}");
                        sb.AppendLine($"  Value      : {info.Value}");
                        sb.AppendLine($"  WholeNumber: {info.WholeNumber}");
                        sb.AppendLine($"  Suffix     : {info.DirectionSuffix}");
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


        // ───────────────────────────────────────────────────────────────
        // COMMAND 2: BLOCKNEARESTSTATION
        // Step 1 → Select Block
        // Step 2 → Select the EXACT AlignmentStationLabeling
        //
        // TOP 2 LOGIC — split by block insertion point:
        //   The label lies along an alignment line.
        //   The block insertion point sits ON that line.
        //   MTexts appear on BOTH sides of the block along the alignment.
        //   So we pick:
        //     - 1 MText from the side where MText.X >= blockPos.X  (right/above)
        //     - 1 MText from the side where MText.X <  blockPos.X  (left/below)
        //   If label is more vertical, split by Y instead.
        //   Within each side pick the nearest to block.
        //   This guarantees 1 from each direction along the alignment.
        // ───────────────────────────────────────────────────────────────
        [CommandMethod("BLOCKNEARESTSTATION")]
        public void BlockNearestStation()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
                // STEP 1: Select the Block
                ed.WriteMessage("\n[Step 1] Select the Block object: ");

                PromptEntityOptions peoBlock = new PromptEntityOptions(
                    "\nSelect a Block object: ");
                peoBlock.SetRejectMessage("\nPlease select a Block.");
                peoBlock.AddAllowedClass(typeof(BlockReference), exactMatch: false);
                peoBlock.AllowNone = false;

                PromptEntityResult perBlock = ed.GetEntity(peoBlock);
                if (perBlock.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nCancelled.");
                    return;
                }

                // STEP 2: Select the EXACT AlignmentStationLabeling
                ed.WriteMessage(
                    "\n[Step 2] Select the AlignmentStationLabeling to read: ");

                PromptEntityOptions peoLabel = new PromptEntityOptions(
                    "\nSelect AlignmentStationLabeling: ");
                peoLabel.AllowNone = false;

                PromptEntityResult perLabel = ed.GetEntity(peoLabel);
                if (perLabel.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nCancelled.");
                    return;
                }

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // STEP 3: Get block position
                    BlockReference selectedBlock = tr.GetObject(
                        perBlock.ObjectId, OpenMode.ForRead) as BlockReference;

                    if (selectedBlock == null)
                    {
                        ed.WriteMessage(
                            "\n✘ Selected object is not a BlockReference.");
                        return;
                    }

                    Point3d blockPos = selectedBlock.Position;
                    ed.WriteMessage($"\nBlock Name     : {selectedBlock.Name}");
                    ed.WriteMessage(
                        $"\nBlock Position : X={blockPos.X:F4}" +
                        $"  Y={blockPos.Y:F4}");

                    // STEP 4: Open the exactly selected label
                    DBObject labelObj = tr.GetObject(
                        perLabel.ObjectId, OpenMode.ForRead);
                    Entity labelEnt = labelObj as Entity;
                    string labelType = labelObj.GetType().Name;

                    ed.WriteMessage($"\nLabel Type     : {labelType}");

                    if (labelEnt == null)
                    {
                        ed.WriteMessage(
                            "\n✘ Selected label is not a valid Entity.");
                        return;
                    }

                    // STEP 5: Collect ALL MText from the selected label
                    List<MTextInfo> allTexts = new List<MTextInfo>();
                    CollectAllMTexts(
                        labelEnt, allTexts, ed,
                        level: 1, blockPos: blockPos);

                    if (allTexts.Count == 0)
                    {
                        ed.WriteMessage(
                            "\n✘ No MText found in the selected label.");
                        tr.Commit();
                        return;
                    }

                    ed.WriteMessage(
                        $"\n  Total MText collected: {allTexts.Count}");

                    // ── DEBUG: show all collected ───────────────────────
                    ed.WriteMessage($"\n\n{'─',65}");
                    ed.WriteMessage("  [DEBUG] All collected MTexts:");
                    foreach (MTextInfo t in allTexts)
                        ed.WriteMessage(
                            $"\n    Clean=\"{t.CleanText}\"" +
                            $"  X={t.Location.X:F2}" +
                            $"  Y={t.Location.Y:F2}" +
                            $"  WholeNo={t.WholeNumber}" +
                            $"  Dist={t.DistanceToBlock:F2}");
                    ed.WriteMessage($"\n{'─',65}");
                    // ────────────────────────────────────────────────────

                    // STEP 6: Split MTexts by block insertion point
                    //
                    // Determine alignment direction:
                    //   Compare X-spread vs Y-spread of all MText locations.
                    //   More X spread → alignment is HORIZONTAL → split by X
                    //   More Y spread → alignment is VERTICAL   → split by Y
                    //
                    // Then split using the BLOCK INSERTION POINT as divider:
                    //   Horizontal: sideA = MText.X >= blockPos.X  (right of block)
                    //               sideB = MText.X <  blockPos.X  (left  of block)
                    //   Vertical:   sideA = MText.Y >= blockPos.Y  (above block)
                    //               sideB = MText.Y <  blockPos.Y  (below block)
                    //
                    // Within each side pick the 1 nearest to block.

                    double xSpread = allTexts.Max(t => t.Location.X)
                                   - allTexts.Min(t => t.Location.X);
                    double ySpread = allTexts.Max(t => t.Location.Y)
                                   - allTexts.Min(t => t.Location.Y);

                    bool isHorizontal = xSpread >= ySpread;

                    ed.WriteMessage(
                        $"\n  X-Spread={xSpread:F4}  Y-Spread={ySpread:F4}" +
                        $"  Split by={(isHorizontal ? "X (HORIZONTAL)" : "Y (VERTICAL)")}");
                    ed.WriteMessage(
                        $"\n  Block split point:" +
                        $"  {(isHorizontal ? $"X={blockPos.X:F4}" : $"Y={blockPos.Y:F4}")}");

                    MTextInfo sideA = null;  // right of block (X>=) or above (Y>=)
                    MTextInfo sideB = null;  // left  of block (X<)  or below (Y<)

                    if (isHorizontal)
                    {
                        // Split by block X position
                        sideA = allTexts
                            .Where(t => t.Location.X >= blockPos.X)
                            .OrderBy(t => t.DistanceToBlock)
                            .FirstOrDefault();

                        sideB = allTexts
                            .Where(t => t.Location.X < blockPos.X)
                            .OrderBy(t => t.DistanceToBlock)
                            .FirstOrDefault();
                    }
                    else
                    {
                        // Split by block Y position
                        sideA = allTexts
                            .Where(t => t.Location.Y >= blockPos.Y)
                            .OrderBy(t => t.DistanceToBlock)
                            .FirstOrDefault();

                        sideB = allTexts
                            .Where(t => t.Location.Y < blockPos.Y)
                            .OrderBy(t => t.DistanceToBlock)
                            .FirstOrDefault();
                    }

                    ed.WriteMessage(
                        $"\n  SideA (right/above block): " +
                        $"\"{(sideA != null ? sideA.CleanText : "none")}\"");
                    ed.WriteMessage(
                        $"\n  SideB (left/below  block): " +
                        $"\"{(sideB != null ? sideB.CleanText : "none")}\"");

                    // Build top2 — one from each side
                    List<MTextInfo> top2 = new List<MTextInfo>();
                    if (sideA != null) top2.Add(sideA);
                    if (sideB != null) top2.Add(sideB);

                    // Fallback: if all MTexts are on the same side of the block
                    // (block is at the very end of the alignment)
                    // fall back to 2 nearest by distance
                    if (top2.Count < 2)
                    {
                        ed.WriteMessage(
                            "\n  [INFO] All MTexts on same side of block —" +
                            " falling back to 2 nearest by distance.");
                        top2 = allTexts
                            .OrderBy(t => t.DistanceToBlock)
                            .Take(2)
                            .ToList();
                    }

                    ed.WriteMessage($"\n\n  Top 2 (one from each side of block):");
                    for (int i = 0; i < top2.Count; i++)
                        ed.WriteMessage(
                            $"\n    [{i + 1}] \"{top2[i].CleanText}\"" +
                            $"  WholeNumber={top2[i].WholeNumber}" +
                            $"  Dist={top2[i].DistanceToBlock:F4}");

                    // STEP 7: From top2 pick lowest whole number
                    MTextInfo lowestResult = top2
                        .Where(t => t.WholeNumber >= 0)
                        .OrderBy(t => t.WholeNumber)
                        .FirstOrDefault() ?? top2[0];

                    // STEP 8: Display result
                    ed.WriteMessage($"\n\n{'─',65}");
                    ed.WriteMessage($"\n  Block          : {selectedBlock.Name}");
                    ed.WriteMessage(
                        $"\n  Block Position : X={blockPos.X:F4}" +
                        $"  Y={blockPos.Y:F4}");
                    ed.WriteMessage($"\n  Label Type     : {labelType}");
                    ed.WriteMessage($"\n{'─',65}");
                    ed.WriteMessage(
                        "\n  Top 2 MText (one from each side of block):");

                    for (int i = 0; i < top2.Count; i++)
                    {
                        MTextInfo t = top2[i];
                        string marker = (t == lowestResult)
                            ? "  ← LOWEST WHOLE NUMBER"
                            : "";
                        ed.WriteMessage(
                            $"\n  [{i + 1}] Clean=\"{t.CleanText}\"" +
                            $"  Value=\"{t.Value}\"" +
                            $"  WholeNo={t.WholeNumber}" +
                            $"  Suffix=\"{t.DirectionSuffix}\"" +
                            $"  Dist={t.DistanceToBlock:F4}{marker}");
                    }

                    ed.WriteMessage($"\n{'─',65}");
                    ed.WriteMessage($"\n  ✔ FINAL RESULT (lowest whole number)");
                    ed.WriteMessage($"\n  Clean Text      : {lowestResult.CleanText}");
                    ed.WriteMessage($"\n  Value           : {lowestResult.Value}");
                    ed.WriteMessage($"\n  Whole Number    : {lowestResult.WholeNumber}");
                    ed.WriteMessage(
                        $"\n  Direction/Suffix: {lowestResult.DirectionSuffix}");
                    ed.WriteMessage(
                        $"\n  Location        : X={lowestResult.Location.X:F4}" +
                        $"  Y={lowestResult.Location.Y:F4}");
                    ed.WriteMessage($"\n{'─',65}");

                    System.Text.StringBuilder sb = new System.Text.StringBuilder();
                    sb.AppendLine("Station Label — Result");
                    sb.AppendLine(new string('─', 48));
                    sb.AppendLine($"Block          : {selectedBlock.Name}");
                    sb.AppendLine(
                        $"Block Position : X={blockPos.X:F4}" +
                        $"  Y={blockPos.Y:F4}");
                    sb.AppendLine($"Label Type     : {labelType}");
                    sb.AppendLine(new string('─', 48));
                    sb.AppendLine(
                        "Top 2 MText (one from each side of block):");

                    for (int i = 0; i < top2.Count; i++)
                    {
                        MTextInfo t = top2[i];
                        string marker = (t == lowestResult)
                            ? "  ← SELECTED (lowest whole no.)"
                            : "";
                        sb.AppendLine(
                            $"  [{i + 1}] Clean       : {t.CleanText}{marker}");
                        sb.AppendLine($"       Value       : {t.Value}");
                        sb.AppendLine($"       Whole Number: {t.WholeNumber}");
                        sb.AppendLine($"       Suffix      : {t.DirectionSuffix}");
                        sb.AppendLine(
                            $"       Distance    : {t.DistanceToBlock:F4}");
                    }

                    sb.AppendLine(new string('─', 48));
                    sb.AppendLine("Final Result (lowest whole number):");
                    sb.AppendLine($"  Clean Text      : {lowestResult.CleanText}");
                    sb.AppendLine($"  Value           : {lowestResult.Value}");
                    sb.AppendLine($"  Whole Number    : {lowestResult.WholeNumber}");
                    sb.AppendLine($"  Direction/Suffix: {lowestResult.DirectionSuffix}");
                    sb.AppendLine(
                        $"  Location        : X={lowestResult.Location.X:F4}" +
                        $"  Y={lowestResult.Location.Y:F4}");

                    Application.ShowAlertDialog(sb.ToString());
                    tr.Commit();
                }
            }
            catch (Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}\n{ex.StackTrace}");
            }
        }


        // ═══════════════════════════════════════════════════════════════
        // SHARED HELPER METHODS
        // ═══════════════════════════════════════════════════════════════

        private void CollectAllMTexts(
            Entity entity,
            List<MTextInfo> results,
            Editor ed,
            int level,
            Point3d? blockPos)
        {
            if (entity == null || level > 5) return;

            DBObjectCollection exploded = new DBObjectCollection();
            try
            {
                entity.Explode(exploded);
                ed.WriteMessage(
                    $"\n  [Level {level}] Exploded '{entity.GetType().Name}'" +
                    $" → {exploded.Count} object(s)");
            }
            catch (Exception ex)
            {
                ed.WriteMessage(
                    $"\n  [Level {level}] Explode failed: {ex.Message}");
                return;
            }

            foreach (DBObject obj in exploded)
            {
                try
                {
                    string raw = string.Empty;
                    Point3d location = Point3d.Origin;
                    string srcType = string.Empty;

                    if (obj is MText mText)
                    {
                        raw = mText.Contents;
                        location = mText.Location;
                        srcType = "MText";
                    }
                    else if (obj is DBText dbText)
                    {
                        raw = dbText.TextString;
                        location = dbText.Position;
                        srcType = "DBText";
                    }
                    else if (obj is AttributeReference attRef)
                    {
                        raw = attRef.TextString;
                        location = attRef.Position;
                        srcType = "AttributeRef";
                    }
                    else if (obj is BlockReference blkRef)
                    {
                        CollectAllMTexts(blkRef, results, ed, level + 1, blockPos);
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(raw)) continue;

                    string clean = StripMTextFormatting(raw);

                    string stationValue, suffix;
                    ParseValueAndSuffix(clean, out stationValue, out suffix);

                    int wholeNumber = ExtractWholeNumber(stationValue);
                    double distToBlock = blockPos.HasValue
                        ? blockPos.Value.DistanceTo(location)
                        : 0.0;

                    results.Add(new MTextInfo
                    {
                        RawText = raw,
                        CleanText = clean,
                        Value = stationValue,
                        DirectionSuffix = suffix,
                        WholeNumber = wholeNumber,
                        Location = location,
                        SourceType = srcType,
                        DistanceToBlock = distToBlock
                    });

                    ed.WriteMessage(
                        $"\n    ✔ {srcType}" +
                        $"  X={location.X:F2}  Y={location.Y:F2}" +
                        $"  Dist={distToBlock:F2}" +
                        $"  Full=\"{clean}\"" +
                        $"  Value=\"{stationValue}\"" +
                        $"  WholeNo={wholeNumber}" +
                        $"  Suffix=\"{suffix}\"");
                }
                catch { }
                finally
                {
                    try { if (!obj.IsDisposed) obj.Dispose(); } catch { }
                }
            }
        }


        private int ExtractWholeNumber(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return -1;

            string part = value.Contains("+")
                ? value.Substring(0, value.IndexOf('+')).Trim()
                : value.Trim();

            int result;
            return int.TryParse(part, out result) ? result : -1;
        }


        private void ParseValueAndSuffix(
            string cleanText,
            out string stationValue,
            out string directionSuffix)
        {
            stationValue = cleanText?.Trim() ?? string.Empty;
            directionSuffix = string.Empty;

            if (string.IsNullOrWhiteSpace(cleanText)) return;

            string input = cleanText.Trim();

            string[] knownSuffixes = new[]
            {
                "NE", "NW", "SE", "SW", "LT", "RT",
                "N", "S", "E", "W", "L", "R"
            };

            foreach (string sfx in knownSuffixes)
            {
                if (input.EndsWith(" " + sfx, StringComparison.OrdinalIgnoreCase))
                {
                    stationValue = input
                        .Substring(0, input.Length - sfx.Length - 1).Trim();
                    directionSuffix = sfx.ToUpper();
                    return;
                }
            }

            foreach (string sfx in knownSuffixes)
            {
                if (string.Equals(input, sfx, StringComparison.OrdinalIgnoreCase))
                {
                    stationValue = string.Empty;
                    directionSuffix = sfx.ToUpper();
                    return;
                }
            }

            foreach (string sfx in knownSuffixes)
            {
                if (input.EndsWith(sfx, StringComparison.OrdinalIgnoreCase))
                {
                    string candidate = input
                        .Substring(0, input.Length - sfx.Length).Trim();
                    if (!string.IsNullOrWhiteSpace(candidate))
                    {
                        stationValue = candidate;
                        directionSuffix = sfx.ToUpper();
                        return;
                    }
                }
            }

            stationValue = input;
            directionSuffix = string.Empty;
        }


        private string StripMTextFormatting(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return string.Empty;

            string result = raw;
            result = Regex.Replace(result, @"\\([HhWwQqTtAaCcFf])[^;]*;", string.Empty);
            result = Regex.Replace(result, @"\\f[^;]*;", string.Empty);
            result = result.Replace(@"\P", "\n");
            result = result.Replace(@"\p", "\n");
            result = Regex.Replace(result, @"\\[LlOoKk]", string.Empty);
            result = Regex.Replace(result, @"\\S([^;]*);", "$1");
            result = result.Replace(@"\~", " ");
            result = result.Replace(@"\X", string.Empty);
            result = result.Replace(@"\n", string.Empty);
            result = result.Replace("{", string.Empty);
            result = result.Replace("}", string.Empty);
            return result.Trim();
        }


        private Point3d GetEntityCenter(Entity entity)
        {
            try
            {
                Extents3d ext = entity.GeometricExtents;
                return new Point3d(
                    (ext.MinPoint.X + ext.MaxPoint.X) / 2.0,
                    (ext.MinPoint.Y + ext.MaxPoint.Y) / 2.0,
                    (ext.MinPoint.Z + ext.MaxPoint.Z) / 2.0);
            }
            catch
            {
                if (entity is BlockReference br) return br.Position;
                return Point3d.Origin;
            }
        }
    }
}
