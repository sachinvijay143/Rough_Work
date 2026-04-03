using IntelliCAD.ApplicationServices;
using IntelliCAD.EditorInput;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Teigha.DatabaseServices;
using Teigha.Geometry;
using Teigha.Runtime;
using static System.Net.Mime.MediaTypeNames;
using Application = IntelliCAD.ApplicationServices.Application;
using Exception = System.Exception;

namespace Rough_Works
{
    public class Class1
    {
        [CommandMethod("ShowForm")]
        public void ShowForm()
        {
            Form1 form = new Form1();            
            form.ShowDialog();
        }

        [CommandMethod("ReadAlignmentStationLabels")]
        public void ReadAlignmentStationLabels()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                    List<AlignmentStationLabelInfo> labelInfoList = new List<AlignmentStationLabelInfo>();

                    foreach (ObjectId objId in btr)
                    {
                        DBObject obj = tr.GetObject(objId, OpenMode.ForRead);

                        // --- Method 1: Match by DXF class name (common for Civil/survey objects) ---
                        string dxfName = obj.GetType().Name;
                        ed.WriteMessage($"\nChecking object with Handle {obj.Handle} and DXF Name: {dxfName}");
                        if (dxfName.Contains("AlignmentStationLabel") ||
                            dxfName.Contains("AlignmentLabel"))
                        {
                            var info = ExtractFromObject(obj, tr);
                            if (info != null) labelInfoList.Add(info);
                        }

                        // --- Method 2: Check XData for alignment station label data ---
                        var xdataInfo = ReadXData(obj, tr);
                        if (xdataInfo != null) labelInfoList.Add(xdataInfo);

                        // --- Method 3: Check Extension Dictionary ---
                        if (obj.ExtensionDictionary != ObjectId.Null)
                        {
                            var extDictInfo = ReadExtensionDictionary(obj, tr);
                            if (extDictInfo != null) labelInfoList.AddRange(extDictInfo);
                        }
                    }

                    // Output results
                    if (labelInfoList.Count == 0)
                    {
                        ed.WriteMessage("\nNo AlignmentStationLabelling found in Model Space.");
                    }
                    else
                    {
                        ed.WriteMessage($"\nFound {labelInfoList.Count} AlignmentStationLabel(s):\n");
                        foreach (var info in labelInfoList)
                        {
                            ed.WriteMessage($"  Handle: {info.Handle}");
                            ed.WriteMessage($"  Station: {info.Station}");
                            ed.WriteMessage($"  Label Text: {info.LabelText}");
                            ed.WriteMessage($"  Position: {info.Position}");
                            ed.WriteMessage($"  Source: {info.Source}\n");
                        }
                    }

                    tr.Commit();
                }
            }
            catch (Exception ex)
            {
                ed.WriteMessage($"\nError reading AlignmentStationLabels: {ex.Message}");
            }
        }

        // --- Method 1: Extract properties via reflection (for custom ODA/ProgeCAD objects) ---
        private AlignmentStationLabelInfo ExtractFromObject(DBObject obj, Transaction tr)
        {
            try
            {
                var info = new AlignmentStationLabelInfo
                {
                    Handle = obj.Handle.ToString(),
                    Source = "DirectObject"
                };

                // Use reflection to read properties dynamically
                var type = obj.GetType();

                var stationProp = type.GetProperty("Station") ?? type.GetProperty("StationValue");
                if (stationProp != null)
                    info.Station = stationProp.GetValue(obj)?.ToString() ?? "N/A";

                var labelProp = type.GetProperty("LabelText") ?? type.GetProperty("Text");
                if (labelProp != null)
                    info.LabelText = labelProp.GetValue(obj)?.ToString() ?? "N/A";

                var posProp = type.GetProperty("Position") ?? type.GetProperty("Location");
                if (posProp != null)
                    info.Position = posProp.GetValue(obj)?.ToString() ?? "N/A";

                return info;
            }
            catch
            {
                return null;
            }
        }

        // --- Method 2: Read XData (Extended Entity Data) ---
        private AlignmentStationLabelInfo ReadXData(DBObject obj, Transaction tr)
        {
            // Common XData app names used by Civil/survey tools in ProgeCAD
            string[] appNames = { "CIVIL_ALIGN", "ALIGNMENT_LABEL", "STATION_LABEL", "PROG_CIVIL" };

            foreach (var appName in appNames)
            {
                ResultBuffer rb = obj.GetXDataForApplication(appName);
                if (rb == null) continue;

                var info = new AlignmentStationLabelInfo
                {
                    Handle = obj.Handle.ToString(),
                    Source = $"XData[{appName}]"
                };

                var values = new List<string>();
                foreach (TypedValue tv in rb)
                    values.Add($"Code {tv.TypeCode}: {tv.Value}");

                info.LabelText = string.Join(" | ", values);
                rb.Dispose();
                return info;
            }

            return null;
        }

        // --- Method 3: Read Extension Dictionary ---
        private List<AlignmentStationLabelInfo> ReadExtensionDictionary(DBObject obj, Transaction tr)
        {
            var results = new List<AlignmentStationLabelInfo>();
            try
            {
                DBDictionary extDict = tr.GetObject(obj.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
                if (extDict == null) return results;

                foreach (DBDictionaryEntry entry in extDict)
                {
                    string key = entry.Key.ToUpper();
                    if (!key.Contains("STATION") && !key.Contains("ALIGN") && !key.Contains("LABEL"))
                        continue;

                    DBObject dictObj = tr.GetObject(entry.Value, OpenMode.ForRead);
                    results.Add(new AlignmentStationLabelInfo
                    {
                        Handle = obj.Handle.ToString(),
                        LabelText = dictObj.ToString(),
                        Source = $"ExtDict[{entry.Key}]"
                    });
                }
            }
            catch { /* Some dictionary entries may not be readable */ }

            return results;
        }
        [CommandMethod("GetEntityDXFName")]
        public void GetEntityDXFName()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // Prompt user to select an entity
            PromptEntityOptions peo = new PromptEntityOptions("\nSelect an entity to inspect: ");
            peo.AllowNone = false;
            PromptEntityResult per = ed.GetEntity(peo);

            if (per.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nNo entity selected.");
                return;
            }

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    DBObject obj = tr.GetObject(per.ObjectId, OpenMode.ForRead);

                    // --- DXF Name & Type Info ---
                    ed.WriteMessage("\n=======================================");
                    ed.WriteMessage($"\n.NET Type Name  : {obj.GetType().Name}");
                    ed.WriteMessage($"\n.NET Full Name  : {obj.GetType().FullName}");
                    ed.WriteMessage($"\nHandle          : {obj.Handle}");
                    ed.WriteMessage($"\nObject ID       : {obj.ObjectId}");

                    // DXF name via RXObject (most reliable way in Teigha/ProgeCAD)
                    ed.WriteMessage($"\nDXF Name (RXClass): {obj.GetRXClass().DxfName}");

                    // Also show class name from RXClass
                    ed.WriteMessage($"\nRXClass Name    : {obj.GetRXClass().Name}");

                    // If it's an Entity, show layer and other common props
                    if (obj is Entity ent)
                    {
                        ed.WriteMessage($"\nLayer           : {ent.Layer}");
                        ed.WriteMessage($"\nColor Index     : {ent.ColorIndex}");
                        ed.WriteMessage($"\nLinetype        : {ent.Linetype}");
                    }

                    ed.WriteMessage("\n=======================================\n");

                    tr.Commit();
                }
            }
            catch (Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
            }
        }
        [CommandMethod("ReadAeccStationLabels")]
        public void ReadAeccStationLabels()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            int found = 0;

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Iterate all handles in the database
                    for (long i = 1; i < db.Handseed.Value; i++)
                    {
                        ObjectId objId;
                        Handle handle = new Handle(i);
                        if (!db.TryGetObjectId(handle, out objId)) continue;
                        if (objId.IsNull || !objId.IsValid) continue;

                        try
                        {
                            DBObject obj = tr.GetObject(objId, OpenMode.ForRead, false, true);
                            if (obj == null) continue;

                            // ✅ Match by DXF name string directly — no RXClass.GetClass() needed
                            if (obj.GetRXClass().DxfName != "AECC_ALIGNMENT_STATION_LABELING")
                                continue;

                            found++;
                            ed.WriteMessage($"\n========== Station Label #{found} ==========");
                            ed.WriteMessage($"\n  Handle         : {obj.Handle}");
                            ed.WriteMessage($"\n  DXF Name       : {obj.GetRXClass().DxfName}");
                            ed.WriteMessage($"\n  .NET Type      : {obj.GetType().FullName}");

                            DisplayAllProperties(obj, ed);
                            DisplayXData(obj, ed);
                            DisplayExtensionDictionary(obj, tr, ed);

                            ed.WriteMessage("\n");
                        }
                        catch { /* skip unreadable objects */ }
                    }

                    ed.WriteMessage(found == 0
                        ? "\nNo AECC_ALIGNMENT_STATION_LABELING entities found."
                        : $"\nTotal found: {found}");

                    tr.Commit();
                }
            }
            catch (Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}\n{ex.StackTrace}");
            }
        }

        [CommandMethod("PickAeccStationLabel")]
        public void PickAeccStationLabel()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptEntityOptions peo = new PromptEntityOptions(
                "\nSelect an AECC_ALIGNMENT_STATION_LABELING entity: ");
            peo.AllowNone = false;
            PromptEntityResult per = ed.GetEntity(peo);

            if (per.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nNo entity selected.");
                return;
            }

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    DBObject obj = tr.GetObject(per.ObjectId, OpenMode.ForRead);

                    if (obj.GetRXClass().DxfName != "AECC_ALIGNMENT_STATION_LABELING")
                    {
                        ed.WriteMessage($"\nWrong entity: '{obj.GetRXClass().DxfName}'");
                        return;
                    }

                    ed.WriteMessage($"\n  Handle   : {obj.Handle}");
                    ed.WriteMessage($"\n  DXF Name : {obj.GetRXClass().DxfName}");

                    // ── Step 1: Write just this entity to a temp DXF file ──
                    string tempDxf = Path.Combine(Path.GetTempPath(), "aecc_label_dump.dxf");

                    try
                    {
                        // Create a temp in-memory DB, copy the object into it, save as DXF
                        using (Database tempDb = new Database(true, true))
                        {
                            using (Transaction tempTr = tempDb.TransactionManager.StartTransaction())
                            {
                                // Copy the selected object into the temp DB's model space
                                IdMapping idMap = new IdMapping();
                                ObjectIdCollection ids = new ObjectIdCollection();
                                ids.Add(per.ObjectId);

                                db.WblockCloneObjects(
                                    ids,
                                    tempDb.CurrentSpaceId,
                                    idMap,
                                    DuplicateRecordCloning.Ignore,
                                    false);

                                tempTr.Commit();
                            }

                            // Save temp DB as DXF (ASCII) so we can read the raw codes
                            tempDb.DxfOut(tempDxf, 16, DwgVersion.Current);
                        }

                        // ── Step 2: Read and display DXF file — filter relevant sections ──
                        ed.WriteMessage("\n\n--- Raw DXF Codes for Selected Entity ---");
                        string[] lines = File.ReadAllLines(tempDxf);
                        bool insideEntity = false;
                        bool foundAecc = false;

                        for (int i = 0; i < lines.Length - 1; i++)
                        {
                            string code  = lines[i].Trim();
                            string value = lines[i + 1].Trim();

                            // Detect start of ENTITIES section
                            if (value == "AECC_ALIGNMENT_STATION_LABELING")
                            {
                                insideEntity = true;
                                foundAecc = true;
                                ed.WriteMessage($"\n  [{code}] {value}  <-- ENTITY START");
                                i++; // skip value line
                                continue;
                            }

                            // Stop at next entity or end of section
                            if (insideEntity)
                            {
                                if ((code == "0" && foundAecc && value != "AECC_ALIGNMENT_STATION_LABELING")
                                    || value == "ENDSEC")
                                {
                                    insideEntity = false;
                                    break;
                                }

                                // ── Highlight text-related DXF group codes ──
                                // Code 1  = primary text string
                                // Code 3  = additional text
                                // Code 300-309 = arbitrary text (Civil stores labels here)
                                bool isTextCode = code == "1"   || code == "3"   ||
                                                  code == "300" || code == "301" ||
                                                  code == "302" || code == "303" ||
                                                  code == "304" || code == "305" ||
                                                  code == "9";   // variable name (DXF header vars)

                                string marker = isTextCode ? "  <<< TEXT" : "";
                                ed.WriteMessage($"\n  [{code,4}] {value}{marker}");
                                i++; // skip value line
                            }
                        }

                        if (!foundAecc)
                            ed.WriteMessage("\n  AECC entity not found in DXF output.");

                        // Clean up temp file
                        if (File.Exists(tempDxf))
                            File.Delete(tempDxf);
                    }
                    catch (Exception dxfEx)
                    {
                        ed.WriteMessage($"\nDXF export error: {dxfEx.Message}");
                    }

                    tr.Commit();
                }
            }
            catch (Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}\n{ex.StackTrace}");
            }
        }

        // ── Proximity scan: find all DBText / MText within radius ──
        private void FindNearbyTextEntities(Database db, Transaction tr, Editor ed,
            Point3d center, double searchRadius)
        {
            bool found = false;

            for (long i = 1; i < db.Handseed.Value; i++)
            {
                ObjectId objId;
                if (!db.TryGetObjectId(new Handle(i), out objId)) continue;
                if (!objId.IsValid || objId.IsNull) continue;

                try
                {
                    DBObject obj = tr.GetObject(objId, OpenMode.ForRead, false, true);
                    if (obj == null) continue;

                    // Only interested in text entities
                    string text = null;
                    Point3d pos = Point3d.Origin;

                    if (obj is DBText dbTxt)
                    {
                        text = dbTxt.TextString;
                        pos = dbTxt.Position;
                    }
                    else if (obj is MText mt)
                    {
                        text = mt.Contents;
                        pos = mt.Location;
                    }
                    else continue;

                    // Check if within search radius
                    double dist = center.DistanceTo(pos);
                    if (dist > searchRadius) continue;

                    ed.WriteMessage($"\n  [{obj.GetRXClass().DxfName}] Handle: {obj.Handle}");
                    ed.WriteMessage($"\n    TextString : {text}");
                    ed.WriteMessage($"\n    Position   : ({pos.X:F4}, {pos.Y:F4})");
                    ed.WriteMessage($"\n    Distance   : {dist:F4}");
                    found = true;
                }
                catch { }
            }

            if (!found)
                ed.WriteMessage("\n  No nearby text entities found within radius 50.");
        }

        // ── Extract text from any known text entity type ──
        private string ExtractText(DBObject obj)
        {
            if (obj is DBText t) return t.TextString;
            if (obj is MText mt) return mt.Contents;
            if (obj is AttributeReference ar) return ar.TextString;
            if (obj is AttributeDefinition ad) return ad.TextString;

            try
            {
                var prop = obj.GetType().GetProperty("TextString",
                               BindingFlags.Public | BindingFlags.Instance)
                           ?? obj.GetType().GetProperty("Contents",
                               BindingFlags.Public | BindingFlags.Instance)
                           ?? obj.GetType().GetProperty("Text",
                               BindingFlags.Public | BindingFlags.Instance);
                return prop?.GetValue(obj)?.ToString();
            }
            catch { return null; }
        }

        // --- Method 1: Reflect known text property names ---
        private void TryReadTextProperties(DBObject obj, Editor ed)
        {
            // Known text-related property names across AutoCAD/Civil APIs
            string[] textPropNames = {
                "TextString", "Text", "Contents", "LabelText",
                "StationText", "StationValue", "Station",
                "TextOverride", "SummaryText", "DisplayText"
            };

            Type type = obj.GetType();
            bool anyFound = false;

            foreach (string propName in textPropNames)
            {
                PropertyInfo prop = type.GetProperty(propName,
                    BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                if (prop == null) continue;

                try
                {
                    object val = prop.GetValue(obj);
                    ed.WriteMessage($"\n  {propName,-25}: {val ?? "(null)"}");
                    anyFound = true;
                }
                catch (Exception ex)
                {
                    ed.WriteMessage($"\n  {propName,-25}: [Error: {ex.InnerException?.Message ?? ex.Message}]");
                }
            }

            if (!anyFound)
                ed.WriteMessage("\n  No direct text properties found.");
        }

        // --- Method 2: Open as BlockReference or iterate OwnedObjects ---
        private void ReadChildObjectTexts(DBObject obj, Transaction tr, Editor ed)
        {
            bool found = false;

            // Case A: Label might be a BlockReference containing text entities
            if (obj is BlockReference bref)
            {
                BlockTableRecord btr = tr.GetObject(
                    bref.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;

                if (btr != null)
                {
                    foreach (ObjectId childId in btr)
                    {
                        try
                        {
                            DBObject child = tr.GetObject(childId, OpenMode.ForRead);
                            string text = ExtractTextFromObject(child);
                            if (text != null)
                            {
                                ed.WriteMessage($"\n  [{child.GetType().Name}] TextString: {text}");
                                found = true;
                            }
                        }
                        catch { }
                    }
                }
            }

            // Case B: Try OwnerId chain — get all objects that share the same owner
            try
            {
                DBObject ownerObj = tr.GetObject(obj.OwnerId, OpenMode.ForRead);
                ed.WriteMessage($"\n  Owner Type: {ownerObj?.GetType().Name ?? "null"}");
            }
            catch { }

            if (!found)
                ed.WriteMessage("\n  No child text objects found.");
        }

        // --- Method 3: Scan entire DB for objects owned by this label handle ---
        private void ReadOwnedObjectsFromDB(DBObject parentObj, Database db, Transaction tr, Editor ed)
        {
            bool found = false;

            for (long i = 1; i < db.Handseed.Value; i++)
            {
                ObjectId childId;
                if (!db.TryGetObjectId(new Handle(i), out childId)) continue;
                if (childId.IsNull || !childId.IsValid) continue;

                try
                {
                    DBObject child = tr.GetObject(childId, OpenMode.ForRead, false, true);
                    if (child == null || child.OwnerId != parentObj.ObjectId) continue;

                    string dxf = child.GetRXClass().DxfName;
                    string text = ExtractTextFromObject(child);

                    ed.WriteMessage($"\n  Owned [{dxf}] Handle: {child.Handle}");

                    if (text != null)
                    {
                        ed.WriteMessage($"  => TextString: {text}");
                        found = true;
                    }
                    else
                    {
                        // Reflect all properties of this child
                        PropertyInfo[] props = child.GetType().GetProperties(
                            BindingFlags.Public | BindingFlags.Instance);
                        foreach (PropertyInfo p in props)
                        {
                            if (p.GetIndexParameters().Length > 0) continue;
                            try
                            {
                                object val = p.GetValue(child);
                                if (val == null) continue;
                                ed.WriteMessage($"\n      {p.Name,-30}: {val}");
                            }
                            catch { }
                        }
                    }
                }
                catch { }
            }

            if (!found)
                ed.WriteMessage("\n  No owned text objects found.");
        }

        // --- Method 4: Write object to a DXF memory stream and read group codes ---
        private void ReadRawDxfCodes(DBObject obj, Editor ed)
        {
            ed.WriteMessage("\n  All String / Numeric properties:");

            PropertyInfo[] props = obj.GetType().GetProperties(
                BindingFlags.Public | BindingFlags.Instance);

            foreach (PropertyInfo p in props)
            {
                if (p.GetIndexParameters().Length > 0) continue;

                // Only show string, double, int, bool types — skip complex objects
                if (p.PropertyType != typeof(string) &&
                    p.PropertyType != typeof(double) &&
                    p.PropertyType != typeof(float) &&
                    p.PropertyType != typeof(int) &&
                    p.PropertyType != typeof(long) &&
                    p.PropertyType != typeof(bool)) continue;

                try
                {
                    object val = p.GetValue(obj);
                    if (val == null) continue;
                    ed.WriteMessage($"\n    {p.Name,-35}: {val}");
                }
                catch (Exception ex)
                {
                    ed.WriteMessage($"\n    {p.Name,-35}: [Error: {ex.InnerException?.Message ?? ex.Message}]");
                }
            }
        }

        // --- Helper: extract TextString from common text entity types ---
        private string ExtractTextFromObject(DBObject obj)
        {
            // DBText (TEXT entity)
            if (obj is DBText dbText)
                return dbText.TextString;

            // MText entity
            if (obj is MText mtext)
                return mtext.Contents;

            // AttributeDefinition
            if (obj is AttributeDefinition attDef)
                return attDef.TextString;

            // AttributeReference
            if (obj is AttributeReference attRef)
                return attRef.TextString;

            // Fallback: reflect TextString or Contents property
            try
            {
                var prop = obj.GetType().GetProperty("TextString",
                    BindingFlags.Public | BindingFlags.Instance)
                    ?? obj.GetType().GetProperty("Contents",
                    BindingFlags.Public | BindingFlags.Instance);

                return prop?.GetValue(obj)?.ToString();
            }
            catch { return null; }
        }

        private void DisplayAllProperties(DBObject obj, Editor ed)
        {
            ed.WriteMessage("\n  --- Reflected Properties ---");
            PropertyInfo[] props = obj.GetType().GetProperties(
                BindingFlags.Public | BindingFlags.Instance);

            foreach (PropertyInfo prop in props)
            {
                if (prop.GetIndexParameters().Length > 0) continue;
                try
                {
                    object val = prop.GetValue(obj);
                    string display = val != null ? val.ToString() : "(null)";
                    if (display.Length > 120)
                        display = display.Substring(0, 120) + "...";
                    ed.WriteMessage($"\n    {prop.Name,-35}: {display}");
                }
                catch (Exception ex)
                {
                    ed.WriteMessage($"\n    {prop.Name,-35}: [Error: {ex.InnerException?.Message ?? ex.Message}]");
                }
            }
        }

        private void DisplayXData(DBObject obj, Editor ed)
        {
            try
            {
                ResultBuffer rb = obj.XData;
                if (rb == null) return;

                ed.WriteMessage("\n  --- XData ---");
                foreach (TypedValue tv in rb)
                    ed.WriteMessage($"\n    Code {tv.TypeCode,4}: {tv.Value}");

                rb.Dispose();
            }
            catch { }
        }

        private void DisplayExtensionDictionary(DBObject obj, Transaction tr, Editor ed)
        {
            if (obj.ExtensionDictionary == ObjectId.Null) return;

            try
            {
                ed.WriteMessage("\n  --- Extension Dictionary ---");
                DBDictionary dict = tr.GetObject(
                    obj.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
                if (dict == null) return;

                foreach (DBDictionaryEntry entry in dict)
                {
                    ed.WriteMessage($"\n    Key: {entry.Key}");
                    try
                    {
                        DBObject dictObj = tr.GetObject(entry.Value, OpenMode.ForRead);
                        ed.WriteMessage($" => Type: {dictObj.GetType().Name}");
                    }
                    catch { }
                }
            }
            catch { }
        }

        private static ObjectId _templateMLeaderId = ObjectId.Null;

        [CommandMethod("PICK_MLEADER_TEMPLATE")]
        public void PickMLeaderTemplate()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // ── Force user to pick from Model Space only ──
            // Switch to model space before picking
            object savedTileMode = Application.GetSystemVariable("TILEMODE");
            Application.SetSystemVariable("TILEMODE", 1); // Switch to model space

            try
            {
                PromptEntityOptions peo = new PromptEntityOptions(
                    "\nSwitch to MODEL SPACE - Select the correct MLeader as template: ");
                peo.SetRejectMessage("\nMust select a MLeader.");
                peo.AddAllowedClass(typeof(MLeader), true);
                PromptEntityResult per = ed.GetEntity(peo);

                if (per.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nTemplate selection cancelled.");
                    return;
                }

                // Verify it is in model space
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    MLeader ml = tr.GetObject(per.ObjectId, OpenMode.ForRead) as MLeader;
                    if (ml == null)
                    {
                        ed.WriteMessage("\nSelected object is not a MLeader.");
                        return;
                    }

                    // Confirm it belongs to model space block table record
                    ObjectId modelSpaceId = SymbolUtilityServices.GetBlockModelSpaceId(db);
                    BlockTableRecord owner = tr.GetObject(ml.OwnerId, OpenMode.ForRead) as BlockTableRecord;

                    if (owner == null || owner.ObjectId != modelSpaceId)
                    {
                        ed.WriteMessage("\n✘ Selected MLeader is NOT in model space.");
                        ed.WriteMessage("\n  Please run CHSPACE on a paper space MLeader first,");
                        ed.WriteMessage("\n  then run PICK_MLEADER_TEMPLATE again and select it.");
                        tr.Commit();
                        return;
                    }

                    _templateMLeaderId = per.ObjectId;

                    ed.WriteMessage($"\n✔ Template stored from MODEL SPACE.");
                    ed.WriteMessage($"\n  Scale      : {ml.Scale}");
                    ed.WriteMessage($"\n  BlockScale : {ml.BlockScale.X}");
                    tr.Commit();
                }
            }
            finally
            {
                // Restore original space
                Application.SetSystemVariable("TILEMODE", savedTileMode);
            }
        }

        // ─────────────────────────────────────────────────────────────
        // Auto-find template from model space — no manual picking needed
        // Scans model space for first MLeader that uses CIRCLE FOR LEADER block
        // ─────────────────────────────────────────────────────────────
        [CommandMethod("AUTO_FIND_TEMPLATE")]
        public void AutoFindTemplate()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Get the block id of CIRCLE FOR LEADER
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (!bt.Has("CIRCLE FOR LEADER"))
                {
                    ed.WriteMessage("\n✘ Block 'CIRCLE FOR LEADER' not found in drawing.");
                    return;
                }
                ObjectId circleBlockId = bt["CIRCLE FOR LEADER"];

                // Scan model space for MLeader using this block
                ObjectId modelSpaceId = SymbolUtilityServices.GetBlockModelSpaceId(db);
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(modelSpaceId, OpenMode.ForRead);

                foreach (ObjectId id in ms)
                {
                    MLeader ml = tr.GetObject(id, OpenMode.ForRead) as MLeader;
                    if (ml == null) continue;
                    if (ml.ContentType != ContentType.BlockContent) continue;
                    if (ml.BlockContentId != circleBlockId) continue;

                    // Found a valid model space MLeader using CIRCLE FOR LEADER
                    _templateMLeaderId = id;

                    ed.WriteMessage($"\n✔ Template auto-found in MODEL SPACE.");
                    ed.WriteMessage($"\n  ObjectId   : {id}");
                    ed.WriteMessage($"\n  Scale      : {ml.Scale}");
                    ed.WriteMessage($"\n  BlockScale : {ml.BlockScale.X}");
                    ed.WriteMessage($"\n  Position   : {ml.BlockPosition}");
                    tr.Commit();
                    return;
                }

                ed.WriteMessage("\n✘ No MLeader with 'CIRCLE FOR LEADER' found in model space.");
                ed.WriteMessage("\n  Either run CHSPACE on one paper space MLeader first,");
                ed.WriteMessage("\n  or manually run PICK_MLEADER_TEMPLATE from model space.");
                tr.Commit();
            }
        }

        // ─────────────────────────────────────────────────────────────
        // Main creation method — uses model space template via CopyFrom
        // ─────────────────────────────────────────────────────────────
        public void CreatePHMLeader(Transaction tr, Database db, Point3d arrowheadPt,
                                     string phValue, ObjectId templateMLeaderId)
        {
            double offsetX = 2.5;
            double offsetY = 10.5;
            Point3d bubblePt = new Point3d(arrowheadPt.X + offsetX,
                                            arrowheadPt.Y + offsetY,
                                            arrowheadPt.Z);

            MLeader template = tr.GetObject(templateMLeaderId, OpenMode.ForRead) as MLeader;
            if (template == null) throw new System.Exception("Template MLeader not found.");

            // Verify template is in model space
            ObjectId modelSpaceId = SymbolUtilityServices.GetBlockModelSpaceId(db);
            BlockTableRecord owner = tr.GetObject(template.OwnerId, OpenMode.ForRead) as BlockTableRecord;
            if (owner == null || owner.ObjectId != modelSpaceId)
                throw new System.Exception(
                    "Template MLeader must be in model space. " +
                    "Run AUTO_FIND_TEMPLATE or PICK_MLEADER_TEMPLATE first.");

            // 1. Append to model space — NO SetDatabaseDefaults
            BlockTableRecord ms = (BlockTableRecord)tr.GetObject(modelSpaceId, OpenMode.ForWrite);
            MLeader ml = new MLeader();
            ms.AppendEntity(ml);
            tr.AddNewlyCreatedDBObject(ml, true);

            // 2. CopyFrom immediately — copies all internal scale state from model space template
            ml.CopyFrom(template);

            // 3. Remove leader geometry copied from template
            try
            {
                System.Collections.ArrayList leaderIndexes = ml.GetLeaderIndexes();
                for (int i = leaderIndexes.Count - 1; i >= 0; i--)
                    ml.RemoveLeader((int)leaderIndexes[i]);
            }
            catch { }

            // 4. Set new geometry at correct positions
            int ldIdx = ml.AddLeader();
            int lnIdx = ml.AddLeaderLine(ldIdx);
            ml.AddFirstVertex(lnIdx, arrowheadPt);
            ml.AddLastVertex(lnIdx, bubblePt);
            ml.BlockPosition = bubblePt;

            // 5. Set PH attribute
            ObjectId blockId = ml.BlockContentId;
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

            ml.RecordGraphicsModified(true);
        }

        // ── Updated caller that uses the template ──
        [CommandMethod("TEST_CREATE_MLEADER")]
        public void TestCreateMLeader()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            if (_templateMLeaderId == ObjectId.Null || !_templateMLeaderId.IsValid)
            {
                ed.WriteMessage("\n✘ No template set. Run PICK_MLEADER_TEMPLATE first.");
                return;
            }

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Test: create at a fixed point
                Point3d arrowPt = new Point3d(0, 0, 0);
                CreatePHMLeader(tr, db, arrowPt, "7.5", _templateMLeaderId);
                tr.Commit();
            }

            ed.WriteMessage("\n✔ Test MLeader created. Check size matches template.");
        }

    }

    // --- Data model for label info ---
    public class AlignmentStationLabelInfo
    {
        public string Handle { get; set; } = "N/A";
        public string Station { get; set; } = "N/A";
        public string LabelText { get; set; } = "N/A";
        public string Position { get; set; } = "N/A";
        public string Source { get; set; } = "N/A";
    }
}
