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
        [CommandMethod("TEST_MLEADER_BLOCK")]
        public void TestMLeaderBlock()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // 👉 Select MLeader
            PromptEntityOptions peo = new PromptEntityOptions("\nSelect MLeader: ");
            peo.SetRejectMessage("\nOnly MLeader allowed.");
            peo.AddAllowedClass(typeof(MLeader), true);

            PromptEntityResult per = ed.GetEntity(peo);

            if (per.Status != PromptStatus.OK)
                return;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                MLeader ml = tr.GetObject(per.ObjectId, OpenMode.ForRead) as MLeader;

                if (ml == null)
                {
                    ed.WriteMessage("\nNot a valid MLeader.");
                    return;
                }

                // ✅ Check content type
                ed.WriteMessage($"\nContentType: {ml.ContentType}");

                if (ml.ContentType != ContentType.BlockContent)
                {
                    ed.WriteMessage("\nThis MLeader does NOT contain a block.");
                    return;
                }

                try
                {
                    // 🔥 Explode to get BlockReference
                    DBObjectCollection objs = new DBObjectCollection();
                    ml.Explode(objs);

                    bool found = false;

                    foreach (DBObject obj in objs)
                    {
                        if (obj is BlockReference br)
                        {
                            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);

                            string blockName = btr.Name;

                            ed.WriteMessage($"\nBlock Name: {blockName}");

                            found = true;
                            break;
                        }
                    }

                    if (!found)
                    {
                        ed.WriteMessage("\nNo BlockReference found after explode.");
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError: {ex.Message}");
                }

                tr.Commit();
            }
        }

        [CommandMethod("TEST_MLEADER_BLOCK_ATT")]
        public void TestMLeaderBlockAttribute()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            PromptEntityOptions peo = new PromptEntityOptions("\nSelect MLeader: ");
            peo.SetRejectMessage("\nOnly MLeader allowed.");
            peo.AddAllowedClass(typeof(MLeader), true);

            PromptEntityResult per = ed.GetEntity(peo);

            if (per.Status != PromptStatus.OK)
                return;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                MLeader ml = tr.GetObject(per.ObjectId, OpenMode.ForRead) as MLeader;

                if (ml == null)
                    return;

                ed.WriteMessage($"\nContentType: {ml.ContentType}");

                if (ml.ContentType != ContentType.BlockContent)
                {
                    ed.WriteMessage("\nNot a block-based MLeader.");
                    return;
                }

                try
                {
                    DBObjectCollection objs = new DBObjectCollection();
                    ml.Explode(objs);

                    foreach (DBObject obj in objs)
                    {
                        if (obj is BlockReference br)
                        {
                            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                            ed.WriteMessage($"\nBlock Name: {btr.Name}");

                            // 🔥 SAFE ATTRIBUTE READ
                            foreach (ObjectId attId in br.AttributeCollection)
                            {
                                DBObject attObj = tr.GetObject(attId, OpenMode.ForRead);

                                if (attObj is AttributeReference attRef)
                                {
                                    ed.WriteMessage($"\nTag: {attRef.Tag}  Value: {attRef.TextString}");

                                    if (attRef.Tag.Equals("PH", StringComparison.OrdinalIgnoreCase))
                                    {
                                        ed.WriteMessage($"\n👉 PH Value: {attRef.TextString}");
                                    }
                                }
                                else
                                {
                                    // Debug: what type is this?
                                    ed.WriteMessage($"\nUnexpected type: {attObj.GetType().Name}");
                                }
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError: {ex.Message}");
                }

                tr.Commit();
            }
        }

        [CommandMethod("MLEADER_PH_SORT_PS")]
        public void GetAllMLeaderPHSorted()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // Store results: (numeric value, original text, ObjectId)
            List<(int num, string raw, ObjectId id)> data = new List<(int, string, ObjectId)>();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord ps = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForRead);

                foreach (ObjectId entId in ps)
                {
                    Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;

                    if (!(ent is MLeader ml))
                        continue;

                    // Only block-based MLeader
                    if (ml.ContentType != ContentType.BlockContent)
                        continue;

                    try
                    {
                        DBObjectCollection objs = new DBObjectCollection();
                        ml.Explode(objs);

                        foreach (DBObject obj in objs)
                        {
                            string textValue = null;

                            // ✅ Extract visible text
                            if (obj is DBText txt)
                                textValue = txt.TextString;

                            else if (obj is MText mt)
                                textValue = mt.Contents;

                            if (string.IsNullOrWhiteSpace(textValue))
                                continue;

                            // 🔥 Extract numeric part (P2 → 2, PH10 → 10)
                            string numberPart = new string(textValue.Where(char.IsDigit).ToArray());

                            if (int.TryParse(numberPart, out int num))
                            {
                                data.Add((num, textValue, entId));
                            }
                        }
                    }
                    catch
                    {
                        // silently ignore problematic MLeaders
                    }
                }

                tr.Commit();
            }

            // ✅ Sort numerically
            var sorted = data.OrderBy(x => x.num).ToList();

            // ✅ Output
            ed.WriteMessage("\n--- Sorted PH Values (PaperSpace) ---");

            foreach (var item in sorted)
            {
                ed.WriteMessage($"\nPH: {item.raw}  →  {item.num} | Id: {item.id}");
            }
        }

        [CommandMethod("GET_MLEADER_PH")]
        public void GetMLeaderPH()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            PromptEntityOptions peo = new PromptEntityOptions("\nSelect MLeader: ");
            peo.SetRejectMessage("\nOnly MLeader allowed.");
            peo.AddAllowedClass(typeof(MLeader), true);

            PromptEntityResult per = ed.GetEntity(peo);

            if (per.Status != PromptStatus.OK)
                return;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                MLeader ml = tr.GetObject(per.ObjectId, OpenMode.ForRead) as MLeader;

                if (ml == null)
                    return;

                if (ml.ContentType != ContentType.BlockContent)
                {
                    ed.WriteMessage("\nNot a block-based MLeader.");
                    return;
                }

                try
                {
                    DBObjectCollection objs = new DBObjectCollection();
                    ml.Explode(objs);

                    foreach (DBObject obj in objs)
                    {
                        // ✅ Read TEXT (this holds your PH value visually)
                        if (obj is DBText txt)
                        {
                            ed.WriteMessage($"\nFound Text: {txt.TextString}");
                        }
                        else if (obj is MText mt)
                        {
                            ed.WriteMessage($"\nFound MText: {mt.Contents}");
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError: {ex.Message}");
                }

                tr.Commit();
            }
        }

        [CommandMethod("READ_MLEADER_FULL")]
        public void ReadMLeaderFull()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            PromptEntityOptions peo = new PromptEntityOptions("\nSelect MLeader: ");
            peo.SetRejectMessage("\nOnly MLeader allowed.");
            peo.AddAllowedClass(typeof(MLeader), true);

            PromptEntityResult per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK)
                return;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                MLeader ml = tr.GetObject(per.ObjectId, OpenMode.ForRead) as MLeader;

                if (ml == null)
                    return;

                ed.WriteMessage($"\nContentType: {ml.ContentType}");

                if (ml.ContentType != ContentType.BlockContent)
                {
                    ed.WriteMessage("\nNot a block-based MLeader.");
                    return;
                }

                try
                {
                    DBObjectCollection exploded = new DBObjectCollection();
                    ml.Explode(exploded);

                    foreach (DBObject obj in exploded)
                    {
                        ed.WriteMessage($"\nType: {obj.GetType().Name}");

                        // ✅ BLOCK NAME
                        if (obj is BlockReference br)
                        {
                            try
                            {
                                BlockTableRecord btr =
                                    (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);

                                ed.WriteMessage($"\nBlock Name: {btr.Name}");

                                // 🔥 TRY ATTRIBUTE READING (SAFE)
                                foreach (ObjectId attId in br.AttributeCollection)
                                {
                                    DBObject attObj = null;

                                    try
                                    {
                                        attObj = tr.GetObject(attId, OpenMode.ForRead);
                                    }
                                    catch
                                    {
                                        continue; // skip invalid objects
                                    }

                                    // ✅ Case 1: AttributeReference (ideal)
                                    if (attObj is AttributeReference attRef)
                                    {
                                        ed.WriteMessage($"\n[ATTR-REF] {attRef.Tag} = {attRef.TextString}");
                                    }
                                    // ✅ Case 2: AttributeDefinition (fallback)
                                    else if (attObj is AttributeDefinition attDef)
                                    {
                                        ed.WriteMessage($"\n[ATTR-DEF] {attDef.Tag} = {attDef.TextString}");
                                    }
                                }
                            }
                            catch (System.Exception ex)
                            {
                                ed.WriteMessage($"\nBlock Error: {ex.Message}");
                            }
                        }

                        // ✅ TEXT (MOST RELIABLE IN YOUR CASE)
                        if (obj is DBText txt)
                        {
                            ed.WriteMessage($"\n[TEXT] {txt.TextString}");
                        }
                        else if (obj is MText mt)
                        {
                            ed.WriteMessage($"\n[MTEXT] {mt.Contents}");
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError: {ex.Message}");
                }

                tr.Commit();
            }
        }

    }
}
