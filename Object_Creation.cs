using IntelliCAD.ApplicationServices;
using IntelliCAD.EditorInput;
using System;
using System.Collections.Generic;
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
    internal class Object_Creation
    {
        [CommandMethod("Circle_Creation", "Circle_Creation", "Circle_Creation", CommandFlags.Modal)]
        public static string Circle_Creation(Point3d ppoint)
        {
            // Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;
            string result = "";
            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                try
                {
                    // Open the Block table for read
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

                    // Open the Block table record Model space for write
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    // Create a circle that is at 2,3 with a radius of 4.25
                    using (Circle acCirc = new Circle())
                    {
                        //Point3d ptStart = Common_functions.ExGetPoint("Enter the centre point for the circle: ");   // Exit if the user presses ESC or cancels the command
                        Point3d ptStart = ppoint;

                        if (ptStart.X > 0 && ptStart.Y > 0)
                        {
                            acCirc.Center = ptStart;
                            double rad = Common_functions.ExGetDouble("Enter the radius: ");
                            if (rad > 0)
                            {
                                acCirc.Radius = rad;
                            }
                            else
                            {
                                MessageBox.Show("Please enter the value for radius");
                                result = ("Please enter the value for radius");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Please pick centre point for the circle");
                            result = ("Please pick centre point for the circle");
                        }
                        // Add the new object to the block table record and the transaction
                        acBlkTblRec.AppendEntity(acCirc);
                        acTrans.AddNewlyCreatedDBObject(acCirc, true);
                    }

                    // Save the new object to the database
                    acTrans.Commit();
                }
                catch (Teigha.Runtime.Exception ex)
                {
                    ed.WriteMessage($"\nError: {ex.Message}");
                    result = ($"\nError: {ex.Message}");
                    acTrans.Abort();
                }

            }
            return result;
        }

        [CommandMethod("CreateMText")]
        public static string CreateMText()
        {
            // Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;
            string result = "";
            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                try
                {
                    // Open the Block table for read
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                                    OpenMode.ForRead) as BlockTable;

                    // Open the Block table record Model space for write
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                    OpenMode.ForWrite) as BlockTableRecord;

                    // Create a multiline text object
                    using (MText acMText = new MText())
                    {
                        Point3d pnt = Common_functions.ExGetPoint("Enter the Insertion point for the Mtext object");
                        if (pnt.X > 0 && pnt.Y > 0)
                        {
                            acMText.Location = pnt;

                            string TextString = Common_functions.ExGetString("Enter the Text String value");
                            if (TextString != "")
                            {
                                acMText.Width = 4;
                                acMText.TextHeight = 4;
                                acMText.Contents = TextString;
                                acBlkTblRec.AppendEntity(acMText);
                                acTrans.AddNewlyCreatedDBObject(acMText, true);
                            }
                            else
                            {
                                MessageBox.Show("Please enter the value for Text String");
                                result = ("Please enter the value for Text String");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Please enter the Insertion point for the Mtext object");
                            result = ("Please enter the Insertion point for the Mtext object");
                        }
                    }
                    // Save the changes and dispose of the transaction
                    acTrans.Commit();
                }
                catch (Teigha.Runtime.Exception ex)
                {
                    ed.WriteMessage($"\nError: {ex.Message}");
                    result = ($"\nError: {ex.Message}");
                    acTrans.Abort();
                }
                return result;
            }
        }
        [CommandMethod("CreateText")]
        public static string CreateText()
        {
            // Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;
            string result = "";

            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                try
                {
                    // Open the Block table for read
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    // Open the Block table record Model space for write
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    // Create a single-line text object
                    using (DBText acText = new DBText())
                    {
                        Point3d pnt = Common_functions.ExGetPoint("Enter the Insertion point for the Text object");
                        if (pnt.X > 0 && pnt.Y > 0)
                        {
                            acText.Position = pnt;
                            acText.Height = 0.5;
                            string TextString = Common_functions.ExGetString("Enter the Text String value");
                            if (TextString != "")
                            {
                                acText.TextString = TextString;
                                acBlkTblRec.AppendEntity(acText);
                                acTrans.AddNewlyCreatedDBObject(acText, true);
                            }
                            else
                            {
                                MessageBox.Show("Please enter the Text String value");
                                result = ("Please enter the Text String value");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Please enter the Insertion point for the Text object");
                            result = ("Please enter the Insertion point for the Text object");
                        }
                    }
                    // Save the changes and dispose of the transaction
                    acTrans.Commit();
                }
                catch (Teigha.Runtime.Exception ex)
                {
                    ed.WriteMessage($"\nError: {ex.Message}");
                    result = ($"\nError: {ex.Message}");
                    acTrans.Abort();
                }
                return result;
            }
        }
        [CommandMethod("ldr")]
        public static string leaderLine_creation(string Layname)
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;
            string rslt = "";

            using (var tran = acDoc.Database.TransactionManager.StartTransaction())
            {
                try
                {
                    BlockTable blkname = (BlockTable)tran.GetObject(acCurDb.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord blkrec = (BlockTableRecord)tran.GetObject(blkname[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                    DBDictionary dict = acCurDb.MLeaderStyleDictionaryId.GetObject(OpenMode.ForRead) as DBDictionary;
                    Point3d startPt = Common_functions.ExGetPoint("Pick starting point: ");
                    Point3d scnPt;
                    Point3d endPt;
                    if (startPt.X > 0 && startPt.Y > 0)
                    {
                        scnPt = Common_functions.ExGetPoint("Pick second point: ");
                        if (scnPt.X > 0 && scnPt.Y > 0)
                        {
                            endPt = Common_functions.ExGetPoint("Pick end point: ");
                            if (endPt.X > 0 && endPt.Y > 0)
                            {
                                using (Leader mldr = new Leader())
                                {
                                    mldr.AppendVertex(startPt);
                                    mldr.AppendVertex(scnPt);
                                    mldr.AppendVertex(endPt);
                                    mldr.HasArrowHead = true;
                                    mldr.Layer = Layname;
                                    mldr.Color = Teigha.Colors.Color.FromColorIndex(Teigha.Colors.ColorMethod.ByLayer, 256);
                                    mldr.Linetype = "ByLayer";
                                    mldr.LinetypeScale = 1.0;
                                    mldr.LineWeight = LineWeight.ByLayer;

                                    mldr.Dimasz = 1.0;
                                    mldr.Annotative = AnnotativeStates.True;
                                    //mldr.Dimsah = 1.0;
                                    // MsgBox("arrow size: " & mldr.Dimasz.ToString)
                                    blkrec.AppendEntity(mldr);
                                    tran.AddNewlyCreatedDBObject(mldr, true);
                                }
                            }
                            else
                            {
                                MessageBox.Show("Please pick the end point");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Please pick the second point");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please pick the starting point");
                    }
                }
                catch (Teigha.Runtime.Exception ex)
                {
                    MessageBox.Show("Unexpected error --" + Environment.NewLine + ex.Message);
                    rslt = ("Unexpected error --" + Environment.NewLine + ex.Message);
                }
                tran.Commit();
                ed.WriteMessage("\nLeader Line Placed.");
                rslt = "\nLeader Line Placed.";
            }
            return rslt;
        }

        //Use this sample function to invoke Dimension creation function
        [CommandMethod("CreateLinearDimension")]
        public static void dim()
        {
            Point3d st = Common_functions.ExGetPoint("pick the fisrt point: ");
            Point3d en = Common_functions.ExGetPoint("pick the second point:");
            CreateLinearDimension(st, en);
        }
        public static void CreateLinearDimension(Point3d startpoint, Point3d endpoint)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Get the BlockTable and BlockTableRecord for the model space
                    BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    // Define the points for the dimension                    
                    Point3d dimLinePoint = Common_functions.Midpoint(startpoint, endpoint);

                    // Create the linear dimension
                    AlignedDimension dim = new AlignedDimension(
                        startpoint, // First definition point
                        endpoint, // Second definition point
                        dimLinePoint, // Dimension line point
                        "Dimension Text", // Dimension text (empty string for default)
                        db.Dimstyle // Dimension style (default)
                    );

                    // Add the dimension to the model space
                    btr.AppendEntity(dim);
                    trans.AddNewlyCreatedDBObject(dim, true);

                    // Commit the transaction
                    trans.Commit();

                    ed.WriteMessage("\nDimension created successfully.");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage("Error: " + ex.Message);
                    trans.Abort();
                }
            }
        }

        [CommandMethod("CreateMLeaderText")]
        public void CreateMLeaderText()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                try
                {
                    BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    // Create the MLeader object
                    MLeader mleader = new MLeader();
                    mleader.ContentType = ContentType.MTextContent;

                    // Define the position for the MLeader
                    Point3d leaderPoint = new Point3d(10, 10, 0);
                    mleader.AddLeaderLine(leaderPoint);
                    mleader.AddFirstVertex(0, leaderPoint);
                    mleader.AddLastVertex(0, new Point3d(15, 15, 0));

                    // Create MText for the MLeader
                    MText mtext = new MText();
                    mtext.Contents = "This is an MLeader text";
                    mtext.Location = new Point3d(15, 15, 0);

                    mleader.MText = mtext;

                    // Add the MLeader to the drawing
                    btr.AppendEntity(mleader);
                    trans.AddNewlyCreatedDBObject(mleader, true);

                    trans.Commit();
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage("Error: " + ex.Message);
                    trans.Abort();
                }
            }
        }

        [CommandMethod("InsBlock")]
        public void mn()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            string blockName = "POLE";
            PromptPointOptions pointOptions = new PromptPointOptions("\nSpecify insertion point: ");
            PromptPointResult pointResult = ed.GetPoint(pointOptions);
            if (pointResult.Status != PromptStatus.OK)
                return;

            Point3d insertionPoint = pointResult.Value;
            double scl = 1;
            PromptPointOptions ppo = new PromptPointOptions("Rotation: ");
            ppo.UseBasePoint = true;
            ppo.BasePoint = insertionPoint;
            PromptPointResult ppr = ed.GetPoint(ppo);
            Point3d insecpnt = ppr.Value;
            double dblRotation = Common_functions.GetAngleBetweenPoints(insertionPoint, insecpnt);
            Block_insertion(blockName, insertionPoint, dblRotation, scl);
        }
        public static void Block_insertion(string blockName, Point3d inspnt, double Rotation_value, double scl, string imm_layer = "0", double imm_ltscl = 0.0)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // Get the insertion point from the user            

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTable blockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                if (blockTable.Has(blockName))
                {
                    BlockTableRecord blockDef = trans.GetObject(blockTable[blockName], OpenMode.ForRead) as BlockTableRecord;

                    BlockTableRecord currentSpace = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                    using (BlockReference blockRef = new BlockReference(inspnt, blockDef.ObjectId))
                    {
                        blockRef.TransformBy(Matrix3d.Scaling(scl, inspnt));
                        Point3d basePt = blockRef.Position;
                        Matrix3d matrix = Matrix3d.Rotation(Rotation_value, Vector3d.ZAxis, basePt);
                        blockRef.TransformBy(matrix);
                        currentSpace.AppendEntity(blockRef);
                        trans.AddNewlyCreatedDBObject(blockRef, true);
                        blockRef.Layer = imm_layer;
                        blockRef.LinetypeScale = imm_ltscl;

                        // Add attribute references if the block has attributes
                        if (blockDef.HasAttributeDefinitions)
                        {
                            foreach (ObjectId id in blockDef)
                            {
                                DBObject obj = trans.GetObject(id, OpenMode.ForRead);
                                if (obj is AttributeDefinition attDef)
                                {
                                    using (AttributeReference attRef = new AttributeReference())
                                    {
                                        attRef.SetAttributeFromBlock(attDef, blockRef.BlockTransform);
                                        attRef.Position = attDef.Position.TransformBy(blockRef.BlockTransform);
                                        attRef.TextString = attDef.TextString;

                                        blockRef.AttributeCollection.AppendAttribute(attRef);
                                        trans.AddNewlyCreatedDBObject(attRef, true);
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    ed.WriteMessage("\nBlock not found.");
                }

                trans.Commit();
            }
        }

        [CommandMethod("DLINE")]
        public static void DraftLine(Document doc, Database db, Editor ed, string layname, Point3d startPoint, Point3d endPoint)
        {
            using (Transaction tx = db.TransactionManager.StartTransaction())
            {
                BlockTable blockTable = (BlockTable)tx.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord blockTableRecord = (BlockTableRecord)tx.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                // Create the line
                Line line = new Line(startPoint, endPoint);
                blockTableRecord.AppendEntity(line);
                tx.AddNewlyCreatedDBObject(line, true);

                // Commit the transaction
                tx.Commit();
            }

            // Update the drawing
            doc.Editor.Regen();
        }

    }
}
