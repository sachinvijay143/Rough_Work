using IntelliCAD.ApplicationServices;
using IntelliCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Teigha.Colors;
using Teigha.DatabaseServices;
using Teigha.Geometry;
using Teigha.Runtime;
using Application = IntelliCAD.ApplicationServices.Application;

namespace Rough_Works
{
    internal class Common_functions
    {
        [CommandMethod("SetLayer")]
        public string SetActiveLayer(string Layname)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            string result = "";
            string layerName = Layname; // Replace with the name of the layer you want to set as current
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Open the Layer table for read
                    LayerTable lt = trans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                    if (lt.Has(layerName))
                    {
                        // Get the ObjectId of the layer
                        ObjectId layerId = lt[layerName];
                        // Open the layer
                        LayerTableRecord ltr = trans.GetObject(layerId, OpenMode.ForRead) as LayerTableRecord;
                        // Set the layer current
                        db.Clayer = layerId;
                        ed.WriteMessage($"\nLayer '{layerName}' has been set as current.");
                        result = ($"\nLayer '{layerName}' has been set as current.");
                    }
                    else
                    {
                        ed.WriteMessage($"\nLayer '{layerName}' does not exist in the drawing.");
                        // Application.ShowAlertDialog("Layername missing: " & layername)
                        LayerTableRecord newLayer = new LayerTableRecord();
                        newLayer.Name = layerName;
                        newLayer.LineWeight = LineWeight.LineWeight005;
                        // newLayer.Description = "This is new layer"
                        newLayer.Color = Color.FromColorIndex(ColorMethod.ByAci, 1);
                        lt.UpgradeOpen();
                        lt.Add(newLayer);
                        trans.AddNewlyCreatedDBObject(newLayer, true);
                        ObjectId layerId = lt[layerName];
                        // Open the layer
                        LayerTableRecord ltr = trans.GetObject(layerId, OpenMode.ForRead) as LayerTableRecord;
                        // Set the layer current
                        db.Clayer = layerId;
                        ed.WriteMessage($"\nLayer '{layerName}' has been set as current.");
                        result = ($"\nLayer '{layerName}' has been set as current.");
                    }
                    trans.Commit();
                }
                catch (Teigha.Runtime.Exception ex)
                {
                    ed.WriteMessage($"\nError: {ex.Message}");
                    result = ($"\nError: {ex.Message}");
                    trans.Abort();
                }
            }
            return result;
        }
        [CommandMethod("CreateAndAssignALayer")]
        public string CreateAndAssignALayer(string Layname)
        {
            // ' Get the current document and database
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            string result = "";
            // ' Start a transaction
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // ' Open the Layer table for read
                LayerTable LyrTbl = trans.GetObject(db.LayerTableId, OpenMode.ForWrite) as LayerTable;
                if (LyrTbl.Has(Layname) == false)
                {
                    LayerTableRecord LyrTblRec = new LayerTableRecord();
                    // ' Assign the layer ACI color 1 and a name
                    //LyrTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci,1);
                    LyrTblRec.Color = Color.FromColorIndex(ColorMethod.ByLayer, 1);
                    LyrTblRec.Name = Layname;
                    // ' Upgrade the layer table for write
                    LyrTbl.UpgradeOpen();
                    // ' Append the new layer to the  layer table and the transaction
                    LyrTbl.Add(LyrTblRec);
                    trans.AddNewlyCreatedDBObject(LyrTblRec, true);
                    result = "New layer " + Layname + " created...";
                }
                else
                {
                    result = Layname + " Already exists...";
                }
                trans.Commit();
            }
            return result;
        }

        [CommandMethod("DOTNET_EXAMPLES", "ExGetKeywords", "ExGetKeywords", CommandFlags.Modal)]
        public string ExGetKeywords()
        {
            Document objDoc = Application.DocumentManager.MdiActiveDocument;
            Editor objEd = objDoc.Editor;
            // With PromptKeywordOptions
            PromptKeywordOptions objPrompt = new PromptKeywordOptions("\nType Keyword");
            objPrompt.AllowArbitraryInput = false;
            objPrompt.AllowNone = true;
            objPrompt.Keywords.Add("Aaa");
            objPrompt.Keywords.Add("Bbb");
            objPrompt.Keywords.Add("Ccc");
            objPrompt.Keywords.Add("Ddd");
            objPrompt.Keywords.Add("Eee");
            objPrompt.AppendKeywordsToMessage = true;
            objPrompt.Keywords.Default = "Ddd";
            PromptResult objResult = objEd.GetKeywords(objPrompt);
            //if (objResult.Status == PromptStatus.OK)
            //{
            //    objEd.WriteMessage("\nKeyword {0}", objResult.StringResult);
            //}
            //else if (objResult.Status == PromptStatus.Keyword)
            //{
            //    objEd.WriteMessage("\nKeyword {0}", objResult.StringResult);
            //}
            //else
            //{
            //    objEd.WriteMessage("\nKeyword Error {0}", objResult.Status);
            //}

            //Without PromptKeywordOptions
            //objResult = objEd.GetKeywords("\nType Keyword", "Aaa", "Bbb", "Ccc", "Ddd", "Eee");
            if (objResult.Status == PromptStatus.OK)
            {
                objEd.WriteMessage("\nKeyword {0}", objResult.StringResult);
            }
            else if (objResult.Status == PromptStatus.Keyword)
            {
                objEd.WriteMessage("\nKeyword {0}", objResult.StringResult);
            }
            else
            {
                objEd.WriteMessage("\nKeyword Error {0}", objResult.Status);
            }
            return objResult.StringResult;
        }
        [CommandMethod("DOTNET_EXAMPLES", "_ExGetDistance", "ExGetDistance", CommandFlags.Modal)]
        public static Double ExGetDistance()
        {
            Document objDoc = Application.DocumentManager.MdiActiveDocument;
            Editor objEd = objDoc.Editor;

            PromptDistanceOptions objOption = new PromptDistanceOptions("\nDotNet:  Starting point for distance: ");
            objOption.DefaultValue = 0;
            objOption.UseDefaultValue = true;
            objOption.AllowNone = true;
            objOption.AllowArbitraryInput = false;
            objOption.AllowZero = true;
            objOption.AllowNegative = false;

            PromptDoubleResult objDistance = objEd.GetDistance(objOption);
            if (objDistance.Status == PromptStatus.OK)
            {
                objEd.WriteMessage("\nDistance = {0}", objDistance.Value);
            }
            return objDistance.Value;
        }
        [CommandMethod("DOTNET_EXAMPLES", "_ExGetAngle", "ExGetAngle", CommandFlags.Modal)]
        public static Double ExGetAngle()
        {
            PromptAngleOptions objOptions = new PromptAngleOptions("\nEnter Angle");
            objOptions.AllowNone = true;
            objOptions.AllowZero = true;
            objOptions.AllowArbitraryInput = true;
            // objOptions.BasePoint =
            objOptions.UseBasePoint = false;
            objOptions.DefaultValue = Math.PI / 2.0;
            objOptions.UseDashedLine = false;
            objOptions.UseAngleBase = false;
            objOptions.UseDefaultValue = true;

            Editor objEditor = Application.DocumentManager.MdiActiveDocument.Editor;
            PromptDoubleResult objResult = objEditor.GetAngle(objOptions);
            objEditor.WriteMessage("\nStatus: {0}\nValue: {1}\nString Result: {2}", objResult.Status, objResult.Value, objResult.StringResult);
            return objResult.Value;
        }

        [CommandMethod("DOTNET_EXAMPLES", "_ExGetCorner", "ExGetCorner", CommandFlags.Modal)]
        public static String ExGetBoundingbox()
        {
            Editor objEditor = Application.DocumentManager.MdiActiveDocument.Editor;
            string corner_val = "";
            PromptPointResult objResultBase = objEditor.GetPoint("\nBase Point");
            if (objResultBase.Status == PromptStatus.OK)
            {
                PromptCornerOptions objOptions = new PromptCornerOptions("\nEnter Corner", objResultBase.Value);
                corner_val = corner_val + objResultBase.Value.ToString();
                objOptions.AllowNone = true;
                objOptions.AllowArbitraryInput = true;
                objOptions.BasePoint = objResultBase.Value;
                // objOptions.BasePoint =
                objOptions.LimitsChecked = false;
                objOptions.UseDashedLine = false;


                PromptPointResult objResult = objEditor.GetCorner(objOptions);
                objEditor.WriteMessage("\nStatus: {0}\nValue: {1}\nString Result: {2}", objResult.Status, objResult.Value, objResult.StringResult);
                corner_val = corner_val + objResult.Value.ToString();
            }
            return corner_val;
        }
        [CommandMethod("DOTNET_EXAMPLES", "_ExGetDouble", "ExGetDouble", CommandFlags.Modal)]
        public static Double ExGetDouble(string prompt)
        {
            //"\nEnter Double"
            PromptDoubleOptions objOptions = new PromptDoubleOptions("\n" + prompt);
            objOptions.AllowNone = true;
            objOptions.AllowZero = true;
            objOptions.AllowArbitraryInput = true;
            objOptions.DefaultValue = 500.0;
            objOptions.UseDefaultValue = true;

            Editor objEditor = Application.DocumentManager.MdiActiveDocument.Editor;
            PromptDoubleResult objResult = objEditor.GetDouble(objOptions);
            objEditor.WriteMessage("\nStatus: {0}\nValue: {1}\nString Result: {2}", objResult.Status, objResult.Value, objResult.StringResult);
            return objResult.Value;
        }
        [CommandMethod("DOTNET_EXAMPLES", "_ExGetInteger", "ExGetInteger", CommandFlags.Modal)]
        public static int ExGetInteger(string prompt)
        {
            //"\nEnter Integer"
            PromptIntegerOptions objOptions = new PromptIntegerOptions("\n" + prompt);
            objOptions.AllowNone = true;
            objOptions.AllowZero = true;
            objOptions.AllowArbitraryInput = true;
            objOptions.DefaultValue = 500;
            objOptions.UseDefaultValue = true;

            Editor objEditor = Application.DocumentManager.MdiActiveDocument.Editor;
            PromptIntegerResult objResult = objEditor.GetInteger(objOptions);
            objEditor.WriteMessage("\nStatus: {0}\nValue: {1}\nString Result: {2}", objResult.Status, objResult.Value, objResult.StringResult);
            return objResult.Value;
        }
        [CommandMethod("DOTNET_EXAMPLES", "_ExGetPoint", "ExGetPoint", CommandFlags.Modal)]
        public static Point3d ExGetPoint(string prompt)
        {
            //"\nEnter Point"
            PromptPointOptions objOptions = new PromptPointOptions("\n" + prompt);
            objOptions.AllowNone = true;
            objOptions.AllowArbitraryInput = true;
            // objOptions.BasePoint
            objOptions.LimitsChecked = false;
            Editor objEditor = Application.DocumentManager.MdiActiveDocument.Editor;
            PromptPointResult objResult = objEditor.GetPoint(objOptions);
            objEditor.WriteMessage("\nStatus: {0}\nValue: {1}\nString Result: {2}", objResult.Status, objResult.Value, objResult.StringResult);
            return objResult.Value;
        }
        [CommandMethod("DOTNET_EXAMPLES", "_ExGetString", "ExGetString", CommandFlags.Modal)]
        public static string ExGetString(string prompt)
        {
            //"\nEnter String"//
            PromptStringOptions objOptions = new PromptStringOptions("\n" + prompt);
            objOptions.AllowSpaces = true;
            objOptions.DefaultValue = "Test";
            objOptions.UseDefaultValue = true;

            Editor objEditor = Application.DocumentManager.MdiActiveDocument.Editor;
            PromptResult objResult = objEditor.GetString(objOptions);
            objEditor.WriteMessage("\nStatus: {0}\nValue: {1}", objResult.Status, objResult.StringResult);
            return objResult.StringResult;
        }

        public static object GetSystemVariable(string variableName)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            object value = null;
            try
            {
                value = Application.GetSystemVariable(variableName);
            }
            catch (Teigha.Runtime.Exception ex)
            {
                // Handle or report the error
                doc.Editor.WriteMessage($"\nError getting system variable '{variableName}': {ex.Message}");
            }
            return value;
        }
        [CommandMethod("DOTNET_EXAMPLES", "_ExDetailSelection", "ExDetailSelection", CommandFlags.Modal)]
        public static void ExDetailSelection()
        {
            Document objDoc = Application.DocumentManager.MdiActiveDocument;
            Editor objEd = objDoc.Editor;
            PromptSelectionOptions objOption = new PromptSelectionOptions();
            objOption.AllowDuplicates = true;
            objOption.AllowSubSelections = true;
            objOption.ForceSubSelections = true;
            objOption.PrepareOptionalDetails = true;

            PromptSelectionResult objResult = objEd.GetSelection(objOption); // objFilter
            if (objResult.Status == PromptStatus.OK)
            {
                objEd.WriteMessage("\nFollowing Entities selected ({0}):", objResult.Value.Count);
                foreach (SelectedObject item in objResult.Value)
                {
                    ObjectId id = item.ObjectId;
                    objEd.WriteMessage("\n  {0}: {1} [Selection Method: {2}]", id, id.ObjectClass.DxfName, item.SelectionMethod);
                    SelectionMethod selMet = item.SelectionMethod;
                    if (item.OptionalDetails != null)
                    {
                        objEd.WriteMessage("\n    Details:");
                        objEd.WriteMessage("\n      Transform: {0}", item.OptionalDetails.Transform);
                        ObjectId[] objectIds = item.OptionalDetails.GetContainers();
                        objEd.WriteMessage("\n      Entities:");
                        foreach (ObjectId objDeytailId in objectIds)
                        {
                            objEd.WriteMessage("\n        {0}: {1}", objDeytailId, objDeytailId.ObjectClass.DxfName);
                        }
                    }
                }
            }
        }
        //Use this method to get angle between two points
        [CommandMethod("GetAnglepoints")]
        public static void proceed()
        {
            Point3d p1 = Common_functions.ExGetPoint("Pick first point");
            Point3d p2 = Common_functions.ExGetPoint("Pick second");
            double value = GetAngleBetweenPoints(p1, p2);
        }
        public static double GetAngleBetweenPoints(Point3d point1, Point3d point2)
        {
            // Create a vector from the first point to the second point
            Vector3d vector = point2 - point1;
            // Get the angle of the vector in the XY plane
            double angle = Vector3d.XAxis.GetAngleTo(vector, Vector3d.ZAxis);

            //Point2d pt1 = new Point2d(point1.X, point1.Y);
            //Point2d pt2 = new Point2d(point2.X, point2.Y);
            //double angle = pt1.GetVectorTo(pt2).Angle;
            //MessageBox.Show(angle.ToString());
            return angle;
        }

        //Use this sample code to get the entity from the dwg file
        [CommandMethod("SelectAndDisplayEntityType")]
        public void SelectAndDisplayEntityType()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            // Use the function to select an entity
            Entity selectedEntity = SelectEntity();

            if (selectedEntity != null)
            {
                // Output the type of the selected entity
                ed.WriteMessage($"\nSelected entity type: {selectedEntity.GetType().Name}");
            }
            else
            {
                ed.WriteMessage("\nNo entity selected or an error occurred.");
            }
        }
        public static Entity SelectEntity()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Entity selectedEntity = null;
            try
            {
                // Prompt the user to select an entity
                PromptEntityOptions peo = new PromptEntityOptions("\nSelect an entity: ");
                PromptEntityResult per = ed.GetEntity(peo);
                // Check the prompt status
                if (per.Status == PromptStatus.OK)
                {
                    // Start a transaction to access the selected entity
                    using (Transaction tr = doc.TransactionManager.StartTransaction())
                    {
                        // Get the selected entity
                        selectedEntity = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Entity;
                        // Commit the transaction
                        tr.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("Error: " + ex.Message);
            }
            return selectedEntity;
        }
        public static Point3d Midpoint(Point3d mnpnt, Point3d mxpnt)
        {
            double midX = (mxpnt.X - mnpnt.X) / 2.0 + mnpnt.X;
            double midY = (mxpnt.Y - mnpnt.Y) / 2.0 + mnpnt.Y;
            double midZ = (mxpnt.Z - mnpnt.Z) / 2.0 + mnpnt.Z;
            return MakePoint(midX, midY, midZ);
        }
        public static Point3d MakePoint(double x, double y, double z)
        {
            return new Point3d(x, y, z);
        }
        [CommandMethod("ZW_enter")]
        public static void zw_execute()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            Point3d st = Common_functions.ExGetPoint("pick the fisrt point: ");
            Point3d en = Common_functions.ExGetPoint("pick the second point:");
            ZoomWin(ed, st, en);
        }
        public static void ZoomWin(Editor ed, Point3d min, Point3d max)
        {
            try
            {
                Point2d min2d = new Point2d(min.X, min.Y);
                Point2d max2d = new Point2d(max.X, max.Y);
                ViewTableRecord view = new ViewTableRecord();

                string lower = $"{min.X},{min.Y},{min.Z}";
                string upper = $"{max.X},{max.Y},{max.Z}";
                string cmd = $"_.ZOOM _W {lower} {upper} ";

                view.CenterPoint = min2d + ((max2d - min2d) / 2.0);
                view.Height = max2d.Y - min2d.Y;
                view.Width = max2d.X - min2d.X;

                ed.SetCurrentView(view);
                ed.Document.SendStringToExecute(cmd, false, false, false);
                ed.Regen();
            }
            catch (Teigha.Runtime.Exception ex)
            {

            }

        }
        //Use this sample snippet code to get the polarpoints when passing a point3d,angle and distance values
        [CommandMethod("Polarpoints")]
        public static void Main()
        {
            Point3d bP = ExGetPoint("Pick point");
            Object_Creation.Circle_Creation(bP);
            Point3d snd = ExGetPoint("Pick secnd point");
            double angl = GetAngleBetweenPoints(bP, snd);
            double dist = 5.0;
            Point3d pp = PolarPoint(bP, angl, dist);
            Object_Creation.Circle_Creation(pp);
        }
        public static Point3d PolarPoint(Point3d basePoint, double angle, double distance)
        {
            return new Point3d(
                basePoint.X + distance * Math.Cos(angle),
                basePoint.Y + distance * Math.Sin(angle),
                basePoint.Z
            );
        }
        [CommandMethod("GetDistanceBetweenPoints")]
        public static double GetDistanceBetweenPoints(Point3d point1, Point3d point2)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            // Calculate the distance between the points
            double distance = point1.DistanceTo(point2);
            // Display the distance
            ed.WriteMessage($"\nDistance between points: {distance}");
            return distance;
        }

        [CommandMethod("GetBlockInsertionPoint")]
        public async void GetBlockInsertionPoint()
        {
            Point3d insertionPoint = await GetBlockInsertionPointStatic();
            if (insertionPoint != Point3d.Origin)
            {
                Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
                ed.WriteMessage($"\nBlock insertion point: {insertionPoint}");
            }
        }

        public static async Task<Point3d> GetBlockInsertionPointStatic()
        {
            await Task.Delay(1);
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // Prompt the user to select a block reference
            MessageBox.Show("Select a block reference:");
            PromptEntityOptions promptOptions = new PromptEntityOptions("\nSelect a block reference: ");
            promptOptions.SetRejectMessage("\nSelected entity is not a block reference.");
            promptOptions.AddAllowedClass(typeof(BlockReference), false);

            PromptEntityResult promptResult = ed.GetEntity(promptOptions);

            if (promptResult.Status != PromptStatus.OK)
                return Point3d.Origin; // Return origin point if selection is not successful

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // Get the selected block reference
                BlockReference blockRef = trans.GetObject(promptResult.ObjectId, OpenMode.ForRead) as BlockReference;

                if (blockRef != null)
                {
                    // Get the insertion point of the block reference
                    Point3d insertionPoint = blockRef.Position;
                    trans.Commit();
                    return insertionPoint;
                }
            }
            ed.Dispose();
            return Point3d.Origin; // Return origin point if block reference is not valid

        }
        [CommandMethod("ResetAutoCADCursor")]
        public static void ResetAutoCADCursor()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            // Clear any active custom cursor
            ed.SetImpliedSelection(new ObjectId[0]);
            ed.PointMonitor -= CustomPointMonitor; // Remove any custom point monitor if it was set

            // Reset cursor to default
            Application.SetSystemVariable("CURSOR", 0);

            ed.WriteMessage("\nCursor reset to AutoCAD default.");
        }
        private static void CustomPointMonitor(object sender, PointMonitorEventArgs e)
        {
            // Custom point monitor logic here
        }
        [CommandMethod("GetInsertionPoint")]
        public void GetInsertionPoint()
        {
            Point3d insertionPoint = GetInsertionPointStatic();
            if (insertionPoint != Point3d.Origin)
            {
                Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
                ed.WriteMessage($"\nBlock insertion point: {insertionPoint}");
            }
        }
        [DllImport("user32.dll")]
        static extern IntPtr LoadCursor(IntPtr hInstance, int lpCursorName);

        [DllImport("user32.dll")]
        static extern IntPtr SetCursor(IntPtr hCursor);

        const int IDC_CROSS = 32515;

        [CommandMethod("ChangeCursorToCrosshair")]
        public void ChangeCursorToCrosshair()
        {
            IntPtr hCrosshair = LoadCursor(IntPtr.Zero, IDC_CROSS);
            SetCursor(hCrosshair);
        }
        public static Point3d GetInsertionPointStatic()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // Prompt the user to select a block reference
            MessageBox.Show("Select a block reference:");
            PromptEntityOptions promptOptions = new PromptEntityOptions("\nSelect a block reference: ");
            promptOptions.SetRejectMessage("\nSelected entity is not a block reference.");
            promptOptions.AddAllowedClass(typeof(BlockReference), false);

            PromptEntityResult promptResult = ed.GetEntity(promptOptions);

            if (promptResult.Status != PromptStatus.OK)
                return Point3d.Origin; // Return origin point if selection is not successful

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // Get the selected block reference
                BlockReference blockRef = trans.GetObject(promptResult.ObjectId, OpenMode.ForRead) as BlockReference;

                if (blockRef != null)
                {
                    // Get the insertion point of the block reference
                    Point3d insertionPoint = blockRef.Position;
                    trans.Commit();
                    return insertionPoint;
                }
            }

            return Point3d.Origin; // Return origin point if block reference is not valid
        }
        [CommandMethod("SelectEntityAndReturnInsertionPoint")]
        public Point3d SelectEntityAndReturnInsertionPoint()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            // Prompt the user to select an entity
            PromptEntityOptions peo = new PromptEntityOptions("\nSelect an entity: ");
            PromptEntityResult per = ed.GetEntity(peo);
            Point3d insertionPoint = new Point3d();
            if (per.Status == PromptStatus.OK)
            {
                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    Entity ent = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Entity;
                    if (ent != null)
                    {
                        // Check if the entity is a type that has an insertion point
                        if (ent is BlockReference blockRef)
                        {
                            insertionPoint = blockRef.Position;
                            ed.WriteMessage($"\nInsertion point of the selected block: {insertionPoint}");
                        }
                        else if (ent is DBText dbText)
                        {
                            insertionPoint = dbText.Position;
                            ed.WriteMessage($"\nInsertion point of the selected text: {insertionPoint}");
                        }
                        else if (ent is MText mText)
                        {
                            insertionPoint = mText.Location;
                            ed.WriteMessage($"\nInsertion point of the selected MText: {insertionPoint}");
                        }
                        else
                        {
                            ed.WriteMessage("\nSelected entity does not have an insertion point.");
                        }
                    }
                    tr.Commit();
                }
            }
            else
            {
                ed.WriteMessage("\nNo entity selected or command canceled.");
            }
            return insertionPoint;
        }


    }
}
