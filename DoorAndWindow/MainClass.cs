using System;
using Autodesk.AutoCAD.Runtime;
using _acAppSer = Autodesk.AutoCAD.ApplicationServices;
using _acGeo = Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Customization;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using System.IO;
using Autodesk.AutoCAD.Geometry;

namespace DoorAndWindow
{
    public class MainClass
    {
        [CommandMethod("InsertDoor")]
        public void door()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            // Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            //Getting lower corner point
            PromptPointOptions pointOptions = new PromptPointOptions("\nPlease select the lower corner point :");
            pointOptions.AllowArbitraryInput = false;
            pointOptions.AllowNone = true;
            PromptPointResult prPtRes1 = ed.GetPoint(pointOptions);
            if (prPtRes1.Status != PromptStatus.OK) return;
            Point3d cornerPointLower = prPtRes1.Value;

            //Getting upper corner point
            pointOptions.Message = "\nPlease select the upper corner point :";
            prPtRes1 = ed.GetPoint(pointOptions);
            if (prPtRes1.Status != PromptStatus.OK) return;
            Point3d cornerPointUper = prPtRes1.Value;


            PromptStringOptions stringPrompt = new PromptStringOptions("\nEnter the start distance from lower corner :");
            stringPrompt.AllowSpaces = false;

            //start distance
            PromptResult distanceString = ed.GetString(stringPrompt);
            if (distanceString.Status != PromptStatus.OK && decimal.TryParse(distanceString.StringResult, out _)) return;
            decimal startDistance;
            decimal.TryParse(distanceString.StringResult, out startDistance);

            //door width
            stringPrompt.Message = "\nEnter the door width :";
            distanceString = ed.GetString(stringPrompt);
            if (distanceString.Status != PromptStatus.OK && decimal.TryParse(distanceString.StringResult, out _)) return;
            decimal doorWidth;
            decimal.TryParse(distanceString.StringResult, out doorWidth);

            //Getting direction
            pointOptions.Message = "\nSelect the direction to place the door";
            prPtRes1 = ed.GetPoint(pointOptions);
            if (prPtRes1.Status != PromptStatus.OK) return;
            Point3d directionPoint = prPtRes1.Value;

            //Calculating direction
            direction direction;
            double xFactor, yFactor;
            xFactor = directionPoint.X- cornerPointLower.X;
            yFactor = directionPoint.Y- cornerPointLower.Y;

            if( Math.Abs(xFactor)>Math.Abs(yFactor)) //X is the determinent direction: right or left
            {
                if(xFactor>0)
                {
                    direction = direction.right;
                }else
                {
                    direction = direction.left;
                }

            }else //Y is the determinent direction: Top or bottom
            {
                if(yFactor>0)
                {
                    direction = direction.top;
                }else
                {
                    direction = direction.bottom;
                }
            }

            //Trim required line
            double wallOffset = Math.Abs(cornerPointUper.Y - cornerPointLower.Y);
            ed.WriteMessage("\noffset distance :" + wallOffset);

            Point3d firstLineLowerPoint = new Point3d(), firstLineUpperPoint= new Point3d(), secondLineLowerPoint=new Point3d(),
                secondLineUpperPoint=new Point3d();
            if (direction==direction.right)
            {
                firstLineLowerPoint = new Point3d(cornerPointLower.X + Convert.ToDouble(startDistance), cornerPointLower.Y, cornerPointLower.Z);
                firstLineUpperPoint = new Point3d(cornerPointLower.X + Convert.ToDouble(startDistance), cornerPointLower.Y + wallOffset, cornerPointLower.Z);

                secondLineLowerPoint = new Point3d(firstLineLowerPoint.X + Convert.ToDouble(doorWidth), cornerPointLower.Y, cornerPointLower.Z);
                secondLineUpperPoint = new Point3d(firstLineUpperPoint.X + Convert.ToDouble(doorWidth), cornerPointLower.Y + wallOffset, cornerPointLower.Z);

            }
            else if(direction==direction.left)
            {
                firstLineLowerPoint = new Point3d(cornerPointLower.X - Convert.ToDouble(startDistance), cornerPointLower.Y, cornerPointLower.Z);
                firstLineUpperPoint = new Point3d(cornerPointLower.X - Convert.ToDouble(startDistance), cornerPointLower.Y + wallOffset, cornerPointLower.Z);

                secondLineLowerPoint = new Point3d(firstLineLowerPoint.X - Convert.ToDouble(doorWidth), cornerPointLower.Y, cornerPointLower.Z);
                secondLineUpperPoint = new Point3d(firstLineUpperPoint.X - Convert.ToDouble(doorWidth), cornerPointLower.Y + wallOffset, cornerPointLower.Z);

            }
            else if(direction == direction.top)
            {
                firstLineLowerPoint = new Point3d(cornerPointLower.X , cornerPointLower.Y + Convert.ToDouble(startDistance), cornerPointLower.Z);
                firstLineUpperPoint = new Point3d(cornerPointLower.X + wallOffset , cornerPointLower.Y + Convert.ToDouble(startDistance), cornerPointLower.Z);

                secondLineLowerPoint = new Point3d(cornerPointLower.X , firstLineLowerPoint.Y + Convert.ToDouble(doorWidth), cornerPointLower.Z);
                secondLineUpperPoint = new Point3d(firstLineUpperPoint.X , firstLineLowerPoint.Y + Convert.ToDouble(doorWidth), cornerPointLower.Z);

            }
            else if(direction==direction.bottom)
            {
                firstLineLowerPoint = new Point3d(cornerPointLower.X, cornerPointLower.Y -Convert.ToDouble(startDistance), cornerPointLower.Z);
                firstLineUpperPoint = new Point3d(cornerPointLower.X + wallOffset, cornerPointLower.Y - Convert.ToDouble(startDistance), cornerPointLower.Z);

                secondLineLowerPoint = new Point3d(cornerPointLower.X, firstLineLowerPoint.Y - Convert.ToDouble(doorWidth), cornerPointLower.Z);
                secondLineUpperPoint = new Point3d(firstLineUpperPoint.X, firstLineLowerPoint.Y - Convert.ToDouble(doorWidth), cornerPointLower.Z);
            }

            //Getting wall lines
            Line lowerWallLine = new Line();
            Line upperWallLine = new Line();


            using (Transaction tr = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Block table for read
                BlockTable acBlkTbl;
                acBlkTbl = tr.GetObject(acCurDb.BlockTableId,
                OpenMode.ForRead) as BlockTable;
                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = tr.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                OpenMode.ForWrite) as BlockTableRecord;

                //Getting the wall lines
                PromptSelectionResult acSSPrompt;
                acSSPrompt = ed.SelectCrossingWindow(firstLineLowerPoint,
                                                          secondLineUpperPoint);
                // If the prompt status is OK, objects were selected
                if (acSSPrompt.Status == PromptStatus.OK)
                {
                    SelectionSet acSSet = acSSPrompt.Value;

                    foreach (ObjectId objectId in acSSet.GetObjectIds())
                    {
                        //Entity ent = tr.GetObject(objectId, OpenMode.ForRead) as Entity;
                        if (objectId.ObjectClass.DxfName == "LINE")
                        {
                            Line pl = tr.GetObject(objectId, OpenMode.ForWrite) as Line;
                            if(direction==direction.right || direction==direction.bottom)  //because line starts from left
                            {
                                if (pl.StartPoint == cornerPointLower)
                                {
                                    lowerWallLine = pl;
                                }
                                else if (pl.StartPoint == cornerPointUper)
                                {
                                    upperWallLine = pl;
                                }

                            }
                            else if(direction==direction.top || direction ==direction.left) //because line starts from top
                            {
                                if (pl.EndPoint == cornerPointLower)
                                {
                                    lowerWallLine = pl;
                                }
                                else if (pl.EndPoint == cornerPointUper)
                                {
                                    upperWallLine = pl;
                                }
                            }
                            
                        }
                    }

                    //Now do the trimming operation by creating two line
                    Line lowerLineTrimBefore = new Line(), upperLineTrimBefore=new Line(), lowerLineTrimAfter=new Line(),
                        upperLineTrimAfter=new Line();

                    if(direction==direction.right || direction == direction.bottom)
                    {
                        lowerLineTrimBefore = new Line(lowerWallLine.StartPoint, firstLineLowerPoint);
                        upperLineTrimBefore = new Line(upperWallLine.StartPoint, firstLineUpperPoint);

                        lowerLineTrimAfter = new Line(secondLineLowerPoint, lowerWallLine.EndPoint);
                        upperLineTrimAfter = new Line(secondLineUpperPoint, upperWallLine.EndPoint);
                    }
                    else if (direction==direction.top || direction == direction.left)
                    {
                        lowerLineTrimBefore = new Line(lowerWallLine.EndPoint, firstLineLowerPoint);
                        upperLineTrimBefore = new Line(upperWallLine.EndPoint, firstLineUpperPoint);

                        lowerLineTrimAfter = new Line(secondLineLowerPoint, lowerWallLine.StartPoint);
                        upperLineTrimAfter = new Line(secondLineUpperPoint, upperWallLine.StartPoint);
                    }
                    

                    lowerLineTrimBefore.SetDatabaseDefaults();
                    upperLineTrimBefore.SetDatabaseDefaults();
                    lowerLineTrimAfter.SetDatabaseDefaults();
                    upperLineTrimAfter.SetDatabaseDefaults();

                    acBlkTblRec.AppendEntity(lowerLineTrimBefore);
                    acBlkTblRec.AppendEntity(upperLineTrimBefore);
                    acBlkTblRec.AppendEntity(lowerLineTrimAfter);
                    acBlkTblRec.AppendEntity(upperLineTrimAfter);

                    tr.AddNewlyCreatedDBObject(lowerLineTrimBefore, true);
                    tr.AddNewlyCreatedDBObject(upperLineTrimBefore, true);
                    tr.AddNewlyCreatedDBObject(lowerLineTrimAfter, true);
                    tr.AddNewlyCreatedDBObject(upperLineTrimAfter, true);

                    //Delete the base line to complete trimming operation
                    lowerWallLine.Erase();
                    upperWallLine.Erase();
                                       
                    tr.Commit();

                }
                else
                {
                    Application.ShowAlertDialog("Number of objects selected: 0");
                }

            }

            

        }


    }

    public enum direction
    {
        right,
        left,
        top,
        bottom
    }
}
