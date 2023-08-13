#define ACAD
#define MY
#define PODGON

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using System.Linq;

#if ACAD
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System.Runtime.InteropServices;

#else
using HostMgd.ApplicationServices;
using Teigha.Runtime;
using HostMgd.EditorInput;
using Teigha.Colors;
using Teigha.DatabaseServices;
using Teigha.Geometry;
#endif


namespace VHLines
{
    public class CadToolsMainClass
    {
#if MY
        static double g_epsilon = 1.0;
        public static bool EqualTo(in double value1, in double value2, double epsilon)
        {
            return Math.Abs(value1 - value2) < epsilon;
        }
        public static int CompareTo(in double value1, in double value2, double epsilon)
        {
            var result = Math.Abs(value1 - value2);
            if (result < epsilon)
                return 0;
            else 
                return (int)result;
        }
        public static void getHeightAndWidth(in Extents3d extents, out double height, out double width)
        {
            height = extents.MaxPoint.Y - extents.MinPoint.Y;
            width = extents.MaxPoint.X - extents.MinPoint.X;
        }
        public static bool isHorizontalOrientation(in Extents3d extents) 
        {
            double height, width;
            getHeightAndWidth(extents, out height, out width);
            return height < width;
        }
        public static Extents2d GetMaximumExtents(ref Layout lay)
        {
            // If the drawing template is imperial, we need to divide by
            // 1" in mm (25.4)
            var div = lay.PlotPaperUnits == PlotPaperUnit.Inches ? 25.4 : 1.0;

            // We need to flip the axes if the plot is rotated by 90 or 270 deg
            var doIt =
              lay.PlotRotation == PlotRotation.Degrees090 ||
              lay.PlotRotation == PlotRotation.Degrees270;

            // Get the extents in the correct units and orientation

            var min = lay.PlotPaperMargins.MinPoint/*.Swap(doIt)*/ / div;

            var max = (lay.PlotPaperSize/*.Swap(doIt)*/ - lay.PlotPaperMargins.MaxPoint/*.Swap(doIt)*/.GetAsVector()) / div;

            return new Extents2d(min, max);
        }
        static void Zoom(Point3d pMin, Point3d pMax, Point3d pCenter, double dFactor)
        {
            // Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            int nCurVport = System.Convert.ToInt32(Application.GetSystemVariable("CVPORT"));

            // Get the extents of the current space when no points 
            // or only a center point is provided
            // Check to see if Model space is current
            if (acCurDb.TileMode == true)
            {
                if (pMin.Equals(new Point3d()) == true &&
                    pMax.Equals(new Point3d()) == true)
                {
                    pMin = acCurDb.Extmin;
                    pMax = acCurDb.Extmax;
                }
            }
            else
            {
                // Check to see if Paper space is current
                if (nCurVport == 1)
                {
                    // Get the extents of Paper space
                    if (pMin.Equals(new Point3d()) == true &&
                        pMax.Equals(new Point3d()) == true)
                    {
                        pMin = acCurDb.Pextmin;
                        pMax = acCurDb.Pextmax;
                    }
                }
                else
                {
                    // Get the extents of Model space
                    if (pMin.Equals(new Point3d()) == true &&
                        pMax.Equals(new Point3d()) == true)
                    {
                        pMin = acCurDb.Extmin;
                        pMax = acCurDb.Extmax;
                    }
                }
            }

            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Get the current view
                using (ViewTableRecord acView = acDoc.Editor.GetCurrentView())
                {
                    Extents3d eExtents;

                    // Translate WCS coordinates to DCS
                    Matrix3d matWCS2DCS;
                    matWCS2DCS = Matrix3d.PlaneToWorld(acView.ViewDirection);
                    matWCS2DCS = Matrix3d.Displacement(acView.Target - Point3d.Origin) * matWCS2DCS;
                    matWCS2DCS = Matrix3d.Rotation(-acView.ViewTwist,
                                                    acView.ViewDirection,
                                                    acView.Target) * matWCS2DCS;

                    // If a center point is specified, define the min and max 
                    // point of the extents
                    // for Center and Scale modes
                    if (pCenter.DistanceTo(Point3d.Origin) != 0)
                    {
                        pMin = new Point3d(pCenter.X - (acView.Width / 2),
                                            pCenter.Y - (acView.Height / 2), 0);

                        pMax = new Point3d((acView.Width / 2) + pCenter.X,
                                            (acView.Height / 2) + pCenter.Y, 0);
                    }

                    // Create an extents object using a line
                    using (Line acLine = new Line(pMin, pMax))
                    {
                        eExtents = new Extents3d(acLine.Bounds.Value.MinPoint,
                                                    acLine.Bounds.Value.MaxPoint);
                    }

                    // Calculate the ratio between the width and height of the current view
                    double dViewRatio;
                    dViewRatio = (acView.Width / acView.Height);

                    // Tranform the extents of the view
                    matWCS2DCS = matWCS2DCS.Inverse();
                    eExtents.TransformBy(matWCS2DCS);

                    double dWidth;
                    double dHeight;
                    Point2d pNewCentPt;

                    // Check to see if a center point was provided (Center and Scale modes)
                    if (pCenter.DistanceTo(Point3d.Origin) != 0)
                    {
                        dWidth = acView.Width;
                        dHeight = acView.Height;

                        if (dFactor == 0)
                        {
                            pCenter = pCenter.TransformBy(matWCS2DCS);
                        }

                        pNewCentPt = new Point2d(pCenter.X, pCenter.Y);
                    }
                    else // Working in Window, Extents and Limits mode
                    {
                        // Calculate the new width and height of the current view
                        dWidth = eExtents.MaxPoint.X - eExtents.MinPoint.X;
                        dHeight = eExtents.MaxPoint.Y - eExtents.MinPoint.Y;

                        // Get the center of the view
                        pNewCentPt = new Point2d(((eExtents.MaxPoint.X + eExtents.MinPoint.X) * 0.5),
                                                    ((eExtents.MaxPoint.Y + eExtents.MinPoint.Y) * 0.5));
                    }

                    // Check to see if the new width fits in current window
                    if (dWidth > (dHeight * dViewRatio)) dHeight = dWidth / dViewRatio;

                    // Resize and scale the view
                    if (dFactor != 0)
                    {
                        acView.Height = dHeight * dFactor;
                        acView.Width = dWidth * dFactor;
                    }

                    // Set the center of the view
                    acView.CenterPoint = pNewCentPt;

                    // Set the current view
                    acDoc.Editor.SetCurrentView(acView);
                }

                // Commit the changes
                acTrans.Commit();
            }
        }
        static void FitContentToViewport(ref Viewport vp, in Extents3d ext, double fac = 1.0)
        {
            // Let's zoom to just larger than the extents
            var point = ext.MinPoint + ((ext.MaxPoint - ext.MinPoint) * 0.5);
            vp.ViewCenter = new Point2d(point.X, point.Y);

            // Get the dimensions of our view from the database extents

            double hgt;
            double wid;
            getHeightAndWidth(ext, out hgt, out wid);

            // We'll compare with the aspect ratio of the viewport itself
            // (which is derived from the page size)

            var aspect = vp.Width / vp.Height;

            // If our content is wider than the aspect ratio, make sure we
            // set the proposed height to be larger to accommodate the
            // content

            if (wid / hgt > aspect)
            {
                hgt = wid / aspect;
            }

            // Set the height so we're exactly at the extents

            vp.ViewHeight = hgt;

            // Set a custom scale to zoom out slightly (could also
            // vp.ViewHeight *= 1.1, for instance)
            vp.CustomScale *= fac;
        }
        static ObjectId CreateViewport(ref ObjectId layoutId, in Extents3d extents)
        {
            var db = Application.DocumentManager.MdiActiveDocument.Database;
            var ed = Application.DocumentManager.MdiActiveDocument.Editor;

            Viewport acVport;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var acLayout = tr.GetObject(layoutId, OpenMode.ForWrite) as Layout;
                if (null == acLayout)
                    return ObjectId.Null;

                double vpWidth = acLayout.PlotPaperSize.Y;
                double vpHeight = acLayout.PlotPaperSize.X;

                var layoutAsBlockTblRec = tr.GetObject(acLayout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;

                if (null == layoutAsBlockTblRec)
                    return ObjectId.Null;

                acVport = new Viewport();

                acVport.CenterPoint = new Point3d(vpWidth * 0.5, vpHeight * 0.5, 0);
                acVport.Width = vpWidth;
                acVport.Height = vpHeight;

                acVport.ColorIndex = 252;
                acVport.Layer = "Defpoints";
                acVport.LineWeight = LineWeight.LineWeight000;

                // Add the new object to the block table record and the transaction
                layoutAsBlockTblRec.AppendEntity(acVport);
                tr.AddNewlyCreatedDBObject(acVport, true);

                // Change the view direction
                acVport.ViewDirection = new Vector3d(0, 0, 1);

                // Enable the viewport
                acVport.On = true;  // CRASH from 24.07.23
                acVport.Locked = true;

                FitContentToViewport(ref acVport, extents);

                if (EqualTo(acVport.CustomScale, 1.0, 0.1))
                    acVport.StandardScale = StandardScaleType.Scale1To1; // если не задать значение то масштаб будет подогнан под объекты

                acVport.DowngradeOpen();

                using (ViewTableRecord acView = ed.GetCurrentView())
                {
                    acView.Width = vpWidth;
                    acView.Height = vpHeight;
                    acView.CenterPoint = new Point2d(acVport.CenterPoint.X, acVport.CenterPoint.Y);
                    ed.SetCurrentView(acView);
                }

                tr.Commit();
            }

            return acVport.Id;
        }
        static ObjectId CreateLayout(string newLayoutName, in Extents3d extents)
        {
            var db = Application.DocumentManager.MdiActiveDocument.Database;

            if (newLayoutName.Length < 1)
                return ObjectId.Null;

            Layout acLayout;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LayoutManager layoutMgr = LayoutManager.Current;

                DBDictionary lays = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForWrite) as DBDictionary;
                List<string> laysNames = new List<string>();

                foreach (var item in lays)
                    laysNames.Add(item.Key);

                newLayoutName = getUniqueName(newLayoutName, ref laysNames);

                // Create the new layout with default settings
                ObjectId objID = layoutMgr.CreateLayout(newLayoutName);

                if (objID == null)
                    return ObjectId.Null;

                try
                {
                    layoutMgr.CurrentLayout = newLayoutName;
                }
                catch
                {
                }

                // Open the layout
                acLayout = tr.GetObject(objID, OpenMode.ForWrite) as Layout;

                if (null == acLayout)
                    return ObjectId.Null;

                { // erease all viewPorts
                    var viewPorts = acLayout.GetViewports();
                    foreach (ObjectId obj in viewPorts)
                    {
                        var viewPort = tr.GetObject(obj, OpenMode.ForWrite) as Viewport;

                        if (viewPort != null && viewPort.GeometricExtents.MinPoint != new Point3d())
                            viewPort.Erase();
                    }
                }

                string pageSize;
                if (isHorizontalOrientation(extents))
                    pageSize = "ISO_full_bleed_A4_(210.00_x_297.00_MM)"; 
                else
                    pageSize = "ISO_full_bleed_A4_(297.00_x_210.00_MM)";

                string styleSheet = "monochrome.ctb";
                string device = "DWG To PDF.pc3";

                using (var ps = new PlotSettings(acLayout.ModelType))
                {
                    ps.CopyFrom(acLayout);

                    var psv = PlotSettingsValidator.Current;
                    // Set the device
                    var devs = psv.GetPlotDeviceList();

                    if (devs.Contains(device))
                    {
                        psv.SetPlotConfigurationName(ps, device, null);
                        psv.RefreshLists(ps);
                    }
                    // Set the media name/size
                    var mns = psv.GetCanonicalMediaNameList(ps);

                    if (mns.Contains(pageSize))
                    {
                        psv.SetCanonicalMediaName(ps, pageSize);
                    }
                    // Set the pen settings
                    var ssl = psv.GetPlotStyleSheetList();

                    if (ssl.Contains(styleSheet))
                    {
                        psv.SetCurrentStyleSheet(ps, styleSheet);
                    }
                    
                    //
                    psv.SetPlotPaperUnits(ps, PlotPaperUnit.Millimeters);
                    psv.SetStdScaleType(ps, StdScaleType.StdScale1To1);
                    psv.SetUseStandardScale(ps, true);
                    psv.SetPlotWindowArea(ps, new Extents2d(extents.MinPoint.X, extents.MinPoint.Y, extents.MaxPoint.X, extents.MaxPoint.Y)); // проверить
                    //

                    // Copy the PlotSettings data back to the Layout
                     var upgraded = false;

                    if (!acLayout.IsWriteEnabled)
                    {
                        acLayout.UpgradeOpen();
                        upgraded = true;
                    }
                    acLayout.CopyFrom(ps);

                     if (upgraded)
                    {
                        acLayout.DowngradeOpen();
                    }
                } // using (var ps = new PlotSettings(acLayout.ModelType))

                tr.Commit();
            }
            return acLayout.Id;
        }

        [CommandMethod("CadTools_Help")]
        public void Help()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            ed.WriteMessage("\nДоступные команды:\n/CadTools_VHLines\nCadTools_CreateLayersFromTXT\n");
        }

        [CommandMethod("CadTools_VHLines")]
        public void VHLines()
        {
            double Vlength = 0;
            double Hlength = 0;
            int VcolorIndex = 1;
            int HcolorIndex = 2;
            string Vname = "! Вертикальные отрезки";
            string Hname = "! Горизонтальные отрезки";

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptSelectionOptions selOptions = new PromptSelectionOptions();
            selOptions.MessageForAdding = "\nВыберите полилинии";
            var sel = ed.GetSelection();

            if (sel.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\n*** Вы выбрали что-то не то! ***\n");
                return;
            }

            ObjectId[] idsPoly = sel.Value.GetObjectIds();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord ms = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                LayerTable acLayerTable = tr.GetObject(db.LayerTableId, OpenMode.ForWrite) as LayerTable;

                LayerTableRecord VLayerTableRecord = new LayerTableRecord();
                LayerTableRecord HLayerTableRecord = new LayerTableRecord();

                if (!acLayerTable.Has(Vname) && !acLayerTable.Has(Hname))
                {
                    VLayerTableRecord.Name = Vname;
                    HLayerTableRecord.Name = Hname;

                    acLayerTable.Add(VLayerTableRecord);
                    acLayerTable.Add(HLayerTableRecord);

                    tr.AddNewlyCreatedDBObject(VLayerTableRecord, true);
                    tr.AddNewlyCreatedDBObject(HLayerTableRecord, true);
                }

                foreach (ObjectId id in idsPoly)
                {
                    Polyline acPoly;
                    Line acLine;
                    try
                    {
                        acPoly = (Polyline)tr.GetObject(id, OpenMode.ForRead);
                    }
                    catch
                    {
                        try
                        {
                            acLine = (Line)tr.GetObject(id, OpenMode.ForRead);
                        }
                        catch
                        {
                            continue;
                        }

                        Line line = new Line();

                        line.StartPoint = acLine.StartPoint;
                        line.EndPoint = acLine.EndPoint;

                        double degAngle = line.Angle * (180 / Math.PI);

                        ed.WriteMessage($"\n{line.Angle} = {degAngle}");
                        if ((degAngle > 45 && degAngle < 135) || (degAngle > 225 && degAngle < 315)) // вертикальные
                        {
                            line.Layer = Vname;
                            line.ColorIndex = VcolorIndex;
                            Vlength += line.Length;
                        }
                        else
                        {
                            line.Layer = Hname;
                            line.ColorIndex = HcolorIndex;
                            Hlength += line.Length;
                        }

                        line.LineWeight = LineWeight.LineWeight030;

                        ms.AppendEntity(line);
                        tr.AddNewlyCreatedDBObject(line, true);

                        continue;
                    }

                    Point2dCollection points = new Point2dCollection();

                    for (int i = 0; i < acPoly.NumberOfVertices; i++)
                        points.Add(acPoly.GetPoint2dAt(i));

                    for (int i = 0; i < points.Count; i++)
                    {
                        Line line = new Line();


                        if ((i == (points.Count - 1)))
                        {
                            if (acPoly.Closed)
                            {
                                line.StartPoint = new Point3d(points[i].X, points[i].Y, 0);
                                line.EndPoint = new Point3d(points[0].X, points[0].Y, 0);
                            }
                        }
                        else
                        {
                            line.StartPoint = new Point3d(points[i].X, points[i].Y, 0);
                            line.EndPoint = new Point3d(points[i + 1].X, points[i + 1].Y, 0);
                        }

                        double degAngle = line.Angle * (180 / Math.PI);

                        ed.WriteMessage($"\n{line.Angle} = {degAngle}");
                        if ((degAngle > 45 && degAngle < 135) || (degAngle > 225 && degAngle < 315)) // вертикальные
                        {
                            line.Layer = Vname;
                            line.ColorIndex = VcolorIndex;
                            Vlength += line.Length;
                        }
                        else
                        {
                            line.Layer = Hname;
                            line.ColorIndex = HcolorIndex;
                            Hlength += line.Length;
                        }

                        line.LineWeight = LineWeight.LineWeight030;

                        ms.AppendEntity(line);
                        tr.AddNewlyCreatedDBObject(line, true);
                    }
                }
                tr.Commit();
            }
            ed.WriteMessage($"\n\nСуммарная длина:\nвертикальных отрезков:\t{Vlength}\nгоризонтальных отрезков:\t{Hlength}"); ;
        }

        [CommandMethod("CadTools_CreateLayersFromTXT")]
        public void CreateLayersFromTxt()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            string path = "";
            string[] layerNames;
            try
            {
                PromptOpenFileOptions openFileOpt = new PromptOpenFileOptions("Файл со списком имён слоёв");
                openFileOpt.Filter = "Текстовый файл (*.txt)|*.txt";

                path = ed.GetFileNameForOpen(openFileOpt).StringResult;

                layerNames = File.ReadAllLines(path);

                if (layerNames.Length < 1)
                    throw new System.Exception();
            }
            catch
            {
                ed.WriteMessage("\nЧто-то пошло не так! Попробуй снова...");
                return;
            }

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord ms = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                LayerTable acLayerTable = tr.GetObject(db.LayerTableId, OpenMode.ForWrite) as LayerTable;
                int counter = 0;
                foreach (var layerName in layerNames)
                {
                    if (!acLayerTable.Has(layerName))
                    {
                        try
                        {
                            LayerTableRecord layerTableRecord = new LayerTableRecord();
                            layerTableRecord.Name = layerName;
                            acLayerTable.Add(layerTableRecord);
                            tr.AddNewlyCreatedDBObject(layerTableRecord, true);
                            counter++;
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }
                tr.Commit();
                ed.WriteMessage($"Успешно создано {counter} слоёв");
            }
        }

        static uint g_ListIterator = 0;
        static uint g_TitleIterator = 0;
        static uint g_ListsCount = 0;

        public class ShtampParams
        {
            public enum Shtamp
            {
                Shtamp1,
                Shtamp2,
                Shtamp3,
                bezShtampa
            }
            public enum Edge
            {
                Left,
                Right
            }

            public Shtamp dynShtamp = Shtamp.bezShtampa;
            public bool dynIsNumerationOn = true;
            public bool dynIsMirrorNum = false;
            public Edge dynEdge = Edge.Left;
            public Extents3d extents = new Extents3d();

            public string attrCompany = ""; // КОМПАНИЯ_1, КОМПАНИЯ_2, КОМПАНИЯ_3
            public string attrSheet = ""; // ЛИСТ_1, ЛИСТ_2, ЛИСТ_3
            public string attrSheetsCount = ""; // ЛИСТОВ_1
            public string attrDescription = ""; // НАИМЕНОВАНИЕ_1, НАИМЕНОВАНИЕ_3, НАИМЕНОВАНИЕ_4
            public string attrDescription2 = ""; // НАИМЕНОВАНИЕ_2,
            public string attrNumber = ""; // НОМЕР
            public string attrStage = ""; // СТАДИЯ_1
            public string attrLegend = ""; // УСЛОВНЫЕ_ОБОЗНАЧЕНИЯ_1
            public string attrAuthorPost = ""; // СПЕЦ_1
            public string attrAuthorPost2 = ""; // СПЕЦ_2
            public string attrAuthorPost3 = ""; // СПЕЦ_3
            public string attrAuthorPost4 = ""; // СПЕЦ_4
            public string attrAuthorPost5 = ""; // СПЕЦ_5
            public string attrAuthor = ""; // ФАМИЛИЯ_1
            public string attrAuthor2 = ""; // ФАМИЛИЯ_2
            public string attrAuthor3 = ""; // ФАМИЛИЯ_3
            public string attrAuthor4 = ""; // ФАМИЛИЯ_4
            public string attrAuthor5 = ""; // ФАМИЛИЯ_5

        }
        public class Zone
        {
            public Point3dCollection points;
            public ShtampParams shtampParams = new ShtampParams();
        }
        public static Point3d getCenterOfExtents(in Extents3d extents)
        {
            if (extents == new Extents3d())
                return new Point3d();

            double cX = (extents.MinPoint.X + extents.MaxPoint.X) * 0.5;
            double cY = (extents.MinPoint.Y + extents.MaxPoint.Y) * 0.5;
            double cZ = (extents.MinPoint.Z + extents.MaxPoint.Z) * 0.5;

            return new Point3d(cX, cY, cZ);
        }
        public static string getUniqueName(string text, ref List<string> uniqueTexts)
        {
            int i = 2;
            while (uniqueTexts.Contains(text))
            {
                Regex regex = new Regex(@" \(\d+\)");
                text = regex.Replace(text, "");
                text = $"{text} ({i++})";
            }
            if (!uniqueTexts.Contains(text))

                uniqueTexts.Add(text);

            return text;
        }
        Dictionary<ItemLocator, Zone> getZonesFromSelectionByArray()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            Dictionary<ItemLocator, Zone> zones = new Dictionary<ItemLocator, Zone>();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
                BlockTableRecord modelSpace = tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                LayoutManager layoutMgr = LayoutManager.Current;

                if (modelSpace == null)
                    return new Dictionary<ItemLocator, Zone>();

                var options = new PromptSelectionOptions();
                options.SingleOnly = true;
                options.SinglePickInSpace = true;
                options.MessageForAdding = "Выберите массив контуров чертежей\n(контур должен состоять из полилинии)";

                // var filter = new SelectionFilter(new TypedValue[] { new TypedValue(0, "ASSOCARRAY") });

                var result = ed.GetSelection(options/*, filter*/);

                if (result.Status != PromptStatus.OK)
                    return new Dictionary<ItemLocator, Zone>();

                foreach (ObjectId id in result.Value.GetObjectIds())
                {
                    if (!AssocArray.IsAssociativeArray(id))
                    {
                        ed.WriteMessage("Выбран не массив! А какая-то хуета!"); // debug
                        return new Dictionary<ItemLocator, Zone>();
                    }

                    AssocArray assocArray = AssocArray.GetAssociativeArray(id);

                    foreach (ObjectId itemId in assocArray.SourceEntities)
                    {
                        var poly = tr.GetObject(itemId, OpenMode.ForRead) as Polyline;

                        if (!poly.Closed)
                            continue;

                        var blockRef = tr.GetObject(id, OpenMode.ForRead) as BlockReference;

                        if (blockRef == null)
                            continue;

                        foreach (var locator in assocArray.getItems(true))
                        {
                            Point3dCollection vertexes = new Point3dCollection();

                            var transform = assocArray.GetItemTransform(locator);

                            for (int i = 0; i < poly.NumberOfVertices; i++)
                            {
                                double dx = 0, 
                                       dy = 0, 
                                       delta = 0.0; // смещение полилинии (увеличение\меньшение контура)

                                if (i == 0)
                                {
                                    dx = -delta;
                                    dy = delta;
                                }
                                else if (i == 1)
                                {
                                    dx = delta;
                                    dy = delta;
                                }
                                else if (i == 2)
                                {
                                    dx = delta;
                                    dy = -delta;
                                }
                                else if (i == 3)
                                {
                                    dx = -delta;
                                    dy = -delta;
                                }

                                var point = poly.GetPoint2dAt(i);
                                var offsetPoint = new Point3d(point.X + blockRef.Position.X + dx, point.Y + blockRef.Position.Y + dy, 0);
                                vertexes.Add(offsetPoint.TransformBy(transform));
                            }

                            if (vertexes.Count > 2)
                            {
                                Zone zone = new Zone();
                                zone.points = vertexes;
                                zones[locator] = zone;
                            }
                        }
                    }
                }
                tr.Commit();
            }
            return zones.OrderBy(x => x.Key.RowIndex).ToDictionary(x => x.Key, x => x.Value);
        }
        public class PointsAndCenter
        {
            public Point3dCollection points = new Point3dCollection();
            public Point2d center = new Point2d(0, 0);
        }
        public class Point2dComparer : IComparer<PointsAndCenter>
        {
            public int Compare(PointsAndCenter p1, PointsAndCenter p2)   
            {
#if PODGON                
                int tempRes = CompareTo(p1.center.Y, p2.center.Y, 50) * (-1); // * (-1) т.к. в СК Y растёт снизу вверх, а чертежи сортируем сверху вниз

                if (tempRes == 0)
                    return CompareTo(p1.center.X, p2.center.X, 1/*g_epsilon*/);
                else
                    return tempRes;
#else
                int tempRes = p1.center.Y.CompareTo(p2.center.Y) * (-1);

                if (tempRes == 0)
                    return p1.center.X.CompareTo(p2.center.X);
                else
                    return tempRes;
#endif
            }
        }
        enum TypeOfSelection
        {
            Null,
            DynArray,
            Polyline
        }
        TypeOfSelection isDynArrayOrPolyByKeywordSelection()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var options = new PromptKeywordOptions("Чем заданы границы печати?");

                const string arrStr = "Массивом";
                const string polyStr = "Полилиниями";

                options.Keywords.Add(arrStr);
                options.Keywords.Add(polyStr);
                options.Keywords.Default = arrStr;

                var result = ed.GetKeywords(options);

                if (result.Status == PromptStatus.Keyword || result.Status == PromptStatus.OK)
                {
                    if (result.StringResult == arrStr)
                        return TypeOfSelection.DynArray;
                    else if (result.StringResult == polyStr)
                        return TypeOfSelection.Polyline;
                    else
                        return TypeOfSelection.Null;
                }
                else
                    return TypeOfSelection.Null;
            }
        }
        Dictionary<ItemLocator, Zone> getZonesFromSelectionByPolyline()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            Dictionary<ItemLocator, Zone> zones = new Dictionary<ItemLocator, Zone>();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
                BlockTableRecord modelSpace = tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                LayoutManager layoutMgr = LayoutManager.Current;

                if (modelSpace == null)
                    return new Dictionary<ItemLocator, Zone>();

                var options = new PromptSelectionOptions();
                options.SingleOnly = false;
                options.SinglePickInSpace = false;
                options.MessageForAdding = "\nВыберите замкнутые полилинии формата А4";

                var filter = new SelectionFilter(new TypedValue[] { new TypedValue(0, "LWPOLYLINE") });

                var result = ed.GetSelection(options, filter);

                if (result.Status != PromptStatus.OK)
                    return new Dictionary<ItemLocator, Zone>();

                int column = 0, row = 0;

                double centerY = 0;

                List<PointsAndCenter> points = new List<PointsAndCenter>();

                foreach (ObjectId id in result.Value.GetObjectIds())
                {
                    Polyline poly = tr.GetObject(id, OpenMode.ForWrite) as Polyline;

                    if (poly == null)
                    {
                        //ed.WriteMessage("Выбраны не полилинии! А какая-то хуета!"); // debug
                        continue;
                    }

                    if (!poly.Closed || poly.NumberOfVertices != 4)
                        continue;

                    Point3dCollection vertexes = new Point3dCollection();
                    Point3d tempCenter = new Point3d();
                    //double tempCenterX = 0, tempCenterY = 0;
                    //g_epsilon
                    for (int i = 0; i < poly.NumberOfVertices; i++)
                    {
                        var point = poly.GetPoint3dAt(i);
                        vertexes.Add(point);
                        tempCenter = point;
                    }

                    var tempCenterVec = tempCenter.GetAsVector() / 4;                    

                    PointsAndCenter pointsAndCenter = new PointsAndCenter();
                    pointsAndCenter.center = new Point2d(tempCenterVec.X, tempCenterVec.Y);
                    pointsAndCenter.points = vertexes;

                    points.Add(pointsAndCenter);
                }

                points.Sort(new Point2dComparer());

                foreach (var pc in points)
                {
                    if (!pc.center.Y.Equals(centerY))
                    {
                        column = 0;
                        row++;
                    }

                    centerY = pc.center.Y;

                    ItemLocator locator = new ItemLocator();
                    locator.ItemIndex = column;
                    locator.RowIndex = row;
                    locator.LevelIndex = 0;

                    Zone zone = new Zone();
                    zone.points = pc.points;
                    zones[locator] = zone;

                    column++;
                }
                tr.Commit();
            }
            return zones;//.OrderBy(x => x.Key.RowIndex).ToDictionary(x => x.Key, x => x.Value);
        }
        public static bool fillParamsFromBlockRef(ref ShtampParams shtampParams, ObjectId objId, bool renumerateLists = false)
        {
            var db = Application.DocumentManager.MdiActiveDocument.Database;

            if (AssocArray.IsAssociativeArray(objId))
                return false;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var selBlockRef = tr.GetObject(objId, OpenMode.ForWrite) as BlockReference;

                if (selBlockRef == null)
                    return false;

                var selBlockTblRec = tr.GetObject(selBlockRef.DynamicBlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;

                if (selBlockTblRec == null || selBlockTblRec.Name != "SG_Форма")
                    return false;

                shtampParams.extents = selBlockRef.GeometricExtents;

                foreach (DynamicBlockReferenceProperty prop in selBlockRef.DynamicBlockReferencePropertyCollection)
                {
                    if (prop.PropertyName == "Штамп")
                    {
                        if (prop.Value.ToString() == "Штамп 1")
                            shtampParams.dynShtamp = ShtampParams.Shtamp.Shtamp1;
                        else if (prop.Value.ToString() == "Штамп 2")
                            shtampParams.dynShtamp = ShtampParams.Shtamp.Shtamp2;
                        else if (prop.Value.ToString() == "Штамп 3")
                            shtampParams.dynShtamp = ShtampParams.Shtamp.Shtamp3;
                        else if (prop.Value.ToString() == "Без штампа")
                            shtampParams.dynShtamp = ShtampParams.Shtamp.bezShtampa;
                    }
                    else if (prop.PropertyName == "Отступ") // int 5 or 20
                    {

                    }
                    else if (prop.PropertyName == "Нумерация")
                    {
                        shtampParams.dynIsNumerationOn = Convert.ToBoolean(prop.Value);
                    }
                    else if (prop.PropertyName == "Отражение нумерации")
                    {
                        shtampParams.dynIsMirrorNum = Convert.ToBoolean(prop.Value);
                    }
                    else if (prop.PropertyName == "Выбор стороны края")
                    {
                        if (prop.Value.ToString() == "Слева")
                            shtampParams.dynEdge = ShtampParams.Edge.Left;
                        else if (prop.Value.ToString() == "Справа")
                            shtampParams.dynEdge = ShtampParams.Edge.Right;
                    }

                }

                foreach (ObjectId attrId in selBlockRef.AttributeCollection)
                {
                    var attrDef = attrId.GetObject(OpenMode.ForWrite) as AttributeReference;

                    if (attrDef != null && !attrDef.IsConstant)
                    {

                        if (shtampParams.dynShtamp == ShtampParams.Shtamp.Shtamp1)
                        {
                            if (attrDef.Tag == "КОМПАНИЯ_1")
                            {
                                shtampParams.attrCompany = attrDef.TextString;
                            }
                            else if (attrDef.Tag == "НАИМЕНОВАНИЕ_1")
                            {
                                shtampParams.attrDescription = attrDef.TextString;
                            }
                            else if (attrDef.Tag == "НАИМЕНОВАНИЕ_2")
                            {
                                shtampParams.attrDescription2 = attrDef.TextString;
                            }
                            else if (attrDef.Tag == "ЛИСТ_1")
                            {
                                if (renumerateLists)
                                    attrDef.TextString = g_ListIterator.ToString();

                                shtampParams.attrSheet = attrDef.TextString;
                            }
                        }
                        else if (shtampParams.dynShtamp == ShtampParams.Shtamp.Shtamp2)
                        {
                            if (attrDef.Tag == "КОМПАНИЯ_2")
                            {
                                shtampParams.attrCompany = attrDef.TextString;
                            }
                            else if (attrDef.Tag == "НАИМЕНОВАНИЕ_3")
                            {
                                shtampParams.attrDescription = attrDef.TextString;
                            }
                            else if (attrDef.Tag == "ЛИСТОВ_1")
                            {
                                if (renumerateLists && g_ListsCount > 0)
                                    attrDef.TextString = g_ListsCount.ToString();

                                shtampParams.attrSheetsCount = attrDef.TextString;
                            }
                            else if (attrDef.Tag == "ЛИСТ_2")
                            {
                                if (renumerateLists)
                                    attrDef.TextString = g_ListIterator.ToString();

                                shtampParams.attrSheet = attrDef.TextString;
                            }
                            else if (attrDef.Tag == "СТАДИЯ_1")
                            {
                                shtampParams.attrStage = attrDef.TextString;
                            }
                            else if (attrDef.Tag == "УСЛОВНЫЕ_ОБОЗНАЧЕНИЯ_1")
                            {
                                shtampParams.attrLegend = attrDef.TextString;
                            }
                            else if (attrDef.Tag == "СПЕЦ_1")
                            {
                                shtampParams.attrAuthorPost = attrDef.TextString;
                            }
                            else if (attrDef.Tag == "СПЕЦ_2")
                            {
                                shtampParams.attrAuthorPost2 = attrDef.TextString;
                            }
                            else if (attrDef.Tag == "СПЕЦ_3")
                            {
                                shtampParams.attrAuthorPost3 = attrDef.TextString;
                            }
                            else if (attrDef.Tag == "СПЕЦ_4")
                            {
                                shtampParams.attrAuthorPost4 = attrDef.TextString;
                            }
                            else if (attrDef.Tag == "СПЕЦ_5")
                            {
                                shtampParams.attrAuthorPost5 = attrDef.TextString;
                            }
                            else if (attrDef.Tag == "ФАМИЛИЯ_1")
                            {
                                shtampParams.attrAuthor = attrDef.TextString;
                            }
                            else if (attrDef.Tag == "ФАМИЛИЯ_2")
                            {
                                shtampParams.attrAuthor2 = attrDef.TextString;
                            }
                            else if (attrDef.Tag == "ФАМИЛИЯ_3")
                            {
                                shtampParams.attrAuthor3 = attrDef.TextString;
                            }
                            else if (attrDef.Tag == "ФАМИЛИЯ_4")
                            {
                                shtampParams.attrAuthor4 = attrDef.TextString;
                            }
                            else if (attrDef.Tag == "ФАМИЛИЯ_5")
                            {
                                shtampParams.attrAuthor5 = attrDef.TextString;
                            }
                        }
                        else if (shtampParams.dynShtamp == ShtampParams.Shtamp.Shtamp3)
                        {
                            if (attrDef.Tag == "КОМПАНИЯ_3")
                            {
                                shtampParams.attrCompany = attrDef.TextString;
                            }
                            else if (attrDef.Tag == "НАИМЕНОВАНИЕ_4")
                            {
                                shtampParams.attrDescription = attrDef.TextString;
                            }
                            else if (attrDef.Tag == "ЛИСТ_3")
                            {
                                if (renumerateLists)
                                    attrDef.TextString = g_ListIterator.ToString();

                                shtampParams.attrSheet = attrDef.TextString;
                            }
                        }

                        if (attrDef.Tag == "НОМЕР")
                        {
                            if (shtampParams.dynIsNumerationOn)
                                shtampParams.attrNumber = attrDef.TextString;
                        }
                    }
                }

                if (shtampParams.dynShtamp == ShtampParams.Shtamp.bezShtampa)
                {
                    shtampParams.attrSheet = "ТИТУЛ";
                    g_TitleIterator++;
                }
                else
                    g_ListIterator++;

                tr.Commit();
                return true;
            }
        }
        public static bool fillParamsFromSome(ref ShtampParams shtampParams, ref ObjectId[] objIds)
        {
            var db = Application.DocumentManager.MdiActiveDocument.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                Extents3d sheetTextExtents = new Extents3d();
                List<Tuple<Extents3d, string>> findedNumsStr = new List<Tuple<Extents3d, string>>();

                foreach (ObjectId objId in objIds)
                {
                    Regex sheetStrRegex = new Regex(@"^\s*[Лл]ист\s*$"); // пофиксить регулярку, попадает в выборку "листОВ", а нужны только "лист"
                    Regex sheetValueRegex = new Regex(@"^\s*\d+\s*$");

                    {
                        var dbText = tr.GetObject(objId, OpenMode.ForWrite) as DBText;

                        if (dbText != null)
                        {
                            var textStr = dbText.TextString.Replace(" ", "");

                            if (sheetStrRegex.IsMatch(textStr))
                                sheetTextExtents = dbText.GeometricExtents;
                            else if (sheetValueRegex.IsMatch(textStr))
                                findedNumsStr.Add(new Tuple<Extents3d, string>(dbText.GeometricExtents, textStr));

                            continue;
                        }
                    }

                    var mText = tr.GetObject(objId, OpenMode.ForWrite) as MText;

                    if (mText != null)
                    {
                        var textStr = mText.Text.Replace(" ", "");

                        if (sheetStrRegex.IsMatch(textStr))
                            sheetTextExtents = mText.GeometricExtents;
                        else if (sheetValueRegex.IsMatch(textStr))
                            findedNumsStr.Add(new Tuple<Extents3d, string>(mText.GeometricExtents, textStr));

                        continue;
                    }
                }

                Point3d sheetTextPoint = getCenterOfExtents(sheetTextExtents);

                double minDistance = double.MaxValue;
                string valueStr = "";

                foreach (var tuple in findedNumsStr)
                {
                    var centerValue = getCenterOfExtents(tuple.Item1);

                    if (centerValue == new Point3d())
                        continue;

                    var distance = new Point2d(centerValue.X, centerValue.Y).GetDistanceTo(new Point2d(sheetTextPoint.X, sheetTextPoint.Y));

                    if (distance < minDistance)
                    {
                        minDistance = distance;
                        valueStr = tuple.Item2;
                    }
                }

                if (valueStr != "")
                {
                    shtampParams.attrSheet = valueStr;
                    g_ListIterator++;
                }
                else
                {
                    shtampParams.attrSheet = "ТИТУЛ";
                    g_TitleIterator++;
                }

                tr.Commit();
                return true;
            }
        }
        public static bool createDBText(string TextStr, double textHeight, Point3d position)
        {
            var db = Application.DocumentManager.MdiActiveDocument.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
                BlockTableRecord modelSpace = tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                DBText text = new DBText();
                //text.SetDatabaseDefaults();

                text.ColorIndex = 1;
                text.Height = textHeight;
                text.HorizontalMode = TextHorizontalMode.TextMid;
                text.VerticalMode = TextVerticalMode.TextVerticalMid;
                text.TextString = TextStr;
                text.Position = Point3d.Origin;
                text.AlignmentPoint = position;

                modelSpace.AppendEntity(text);
                tr.AddNewlyCreatedDBObject(text, true);

                tr.Commit();
            }
            return true;
        }
        public static Point3d getCenterOfPoints(Point3dCollection points)
        {
            Vector3d sum = new Vector3d();

            foreach (Point3d point in points)
            {
                sum += (point.GetAsVector());
            }
            return new Point3d() + (sum / points.Count);
        }
        public static void SetModelPaperSpace(bool paperSpace = true)
        {
            try
            {
                if (paperSpace)
                {
                    if (Application.GetSystemVariable("TILEMODE").ToString() == "1")
                        Application.SetSystemVariable("TILEMODE", 0);
                    Application.DocumentManager.MdiActiveDocument.Editor.SwitchToPaperSpace();
                }
                else
                {
                    if (Application.GetSystemVariable("TILEMODE").ToString() != "1")
                        Application.SetSystemVariable("TILEMODE", 1);
                    Application.DocumentManager.MdiActiveDocument.Editor.SwitchToModelSpace();
                }
            }
            catch
            {
            }
        }

        [CommandMethod("CadTools_Test2", CommandFlags.Modal)]
        public void Test2()
        {
            Dictionary<ItemLocator, Zone> zones = new Dictionary<ItemLocator, Zone>();

            var typeOfSel = isDynArrayOrPolyByKeywordSelection();
            if (typeOfSel == TypeOfSelection.DynArray)
            {
                zones = getZonesFromSelectionByArray();
            }
            else if (typeOfSel == TypeOfSelection.Polyline)
            {
                zones = getZonesFromSelectionByPolyline();
            }
            else
                return;

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            bool renumerateLists = false;
            if (typeOfSel == TypeOfSelection.DynArray)
            {
                PromptKeywordOptions options = new PromptKeywordOptions("");
                options.Message = "Пронумеровать листы? ";
                options.Keywords.Add("ДА");
                options.Keywords.Add("НЕТ");
                options.Keywords.Default = "НЕТ";

                var result = ed.GetKeywords(options);

                if (result.StringResult == "ДА")
                    renumerateLists = true;
            }


            string curLayoutName = null;
            ViewTableRecord curView = null;
            bool isCurViewSaved = false;
            LayoutManager layoutMgr = LayoutManager.Current;

            if (!isCurViewSaved)
            {
                curLayoutName = layoutMgr.CurrentLayout;
                curView = ed.GetCurrentView();
                isCurViewSaved = true;
            }
#if DEBUG
            {
                int i = 0;
                foreach (var locator in zones.Keys)
                {
                    var zone = zones[locator];

                    createDBText($"i: {i++}, row: {locator.RowIndex}, col: {locator.ItemIndex}", 5, getCenterOfPoints(zone.points));
                }
            }
#endif

            if (renumerateLists)
            {
                g_ListIterator = 1;
                g_TitleIterator = 1;
            }
            else
            {
                g_ListIterator = 0;
                g_TitleIterator = 0;
            }

            SetModelPaperSpace(false);

            // lists count
            foreach (var locator in zones.Keys)
            {
                var zone = zones[locator];

                var polySelectionResult = ed.SelectCrossingPolygon(zone.points);

                if (polySelectionResult.Status != PromptStatus.OK)
                {
                    ed.WriteMessage($"R:{locator.RowIndex}, I:{locator.ItemIndex}; EMPTY AREA\n");
                    continue;
                }

                if (typeOfSel == TypeOfSelection.DynArray)
                {
                    var selResObjIds = polySelectionResult.Value.GetObjectIds();

                    foreach (ObjectId localObjId in selResObjIds)
                    {
                        if (!fillParamsFromBlockRef(ref zone.shtampParams, localObjId, renumerateLists))
                            continue;
                    }
                }
                else if (typeOfSel == TypeOfSelection.Polyline)
                {
                    foreach (Point3d point in zone.points)
                        zone.shtampParams.extents.AddPoint(new Point3d(point.X, point.Y, 0));

                    var objs = polySelectionResult.Value.GetObjectIds();

                    if (!fillParamsFromSome(ref zone.shtampParams, ref objs))
                        continue;
                }

                zone.shtampParams = new ShtampParams();
            }

            if (renumerateLists)
            {
                g_ListsCount = g_ListIterator - 1;
                g_ListIterator = 1;
                g_TitleIterator = 1;
            }
            else
            {
                g_ListsCount = g_ListIterator + g_TitleIterator;
                g_ListIterator = 0;
                g_TitleIterator = 0;
            }


            foreach (var locator in zones.Keys)
            {
                var zone = zones[locator];

                var polySelectionResult = ed.SelectCrossingPolygon(zone.points);

                if (polySelectionResult.Status != PromptStatus.OK)
                {
                    ed.WriteMessage($"R:{locator.RowIndex}, I:{locator.ItemIndex}; EMPTY AREA\n");
                    continue;
                }


                if (typeOfSel == TypeOfSelection.DynArray)
                {
                    var selResObjIds = polySelectionResult.Value.GetObjectIds();

                    foreach (ObjectId localObjId in selResObjIds)
                    {
                        if (!fillParamsFromBlockRef(ref zone.shtampParams, localObjId, renumerateLists))
                            continue;
                    }
                }
                else if (typeOfSel == TypeOfSelection.Polyline)
                {
                    zone.shtampParams.extents = new Extents3d();

                    foreach (Point3d point in zone.points)
                    {
                        zone.shtampParams.extents.AddPoint(new Point3d(point.X, point.Y, 0));
                    }

                    var objs = polySelectionResult.Value.GetObjectIds();

                    if (!fillParamsFromSome(ref zone.shtampParams, ref objs))
                        continue;
                }
            }

            SetModelPaperSpace(true);

            var temp_saveTimeVar = Application.GetSystemVariable("SAVETIME");
            Application.SetSystemVariable("SAVETIME", 0);
            var temp_regenModeVar = Application.GetSystemVariable("REGENMODE");
            Application.SetSystemVariable("REGENMODE", 0);
            var temp_fieldevalVar = Application.GetSystemVariable("FIELDEVAL");
            Application.SetSystemVariable("FIELDEVAL", 0);
            var temp_layoutRegenVar = Application.GetSystemVariable("LAYOUTREGENCTL");
            Application.SetSystemVariable("LAYOUTREGENCTL", 1);
            var temp_layoutCreateViewportVar = Application.GetSystemVariable("LAYOUTCREATEVIEWPORT");
            Application.SetSystemVariable("LAYOUTCREATEVIEWPORT", 0);

            uint s = g_ListsCount;
            uint count = g_ListIterator + g_TitleIterator;

            if (renumerateLists)
                count -= 2;

            uint listIter = 1;

            foreach (var locator in zones.Keys) // new
            {
                var zone = zones[locator];
                var layoutId = CreateLayout(zone.shtampParams.attrSheet, zone.shtampParams.extents);

                if (layoutId.IsNull)
                {
                    //ed.WriteMessage($"НЕ УДАЛОСЬ СОЗДАТЬ Layout {zone.shtampParams.attrSheet} в позиции Row:{ locator.RowIndex}, Col: { locator.ItemIndex}\n");
                    continue;
                }
                else
                {
                    var viewportId = CreateViewport(ref layoutId, zone.shtampParams.extents);
                    if (viewportId.IsNull)
                    {
                        //ed.WriteMessage($"НЕ УДАЛОСЬ СОЗДАТЬ Viewport {zone.shtampParams.attrSheet} в позиции Row:{ locator.RowIndex}, Col: { locator.ItemIndex}\n");
                        continue;
                    }
                }

                //ed.WriteMessage($"\nСоздано {listIter++} листов из {count}, осталось {count - listIter + 1}\n");
            }

            SetModelPaperSpace(false);

            Application.SetSystemVariable("SAVETIME", temp_saveTimeVar);
            Application.SetSystemVariable("REGENMODE", temp_regenModeVar);
            Application.SetSystemVariable("FIELDEVAL", temp_fieldevalVar);
            Application.SetSystemVariable("LAYOUTREGENCTL", temp_layoutRegenVar);
            Application.SetSystemVariable("LAYOUTCREATEVIEWPORT", temp_layoutCreateViewportVar);

            ed.Regen();
        }

        [CommandMethod("CadTools_Test")]
        public void Test()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord ms = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                tr.Commit();
                ed.WriteMessage($"Hello world {1} !");
            }
        }

#else

        /// <summary>
        /// Reverses the order of the X and Y properties of a Point2d.
        /// </summary>
        /// <param name="flip">Boolean indicating whether to reverse or not.</param>
        /// <returns>The original Point2d or the reversed version.</returns>

        public static Point2d Swap(ref Point2d pt, bool flip = true)
        {
            return flip ? new Point2d(pt.Y, pt.X) : pt;
        }


        /// <summary>
        /// Pads a Point2d with a zero Z value, returning a Point3d.
        /// </summary>
        /// <param name="pt">The Point2d to pad.</param>
        /// <returns>The padded Point3d.</returns>

        public static Point3d Pad(ref Point2d pt)
        {
            return new Point3d(pt.X, pt.Y, 0);
        }

        /// <summary>
        /// Strips a Point3d down to a Point2d by simply ignoring the Z ordinate.
        /// </summary>
        /// <param name="pt">The Point3d to strip.</param>
        /// <returns>The stripped Point2d.</returns>
        public static Point2d Strip(ref Point3d pt)
        {
            return new Point2d(pt.X, pt.Y);
        }

        /// <summary>
        /// Creates a layout with the specified name and optionally makes it current.
        /// </summary>
        /// <param name="name">The name of the viewport.</param>
        /// <param name="select">Whether to select it.</param>
        /// <returns>The ObjectId of the newly created viewport.</returns>
        public static ObjectId CreateAndMakeLayoutCurrent(ref LayoutManager lm, string name, bool select = true)
        {
            // First try to get the layout
            var id = lm.GetLayoutId(name);

            // If it doesn't exist, we create it
            if (!id.IsValid)
            {
                id = lm.CreateLayout(name);
            }

            // And finally we select it
            if (select)
            {
                lm.CurrentLayout = name;
            }
            return id;
        }

        /// <summary>
        /// Applies an action to the specified viewport from this layout.
        /// Creates a new viewport if none is found withthat number.
        /// </summary>
        /// <param name="tr">The transaction to use to open the viewports.</param>
        /// <param name="vpNum">The number of the target viewport.</param>
        /// <param name="f">The action to apply to each of the viewports.</param>
        public static void ApplyToViewport(ref Layout lay, Transaction tr, int vpNum, Action<Viewport> f)
        {
            var vpIds = lay.GetViewports();
            Viewport vp = null;

            foreach (ObjectId vpId in vpIds)
            {
                var vp2 = tr.GetObject(vpId, OpenMode.ForWrite) as Viewport;

                if (vp2 != null && vp2.Number == vpNum)
                {
                    // We have found our viewport, so call the action
                    vp = vp2;
                    break;
                }
            }

            if (vp == null)
            {
                // We have not found our viewport, so create one
                var btr = (BlockTableRecord)tr.GetObject(lay.BlockTableRecordId, OpenMode.ForWrite);
                vp = new Viewport();

                // Add it to the database
                btr.AppendEntity(vp);
                tr.AddNewlyCreatedDBObject(vp, true);

                // Turn it - and its grid - on
                vp.On = true;
                vp.GridOn = true;
            }

            // Finally we call our function on it
            f(vp);
        }

        /// <summary>
        /// Apply plot settings to the provided layout.
        /// </summary>
        /// <param name="pageSize">The canonical media name for our page size.</param>
        /// <param name="styleSheet">The pen settings file (ctb or stb).</param>
        /// <param name="devices">The name of the output device.</param>
        public static void SetPlotSettings(ref Layout lay, string pageSize, string styleSheet, string device)
        {
            using (var ps = new PlotSettings(lay.ModelType))
            {
                ps.CopyFrom(lay);
                var psv = PlotSettingsValidator.Current;

                // Set the device
                var devs = psv.GetPlotDeviceList();

                if (devs.Contains(device))
                {
                    psv.SetPlotConfigurationName(ps, device, null);
                    psv.RefreshLists(ps);
                }

                // Set the media name/size
                var mns = psv.GetCanonicalMediaNameList(ps);

                if (mns.Contains(pageSize))
                {
                    psv.SetCanonicalMediaName(ps, pageSize);

                }

                // Set the pen settings
                var ssl = psv.GetPlotStyleSheetList();

                if (ssl.Contains(styleSheet))
                {
                    psv.SetCurrentStyleSheet(ps, styleSheet);
                }

                // Copy the PlotSettings data back to the Layout
                var upgraded = false;
                if (!lay.IsWriteEnabled)
                {
                    lay.UpgradeOpen();
                    upgraded = true;
                }

                lay.CopyFrom(ps);

                if (upgraded)
                {
                    lay.DowngradeOpen();
                }
            }
        }

        /// <summary>
        /// Determine the maximum possible size for this layout.
        /// </summary>
        /// <returns>The maximum extents of the viewport on this layout.</returns>
        public static Extents2d GetMaximumExtents(ref Layout lay)
        {
            // If the drawing template is imperial, we need to divide by
            // 1" in mm (25.4)
            var div = lay.PlotPaperUnits == PlotPaperUnit.Inches ? 25.4 : 1.0;

            // We need to flip the axes if the plot is rotated by 90 or 270 deg
            var doIt =
              lay.PlotRotation == PlotRotation.Degrees090 ||
              lay.PlotRotation == PlotRotation.Degrees270;

            // Get the extents in the correct units and orientation
            var minPnt = lay.PlotPaperMargins.MinPoint;
            var min = Swap(ref minPnt, doIt) / div;

            var psize = lay.PlotPaperSize;
            var maxPnt = lay.PlotPaperMargins.MaxPoint;
            var max = (Swap(ref psize, doIt) - Swap(ref maxPnt, doIt).GetAsVector()) / div;

            return new Extents2d(min, max);
        }

        /// <summary>
        /// Sets the size of the viewport according to the provided extents.
        /// </summary>
        /// <param name="ext">The extents of the viewport on the page.</param>
        /// <param name="fac">Optional factor to provide padding.</param>
        public static void ResizeViewport(ref Viewport vp, Extents2d ext, double fac = 1.0)
        {
            vp.Width = (ext.MaxPoint.X - ext.MinPoint.X) * fac;
            vp.Height = (ext.MaxPoint.Y - ext.MinPoint.Y) * fac;
            var temp = Point2d.Origin + (ext.MaxPoint - ext.MinPoint) * 0.5;
            vp.CenterPoint = Pad(ref temp);
        }

        /// <summary>
        /// Sets the view in a viewport to contain the specified model extents.
        /// </summary>
        /// <param name="ext">The extents of the content to fit the viewport.</param>
        /// <param name="fac">Optional factor to provide padding.</param>
        public static void FitContentToViewport(ref Viewport vp, Extents3d ext, double fac = 1.0)
        {
            // Let's zoom to just larger than the extents
            var temp = ext.MinPoint + ((ext.MaxPoint - ext.MinPoint) * 0.5);
            vp.ViewCenter = Strip(ref temp);

            // Get the dimensions of our view from the database extents
            var hgt = ext.MaxPoint.Y - ext.MinPoint.Y;
            var wid = ext.MaxPoint.X - ext.MinPoint.X;

            // We'll compare with the aspect ratio of the viewport itself
            // (which is derived from the page size)
            var aspect = vp.Width / vp.Height;

            // If our content is wider than the aspect ratio, make sure we
            // set the proposed height to be larger to accommodate the
            // content

            if (wid / hgt > aspect)
                hgt = wid / aspect;

            // Set the height so we're exactly at the extents
            vp.ViewHeight = hgt;

            // Set a custom scale to zoom out slightly (could also
            // vp.ViewHeight *= 1.1, for instance)
            vp.CustomScale *= fac;
        }


        [CommandMethod("CL")]
        public void CreateLayout()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;

            if (doc == null)
                return;

            var db = doc.Database;
            var ed = doc.Editor;
            var ext = new Extents2d();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                // Create and select a new layout tab
                var temp = LayoutManager.Current;
                var id = CreateAndMakeLayoutCurrent(ref temp, "NewLayout"); //LayoutManager.Current.CreateAndMakeLayoutCurrent("NewLayout");

                // Open the created layout
                var lay = (Layout)tr.GetObject(id, OpenMode.ForWrite);

                // Make some settings on the layout and get its extents
                SetPlotSettings(ref lay, "ANSI_B_(11.00_x_17.00_Inches)", "monochrome.ctb", "DWF6 ePlot.pc3");

                ext = GetMaximumExtents(ref lay);
                ApplyToViewport(ref lay,
                  tr, 2,
                  vp =>
                  {
                      // Size the viewport according to the extents calculated when
                      // we set the PlotSettings (device, page size, etc.)
                      // Use the standard 10% margin around the viewport
                      // (found by measuring pixels on screenshots of Layout1, etc.)
                      ResizeViewport(ref vp, ext, 0.8);

                      // Adjust the view so that the model contents fit
                      if (ValidDbExtents(db.Extmin, db.Extmax))
                          FitContentToViewport(ref vp, new Extents3d(db.Extmin, db.Extmax), 0.9);

                      // Finally we lock the view to prevent meddling
                      vp.Locked = true;
                  }
                );

                // Commit the transaction
                tr.Commit();
            }

            // Zoom so that we can see our new layout, again with a little padding
            //ed.Command("_.ZOOM", "_E");
            //ed.Command("_.ZOOM", ".7X");
            ed.Regen();
        }

        // Returns whether the provided DB extents - retrieved from
        // Database.Extmin/max - are "valid" or whether they are the default
        // invalid values (where the min's coordinates are positive and the
        // max coordinates are negative)
        private bool ValidDbExtents(Point3d min, Point3d max)
        {
            return !(min.X > 0 && min.Y > 0 && min.Z > 0 && max.X < 0 && max.Y < 0 && max.Z < 0);
        }
#endif
    }
}