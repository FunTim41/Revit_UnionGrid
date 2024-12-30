using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using OperationCanceledException = Autodesk.Revit.Exceptions.OperationCanceledException;

namespace 合并轴线
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements
        )
        {
            //ui应用程序
            UIApplication uiapp = commandData.Application;
            //应用程序
            Application app = uiapp.Application;
            //ui文档
            UIDocument uidoc = uiapp.ActiveUIDocument;
            //文档
            Document doc = uidoc.Document;
            GridFilter gridFilter = new GridFilter();
            try
            {
                using (Transaction trans = new Transaction(doc, "transaction"))
                {
                    trans.Start();
                    List<Reference> references = uidoc
                        .Selection.PickObjects(ObjectType.Element, gridFilter)
                        .ToList();
                    List<Grid> girds = new List<Grid>();
                    references.ForEach(i => girds.Add(doc.GetElement(i) as Grid));

                    XYZ InterPoint = GetLineIntersection(
                        girds[0].Curve as Line,
                        girds[1].Curve as Line
                    );
                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    collector
                        .OfClass(typeof(FamilySymbol))
                        .OfCategory(BuiltInCategory.OST_StructuralColumns);
                    FamilySymbol columnType =
                        collector.FirstOrDefault(f => f.Name == "300 x 450mm") as FamilySymbol;
                    FamilyInstance columnInstance = doc.Create.NewFamilyInstance(
                        InterPoint,
                        columnType,
                        doc.ActiveView.GenLevel,
                        StructuralType.Column
                    );
                    //CurveLoop curves = new CurveLoop();

                    CreatNewLines(girds, InterPoint, out CurveLoop curves);
                    SketchPlane sketchPlane = SketchPlane.Create(
                        doc,
                        Plane.CreateByNormalAndOrigin(XYZ.BasisZ, girds[0].Curve.GetEndPoint(0))
                    );

                    var multiSegmentGrid = MultiSegmentGrid.Create(
                        doc,
                        girds[0].GetTypeId(),
                        curves,
                        sketchPlane.Id
                    );
                    var newgird = doc.GetElement(multiSegmentGrid) as MultiSegmentGrid;
                    newgird.Text = girds[0].Name;
                    girds.ForEach(i => doc.Delete(i.Id));
                    trans.Commit();
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Tip", ex.Message);
                return Result.Failed;
            }
            return Result.Succeeded;
        }

        private void CreatNewLines(List<Grid> girds, XYZ interPoint, out CurveLoop curveloop)
        {
            CurveLoop curves = new CurveLoop();
            XYZ startpoint;
            XYZ endpoint;
            var point0 = girds[0].Curve.GetEndPoint(0);
            var point1 = girds[0].Curve.GetEndPoint(1);

            startpoint =
                point0.DistanceTo(interPoint) > point1.DistanceTo(interPoint) ? point0 : point1;
            point0 = girds[1].Curve.GetEndPoint(0);
            point1 = girds[1].Curve.GetEndPoint(1);
            endpoint =
                point0.DistanceTo(interPoint) > point1.DistanceTo(interPoint) ? point0 : point1;

            Line line1 = Line.CreateBound(startpoint, interPoint);
            Line line2 = Line.CreateBound(interPoint, endpoint);
            curves.Append(line1);
            curves.Append(line2);
            curveloop = curves;
        }

        public XYZ GetLineIntersection(Line line1, Line line2)
        { // 获取直线的方向向量
            XYZ direction1 = line1.Direction;
            XYZ direction2 = line2.Direction;
            // 获取直线的起点
            XYZ point1 = line1.GetEndPoint(0);
            XYZ point2 = line2.GetEndPoint(0);
            //计算两条直线的参数方程
            //  line1: P1 + t * D1
            // line2: P2 + s * D2
            // 解方程 P1 +t * D1 = P2 + s * D2
            double a = direction1.X;
            double b = -direction2.X;
            double c = direction1.Y;
            double d = -direction2.Y;
            double det = a * d - b * c;
            if (Math.Abs(det) < 1e-9)
            { // 平行或重合，无交点
                return null;
            }
            double dx = point2.X - point1.X;
            double dy = point2.Y - point1.Y;
            double t = (dx * d - dy * b) / det;
            //double s = (dx * c - dy * a) / det;
            // 计算交点
            XYZ intersection = point1 + t * direction1;
            //TaskDialog.Show("交点", $"两条直线的交点为: ({intersection.X}, {intersection.Y}, {intersection.Z})");
            return intersection;
        }
    }

    internal class GridFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            if (elem.Category?.Name == "轴网")
                return true;
            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}
