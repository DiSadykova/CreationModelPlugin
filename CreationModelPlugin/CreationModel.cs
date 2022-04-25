using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreationModelPlugin
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class CreationModel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            Level level1, level2;
            GetLevels(doc, out level1, out level2);
            List<Wall> walls = CreateWalls(doc, level1, level2);
            AddDoors(doc, level1, walls[0]);
            List<FamilyInstance> windows = AddWindows(doc, level1, walls);
            return Result.Succeeded;
        }

        private static List<FamilyInstance> AddWindows(Document doc, Level level1, List<Wall> walls)
        {
            FamilySymbol windowType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Windows)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0915 x 1830 мм"))
                .Where(x => x.FamilyName.Equals("Фиксированные"))
                .FirstOrDefault();
            List<FamilyInstance> windows = new List<FamilyInstance>();
            Transaction transaction = new Transaction(doc, "Построение окон");
            transaction.Start();
            for (int i = 1; i < walls.Count; i++)
            {
                Wall wall = walls[i];
                LocationCurve hostCurve = wall.Location as LocationCurve;
                XYZ point1 = hostCurve.Curve.GetEndPoint(0);
                XYZ point2 = hostCurve.Curve.GetEndPoint(1);
                XYZ point = (point1 + point2) / 2;

                if (!windowType.IsActive)
                    windowType.Activate();
                FamilyInstance window = doc.Create.NewFamilyInstance(point, windowType, wall, level1, StructuralType.NonStructural);
                windows.Add(window);
                window.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM).Set(UnitUtils.ConvertToInternalUnits(915, UnitTypeId.Millimeters));
            }
            transaction.Commit();

            return windows;
        }

        private static List<Wall> CreateWalls(Document doc, Level level1, Level level2, double widthSet = 10000, double depthSet = 5000)
        {
            double width = UnitUtils.ConvertToInternalUnits(widthSet, UnitTypeId.Millimeters);
            double depth = UnitUtils.ConvertToInternalUnits(depthSet, UnitTypeId.Millimeters);
            double dx = width / 2;
            double dy = depth / 2;

            List<XYZ> points = new List<XYZ>();
            points.Add(new XYZ(-dx, -dy, 0));
            points.Add(new XYZ(dx, -dy, 0));
            points.Add(new XYZ(dx, dy, 0));
            points.Add(new XYZ(-dx, +dy, 0));
            points.Add(new XYZ(-dx, -dy, 0));

            List<Wall> walls = new List<Wall>();

            Transaction transaction = new Transaction(doc, "Построение стен");
            transaction.Start();
            for (int i = 0; i < 4; i++)
            {
                Line line = Line.CreateBound(points[i], points[i + 1]);
                Wall wall = Wall.Create(doc, line, level1.Id, false);
                walls.Add(wall);
                wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(level2.Id);
            }
            transaction.Commit();
            return walls;
        }

        private static void AddDoors(Document doc, Level level1, Wall wall)
        {
            FamilySymbol doorType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0915 x 2134 мм"))
                .Where(x => x.FamilyName.Equals("Одиночные-Щитовые"))
                .FirstOrDefault();

            LocationCurve hostCurve = wall.Location as LocationCurve;
            XYZ point1 = hostCurve.Curve.GetEndPoint(0);
            XYZ point2 = hostCurve.Curve.GetEndPoint(1);
            XYZ point = (point1 + point2) / 2;
            Transaction transaction = new Transaction(doc, "Построение дверей");
            transaction.Start();
            if (!doorType.IsActive)
                doorType.Activate();
            doc.Create.NewFamilyInstance(point, doorType, wall, level1, StructuralType.NonStructural);
            transaction.Commit();
        }

        private static void GetLevels(Document doc, out Level level1, out Level level2)
        {
            List<Level> listLevel = new FilteredElementCollector(doc)
                            .OfClass(typeof(Level))
                            .OfType<Level>()
                            .ToList();

            level1 = listLevel
                .Where(x => x.Name.Equals("Уровень 1"))
                .FirstOrDefault();
            level2 = listLevel
                .Where(x => x.Name.Equals("Уровень 2"))
                .FirstOrDefault();
        }
    }
}
