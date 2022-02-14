using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPITraining_4_2
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var Pipes = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PipeCurves)
                .ToList();

            string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "pipes.xlsx");

            using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workBook = new XSSFWorkbook();
                ISheet sheet = workBook.CreateSheet("Лист 1");

                int rowIndex = 0;
                foreach (var pipe in Pipes)
                {
                    string pipeName = pipe.Name;
                    double outDiam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble();
                    double outDiam1 = UnitUtils.ConvertFromInternalUnits(outDiam, UnitTypeId.Millimeters);
                    double innDiam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM).AsDouble();
                    double inDiam1 = UnitUtils.ConvertFromInternalUnits(innDiam, UnitTypeId.Millimeters);
                    double length = pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();
                    double length1 = UnitUtils.ConvertFromInternalUnits(length, UnitTypeId.Millimeters);
                    sheet.SetCellValue(rowIndex, columnIndex: 0, pipeName);
                    sheet.SetCellValue(rowIndex, columnIndex: 1, outDiam1);
                    sheet.SetCellValue(rowIndex, columnIndex: 2, inDiam1);
                    sheet.SetCellValue(rowIndex, columnIndex: 3, length);
                    rowIndex++;
                }
                workBook.Write(stream);
                workBook.Close();
            }
            return Result.Succeeded;
        }
    }
}
