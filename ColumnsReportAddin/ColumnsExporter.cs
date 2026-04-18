using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ClosedXML.Excel;

namespace ColumnsReportAddin
{
    [Transaction(TransactionMode.ReadOnly)]
    public class ColumnsExporter : IExternalCommand
    {
        private const double FtToM = 0.3048;
        private const double CuFtToM3 = 0.0283168466;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            using (ExportForm form = new ExportForm())
            {
                if (form.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                    return Result.Cancelled;

                string folder = form.SelectedPath;
                string rvtName = Path.GetFileNameWithoutExtension(doc.Title);
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd_hh-mm-ss tt");
                string fileName = $"{rvtName}-Columns Report [{timestamp}].xlsx";
                string fullPath = Path.Combine(folder, fileName);

                try
                {
                    List<ColumnData> data = CollectColumns(doc);
                    WriteExcel(data, fullPath);
                    Process.Start(fullPath);
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Error", ex.Message);
                    return Result.Failed;
                }
            }

            return Result.Succeeded;
        }

        private List<ColumnData> CollectColumns(Document doc)
        {
            List<ColumnData> result = new List<ColumnData>();

            var structural = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_StructuralColumns)
                .WhereElementIsNotElementType();

            var architectural = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Columns)
                .WhereElementIsNotElementType();

            List<Element> allColumns = new List<Element>();
            allColumns.AddRange(structural);
            allColumns.AddRange(architectural);

            foreach (Element elem in allColumns)
            {
                FamilyInstance fi = elem as FamilyInstance;
                if (fi == null) continue;

                ColumnData col = new ColumnData
                {
                    Family = fi.Symbol.FamilyName,
                    Type = fi.Symbol.Name,
                    Id = fi.Id.IntegerValue
                };

                LocationPoint loc = fi.Location as LocationPoint;
                if (loc != null)
                {
                    col.Easting = Math.Round(loc.Point.X * FtToM, 3);
                    col.Northing = Math.Round(loc.Point.Y * FtToM, 3);
                }

                col.BaseLevel = GetLevelName(fi, BuiltInParameter.FAMILY_BASE_LEVEL_PARAM);
                col.BaseOffset = GetParamInMeters(fi, BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM);
                col.TopLevel = GetLevelName(fi, BuiltInParameter.FAMILY_TOP_LEVEL_PARAM);
                col.TopOffset = GetParamInMeters(fi, BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM);
                col.Height = GetParamInMeters(fi, BuiltInParameter.INSTANCE_LENGTH_PARAM);

                Parameter volParam = fi.LookupParameter("Volume");
                if (volParam != null)
                    col.Volume = Math.Round(volParam.AsDouble() * CuFtToM3, 3);

                result.Add(col);
            }

            return result;
        }

        private string GetLevelName(FamilyInstance fi, BuiltInParameter bip)
        {
            Parameter p = fi.get_Parameter(bip);
            if (p == null) return string.Empty;
            Level lvl = fi.Document.GetElement(p.AsElementId()) as Level;
            return lvl?.Name ?? string.Empty;
        }

        private double GetParamInMeters(FamilyInstance fi, BuiltInParameter bip)
        {
            Parameter p = fi.get_Parameter(bip);
            if (p == null) return 0;
            return Math.Round(p.AsDouble() * FtToM, 3);
        }

        private void WriteExcel(List<ColumnData> columns, string path)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Columns");

                string[] headers =
                {
                    "Family", "Type", "ID",
                    "Easting (m)", "Northing (m)",
                    "Base Level", "Base Offset (m)",
                    "Top Level", "Top Offset (m)",
                    "Height (m)", "Volume (m\u00B3)"
                };

                for (int i = 0; i < headers.Length; i++)
                    ws.Cell(1, i + 1).Value = headers[i];

                var headerRange = ws.Range(1, 1, 1, headers.Length);
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

                for (int r = 0; r < columns.Count; r++)
                {
                    var c = columns[r];
                    int row = r + 2;
                    ws.Cell(row, 1).Value = c.Family;
                    ws.Cell(row, 2).Value = c.Type;
                    ws.Cell(row, 3).Value = c.Id;
                    ws.Cell(row, 4).Value = c.Easting;
                    ws.Cell(row, 5).Value = c.Northing;
                    ws.Cell(row, 6).Value = c.BaseLevel;
                    ws.Cell(row, 7).Value = c.BaseOffset;
                    ws.Cell(row, 8).Value = c.TopLevel;
                    ws.Cell(row, 9).Value = c.TopOffset;
                    ws.Cell(row, 10).Value = c.Height;
                    ws.Cell(row, 11).Value = c.Volume;
                }

                ws.Columns().AdjustToContents();
                wb.SaveAs(path);
            }
        }
    }

    internal class ColumnData
    {
        public string Family { get; set; }
        public string Type { get; set; }
        public int Id { get; set; }
        public double Easting { get; set; }
        public double Northing { get; set; }
        public string BaseLevel { get; set; }
        public double BaseOffset { get; set; }
        public string TopLevel { get; set; }
        public double TopOffset { get; set; }
        public double Height { get; set; }
        public double Volume { get; set; }
    }
}
