using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace VirtoCommerce.CatalogModule.Web.Utilities
{
    public static class DemoXlsxUtilities
    {
        public static void CreateFileWithData<T>(Stream outStream, string sheetName, IEnumerable<T> products)
        {
            using (var spreadSheet = SpreadsheetDocument.Create(outStream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = spreadSheet.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                var sheets = spreadSheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                var sheet = new Sheet
                {
                    Id = spreadSheet.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = sheetName
                };
                sheets.AppendChild(sheet);

                WriteColumnHeaders(typeof(T), worksheetPart);

                var count = 0;
                foreach (var product in products)
                {
                    WriteRow(product, worksheetPart);
                    count++;
                }
            }
        }

        private static void WriteColumnHeaders(Type type, WorksheetPart worksheetPart)
        {
            var rowValues = new List<OpenXmlElement>();

            // Note: keys don't have to be sorted, they just need to appear in the same order everytime they're iterated.
            foreach (var property in type.GetProperties())
            {
                rowValues.Add(new Cell { CellValue = new CellValue(property.Name), DataType = CellValues.String });
            }
            worksheetPart.Worksheet.First().AppendChild(new Row(rowValues.ToArray()));
        }



        private static void WriteRow<T>(T product, WorksheetPart worksheetPart)
        {
            var inv = CultureInfo.InvariantCulture;

            var rowValues = new List<OpenXmlElement>();
            foreach (var value in typeof(T).GetProperties().Select(property => property.GetValue(product, null)))
            {
                if (value != null)
                {
                    var formattable = value as IFormattable;
                    rowValues.Add(new Cell { CellValue = new CellValue(formattable?.ToString(null, inv) ?? value.ToString()), DataType = CellValues.String });
                }
                else
                {

                    rowValues.Add(new Cell { CellValue = new CellValue(""), DataType = CellValues.String });
                }
            }
            worksheetPart.Worksheet.First().AppendChild(new Row(rowValues.ToArray()));
        }
    }
}
