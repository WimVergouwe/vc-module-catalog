using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using VirtoCommerce.Domain.Catalog.Model;

namespace VirtoCommerce.CatalogModule.Web.Utilities
{
    public static class DemoXlsxUtilities
    {
        public static void CreateFileWithData(Stream outStream, string sheetName, ExportDefinition<CatalogProduct> exportDefinition, IEnumerable<CatalogProduct> products)
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

                WriteColumnHeaders(exportDefinition, worksheetPart);

                var count = 0;
                foreach (var product in products)
                {
                    WriteRow(worksheetPart, exportDefinition, product);
                    count++;
                }
            }
        }

        private static void WriteColumnHeaders(ExportDefinition<CatalogProduct> exportDefinition, WorksheetPart worksheetPart)
        {
            var rowValues = new List<OpenXmlElement>();

            // Note: keys don't have to be sorted, they just need to appear in the same order everytime they're iterated.
            foreach (var exportColumnDefinition in exportDefinition)
            {
                rowValues.Add(new Cell { CellValue = new CellValue(exportColumnDefinition.Name), DataType = CellValues.String });
            }
            worksheetPart.Worksheet.First().AppendChild(new Row(rowValues.ToArray()));
        }
        
        private static void WriteRow(WorksheetPart worksheetPart, ExportDefinition<CatalogProduct> exportDefinition, CatalogProduct product)
        {
            var rowValues = new List<OpenXmlElement>();
            foreach (var columnExportDefinition in exportDefinition)
            {
                rowValues.Add(new Cell { CellValue = new CellValue(columnExportDefinition.GetValue(product)), DataType = CellValues.String });
            }
            worksheetPart.Worksheet.First().AppendChild(new Row(rowValues.ToArray()));
        }
    }
}
