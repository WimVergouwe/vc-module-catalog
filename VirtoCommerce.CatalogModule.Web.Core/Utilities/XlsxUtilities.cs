using System;
using System.Collections;
using System.Collections.Generic;
using System.Dynamic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using VirtoCommerce.CatalogModule.Web.ExportImport;
using VirtoCommerce.Domain.Catalog.Model;

namespace VirtoCommerce.CatalogModule.Web.Utilities
{
    public static class DemoXlsxUtilities
    {
        public static void CreateFileWithData(Stream outStream, string sheetName, ExportDefinition<XlsxProduct> exportDefinition, IEnumerable<XlsxProduct> products)
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

                var count = 0u;
                foreach (var product in products)
                {
                    WriteRow(worksheetPart, exportDefinition, product, count+2); // Count is not 0-based, and we already have 1 row for the headers.
                    count++;
                }
            }
        }

        private static void WriteColumnHeaders(ExportDefinition<XlsxProduct> exportDefinition, WorksheetPart worksheetPart)
        {
            var rowValues = new List<OpenXmlElement>();
            var index = 1;
            foreach (var exportColumnDefinition in exportDefinition)
            {
                rowValues.Add(new Cell
                {
                    CellValue = new CellValue(exportColumnDefinition.Name),
                    DataType = CellValues.String,
                    CellReference = $"{GetColumnName(index)}1"
                });
                index++;
            }
            worksheetPart.Worksheet.First().AppendChild(new Row(rowValues.ToArray()));
        }
        
        private static void WriteRow(WorksheetPart worksheetPart, ExportDefinition<XlsxProduct> exportDefinition, XlsxProduct product, uint rowIndex)
        {
            var rowValues = new List<OpenXmlElement>();

            var index = 1;
            foreach (var columnExportDefinition in exportDefinition)
            {
                var valueType = CellValues.String;
                if (columnExportDefinition.PropertyType == typeof (bool))
                {
                    valueType = CellValues.Boolean;
                }
                else if (columnExportDefinition.PropertyType == typeof (DateTime))
                {
                    valueType = CellValues.Date;
                }
                else if (columnExportDefinition.PropertyType == typeof (int) || columnExportDefinition.PropertyType == typeof(float) || columnExportDefinition.PropertyType == typeof(double) || columnExportDefinition.PropertyType == typeof(decimal))
                {
                    valueType = CellValues.Number;
                }
                rowValues.Add(new Cell
                {
                    CellValue = new CellValue(columnExportDefinition.GetValue(product)),
                    DataType = valueType,
                    CellReference = $"{GetColumnName(index)}{rowIndex}"
                });
                index++;
            }
            worksheetPart.Worksheet.First().AppendChild(new Row(rowValues.ToArray()) { RowIndex = UInt32Value.FromUInt32(rowIndex) });
        }

        private static string GetColumnName(int columnIndex)
        {
            columnIndex--;
            if (columnIndex >= 0 && columnIndex < 26)
                return ((char)('A' + columnIndex)).ToString();

            if (columnIndex > 25)
                return GetColumnName(columnIndex / 26) + GetColumnName(columnIndex % 26 + 1);

            throw new Exception("Invalid Column #" + (columnIndex + 1).ToString());
        }

        public static IEnumerable<CatalogProduct> GetProductsFromFile(Stream xlsxStream, ImportDefinition<XlsxProduct> importDefinition, string sheetName = null)
        {
            var sheetData = GetWorkSheetData(xlsxStream, sheetName);
            if (sheetData == null) yield break;

            var columnNames = ParseColumnNames(sheetData.Descendants<Row>().FirstOrDefault());
            if (columnNames == null) yield break;
            
            foreach (var data in ParseData(columnNames, sheetData.Descendants<Row>().Skip(1), ignoreEmptyCells: false))
            {
                var product = new XlsxProduct
                {
                     PropertyValues = new List<PropertyValue>(columnNames.Count - importDefinition.Count())
                };
                foreach (var columnName in columnNames)
                {
                    var columnDefinition = importDefinition.FirstOrDefault(x => x.Name.Equals(columnName, StringComparison.OrdinalIgnoreCase));
                    if (columnDefinition == null)
                    {
                        
                        product.PropertyValues.Add(new PropertyValue
                        {
                            PropertyName = columnName,
                            Value = data[columnName],
                            ValueType = ParseType(data[columnName])
                        });
                    }
                    else
                    {
                        columnDefinition.SetValue(product, data[columnName]);
                    }
                }

                yield return product;
            }
        }

        public static PropertyValueType ParseType(object value)
        {
            PropertyValueType parsedType = 0;
            bool parsedBool;
            int parsedInt;
            float parsedFloat;
            double parsedDouble;
            decimal parsedDecimal;
            DateTime parsedDateTime;
            if (bool.TryParse(value.ToString(), out parsedBool))
            {
                parsedType = PropertyValueType.Boolean;
            }
            else if (int.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out parsedInt) || float.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out parsedFloat) || double.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out parsedDouble) || decimal.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out parsedDecimal))
            {
                parsedType = PropertyValueType.Number;
            }
            else if (DateTime.TryParse(value.ToString(), CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal, out parsedDateTime))
            {
                parsedType = PropertyValueType.DateTime;
            }
            return parsedType;
        }

        private static IEnumerable<DynamicData> ParseData(IReadOnlyList<string> columnNames, IEnumerable<Row> rows, bool ignoreEmptyCells)
        {
            foreach (var row in rows)
            {
                var dynamicObject = new CaseInsensitiveDynamicObject();

                var columnIndex = 0;
                var rowHasData = false;
                foreach (var cell in row.Descendants<Cell>())
                {
                    var cellIndex = XlsxUtility.GetCellIndex(cell.CellReference);
                    if (cellIndex > columnIndex)
                        columnIndex = cellIndex;

                    var value = cell.CellValueTyped();

                    var cellHasData = false;
                    if (value != null && !string.Empty.Equals(value))
                    {
                        rowHasData = cellHasData = true;
                    }

                    if (cellHasData || !ignoreEmptyCells)
                    {
                        var columnName = columnNames[columnIndex];
                        dynamicObject[columnName] = value;
                    }

                    columnIndex++;
                }

                if (rowHasData)
                {
                    var data = new DynamicData(dynamicObject, (int)row.RowIndex.Value);
                    yield return data;
                }
            }
        }

        private static List<string> ParseColumnNames(Row firstRow)
        {
            var result = new List<string>();

            var columnIndex = 0;
            if (firstRow != null)
            {
                foreach (var cell in firstRow.Descendants<Cell>())
                {
                    var cellIndex = XlsxUtility.GetCellIndex(cell.CellReference);

                    if (cellIndex > columnIndex)
                    {
                        for (int i = 0; i < cellIndex - columnIndex; i++)
                        {
                            result.Add("undefined column " + columnIndex);
                            columnIndex += i + 1;
                        }
                    }

                    var unprocessedKey = cell.CellTextValue();
                    result.Add(unprocessedKey);

                    columnIndex++;
                }
            }
            else
            {
                return null;
            }
            return result;
        }

        private static SheetData GetWorkSheetData(Stream stream, string sheetToImport = null)
        {
            var spreadSheetDocument = SpreadsheetDocument.Open(stream, false);
            var workBookPart = spreadSheetDocument.WorkbookPart;

            WorksheetPart workSheetPart;
            if (!string.IsNullOrEmpty(sheetToImport))
            {
                workSheetPart = GetWorkSheetPart(workBookPart, sheetToImport);
                if (workSheetPart == null)
                {
                    return null;
                }
            }
            else
            {
                workSheetPart = workBookPart.WorksheetParts.FirstOrDefault();
                if (workSheetPart == null)
                {
                    return null;
                }
            }

            var sheetData = workSheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
            return sheetData;
        }

        private static WorksheetPart GetWorkSheetPart(WorkbookPart workbookPart, string sheetName)
        {
            var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => sheetName.Equals(s.Name));
            if (sheet == null)
                return null;

            return (WorksheetPart)workbookPart.GetPartById(sheet.Id);
        }
    }
    public static class XlsxUtility
    {
        public static string GetColumnName(string cellReference)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellReference);

            return match.Value;
        }

        public static int GetOffsetOfColumnName(string code)
        {
            var offset = 0;
            var byteArray = Encoding.ASCII.GetBytes(code).Reverse().ToArray();
            for (var i = 0; i < byteArray.Length; i++)
            {
                offset += (byteArray[i] - 65 + 1) * Convert.ToInt32(Math.Pow(26.0, Convert.ToDouble(i)));
            }
            return offset - 1;
        }

        public static int GetCellIndex(string cellReference)
        {
            var columnName = GetColumnName(cellReference);
            return GetOffsetOfColumnName(columnName);
        }

        //see http://stackoverflow.com/questions/2624333/how-do-i-read-data-from-a-spreadsheet-using-the-openxml-format-sdk
        public static string CellTextValue(this Cell cell)
        {
            if (cell == null) return null;
            if (cell.DataType == null) return cell.CellValue?.InnerText;

            if (cell.DataType == CellValues.SharedString)
            {
                // For shared strings, look up the value in the shared strings table.
                var worksheetPart = ((Worksheet)cell.Parent.Parent.Parent).WorksheetPart;
                var workbookPart = ((SpreadsheetDocument)worksheetPart.OpenXmlPackage).WorkbookPart;
                var sstPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                if (sstPart?.SharedStringTable != null)
                {
                    var index = Int32.Parse(cell.InnerText);
                    return sstPart.SharedStringTable.ElementAt(index).InnerText;
                }
            }

            return cell.CellValue?.InnerText;
        }

        //see http://stackoverflow.com/questions/2624333/how-do-i-read-data-from-a-spreadsheet-using-the-openxml-format-sdk
        public static object CellValueTyped(this Cell cell)
        {
            if (cell == null) return null;

            var worksheetPart = ((Worksheet)cell.Parent.Parent.Parent).WorksheetPart;
            var workbookPart = ((SpreadsheetDocument)worksheetPart.OpenXmlPackage).WorkbookPart;

            if (cell.DataType == null)
            {
                // Try to get date before trying to parse a number, because date is stored as long.
                DateTime dateTime;
                if (TryGetDate(cell, workbookPart.WorkbookStylesPart.Stylesheet, out dateTime)) return dateTime;

                // Try to get number.
                decimal number;
                if (decimal.TryParse(cell.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture, out number)) return number;

                return cell.CellValue?.InnerText;
            }

            switch (cell.DataType.Value)
            {
                case CellValues.Boolean:
                    switch (cell.InnerText)
                    {
                        case "0":
                            return false;
                        default:
                            return true;
                    }

                case CellValues.Number:
                    return decimal.Parse(cell.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture);

                case CellValues.SharedString:
                    // For shared strings, look up the value in the shared strings table.
                    var sstPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    if (sstPart?.SharedStringTable != null)
                    {
                        var index = Int32.Parse(cell.InnerText);
                        return sstPart.SharedStringTable.ElementAt(index).InnerText;
                    }
                    break;
            }

            return cell.CellValue?.InnerText;
        }

        #region DateTime parsing

        /// <summary>
        /// Set of standard date(time) formats.
        /// See: http://closedxml.codeplex.com/wikipage?title=NumberFormatId%20Lookup%20Table
        /// </summary>
        private static readonly HashSet<uint> StandardDateFormats = new HashSet<uint>
        {
            14,     // d/m/yyyy
            15,     // d-mmm-yy
            16,     // d-mmmm
            17,     // mmm-yy
            22,     // m/d/yyyy H:mm
        };

        /// <summary>
        /// See: http://stackoverflow.com/questions/4730152/what-indicates-an-office-open-xml-cell-contains-a-date-time-value
        /// </summary>
        private static bool TryGetDate(Cell cell, Stylesheet stylesheet, out DateTime dateTime)
        {
            dateTime = default(DateTime);

            // If cell has a specific datatype, it cannot be a Date.
            if (cell.DataType != null) return false;

            // Cell has no styleindex, it cannot be a Date.
            if (cell.StyleIndex == null) return false;

            var styleIndex = (int)cell.StyleIndex.Value;
            var cellFormat = (CellFormat)stylesheet.CellFormats.ElementAt(styleIndex);

            // If cellFormat has no NumberFormatId, it's definitly not a date.
            if (cellFormat.NumberFormatId == null) return false;

            var numberFormatId = cellFormat.NumberFormatId.Value;

            // Check if the numberFormatId is a standard Date FormatId
            if (!StandardDateFormats.Contains(numberFormatId))
            {
                // Lookup the custom NumberingFormat in the NumberingFormats that are include in the excel sheet.
                var numberingFormat = stylesheet.NumberingFormats?
                    .Cast<NumberingFormat>()
                    .SingleOrDefault(f => f.NumberFormatId.Value == numberFormatId);

                // If no custom format was found, we cannot be sure it's a Date.
                if (numberingFormat == null) return false;
                if (!IsDateFormat(numberingFormat)) return false;
            }

            double parsedValue;
            if (!double.TryParse(cell.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture, out parsedValue)) return false;

            dateTime = DateTime.FromOADate(parsedValue);
            return true;
        }

        /// <summary>
        /// Check if the given NumberingFormat is a Date format.
        /// </summary>
        private static bool IsDateFormat(NumberingFormat numberingFormat)
        {
            // A Date should always contain a year and month. Year should be encoded as 'yy' or 'yyyy'.
            // Month should always be formatted as 'm' or 'mmm' ('mm' stands for minutes).
            // The order of month and year can be changed, hence the two regexes.

            var dateRegex1 = new Regex("^.*m[mm]?.*yy[yy]?.*$");
            if (dateRegex1.IsMatch(numberingFormat.FormatCode.Value))
            {
                return true;
            }

            var dateRegex2 = new Regex("^.*yy[yy]?.*m[mm]?.*$");
            if (dateRegex2.IsMatch(numberingFormat.FormatCode.Value))
            {
                return true;
            }

            return false;
        }

        #endregion
    }
    public class CaseInsensitiveDynamicObject : DynamicObject, IDictionary<string, object>
    {
        public IDictionary<string, object> Dictionary { get; private set; }

        public CaseInsensitiveDynamicObject(IDictionary<string, object> dictionary)
        {
            Dictionary = new Dictionary<string, object>(dictionary, StringComparer.OrdinalIgnoreCase);
        }

        public CaseInsensitiveDynamicObject()
        {
            Dictionary = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        }

        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            return Dictionary.TryGetValue(binder.Name, out result);
        }

        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            Dictionary[binder.Name] = value;
            return true;
        }

        public override bool TryDeleteMember(DeleteMemberBinder binder)
        {
            Dictionary.Remove(binder.Name);
            return true;
        }

        public override IEnumerable<string> GetDynamicMemberNames()
        {
            return Dictionary.Keys;
        }


        public IEnumerator<KeyValuePair<string, object>> GetEnumerator()
        {
            return Dictionary.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable)Dictionary).GetEnumerator();
        }

        public void Add(KeyValuePair<string, object> item)
        {
            Dictionary.Add(item);
        }

        public void Clear()
        {
            Dictionary.Clear();
        }

        public bool Contains(KeyValuePair<string, object> item)
        {
            return Dictionary.Contains(item);
        }

        public void CopyTo(KeyValuePair<string, object>[] array, int arrayIndex)
        {
            Dictionary.CopyTo(array, arrayIndex);
        }

        public bool Remove(KeyValuePair<string, object> item)
        {
            return Dictionary.Remove(item);
        }

        public int Count => Dictionary.Count;

        public bool IsReadOnly => Dictionary.IsReadOnly;

        public bool ContainsKey(string key)
        {
            return Dictionary.ContainsKey(key);
        }

        public void Add(string key, object value)
        {
            Dictionary.Add(key, value);
        }

        public bool Remove(string key)
        {
            return Dictionary.Remove(key);
        }

        public bool TryGetValue(string key, out object value)
        {
            return Dictionary.TryGetValue(key, out value);
        }

        public object this[string key]
        {
            get { return Dictionary[key]; }
            set { Dictionary[key] = value; }
        }

        public ICollection<string> Keys => Dictionary.Keys;

        public ICollection<object> Values => Dictionary.Values;
    }

    public class DynamicData : IEnumerable<KeyValuePair<string, object>>
    {
        private readonly IDictionary<string, object> _dictionary;

        /// <summary>
        /// Dynamic object holding the data.
        /// </summary>
        public dynamic Data { get; private set; }

        /// <summary>
        /// Diagnostic source information for this data, e.g. a row number.
        /// </summary>
        public object SourceInfo { get; private set; }

        public DynamicData(dynamic data = null, object sourceInfo = null)
        {
            Data = data ?? new CaseInsensitiveDynamicObject();
            _dictionary = (IDictionary<string, object>)Data;
            SourceInfo = sourceInfo;
        }

        public bool TryGetValue(string name, out object value)
        {
            if (name.Equals("SourceInfo", StringComparison.OrdinalIgnoreCase))
            {
                value = SourceInfo;
                return true;
            }
            return _dictionary.TryGetValue(name, out value);
        }

        public object this[string name]
        {
            get
            {
                object value;
                return TryGetValue(name, out value) ? value : null;
            }
            set { _dictionary[name] = value; }
        }

        public bool HasValue(string name)
        {
            return _dictionary.ContainsKey(name);
        }

        public ICollection<object> Values()
        {
            return _dictionary.Values;
        }

        public ICollection<string> FieldNames()
        {
            return _dictionary.Keys;
        }

        public void Remove(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentNullException("name");

            _dictionary.Remove(name);
        }

        public void Add(string name, object value)
        {
            _dictionary.Add(name, value);
        }

        public override string ToString()
        {
            return ((IDictionary<string, object>)Data).Values.ToCommaDelimitedString();
        }

        public IEnumerator<KeyValuePair<string, object>> GetEnumerator()
        {
            return _dictionary.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
