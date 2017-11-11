using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlFun.Excel.Parser
{
    public class ExcelParser<T> : IDisposable
        where T : new()
    {
        private const string Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        // ReSharper disable once StaticMemberInGenericType
        private static readonly string[] ExcelColumnNames = new string[Alphabet.Length * 2];

        private readonly SpreadsheetDocument _spreadsheetDocument;
        private readonly Worksheet _worksheet;
        private readonly SharedStringTable _ssTable;

        static ExcelParser()
        {
            for (int i = 0; i < Alphabet.Length; i++)
            {
                ExcelColumnNames[i] = Alphabet[i].ToString();
            }

            for (int i = 0; i < Alphabet.Length; i++)
            {
                ExcelColumnNames[i + Alphabet.Length] = $"{Alphabet[0]}{Alphabet[i]}";
            }
        }

        public ExcelParser(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                throw new FileNotFoundException($"{nameof(filePath)} argument is invalid.");
            }

            _spreadsheetDocument = SpreadsheetDocument.Open(filePath, true);
            Sheet sheet = _spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().First();
            if (sheet == null)
            {
                throw new InvalidOperationException("There are no sheets in document.");
            }

            _worksheet = ((WorksheetPart)_spreadsheetDocument.WorkbookPart.GetPartById(sheet.Id.Value)).Worksheet;
            _ssTable = _spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First().SharedStringTable;
        }

        public List<T> Parse(bool ignoreFirstRow = false)
        {
            List<T> list = new List<T>();
            Type stringType = typeof(string);
            Type datetimeType = typeof(DateTime?);
            Type decimalType = typeof(decimal);

            Type type = typeof(T);
            PropertyInfo[] properties = type.GetProperties();
            Dictionary<string, int> numbers = new Dictionary<string, int>();
            Dictionary<string, Type> types = new Dictionary<string, Type>();

            List<Row> rows = CalculateFormulas();

            foreach (Row row in rows)
            {
                List<Cell> cells = row.Elements<Cell>().ToList();

                T item = new T();
                foreach (PropertyInfo property in properties)
                {
                    if (!numbers.ContainsKey(property.Name))
                    {
                        object tmp = property.GetCustomAttributes(typeof(OrderParseAttribute), false).FirstOrDefault();
                        numbers[property.Name] = ((OrderParseAttribute)tmp)?.Number ?? -1;
                    }
                    var tempNumber = numbers[property.Name];
                    if (tempNumber == -1)
                    {
                        continue;
                    }

                    if (!types.ContainsKey(property.Name))
                    {
                        types[property.Name] = property.PropertyType;
                    }
                    Type tempType = types[property.Name];

                    string tempCellValue = GetDataFromCell(cells, row.RowIndex.Value, tempNumber);
                    if (tempType == stringType)
                    {
                        property.SetValue(item, tempCellValue ?? string.Empty, null);
                        continue;
                    }

                    if (tempType == datetimeType)
                    {
                        DateTime? tempDate;
                        try
                        {
                            tempDate = DateTime.TryParse(tempCellValue, out DateTime date)
                                ? date
                                : DateTime.FromOADate(double.Parse(tempCellValue));
                        }
                        catch
                        {
                            tempDate = null;
                        }
                        property.SetValue(item, tempDate, null);
                        continue;
                    }

                    if (tempType == decimalType)
                    {
                        decimal tempDecimal;
                        try
                        {
                            tempDecimal = ParseDecimal(tempCellValue);
                        }
                        catch
                        {
                            tempDecimal = 0;
                        }
                        property.SetValue(item, tempDecimal, null);
                    }
                }
                list.Add(item);
            }

            PropertyInfo keyProperty = properties.SingleOrDefault(p => p.GetCustomAttributes(typeof(UniqueKeyParseAttribute), false).FirstOrDefault() != null);
            if (keyProperty != null)
            {
                list.RemoveAll(x => string.IsNullOrEmpty((string)keyProperty.GetValue(x, null)));
            }

            if (ignoreFirstRow)
            {
                list.RemoveAt(0);
            }
            return list;
        }

        private string GetDataFromCell(List<Cell> cells, uint rowNumber, int itemNumber)
        {
            Cell cell = cells.FirstOrDefault(x => x.CellReference.Value.Equals(ExcelColumnNames[itemNumber] + rowNumber.ToString()));
            if (cell == null)
            {
                return null;
            }

            return cell.CellFormula == null
                ? (cell.DataType == null
                    ? cell.InnerText
                    : (cell.CellValue == null
                        ? null
                        : _ssTable.ChildElements[int.Parse(cell.CellValue.InnerText)].InnerText))
                : cell.CellValue.InnerText;
        }

        private List<Row> CalculateFormulas()
        {
            List<Row> rows = _worksheet.Descendants<Row>().ToList();
            foreach (Row row in rows)
            {
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellFormula != null)
                    {
                        cell.CellFormula.CalculateCell = true;
                    }
                }
            }
            _worksheet.Save();
            return rows;
        }

        private static decimal ParseDecimal(string number)
        {
            decimal.TryParse(number.Replace(".", ","), out decimal num);
            return num;
        }

        public void Dispose()
        {
            _spreadsheetDocument.Dispose();
        }
    }
}
