using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlFun.Excel.Parser
{
    /// <inheritdoc />
    /// <summary>
    /// Parse excel tables to list of enities.
    /// </summary>
    /// <typeparam name="T">Type of class which contains metadata required for parsing.</typeparam>
    public sealed class ExcelParser<T> : IDisposable
        where T : new()
    {
        private const string Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        // ReSharper disable once StaticMemberInGenericType
        private static readonly string[] ExcelColumnNames = new string[Alphabet.Length * 2];

        private readonly SpreadsheetDocument _spreadsheetDocument;
        private readonly Worksheet _worksheet;
        private readonly SharedStringTable _ssTable;

        // ReSharper disable once StaticMemberInGenericType
        private static readonly Dictionary<Type, (Func<string, object> parse, object defaultValue)> Parsers = 
            new Dictionary<Type, (Func<string, object> parse, object defaultValue)>
            {
                {typeof(int), (str => (int)double.Parse(str, CultureInfo.InvariantCulture), 0)},
                {typeof(float), (str => float.Parse(str, CultureInfo.InvariantCulture), 0)},
                {typeof(double), (str => double.Parse(str, CultureInfo.InvariantCulture), 0)},
                {typeof(decimal), (str => decimal.Parse(str, CultureInfo.InvariantCulture), 0M)},
                {typeof(DateTime), (str => DateTime.Parse(str, CultureInfo.InvariantCulture), DateTime.MinValue)},
                {typeof(string), (str => str, string.Empty)}
            };

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
                throw new FileNotFoundException($"{nameof(filePath)} empty or doesn't exist.");
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
            Type type = typeof(T);
            PropertyInfo[] properties = type.GetProperties();
            var numbers = new Dictionary<string, int>();

            var list = new List<T>();
            List<Row> rows = CalculateFormulas();

            for (int i = 0; i < rows.Count; i++)
            {
                if (i == 0 && ignoreFirstRow)
                {
                    continue;
                }

                Row row = rows[i];
                List<Cell> cells = row.Elements<Cell>().ToList();

                T item = new T();
                foreach (PropertyInfo property in properties)
                {
                    if (!numbers.ContainsKey(property.Name))
                    {
                        object tmp = property.GetCustomAttributes(typeof(ParseDetailsAttribute), false)
                            .FirstOrDefault();
                        numbers[property.Name] = ((ParseDetailsAttribute) tmp)?.Order ?? -1;
                    }
                    int propertyOrder = numbers[property.Name];
                    if (propertyOrder == -1)
                    {
                        continue;
                    }

                    Type propertyType = property.PropertyType;
                    if (!Parsers.ContainsKey(propertyType))
                    {
                        throw new NotSupportedException($@"{nameof(propertyType.Name)} is not supported. Supported types: 
{nameof(Int32)}, {nameof(Single)}, {nameof(Double)}, {nameof(Decimal)}, {nameof(DateTime)}, {nameof(String)}");
                    }

                    string cellValue = GetDataFromCell(cells, row.RowIndex.Value, propertyOrder);
                    object parsedValue;
                    try
                    {
                        parsedValue = Parsers[propertyType].parse.Invoke(cellValue);
                    }
                    catch
                    {
                        parsedValue = Parsers[propertyType].defaultValue;
                    }
                    property.SetValue(item, parsedValue, null);
                }
                list.Add(item);
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

        public void Dispose()
        {
            _spreadsheetDocument.Dispose();
        }
    }
}
