using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
// ReSharper disable PossiblyMistakenUseOfParamsMethod

namespace OpenXmlFun.Excel.Writer
{
    public class ExcelSheet
    {
        private readonly Worksheet _sheet;
        private readonly SheetData _sheetData;
        private readonly ExcelStylesheetProvider _excelStylesheetProvider;
        private int _rowIndex = 1;

        private const string Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        // ReSharper disable once StaticMemberInGenericType
        private static readonly string[] ExcelColumnNames = new string[Alphabet.Length * 2];
        
        static ExcelSheet()
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

        internal ExcelSheet(WorksheetPart sheetPart, ExcelStylesheetProvider excelStylesheetProvider, double[] columnWidths)
        {
            _sheet = sheetPart.Worksheet;
            _sheetData = sheetPart.Worksheet.GetFirstChild<SheetData>();
            _excelStylesheetProvider = excelStylesheetProvider;

            ApplyColumnWidths(columnWidths);
        }

        public ExcelSheet AddHeader(params string[] columnNames)
        {
            return AddRow(columnNames.Select(cn => new ExcelCell
            {
                Value = cn,
                Bold = true,
                FontColor = ExcelColors.White,
                BackgroundColor = ExcelColors.Black
            }).ToArray());
        }

        public ExcelSheet AddRow(params object[] values)
        {
            return AddRow(values.Select(v => new ExcelCell
            {
                Value = v
            }).ToArray());
        }

        public ExcelSheet AddRow(params ExcelCell[] cells)
        {
            var row = new Row { RowIndex = (UInt32)_rowIndex };
            for (int i = 0; i < cells.Length; i++)
            {
                row.AppendChild(CreateCell(cells[i], i));
            }

            _sheetData.AppendChild(row);
            _rowIndex++;

            return this;
        }

        public ExcelSheet Save()
        {
            _sheet.Save();
            return this;
        }

        private Cell CreateCell(ExcelCell sourceCell, int index)
        {
            Cell cell;
            if (sourceCell.Value != null &&
                SupportedTypesDetails.Data.TryGetValue(sourceCell.Value.GetType(), 
                    out (uint NumberFormatId, Func<object, Cell> Factory, Func<object, bool> IsDefault) typeDetails))
            {
                cell = typeDetails.Factory.Invoke(sourceCell.Value);
                if (typeDetails.IsDefault(sourceCell.Value) && sourceCell.EmptyOnDefault)
                {
                    cell.CellValue = new CellValue(string.Empty);
                }
                cell.StyleIndex = _excelStylesheetProvider.GetStyleId(sourceCell);
            }
            else
            {
                sourceCell.Value = string.Empty;
                cell = new Cell
                {
                    DataType = CellValues.String,
                    CellValue = new CellValue(string.Empty),
                    StyleIndex = _excelStylesheetProvider.GetStyleId(sourceCell)
                };
            }
            cell.CellReference = $"{ExcelColumnNames[index]}{_rowIndex}";
            if (!string.IsNullOrWhiteSpace(sourceCell.Hyperlink))
            {
                cell.CellFormula = new CellFormula
                {
                    Space = SpaceProcessingModeValues.Preserve,
                    Text = $@"HYPERLINK(""{sourceCell.Hyperlink}"", ""{sourceCell.Value.ToString()}"")"
                };
            }
            return cell;
        }
        
        private void ApplyColumnWidths(double[] columnWidths)
        {
            if (columnWidths == null || columnWidths.Length == 0)
            {
                return;
            }

            Columns customColumns = new Columns();
            for (uint columnIndex = 0; columnIndex < columnWidths.Length; columnIndex++)
            {
                customColumns.Append(new Column
                {
                    Min = new UInt32Value(columnIndex + 1),
                    Max = new UInt32Value(columnIndex + 1),
                    Width = new DoubleValue(columnWidths[columnIndex]),
                    CustomWidth = true
                });
            }
            for (uint columnIndex = (uint)columnWidths.Length; columnIndex < columnWidths.Length + 25; columnIndex++)
            {
                customColumns.Append(new Column
                {
                    Min = new UInt32Value(columnIndex + 1),
                    Max = new UInt32Value(columnIndex + 1),
                    Width = new DoubleValue(15d),
                    CustomWidth = true
                });
            }
            _sheet.InsertAt(customColumns, 0);
        }
    }
}
