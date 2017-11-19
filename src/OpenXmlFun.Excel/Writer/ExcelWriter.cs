using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Border = DocumentFormat.OpenXml.Spreadsheet.Border;
using BottomBorder = DocumentFormat.OpenXml.Spreadsheet.BottomBorder;
using Color = DocumentFormat.OpenXml.Spreadsheet.Color;
using Column = DocumentFormat.OpenXml.Spreadsheet.Column;
using Columns = DocumentFormat.OpenXml.Spreadsheet.Columns;
using Font = DocumentFormat.OpenXml.Spreadsheet.Font;
using Fonts = DocumentFormat.OpenXml.Spreadsheet.Fonts;
using HorizontalAlignmentValues = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues;
using LeftBorder = DocumentFormat.OpenXml.Spreadsheet.LeftBorder;
using NumberingFormat = DocumentFormat.OpenXml.Spreadsheet.NumberingFormat;
using RightBorder = DocumentFormat.OpenXml.Spreadsheet.RightBorder;
using Strike = DocumentFormat.OpenXml.Spreadsheet.Strike;
using TopBorder = DocumentFormat.OpenXml.Spreadsheet.TopBorder;
using VerticalAlignmentValues = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues;

// ReSharper disable PossiblyMistakenUseOfParamsMethod

namespace OpenXmlFun.Excel.Writer
{
    /// <summary>
    /// https://stackoverflow.com/questions/2792304/how-to-insert-a-date-to-an-open-xml-worksheet
    /// http://www.dispatchertimer.com/tutorial/how-to-create-an-excel-file-in-net-using-openxml-part-3-add-stylesheet-to-the-spreadsheet/
    /// </summary>
    public class ExcelWriter : IDisposable
    {
        private readonly SpreadsheetDocument _spreadsheetDocument;
        private readonly Dictionary<string, ExcelSheet> _sheets;

        public ExcelWriter(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException($"Specify {nameof(filePath)}.");
            }
            _spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
            _spreadsheetDocument.AddWorkbookPart();
            _spreadsheetDocument.WorkbookPart.Workbook = new Workbook();

            _sheets = new Dictionary<string, ExcelSheet>();
        }

        public ExcelSheet CreateSheet(string name, params double[] columnWidths)
        {
            foreach (var key in _sheets.Keys)
            {
                _sheets[key].Save();
            }

            var worksheetPart = _spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet();
            worksheetPart.Worksheet.Append(new SheetData());

            if (_spreadsheetDocument.WorkbookPart.Workbook.Sheets == null)
            {
                _spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
            }

            _spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().AppendChild(new Sheet
            {
                Id = _spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = (uint)_sheets.Keys.Count + 1,
                Name = name
            });

            _spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            _spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet = CreateStylesheet();

            _sheets[name] = new ExcelSheet(worksheetPart, columnWidths);
            return _sheets[name];
        }

        private Stylesheet CreateStylesheet()
        {
            var styleSheet = new Stylesheet();
            styleSheet.CellFormats = new CellFormats
            (
                new CellFormat
                {
                    ApplyAlignment = true,
                    Alignment = new Alignment { WrapText = true }
                },
                //date
                new CellFormat
                {
                    ApplyNumberFormat = true,
                    NumberFormatId = 14,
                    ApplyAlignment = true,
                    Alignment = new Alignment { WrapText = true }
                },
                //text
                new CellFormat
                {
                    ApplyNumberFormat = true,
                    NumberFormatId = 49,
                    ApplyAlignment = true,
                    Alignment = new Alignment { WrapText = true }
                },
                //decimal
                new CellFormat
                {
                    ApplyNumberFormat = true,
                    NumberFormatId = 4,
                    ApplyAlignment = true,
                    Alignment = new Alignment { WrapText = true }
                },
                //int
                new CellFormat
                {
                    ApplyNumberFormat = true,
                    NumberFormatId = 1,
                    ApplyAlignment = true,
                    Alignment = new Alignment { WrapText = true }
                }
            );
            return styleSheet;
        }

        public void Dispose()
        {
            _spreadsheetDocument.Dispose();
        }
    }

    public class ExcelSheet
    {
        private readonly Worksheet _sheet;
        private readonly SheetData _sheetData;
        private int _rowIndex = 1;

        private const string Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        // ReSharper disable once StaticMemberInGenericType
        private static readonly string[] ExcelColumnNames = new string[Alphabet.Length * 2];

        private static readonly Dictionary<Type, Func<object, Cell>> CellFactories = new Dictionary<Type, Func<object, Cell>>
        {
            {
                typeof(int), value => new Cell
                {
                    DataType = CellValues.Number,
                    CellValue = new CellValue(value.ToString()),
                    StyleIndex = 4
                }
            },
            {
                typeof(decimal), value => new Cell
                {
                    DataType = CellValues.Number,
                    CellValue = new CellValue(value.ToString()),
                    StyleIndex = 3
                }
            },
            {
                typeof(DateTime), value => new Cell
                {
                    DataType = CellValues.Number,
                    CellValue = new CellValue(((DateTime)value).ToOADate().ToString(CultureInfo.InvariantCulture)),
                    StyleIndex = 1
                }
            },
            {
                typeof(string), value => new Cell
                {
                    DataType = CellValues.String,
                    CellValue = new CellValue(value.ToString()),
                    StyleIndex = 2
                }
            }
        };

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

        internal ExcelSheet(WorksheetPart sheetPart, double[] columnWidths)
        {
            _sheet = sheetPart.Worksheet;
            _sheetData = sheetPart.Worksheet.GetFirstChild<SheetData>();

            ApplyColumnWidths(columnWidths);
        }

        public ExcelSheet AddHeader(params string[] columnNames)
        {
            return AddRow(columnNames.Select(cn => new ExcelCell{Value = cn}).ToArray());
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

        private Cell CreateCell(ExcelCell sourceSell, int index)
        {
            Cell cell;
            if (CellFactories.ContainsKey(sourceSell.Value.GetType()))
            {
                cell = CellFactories[sourceSell.Value.GetType()].Invoke(sourceSell.Value);
            }
            else
            {
                cell = new Cell
                {
                    DataType = CellValues.String,
                    CellValue = new CellValue(sourceSell.Value.ToString())
                };
            }
            cell.CellReference = $"{ExcelColumnNames[index]}{_rowIndex}";
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
            _sheet.Append(customColumns);
        }
    }

    public class ExcelCell
    {
        public Object Value { get; set; }
        public string Formula { get; set; }
        public bool IsStroked { get; set; }
        public bool IsBold { get; set; }
        public string FontColor { get; set; } 
        public string BackgroundColor { get; set; }
        public string Hyperlink { get; set; }
    }
}
