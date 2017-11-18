using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlFun.Excel.Writer;
// ReSharper disable PossiblyMistakenUseOfParamsMethod

namespace OpenXmlFun.Excel.Writer
{
    //todo loks like shit, need rewrite with good api
    public class ExcelWriter : IDisposable
    {
        private const double DefaultColumnWidth = 15;
        private const string DecimalTypeName = "Decimal";
        private const string Int32TypeName = "Int32";
        private const string DatetimeTypeName = "DateTime";
        private const string StringTypeName = "String";

        private readonly SpreadsheetDocument _spreadsheetDocument;
        private SheetData _sheetData;
        Worksheet _worksheet;

        private static readonly Dictionary<string, uint> CellFormatDictionary = new Dictionary<string, uint>
        {
            {"01", 1},
            {"02", 2},
            {"03", 3},
            {"11", 4},
            {"12", 5},
            {"13", 6},
            {"21", 7},
            {"22", 8},
            {"23", 9},
            {"31", 10},
            {"32", 11},
            {"33", 12},
            {"41", 13},
            {"42", 14},
            {"43", 15},
            {"51", 16},
            {"52", 17},
            {"53", 18},
            {"04", 19},
            {"05", 16}
        };

        private int _startIndex = 1;
        private int _sheetId;

        private const string Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        // ReSharper disable once StaticMemberInGenericType
        private static readonly string[] ExcelColumnNames = new string[Alphabet.Length * 2];

        static ExcelWriter()
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

        public ExcelWriter(string filePath)
        {
            _spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
            _spreadsheetDocument.AddWorkbookPart();
            _spreadsheetDocument.WorkbookPart.Workbook = new Workbook();
        }

        public void AddSheet(string sheetName, params double[] columnWidths)
        {
            var worksheetPart = _spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();

            _worksheet?.Save();
            _worksheet = new Worksheet();

            SetWidths(columnWidths);

            worksheetPart.Worksheet = _worksheet;
            worksheetPart.Worksheet.Append(new SheetData());

            if (_spreadsheetDocument.WorkbookPart.Workbook.Sheets == null)
            {
                _spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
            }

            _sheetId = _sheetId + 1;
            _spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().AppendChild(new Sheet
            {
                Id = _spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = (uint)_sheetId,
                Name = sheetName
            });
            _sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            _startIndex = 1;

            if (_spreadsheetDocument.WorkbookPart.WorkbookStylesPart == null)
            {
                _spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                _spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet = CreateStylesheet();
            }
        }

        public void AddAcrossHeader(params string[] columnNames)
        {
            var row = new Row { RowIndex = (UInt32)_startIndex };
            for (int dataItemIndex = 0; dataItemIndex < columnNames.Length; dataItemIndex++)
                row.AppendChild(AddCell(new ExcelCell { Text = columnNames[dataItemIndex] }, _startIndex, ExcelColumnNames[dataItemIndex], type: ExcelDataTypes.Header));
            _sheetData.AppendChild(row);

            var autoFilter = new AutoFilter { Reference = string.Format("A1:{0}1", ExcelColumnNames[columnNames.Length - 1]) };
            _worksheet.Append(autoFilter);

            _startIndex++;
        }

        public void AddRow(ExcelColors color, params ExcelCell[] data)
        {
            Func<ExcelColors, ExcelCell, int, string, Cell> newCell = (cellcolor, cellData, index, headerLetter) =>
                {
                    if (cellData == null)
                        cellData = new ExcelCell { Text = string.Empty };
                    if (cellData.Text == null)
                        return AddCell(new ExcelCell { Text = string.Empty }, index, headerLetter, cellcolor, ExcelDataTypes.Text);

                    string dataType = cellData.Text.GetType().Name;
                    string text = string.Empty;
                    var type = ExcelDataTypes.Text;
                    switch (dataType)
                    {
                        case DecimalTypeName:
                            type = ExcelDataTypes.Money;
                            var number = (decimal)cellData.Text;
                            text = number == 0 ? string.Empty : GetExcelDecimalString(number);
                            break;
                        case Int32TypeName:
                            type = ExcelDataTypes.Text;
                            text = ((int)cellData.Text).ToString();
                            break;
                        case DatetimeTypeName:
                            type = ExcelDataTypes.Date;
                            text = ((DateTime)cellData.Text).ToShortDateString();
                            break;
                        case StringTypeName:
                            type = ExcelDataTypes.Text;
                            text = (string)cellData.Text;
                            break;
                    }
                    return AddCell(cellData, index, headerLetter, cellcolor, type);
                };
            var row = new Row { RowIndex = (UInt32)_startIndex };
            for (int dataItemIndex = 0; dataItemIndex < data.Length; dataItemIndex++)
            {
                row.AppendChild(newCell(data[dataItemIndex].Color == 0 ? color : data[dataItemIndex].Color, data[dataItemIndex], _startIndex, ExcelColumnNames[dataItemIndex]));
            }

            _sheetData.AppendChild(row);
            _startIndex++;
        }

        private Cell AddCell(ExcelCell text, int index, string header, ExcelColors color = ExcelColors.Black, ExcelDataTypes type = ExcelDataTypes.Text)
        {
            Cell cell = null;
            type = text.IsStrike ? ExcelDataTypes.Strike : type;

            switch (type)
            {
                case ExcelDataTypes.Money:
                    cell = new Cell
                    {
                        DataType = CellValues.Number,
                        CellReference = header + index.ToString(),
                        CellValue = new CellValue() { Text = text.Text.ToString() },
                        StyleIndex = CellFormatDictionary[((int)color).ToString() + ((int)type).ToString()]
                    };
                    break;
                case ExcelDataTypes.Strike:
                    cell = new Cell
                    {
                        DataType = CellValues.InlineString,
                        CellReference = header + index.ToString(),
                        StyleIndex = CellFormatDictionary[((int)color).ToString() + ((int)type).ToString()]
                    };
                    if (string.IsNullOrEmpty(text.Hyperlink))
                        cell.InlineString = new InlineString() { Text = new Text(text.Text.ToString()) };
                    break;
                case ExcelDataTypes.Header:
                case ExcelDataTypes.Text:
                    cell = new Cell()
                    {
                        DataType = CellValues.InlineString,
                        CellReference = header + index.ToString(),
                        StyleIndex = CellFormatDictionary[((int)color).ToString() + ((int)type).ToString()]
                    };

                    if (string.IsNullOrEmpty(text.Hyperlink))
                        cell.InlineString = new InlineString() { Text = new Text(text.Text.ToString()) };
                    break;
            }

            if (!string.IsNullOrEmpty(text.Hyperlink))
            {
                var cellValue1 = new CellValue();
                var cellFormula1 = new CellFormula
                    {
                        Space = SpaceProcessingModeValues.Preserve,
                        Text = @"HYPERLINK(""" + text.Hyperlink + @""", """ + text.Text.ToString().Replace("\"", "'") + @""")"
                    };
                cellValue1.Text = text.Text.ToString();
                cell.Append(cellFormula1);
                cell.Append(cellValue1);
            }


            return cell;

        }

        private void SetWidths(params double[] columnWidths)
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
                    Width = new DoubleValue(DefaultColumnWidth),
                    CustomWidth = true
                });
            }
            _worksheet.Append(customColumns);
        }

        private Stylesheet CreateStylesheet()
        {
            Stylesheet styleSheet = new Stylesheet();

            var fonts = new Fonts();
            fonts.Append(new Font());
            //red
            fonts.Append(new Font
            {
                Color = new Color
                {
                    Rgb = "FF003C"
                }
            });
            //green
            fonts.Append(new Font
            {
                Color = new Color
                {
                    Rgb = "32CD32"
                }
            });
            //blue
            fonts.Append(new Font
            {
                Color = new Color
                {
                    Rgb = "4300FF"
                }
            });

            fonts.Append(new Font
            {
                Strike = new Strike(),
                Color = new Color
                {
                    Rgb = "aaaaaa"
                }
            });
            fonts.Count = (uint)fonts.ChildElements.Count;

            Fills fills = new Fills();
            fills.Append(new Fill
            {
                PatternFill = new PatternFill { PatternType = PatternValues.None }
            });
            fills.Append(new Fill
            {
                PatternFill = new PatternFill { PatternType = PatternValues.Gray125 }
            });

            fills.Append(new Fill
            {
                PatternFill = new PatternFill
                {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("d1edee") },
                    BackgroundColor = new BackgroundColor { Rgb = HexBinaryValue.FromString("d1edee") }
                }
            });
            fills.Count = (uint)fills.ChildElements.Count;

            Borders borders = new Borders();
            borders.Append(new Border());
            borders.Append(new Border
            {
                LeftBorder = new LeftBorder
                {
                    Style = BorderStyleValues.Medium,
                    Color = new Color
                    {
                        Indexed = (UInt32Value)64U
                    }
                },
                RightBorder = new RightBorder
                {
                    Style = BorderStyleValues.Medium,
                    Color = new Color
                    {
                        Indexed = (UInt32Value)64U
                    }
                },
                TopBorder = new TopBorder
                {
                    Style = BorderStyleValues.Medium,
                    Color = new Color
                    {
                        Indexed = (UInt32Value)64U
                    }
                },
                BottomBorder = new BottomBorder
                {
                    Style = BorderStyleValues.Medium,
                    Color = new Color
                    {
                        Indexed = (UInt32Value)64U
                    }
                },
                DiagonalBorder = new DiagonalBorder()
            });
            borders.Count = (uint)borders.ChildElements.Count;

            var csFormats = new CellStyleFormats();
            csFormats.Append(new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                Alignment = new Alignment { WrapText = true }
            });
            csFormats.Count = (uint)csFormats.ChildElements.Count;

            var numFormats = new NumberingFormats();
            var cellFormats = new CellFormats();

            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                Alignment = new Alignment { WrapText = true }

            });

            numFormats.Append(new NumberingFormat()
            {
                NumberFormatId = 164,
                FormatCode = "#,##0.00"
            });
            #region cellFormats
            #region black
            //header
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 1,
                FormatId = 0,
                ApplyAlignment = true,
                Alignment = new Alignment
                {
                    Vertical = VerticalAlignmentValues.Center,
                    Horizontal = HorizontalAlignmentValues.Center,
                    WrapText = true
                }
            });
            //text
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 1,
                FormatId = 0,
                ApplyAlignment = true,
                Alignment = new Alignment
                {
                    Vertical = VerticalAlignmentValues.Center,
                    WrapText = true
                }
            });
            //money
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 164,
                FontId = 0,
                FillId = 0,
                BorderId = 1,
                FormatId = 0,
                ApplyAlignment = true,
                Alignment = new Alignment
                {
                    Vertical = VerticalAlignmentValues.Center,
                    WrapText = true
                }
            });
            #endregion
            #region red
            //header
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 0,
                FontId = 1,
                FillId = 0,
                BorderId = 1,
                FormatId = 0,
                ApplyAlignment = true,
                Alignment = new Alignment
                {
                    Vertical = VerticalAlignmentValues.Center,
                    Horizontal = HorizontalAlignmentValues.Center,
                    WrapText = true
                }
            });
            //text
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 0,
                FontId = 1,
                FillId = 0,
                BorderId = 1,
                FormatId = 0,
                ApplyAlignment = true,
                Alignment = new Alignment
                {
                    Vertical = VerticalAlignmentValues.Center,
                    WrapText = true
                }
            });
            //money
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 164,
                FontId = 1,
                FillId = 0,
                BorderId = 1,
                FormatId = 0,
                ApplyAlignment = true,
                Alignment = new Alignment
                {
                    Vertical = VerticalAlignmentValues.Center,
                    WrapText = true
                }
            });
            #endregion
            #region green
            //header
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 0,
                FontId = 2,
                FillId = 0,
                BorderId = 1,
                FormatId = 0,
                ApplyAlignment = true,
                Alignment = new Alignment
                {
                    Vertical = VerticalAlignmentValues.Center,
                    Horizontal = HorizontalAlignmentValues.Center,
                    WrapText = true
                }
            });
            //text
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 0,
                FontId = 2,
                FillId = 0,
                BorderId = 1,
                FormatId = 0,
                ApplyAlignment = true,
                Alignment = new Alignment
                {
                    Vertical = VerticalAlignmentValues.Center,
                    WrapText = true
                }
            });
            //money
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 164,
                FontId = 2,
                FillId = 0,
                BorderId = 1,
                FormatId = 0,
                ApplyAlignment = true,
                Alignment = new Alignment
                {
                    Vertical = VerticalAlignmentValues.Center,
                    WrapText = true
                }
            });
            #endregion
            #region blue
            //header
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 0,
                FontId = 3,
                FillId = 0,
                BorderId = 1,
                FormatId = 0,
                ApplyAlignment = true,
                Alignment = new Alignment
                {
                    Vertical = VerticalAlignmentValues.Center,
                    Horizontal = HorizontalAlignmentValues.Center,
                    WrapText = true
                }
            });
            //text
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 0,
                FontId = 3,
                FillId = 0,
                BorderId = 1,
                FormatId = 0,
                ApplyAlignment = true,
                Alignment = new Alignment
                {
                    Vertical = VerticalAlignmentValues.Center,
                    WrapText = true
                }
            });
            //money
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 164,
                FontId = 3,
                FillId = 0,
                BorderId = 1,
                FormatId = 0,
                ApplyAlignment = true,
                Alignment = new Alignment
                {
                    Vertical = VerticalAlignmentValues.Center,
                    WrapText = true
                }
            });
            #endregion
            #region background
            //header
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 2,
                BorderId = 1,
                FormatId = 0,
                ApplyAlignment = true,
                Alignment = new Alignment
                {
                    Vertical = VerticalAlignmentValues.Center,
                    Horizontal = HorizontalAlignmentValues.Center,
                    WrapText = true
                }
            });
            //text
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 2,
                BorderId = 1,
                FormatId = 0,
                ApplyAlignment = true,
                Alignment = new Alignment
                {
                    Vertical = VerticalAlignmentValues.Center,
                    WrapText = true
                }
            });
            //money
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 164,
                FontId = 0,
                FillId = 2,
                BorderId = 1,
                FormatId = 0,
                ApplyAlignment = true,
                Alignment = new Alignment
                {
                    Vertical = VerticalAlignmentValues.Center,
                    WrapText = true
                }
            });
            //Strike
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 0,
                FontId = 4,
                FillId = 0,
                BorderId = 1,
                FormatId = 0,
                ApplyAlignment = true,
                Alignment = new Alignment
                {
                    Vertical = VerticalAlignmentValues.Center,
                    WrapText = true
                }
            });
            #endregion
            #endregion

            numFormats.Count = (uint)numFormats.ChildElements.Count;
            cellFormats.Count = (uint)cellFormats.ChildElements.Count;

            styleSheet.Append(numFormats);
            styleSheet.Append(fonts);
            styleSheet.Append(fills);
            styleSheet.Append(borders);
            styleSheet.Append(csFormats);
            styleSheet.Append(cellFormats);

            return styleSheet;
        }

        private static string GetExcelDecimalString(decimal dec)
        {
            if (dec == 0)
                return string.Empty;
            var regex = new Regex(@"\s+");
            return regex.Replace(Math.Round(dec, 2).ToString("N", new System.Globalization.CultureInfo("ru-RU")).Replace(',', '.'), string.Empty);
        }

        public void Dispose()
        {
            _spreadsheetDocument.Dispose();
        }
    }
}