﻿using System;
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

        internal ExcelSheet(WorksheetPart sheetPart, ExcelStylesheetProvider excelStylesheetProvider)
        {
            _sheet = sheetPart.Worksheet;
            _sheetData = sheetPart.Worksheet.GetFirstChild<SheetData>();
            _excelStylesheetProvider = excelStylesheetProvider;
        }

        public ExcelSheet AddHeader(params string[] columnNames)
        {
            return AddRow(columnNames?.Select(cn => new ExcelCell
            {
                Value = cn,
                Bold = true,
                FontColor = ExcelColors.White,
                BackgroundColor = ExcelColors.Black
            }).ToArray());
        }

        public ExcelSheet AddRow(params object[] values)
        {
            return AddRow(values?.Select(v => new ExcelCell
            {
                Value = v
            }).ToArray());
        }

        public ExcelSheet AddRow(params ExcelCell[] cells)
        {
            var row = new Row { RowIndex = (UInt32)_rowIndex };
            if (cells != null && cells.Length > 0)
            {
                for (int i = 0; i < cells.Length; i++)
                {
                    row.AppendChild(CreateCell(cells[i], i));
                }
            }

            _sheetData.AppendChild(row);
            _rowIndex++;

            return this;
        }

        public ExcelSheet ApplyColumnWidths(params double[] columnWidths)
        {
            var columns = _sheet.GetFirstChild<Columns>();
            if (columns != null)
            {
                columns.RemoveAllChildren();
            }
            else
            {
                columns = new Columns();
                _sheet.InsertAt(columns, 0);
            }

            if (columnWidths != null && columnWidths.Length > 0)
            {
                for (uint columnIndex = 0; columnIndex < columnWidths.Length; columnIndex++)
                {
                    columns.Append(new Column
                    {
                        Min = new UInt32Value(columnIndex + 1),
                        Max = new UInt32Value(columnIndex + 1),
                        Width = new DoubleValue(columnWidths[columnIndex]),
                        CustomWidth = true
                    });
                }
                for (uint columnIndex = (uint)columnWidths.Length; columnIndex < columnWidths.Length + 25; columnIndex++)
                {
                    columns.Append(new Column
                    {
                        Min = new UInt32Value(columnIndex + 1),
                        Max = new UInt32Value(columnIndex + 1),
                        Width = new DoubleValue(15d),
                        CustomWidth = true
                    });
                }
            }

            return this;
        }

        public ExcelSheet FreezeTopNRows(int firstNRows)
        {
            if (firstNRows <= 0)
            {
                return this;
            }

            var sheetViews = _sheet.GetFirstChild<SheetViews>();
            if (sheetViews != null)
            {
                sheetViews.RemoveAllChildren();
            }
            else
            {
                sheetViews = new SheetViews();
                var columns = _sheet.GetFirstChild<Columns>();
                if (columns == null)
                {
                    ApplyColumnWidths();
                    columns = _sheet.GetFirstChild<Columns>();
                }
                _sheet.InsertBefore(sheetViews, columns);
            }

            var sheetView = new SheetView { TabSelected = true, WorkbookViewId = 0U };
            var pane = new Pane
            {
                VerticalSplit = firstNRows,
                TopLeftCell = $"A{firstNRows + 1}",
                ActivePane = PaneValues.BottomLeft,
                State = PaneStateValues.Frozen
            };
            var selection = new Selection
            {
                Pane = PaneValues.BottomLeft,
                ActiveCell = "A1",
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1" }
            };
            sheetView.Append(pane);
            sheetView.Append(selection);
            sheetViews.Append(sheetView);

            return this;
        }

        internal ExcelSheet Save()
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
                if (typeDetails.IsDefault.Invoke(sourceCell.Value) && sourceCell.EmptyOnDefault)
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
    }
}
