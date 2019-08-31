using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlFun.Excel.Writer.Cells;

// ReSharper disable PossiblyMistakenUseOfParamsMethod

namespace OpenXmlFun.Excel.Writer
{
    public class ExcelSheet
    {
        private readonly Worksheet _sheet;
        private readonly SheetData _sheetData;
        private readonly ExcelStylesheetProvider _styles;
        private uint _rowIndex = 1;

        internal ExcelSheet(WorksheetPart sheetPart, ExcelStylesheetProvider styles)
        {
            _sheet = sheetPart.Worksheet;
            _sheetData = sheetPart.Worksheet.GetFirstChild<SheetData>();
            _styles = styles;

            //set default columns' widths
            ApplyColumnWidths(20);
        }

        public ExcelSheet AddHeader(IEnumerable<string> columnNames)
        {
            return AddHeader(columnNames?.ToArray());
        }

        public ExcelSheet AddHeader(params string[] columnNames)
        {
            return AddRow(columnNames?.Select(cn => new StringCell(cn)
            {
                Bold = true,
                FontColor = ExcelColors.White,
                BackgroundColor = ExcelColors.Black
            }));
        }

        public ExcelSheet AddRow(IEnumerable<CellBase> cells)
        {
            return AddRow(cells?.ToArray());
        }

        public ExcelSheet AddRow(params CellBase[] cells)
        {
            var row = new Row { RowIndex = _rowIndex };
            if (cells != null && cells.Length > 0)
            {
                for (int i = 0; i < cells.Length; i++)
                {
                    var cell = new Cell
                    {
                        CellReference = $"{ColumnAliases.ExcelColumnNames[i]}{_rowIndex}",
                        StyleIndex = _styles.GetStyleId(cells[i])
                    };

                    cells[i].Apply(cell, i, _rowIndex);
                    row.AppendChild(cell);
                }
            }

            _sheetData.AppendChild(row);
            _rowIndex++;

            return this;
        }

        public ExcelSheet ApplyColumnWidths(params double[] columnWidths)
        {
            if (columnWidths == null || columnWidths.Length == 0)
            {
                throw new ArgumentException($"{nameof(columnWidths)} should not be empty.");
            }

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

            return this;
        }

        public ExcelSheet FreezeTopNRows(int firstNRows)
        {
            if (firstNRows <= 0)
            {
                throw new ArgumentException($"{nameof(firstNRows)} must be greater than zero.");
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

        public ExcelSheet AddColumnFilter(int columnNumber, int belowRowNumber)
        {
            AddColumnFilters(columnNumber, columnNumber, belowRowNumber);
            return this;
        }

        public ExcelSheet AddColumnFilters(int fromColumn, int toColumn, int belowRow)
        {
            if (fromColumn <= 0)
            {
                throw new ArgumentException($"{nameof(fromColumn)} must be greater than zero.");
            }

            if (toColumn <= 0)
            {
                throw new ArgumentException($"{nameof(toColumn)} must be greater than zero.");
            }

            if (belowRow <= 0)
            {
                throw new ArgumentException($"{nameof(belowRow)} must be greater than zero.");
            }

            _sheet.Append(new AutoFilter { Reference = $"{ColumnAliases.ExcelColumnNames[fromColumn - 1]}{belowRow}:{ColumnAliases.ExcelColumnNames[toColumn - 1]}{belowRow}" });
            return this;
        }

        public ExcelSheet MergeCells(int fromRow, int toRow, int fromColumn, int toColumn)
        {
            var mergeCells = _sheet.Elements<MergeCells>().FirstOrDefault();
            if (mergeCells == null)
            {
                mergeCells = new MergeCells();
                _sheet.InsertAfter(mergeCells, _sheet.Elements<SheetData>().First());
            }
            
            mergeCells.Append(new MergeCell { Reference = new StringValue($"{ColumnAliases.ExcelColumnNames[fromColumn - 1]}{fromRow}:{ColumnAliases.ExcelColumnNames[toColumn - 1]}{toRow}") });
            return this;
        }

        internal ExcelSheet Save()
        {
            _sheet.Save();
            return this;
        }
    }
}
