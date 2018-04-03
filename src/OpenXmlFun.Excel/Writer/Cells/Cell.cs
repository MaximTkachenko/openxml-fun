using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlFun.Excel.Writer.Cells
{
    public abstract class Cell<T> : CellBase
    {
        protected Cell(T value)
        {
            Value = value;
            FontColor = ExcelColors.Black;
            BackgroundColor = ExcelColors.White;
            EmptyOnDefault = true;
        }

        public T Value { get; }
        public bool Strike { get; set; }
        public bool Bold { get; set; }
        public ExcelColors FontColor { get; set; }
        public ExcelColors BackgroundColor { get; set; }
        public bool EmptyOnDefault { get; set; }

        internal override Cell Create(int columnIndex, uint rowIndex)
        {
            var cell =  new Cell
            {
                StyleIndex = ExcelStylesheetProvider.GetStyleId(this),
                CellReference = $"{ColumnAliases.ExcelColumnNames[columnIndex]}{rowIndex}"
            };

            Apply(cell, columnIndex, rowIndex);

            return cell;
        }

        protected abstract void Apply(Cell cell, int columnIndex, uint rowIndex);

        protected void CheckIndex(int index)
        {
            if (index <= 0)
            {
                throw new IndexOutOfRangeException("Row or column columnIndex must be more than zero.");
            }
        }
    }
}
