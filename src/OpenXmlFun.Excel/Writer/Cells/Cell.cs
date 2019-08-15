using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlFun.Excel.Writer.Cells
{
    public enum HorizontalAlignment
    {
        Left,
        Center,
        Right
    }

    public enum VerticalAlignment
    {
        Top,
        Center,
        Bottom
    }

    public abstract class Cell<T> : CellBase
    {
        protected Cell(T value)
        {
            Value = value;
            FontColor = ExcelColors.Black;
            BackgroundColor = ExcelColors.White;
            HorizontalAlignment = value is decimal || value is DateTime ? HorizontalAlignment.Right : HorizontalAlignment.Left;
            VerticalAlignment = VerticalAlignment.Top;
            EmptyOnDefault = true;
        }

        public T Value { get; }

        internal override Type TypeOfValue => typeof(T);

        internal override Cell Create(int columnIndex, uint rowIndex)
        {
            var cell =  new Cell
            {
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
