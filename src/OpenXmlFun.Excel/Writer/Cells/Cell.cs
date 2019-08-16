using System;

namespace OpenXmlFun.Excel.Writer.Cells
{
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

        protected void CheckIndex(int index)
        {
            if (index <= 0)
            {
                throw new IndexOutOfRangeException("Row or column columnIndex must be more than zero.");
            }
        }
    }
}
