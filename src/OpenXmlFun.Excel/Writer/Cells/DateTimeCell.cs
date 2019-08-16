using System;
using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlFun.Excel.Writer.Cells
{
    public class DateTimeCell : Cell<DateTime>
    {
        public DateTimeCell(DateTime value) : base(value)
        { }

        internal override void Apply(Cell cell, int columnIndex, uint rowIndex)
        {
            cell.DataType = CellValues.Number;
            cell.CellValue = new CellValue(Value == DateTime.MinValue ? string.Empty : Value.ToOADate().ToString(CultureInfo.InvariantCulture));
        }
    }
}
