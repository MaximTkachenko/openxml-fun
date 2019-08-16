using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlFun.Excel.Writer.Cells
{
    public class IntegerCell : Cell<int>
    {
        public IntegerCell(int value) : base(value)
        { }

        internal override void Apply(Cell cell, int columnIndex, uint rowIndex)
        {
            cell.DataType = CellValues.Number;
            cell.CellValue = new CellValue(Value == 0 ? string.Empty : Value.ToString(CultureInfo.InvariantCulture));
        }
    }
}
