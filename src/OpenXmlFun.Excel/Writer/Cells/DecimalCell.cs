using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlFun.Excel.Writer.Cells
{
    public class DecimalCell : Cell<decimal>
    {
        public DecimalCell(decimal value) : base(value)
        { }

        protected override void Apply(Cell cell, int columnIndex, uint rowIndex)
        {
            cell.DataType = CellValues.Number;
            cell.CellValue = new CellValue(Value == 0M ? string.Empty : Value.ToString(CultureInfo.InvariantCulture));
        }
    }
}
