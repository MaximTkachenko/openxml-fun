using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlFun.Excel.Writer.Cells
{
    public class StringCell : Cell<string>
    {
        public StringCell(string value) : base(value)
        { }

        internal override void Apply(Cell cell, int columnIndex, uint rowIndex)
        {
            cell.DataType = CellValues.String;
            cell.CellValue = new CellValue(string.IsNullOrWhiteSpace(Value) ? string.Empty : Value);
        }
    }
}
