using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlFun.Excel.Writer.Cells
{
    public class DecimalHorizontalSumCell : DecimalCell
    {
        public DecimalHorizontalSumCell(int fromColumnNumber, int toColumnNumber) : base(0M)
        {
            CheckIndex(fromColumnNumber);
            FromColumnNumber = fromColumnNumber;

            CheckIndex(toColumnNumber);
            ToColumnNumber = toColumnNumber;
        }

        public int FromColumnNumber { get; }
        public int ToColumnNumber { get; }

        internal override void Apply(Cell cell, int columnIndex, uint rowIndex)
        {
            cell.CellFormula = new CellFormula($"SUBTOTAL(9,{ColumnAliases.ExcelColumnNames[FromColumnNumber - 1]}{rowIndex}:{ColumnAliases.ExcelColumnNames[ToColumnNumber - 1]}{rowIndex})")
            {
                CalculateCell = true
            };
        }
    }
}
