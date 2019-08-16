using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlFun.Excel.Writer.Cells
{
    public class DecimalVerticalSumCell : DecimalCell
    {
        public DecimalVerticalSumCell(int fromRowNumber, int toRowNumber) : base(0M)
        {
            CheckIndex(fromRowNumber);
            FromRowNumber = fromRowNumber;

            CheckIndex(toRowNumber);
            ToRowNumber = toRowNumber;
        }

        public int FromRowNumber { get; }
        public int ToRowNumber { get; }

        internal override void Apply(Cell cell, int columnIndex, uint rowIndex)
        {
            string columnAlias = ColumnAliases.ExcelColumnNames[columnIndex];
            cell.CellFormula = new CellFormula($"SUBTOTAL(9,{columnAlias}{FromRowNumber}:{columnAlias}{ToRowNumber})")
            {
                CalculateCell = true
            };
        }
    }
}
