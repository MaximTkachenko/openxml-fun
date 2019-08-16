using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlFun.Excel.Writer.Cells
{
    public class DecimalSumCell : DecimalCell
    {
        public DecimalSumCell(int fromColumnNumber, int toColumnNumber, int fromRowNumber, int toRowNumber) : base(0M)
        {
            CheckIndex(fromColumnNumber);
            FromColumnNumber = fromColumnNumber;

            CheckIndex(toColumnNumber);
            ToColumnNumber = toColumnNumber;

            CheckIndex(fromRowNumber);
            FromRowNumber = fromRowNumber;

            CheckIndex(toRowNumber);
            ToRowNumber = toRowNumber;
        }

        public int FromColumnNumber { get; }
        public int ToColumnNumber { get; }
        public int FromRowNumber { get; }
        public int ToRowNumber { get; }

        internal override void Apply(Cell cell, int columnIndex, uint rowIndex)
        {
            cell.CellFormula = new CellFormula($"SUBTOTAL(9,{ColumnAliases.ExcelColumnNames[FromColumnNumber - 1]}{FromRowNumber}" +
                                               $":{ColumnAliases.ExcelColumnNames[ToColumnNumber - 1]}{FromRowNumber}" +
                                               $":{ColumnAliases.ExcelColumnNames[FromColumnNumber - 1]}{FromRowNumber}" +
                                               $":{ColumnAliases.ExcelColumnNames[FromColumnNumber - 1]}{ToRowNumber})")
            {
                CalculateCell = true
            };
        }
    }
}
