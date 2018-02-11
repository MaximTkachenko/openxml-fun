using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlFun.Excel.Writer
{
    public class ExcelCell
    {
        public ExcelCell(object value)
        {
            Value = value;
            FontColor = ExcelColors.Black;
            BackgroundColor = ExcelColors.White;
            EmptyOnDefault = true;
        }

        public object Value { get; }
        public bool Strike { get; set; }
        public bool Bold { get; set; }
        public ExcelColors FontColor { get; set; }
        public ExcelColors BackgroundColor { get; set; }
        public bool EmptyOnDefault { get; set; }

        internal virtual void Apply(Cell cell, string columnAlias, uint rowIndex) { }

        protected void CheckIndex(int index)
        {
            if (index <= 0)
            {
                throw new IndexOutOfRangeException("Row or column index must be more than zero.");
            }
        }
    }

    public class DecimalVerticalSumExcelCell : ExcelCell
    {
        public DecimalVerticalSumExcelCell(int fromRowNumber, int toRowNumber) : base(0M)
        {
            CheckIndex(fromRowNumber);
            FromRowNumber = fromRowNumber;

            CheckIndex(toRowNumber);
            ToRowNumber = toRowNumber;
        }

        public int FromRowNumber { get; }
        public int ToRowNumber { get; }

        internal override void Apply(Cell cell, string columnAlias, uint rowIndex)
        {
            cell.CellFormula = new CellFormula($"SUBTOTAL(9,{columnAlias}{FromRowNumber}:{columnAlias}{ToRowNumber})")
            {
                CalculateCell = true
            };
        }
    }

    public class DecimalHorizontalSumExcelCell : ExcelCell
    {
        public DecimalHorizontalSumExcelCell(int fromColumnNumber, int toColumnNumber) : base(0M)
        {
            CheckIndex(fromColumnNumber);
            FromColumnNumber = fromColumnNumber;

            CheckIndex(toColumnNumber);
            ToColumnNumber = toColumnNumber;
        }

        public int FromColumnNumber { get; }
        public int ToColumnNumber { get; }

        internal override void Apply(Cell cell, string columnAlias, uint rowIndex)
        {
            cell.CellFormula = new CellFormula($"SUBTOTAL(9,{ColumnAliases.ExcelColumnNames[FromColumnNumber - 1]}{rowIndex}:{ColumnAliases.ExcelColumnNames[ToColumnNumber - 1]}{rowIndex})")
            {
                CalculateCell = true
            };
        }
    }

    public class DecimalSumExcelCell : ExcelCell
    {
        public DecimalSumExcelCell(int fromColumnNumber, int toColumnNumber, int fromRowNumber, int toRowNumber) : base(0M)
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

        internal override void Apply(Cell cell, string columnAlias, uint rowIndex)
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
