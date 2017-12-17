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

        internal virtual void Apply(Cell cell) { }
    }

    public class SumExcelCell : ExcelCell
    {
        public SumExcelCell(int fromRowNumber, int toRowNumber) : base(0M)
        {
            FromRowNumber = fromRowNumber;
            ToRowNumber = toRowNumber;
        }

        public int FromRowNumber { get; }
        public int ToRowNumber { get; }

        internal override void Apply(Cell cell)
        {
            string columnAlias = GetColumnAlias(cell.CellReference);
            cell.CellFormula = new CellFormula($"SUBTOTAL(9,{columnAlias}{FromRowNumber}:{columnAlias}{ToRowNumber})")
            {
                CalculateCell = true
            };
        }

        private string GetColumnAlias(string cellReference)
        {
            string result = null;
            foreach (var cri in cellReference)
            {
                if (!int.TryParse(cri.ToString(), out var _))
                {
                    result += cri;
                    continue;
                }

                break;
            }

            return result;
        }
    }
}
