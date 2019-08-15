using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlFun.Excel.Writer.Cells
{
    public abstract class CellBase
    {
        public bool Strike { get; set; }
        public bool Bold { get; set; }
        public ExcelColors FontColor { get; set; }
        public ExcelColors BackgroundColor { get; set; }
        public HorizontalAlignment HorizontalAlignment { get; set; }
        public VerticalAlignment VerticalAlignment { get; set; }
        public bool EmptyOnDefault { get; set; }

        internal abstract Cell Create(int columnIndex, uint rowIndex);
        internal abstract Type TypeOfValue { get; }
    }
}
