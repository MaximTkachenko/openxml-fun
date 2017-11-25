using System;

namespace OpenXmlFun.Excel.Writer
{
    public class ExcelCell
    {
        public ExcelCell()
        {
            FontColor = ExcelColors.Black;
            BackgroundColor = ExcelColors.White;
            EmptyOnDefault = true;
        }

        public Object Value { get; set; }
        public string Formula { get; set; }
        public bool Strike { get; set; }
        public bool Bold { get; set; }
        public ExcelColors FontColor { get; set; }
        public ExcelColors BackgroundColor { get; set; }
        public string Hyperlink { get; set; }
        public bool EmptyOnDefault { get; set; }
    }
}
