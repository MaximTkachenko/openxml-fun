using System;

namespace OpenXmlFun.Excel.Writer
{
    public class ExcelCell
    {
        public Object Value { get; set; }
        public string Formula { get; set; }
        public bool IsStroked { get; set; }
        public bool IsBold { get; set; }
        public ExcelColors FontColor { get; set; }
        public ExcelColors BackgroundColor { get; set; }
        public string Hyperlink { get; set; }
    }
}
