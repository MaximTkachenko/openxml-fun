using System;

namespace OpenXmlFun.Excel.Writer
{
    public class ExcelCell
    {
        public Object Value { get; set; }
        public string Formula { get; set; }
        public bool IsStroked { get; set; }
        public bool IsBold { get; set; }
        public string FontColor { get; set; }
        public string BackgroundColor { get; set; }
        public string Hyperlink { get; set; }
    }
}
