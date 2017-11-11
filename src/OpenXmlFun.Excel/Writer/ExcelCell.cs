namespace OpenXmlFun.Excel.Writer
{
   public class ExcelCell
    {
       public object Text { get; set; }
       public string Hyperlink { get; set; }
       public string CellReference { get; set; }
       public ExcelColors Color { get; set; }
       public bool IsStrike { get; set; }
    }
}
