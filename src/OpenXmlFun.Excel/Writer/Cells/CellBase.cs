using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlFun.Excel.Writer.Cells
{
    public abstract class CellBase
    {
        internal abstract Cell Create(int columnIndex, uint rowIndex);
    }
}
