using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace OpenXmlFun.Excel.Writer
{
    internal static class SupportedTypesDetails
    {
        //list of predefined NumberFormatId values https://github.com/closedxml/closedxml/wiki/NumberFormatId-Lookup-Table
        public static readonly Dictionary<Type, (uint NumberFormatId, Func<object, Cell> Factory)> Data = new Dictionary<Type, (uint NumberFormatId, Func<object, Cell> Factory)>
        {
            { typeof(string),
                (49, value => new Cell
                {
                    DataType = CellValues.String,
                    CellValue = new CellValue((string)value)
                }) },
            { typeof(DateTime),
                (14, value => new Cell
                {
                    DataType = CellValues.Number,
                    CellValue = new CellValue(((DateTime)value).ToOADate().ToString(CultureInfo.InvariantCulture))
                }) },
            { typeof(decimal),
                (2, value => new Cell
                {
                    DataType = CellValues.Number,
                    CellValue = new CellValue(value.ToString())
                }) },
            { typeof(int),
                (1, value => new Cell
                {
                    DataType = CellValues.Number,
                    CellValue = new CellValue(value.ToString())
                }) }
        };
    }
}
