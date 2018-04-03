using System;
using System.Collections.Generic;

namespace OpenXmlFun.Excel.Writer
{
    internal static class NumberFormats
    {
        //list of predefined NumberFormatId values https://github.com/closedxml/closedxml/wiki/NumberFormatId-Lookup-Table
        public static readonly Dictionary<Type, uint> Data = new Dictionary<Type, uint>
        {
            { typeof(string), 49 },
            { typeof(DateTime), 14 },
            { typeof(decimal), 4 },
            { typeof(int), 1 }
        };
    }
}
