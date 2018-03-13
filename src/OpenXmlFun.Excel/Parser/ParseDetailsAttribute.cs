using System;

namespace OpenXmlFun.Excel.Parser
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ParseDetailsAttribute : Attribute
    {
        public ParseDetailsAttribute(int order)
        {
            Order = order < 0
                ? throw new ArgumentException($"{nameof(order)} can't be less than zero.")
                : order;
        }

        public int Order { get; }
    }
}
