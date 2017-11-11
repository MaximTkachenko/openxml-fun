using System;

namespace OpenXmlFun.Excel.Parser
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ParseDetailsAttribute : Attribute
    {
        public ParseDetailsAttribute(int order)
        {
            if (order < 0)
            {
                throw new ArgumentException($"{nameof(order)} can't be less than zero.");
            }
            Order = order;
        }

        public int Order { get; }
    }
}
