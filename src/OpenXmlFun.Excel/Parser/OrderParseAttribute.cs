using System;

namespace OpenXmlFun.Excel.Parser
{
    [AttributeUsage(AttributeTargets.Property)]
    public class OrderParseAttribute : Attribute
    {
        public OrderParseAttribute(int number)
        {
            if (number < 0)
            {
                throw new ArgumentException($"{nameof(number)} can't be less than zero.");
            }
            Number = number;
        }

        public int Number { get; }
    }
}
