using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
// ReSharper disable PossiblyMistakenUseOfParamsMethod

namespace OpenXmlFun.Excel.Writer
{
    internal class ExcelStylesheetProvider
    {
        private static readonly Dictionary<ExcelColors, string> Colors = new Dictionary<ExcelColors, string>
        {
            { ExcelColors.White, "FFFFFF" },
            { ExcelColors.Black, "4300FF" },
            { ExcelColors.Red, "FF003C" },
            { ExcelColors.Green, "32CD32" },
            { ExcelColors.Blue, "4300FF" },
            { ExcelColors.Grey, "AAAAAA" }
        };

        //list of predefined NumberFormatId values https://github.com/closedxml/closedxml/wiki/NumberFormatId-Lookup-Table
        private static readonly Dictionary<Type, uint> Formats = new Dictionary<Type, uint>
        {
            { typeof(string), 49 },
            { typeof(DateTime), 14 },
            { typeof(decimal), 2 },
            { typeof(int), 1 }
        };

        //private static Dictionary<Type, Func<ExcelColors, ExcelColors, bool, bool, >>

        private Dictionary<string, uint> _styles = new Dictionary<string, uint>();

        public ExcelStylesheetProvider(bool wrapText)
        {
            Stylesheet = new Stylesheet();

            Stylesheet.Fonts = new Fonts();
            //default
            Stylesheet.Fonts.AppendChild(new Font());

            Stylesheet.Fills = new Fills();
            //default
            Stylesheet.Fills.AppendChild(new Fill());

            foreach (var color in Colors.Values)
            {
                Stylesheet.Fonts.AppendChild(new Font
                {
                    Color = new Color { Rgb = color }
                });
                Stylesheet.Fills.AppendChild(new Fill
                {
                    PatternFill = new PatternFill(new ForegroundColor
                    {
                        Rgb = new HexBinaryValue { Value = color }
                    })
                    {
                        PatternType = PatternValues.Solid
                    }
                });
            }

            Stylesheet.Borders = new Borders();
            //default
            Stylesheet.Borders.AppendChild(new Border());
            Stylesheet.Borders.Append(new Border
            {
                LeftBorder = new LeftBorder
                {
                    Style = BorderStyleValues.Medium,
                    Color = new Color
                    {
                        Indexed = (UInt32Value)64U
                    }
                },
                RightBorder = new RightBorder
                {
                    Style = BorderStyleValues.Medium,
                    Color = new Color
                    {
                        Indexed = (UInt32Value)64U
                    }
                },
                TopBorder = new TopBorder
                {
                    Style = BorderStyleValues.Medium,
                    Color = new Color
                    {
                        Indexed = (UInt32Value)64U
                    }
                },
                BottomBorder = new BottomBorder
                {
                    Style = BorderStyleValues.Medium,
                    Color = new Color
                    {
                        Indexed = (UInt32Value)64U
                    }
                },
                DiagonalBorder = new DiagonalBorder()
            });

            Stylesheet.CellFormats = new CellFormats();
            //default
            Stylesheet.CellFormats.AppendChild(new CellFormat());

            foreach (var fontColorKey in Colors)
            {
                foreach (var backgroundColorKey in Colors)
                {
                    foreach (var format in Formats)
                    {
                        Stylesheet.CellFormats.AppendChild(new CellFormat
                        {
                            ApplyNumberFormat = true,
                            NumberFormatId = format.Value,
                            ApplyAlignment = true,
                            Alignment = new Alignment { WrapText = wrapText },
                            ApplyBorder = true,
                            BorderId = 1,
                            ApplyFont = true,
                            FontId = 1,
                            ApplyFill = true,
                            FillId = 1
                        });
                    }
                }
            }
        }

        public Stylesheet Stylesheet { get; }

        public uint GetStyleId(ExcelCell cell)
        {
            return _styles[GetKey(cell.Value.GetType(), cell.FontColor, cell.BackgroundColor, cell.IsBold, cell.IsStroked)];
        }

        private string GetKey(Type type, ExcelColors fontColor, ExcelColors backgroundColor, bool isBold, bool isStroked)
        {
            return $"{type.Name}:{fontColor}:{backgroundColor}:{isBold}:{isStroked}";
        }
    }
}
