using System;
using System.Collections.Generic;
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
            { ExcelColors.Black, "003300" },
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

        private readonly Dictionary<string, uint> _styles = new Dictionary<string, uint>();

        //todo support alignment, bold, stroke
        public ExcelStylesheetProvider(bool wrapText)
        {
            Stylesheet = new Stylesheet();

            Stylesheet.Fonts = new Fonts();
            //default Font
            Stylesheet.Fonts.AppendChild(new Font());

            Stylesheet.Fills = new Fills();
            //default Fill
            Stylesheet.Fills.AppendChild(new Fill());

            foreach (var color in Colors.Values)
            {
                Stylesheet.Fonts.AppendChild(new Font
                {
                    Color = new Color { Rgb = color }
                });
                Stylesheet.Fonts.AppendChild(new Font
                {
                    Color = new Color { Rgb = color },
                    Bold = new Bold()
                });
                Stylesheet.Fonts.AppendChild(new Font
                {
                    Color = new Color { Rgb = color },
                    Strike = new Strike()
                });
                Stylesheet.Fonts.AppendChild(new Font
                {
                    Color = new Color { Rgb = color },
                    Bold = new Bold(),
                    Strike = new Strike()
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
            Stylesheet.Fonts.Count = (uint)(Colors.Count + 1);
            Stylesheet.Fills.Count = (uint)(Colors.Count + 1);

            Stylesheet.Borders = new Borders();
            //default Border
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
            Stylesheet.Borders.Count = 2;

            Stylesheet.CellFormats = new CellFormats();
            //default CellFormat
            Stylesheet.CellFormats.AppendChild(new CellFormat());

            uint fontIndex = 0;
            uint csIndex = 0;
            foreach (Font font in Stylesheet.Fonts)
            {
                if (fontIndex == 0)
                {
                    fontIndex++;
                    continue;
                }

                uint fillIndex = 0;
                foreach (Fill fill in Stylesheet.Fills)
                {
                    if (fillIndex == 0)
                    {
                        fillIndex++;
                        continue;
                    }

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
                            FontId = fontIndex,
                            ApplyFill = true,
                            FillId = fillIndex
                        });

                        csIndex++;
                        _styles[GetKey(format.Key, 
                            font.Color.Rgb, 
                            fill.PatternFill.ForegroundColor.Rgb, 
                            font.Bold != null, 
                            font.Strike != null)] = csIndex;
                    }

                    fillIndex++;
                }

                fontIndex++;
            }
            Stylesheet.CellFormats.Count = csIndex + 1;
        }

        public Stylesheet Stylesheet { get; }

        public uint GetFormatId(Type type)
        {
            return Formats[type];
        }

        public uint GetStyleId(ExcelCell cell)
        {
            return _styles[GetKey(cell.Value.GetType(), 
                Colors[cell.FontColor], 
                Colors[cell.BackgroundColor], 
                cell.Bold, 
                cell.Strike)];
        }

        private string GetKey(Type type, string fontColorHex, string backgroundColorHex, bool isBold, bool isStroked)
        {
            return $"{type.Name}:{fontColorHex}:{backgroundColorHex}:{isBold}:{isStroked}";
        }
    }
}