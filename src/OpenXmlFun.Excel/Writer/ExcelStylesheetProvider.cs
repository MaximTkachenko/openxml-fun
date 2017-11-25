﻿using System;
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
            { ExcelColors.Black, "000000" },
            { ExcelColors.Red, "FF003C" },
            { ExcelColors.Green, "32CD32" },
            { ExcelColors.Blue, "4300FF" },
            { ExcelColors.Grey, "AAAAAA" }
        };

        private readonly Dictionary<string, uint> _styles = new Dictionary<string, uint>();

        public ExcelStylesheetProvider(bool wrapText)
        {
            var fonts = new Fonts();
            //default Font
            fonts.AppendChild(new Font { Color = new Color() });
            uint defaultFontsCount = (uint)fonts.ChildElements.Count;

            var fills = new Fills();
            //default Fills
            fills.AppendChild(new Fill(new PatternFill { PatternType = PatternValues.None }));
            fills.AppendChild(new Fill(new PatternFill { PatternType = PatternValues.Gray125 }));
            uint defaultFillsCount = (uint)fills.ChildElements.Count;

            foreach (var color in Colors.Values)
            {
                fonts.AppendChild(new Font
                {
                    Color = new Color { Rgb = color }
                });
                fonts.AppendChild(new Font
                {
                    Color = new Color { Rgb = color },
                    Bold = new Bold()
                });
                fonts.AppendChild(new Font
                {
                    Color = new Color { Rgb = color },
                    Strike = new Strike()
                });
                fonts.AppendChild(new Font
                {
                    Color = new Color { Rgb = color },
                    Bold = new Bold(),
                    Strike = new Strike()
                });

                fills.AppendChild(new Fill
                (
                    new PatternFill
                    {
                        ForegroundColor = new ForegroundColor { Rgb = color },
                        PatternType = PatternValues.Solid
                    }
                ));
            }
            fonts.Count = (uint)fonts.ChildElements.Count;
            fills.Count = (uint)fills.ChildElements.Count;

            var borders = new Borders();
            //default Border
            borders.AppendChild(new Border());
            borders.Append(new Border
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
            borders.Count = (uint)borders.ChildElements.Count;

            var cellFormats = new CellFormats();
            //default CellFormat
            cellFormats.AppendChild(new CellFormat { FontId = 0, FillId = 0, BorderId = 0 });

            uint fontIndex = 0;
            uint csIndex = 0;
            foreach (Font font in fonts.ChildElements)
            {
                if (fontIndex < defaultFontsCount)
                {
                    fontIndex++;
                    continue;
                }

                uint fillIndex = 0;
                foreach (Fill fill in fills.ChildElements)
                {
                    if (fillIndex < defaultFillsCount)
                    {
                        fillIndex++;
                        continue;
                    }

                    foreach (var typeDetails in SupportedTypesDetails.Data)
                    {
                        cellFormats.AppendChild(new CellFormat
                        {
                            ApplyNumberFormat = true,
                            NumberFormatId = typeDetails.Value.NumberFormatId,
                            ApplyAlignment = true,
                            Alignment = new Alignment { WrapText = wrapText },
                            ApplyBorder = true,
                            BorderId = 1,
                            ApplyFont = true,
                            FontId = fontIndex,
                            ApplyFill = true,
                            FillId = fillIndex,
                            FormatId = 0
                        });

                        csIndex++;
                        _styles[GetKey(typeDetails.Key, 
                            font.Color.Rgb, 
                            fill.PatternFill.ForegroundColor.Rgb.Value, 
                            font.Bold != null, 
                            font.Strike != null)] = csIndex;
                    }

                    fillIndex++;
                }

                fontIndex++;
            }
            cellFormats.Count = (uint)cellFormats.ChildElements.Count;

            Stylesheet = new Stylesheet(fonts, fills, borders, cellFormats);
        }

        public Stylesheet Stylesheet { get; }

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