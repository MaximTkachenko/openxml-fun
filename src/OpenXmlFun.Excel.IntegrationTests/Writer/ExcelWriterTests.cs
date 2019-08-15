﻿using System;
using System.Globalization;
using System.IO;
using NUnit.Framework;
using OpenXmlFun.Excel.Writer;
using OpenXmlFun.Excel.Writer.Cells;

namespace OpenXmlFun.Excel.IntegrationTests.Writer
{
    [TestFixture]
    class ExcelWriterTests
    {
        [Test]
        public void BasicTest()
        {
            string filePath = Path.Combine(TestContext.CurrentContext.TestDirectory,
                $@"{DateTime.Now.ToString(CultureInfo.InvariantCulture).GetSafeFileName()}.xlsx");

            CellBase[] nullArray = null;
            
            using (var writer = new ExcelWriter(filePath))
            {
                writer.AddSheet("test_1")
                    .ApplyColumnWidths(100, 20, 20, 20)
                    .FreezeTopNRows(1)
                    .AddColumnFilter(1, 1)
                    .AddHeader("text_1", "datetime_1", "money_1", "count_1")
                    .AddRow(new StringCell(@"some /\ text" ),
                        new DateTimeCell(DateTime.Now) { Bold = true, Strike = true, FontColor = ExcelColors.Red, BackgroundColor = ExcelColors.Green },
                        new DecimalCell(555.77M) { BackgroundColor = ExcelColors.Blue },
                        new IntegerCell(55))
                    .AddRow(new StringCell("more staff"),
                        new DateTimeCell(DateTime.UtcNow) { Bold = true, FontColor = ExcelColors.Red, BackgroundColor = ExcelColors.Green },
                        new DecimalCell(999.77M) { BackgroundColor = ExcelColors.Blue },
                        new IntegerCell(999))
                    .AddRow(nullArray)
                    .AddRow(null);

                writer.AddSheet("test_2")
                    .AddHeader("text_2", "datetime_2", "money_2", "count_2")
                    .AddColumnFilters(1, 3, 1)
                    .AddRow(new StringCell("hi im here") { Bold = true },
                        new DateTimeCell(DateTime.UtcNow) { FontColor = ExcelColors.Red },
                        new DecimalCell(222.88M) { FontColor = ExcelColors.Green },
                        new IntegerCell(1277));

                writer.AddSheet("test_3")
                    .AddRow(new DecimalCell(222.88M), new DecimalCell(666M), new DecimalHorizontalSumCell(1, 2))
                    .AddRow(new DecimalCell(11M), new DecimalCell(22M), new DecimalHorizontalSumCell(1, 2))
                    .AddRow(new DecimalVerticalSumCell(1, 2), new DecimalVerticalSumCell(1, 2));
            }
        }
    }

    internal static class StringExt
    {
        public static string GetSafeFileName(this string fileName)
        {
            return string.Join("_", fileName.Split(Path.GetInvalidFileNameChars()))
                .Replace(" ", "_");
        }
    }
}
