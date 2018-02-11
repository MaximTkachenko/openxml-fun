using System;
using System.Globalization;
using System.IO;
using NUnit.Framework;
using OpenXmlFun.Excel.Writer;

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

            DateTime? nullDt = null;
            ExcelCell[] nullArray = null;
            
            using (var writer = new ExcelWriter(filePath))
            {
                writer.AddSheet("Договоры_1")
                    .ApplyColumnWidths(100, 20, 20, 20)
                    .FreezeTopNRows(1)
                    .AddHeader("text_1", "datetime_1", "money_1", "count_1")
                    .AddRow(DateTime.Now, DateTime.MinValue, 555.77M, 55, null, nullDt, 0M)
                    .AddRow(nullArray)
                    .AddRow(null)
                    .AddRow(new ExcelCell(@"some /\ text" ),
                        new ExcelCell(DateTime.Now) { Bold = true, Strike = true, FontColor = ExcelColors.Red, BackgroundColor = ExcelColors.Green },
                        new ExcelCell(555.77M) { BackgroundColor = ExcelColors.Blue },
                        new ExcelCell(55));

                writer.AddSheet("Договоры_2")
                    .AddHeader("text_2", "datetime_2", "money_2", "count_2")
                    .AddRow(new ExcelCell("hi im here") { Bold = true },
                        new ExcelCell(DateTime.UtcNow) { FontColor = ExcelColors.Red },
                        new ExcelCell(222.88M) { FontColor = ExcelColors.Green },
                        new ExcelCell(1277));
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
