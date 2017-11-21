using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using NUnit.Framework;
using OpenXmlFun.Excel.Writer;
using OpenXmlFun.Excel.Writer.Crap;

namespace OpenXmlFun.Excel.IntegrationTests.Writer
{
    [TestFixture]
    class ExcelWriterTests
    {
        [Test]
        public void Test()
        {
            string filePath = Path.Combine(TestContext.CurrentContext.TestDirectory, 
                $@"{DateTime.Now.ToString(CultureInfo.InvariantCulture).GetSafeFileName()}.xlsx");

            using (var writer = new CrapExcelWriter(filePath))
            {
                writer.AddSheet("Договоры");
                writer.AddAcrossHeader("text", "datetime", "money");
                writer.AddRow(CrapExcelColors.Black, new[]
                {
                    new CrapExcelCell{Value = "some text"},
                    new CrapExcelCell{ Value = DateTime.Now},
                    new CrapExcelCell{ Value = 555.77M }
                });
            }
        }

        [Test]
        public void NewTest()
        {
            string filePath = Path.Combine(TestContext.CurrentContext.TestDirectory,
                $@"{DateTime.Now.ToString(CultureInfo.InvariantCulture).GetSafeFileName()}.xlsx");

            using (var writer = new ExcelWriter(filePath))
            {
                writer.AddSheet("Договоры_1", 20, 20, 20, 20)
                    .AddHeader("text_1", "datetime_1", "money_1", "count_1")
                    .AddRow(new[]
                    {
                        new ExcelCell{Value = "some text", Hyperlink = "http://google.com"},
                        new ExcelCell{ Value = DateTime.Now},
                        new ExcelCell{ Value = 555.77M },
                        new ExcelCell{ Value = 55 }
                    });
                writer.AddSheet("Договоры_2", 20, 20, 20, 20)
                    .AddHeader("text_2", "datetime_2", "money_2", "count_2")
                    .AddRow(new[]
                    {
                        new ExcelCell{Value = "hi i'm here"},
                        new ExcelCell{ Value = DateTime.UtcNow, BackgroundColor = ExcelColors.Blue},
                        new ExcelCell{ Value = 222.88M },
                        new ExcelCell{ Value = 1277 }
                    });
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
