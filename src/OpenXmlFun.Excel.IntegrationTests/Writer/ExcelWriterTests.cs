using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using NUnit.Framework;
using OpenXmlFun.Excel.Writer;

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

            using (var writer = new ExcelWriter(filePath))
            {
                writer.AddSheet("Договоры");
                writer.AddAcrossHeader("text", "datetime", "money");
                writer.AddRow(ExcelColors.Black, 
                    new[] {new ExcelCell{ Text = "some text"}, new ExcelCell { Text = DateTime.Now }, new ExcelCell { Text = 555.77M } });
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
