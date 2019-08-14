using NUnit.Framework;
using OpenXmlFun.Word.Writer;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;


namespace OpenXmlFun.Word.IntegrationTests.Writer
{
    [TestFixture]
    class WordWriterTest
    {
        private readonly string filePath = Path.Combine(TestContext.CurrentContext.TestDirectory,
                $@"{DateTime.Now.ToString(CultureInfo.InvariantCulture).GetSafeFileName()}.docx");
        [Test]
        public void BasicTest()
        {
            var wordWriter = new WordWriter(filePath);
            wordWriter.Dispose();
            Assert.That(filePath, Does.Exist);
        }

        [Test]
        public void ShouldThrowExceptionWhenFilepathIsNullOrWitheSpace()
        {
            Assert.Throws<ArgumentException>(
                () => new WordWriter(null)
            );

            Assert.Throws<ArgumentException>(
                () => new WordWriter(string.Empty)
            );
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
