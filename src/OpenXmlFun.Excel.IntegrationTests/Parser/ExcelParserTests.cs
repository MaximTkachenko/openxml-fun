using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using FluentAssertions;
using NUnit.Framework;
using OpenXmlFun.Excel.Parser;

namespace OpenXmlFun.Excel.IntegrationTests.Parser
{
    [TestFixture]
    class ExcelParserTests
    {
        public class ItemToParse
        {
            [ParseDetails(0)]
            public int Number { get; set; }

            [ParseDetails(1)]
            public string Description { get; set; }

            [ParseDetails(2)]
            public DateTime Created { get; set; }

            [ParseDetails(3)]
            public decimal Sum { get; set; }
        }

        [Test]
        [TestCase("en-US")]
        [TestCase("en-GB")]
        [TestCase("ru-RU")]
        public void Parse_CorrectDocument_WorksOk(string culture)
        {
            var cultureInfo = new CultureInfo(culture);
            CultureInfo.DefaultThreadCurrentCulture = cultureInfo;
            CultureInfo.DefaultThreadCurrentUICulture = cultureInfo;
            string filePath = Path.Combine(TestContext.CurrentContext.TestDirectory, "Parser", "ItemsToParse.xlsx");

            List<ItemToParse> items;
            using (var parser = new ExcelParser<ItemToParse>(filePath))
            {
                items = parser.Parse(true);
            }

            items.Count.Should().Be(3);
            items[0].Number.Should().Be(1);
            items[0].Description.Should().Be("description one");
            items[0].Created.Should().Be(new DateTime(2018, 2, 3));
            Math.Round(items[0].Sum, 2).Should().Be(33000.66M);
        }
    }
}
