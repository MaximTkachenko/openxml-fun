![.NET Core](https://github.com/MaximTkachenko/openxml-fun/workflows/.NET%20Core/badge.svg)

# openxml-fun
Create excel tables using [ExcelWriter](https://github.com/MaximTkachenko/openxml-fun/blob/master/src/OpenXmlFun.Excel/Writer/ExcelWriter.cs) and parse excel tables to list of entities using [ExcelParser](https://github.com/MaximTkachenko/openxml-fun/blob/master/src/OpenXmlFun.Excel/Parser/ExcelParser.cs). They use [DocumentFormat.OpenXml](https://github.com/OfficeDev/Open-XML-SDK) package internally. You can find examples of usage in [OpenXmlFun.Excel.IntegrationTests](https://github.com/MaximTkachenko/openxml-fun/tree/master/src/OpenXmlFun.Excel.IntegrationTests) project

It supports following .net types: `String`, `Int32`, `Decimal`, `DateTime`. Summary columns are also supported.

Target framework: `netstandard2.0`.

