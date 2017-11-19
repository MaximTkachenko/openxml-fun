using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

// ReSharper disable PossiblyMistakenUseOfParamsMethod

namespace OpenXmlFun.Excel.Writer
{
    /// <inheritdoc />
    /// <summary>
    /// https://stackoverflow.com/questions/2792304/how-to-insert-a-date-to-an-open-xml-worksheet
    /// http://www.dispatchertimer.com/tutorial/how-to-create-an-excel-file-in-net-using-openxml-part-3-add-stylesheet-to-the-spreadsheet/
    /// </summary>
    public class ExcelWriter : IDisposable
    {
        private readonly SpreadsheetDocument _spreadsheetDocument;
        private readonly Dictionary<string, ExcelSheet> _sheets;

        public ExcelWriter(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException($"Specify {nameof(filePath)}.");
            }
            _spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
            _spreadsheetDocument.AddWorkbookPart();
            _spreadsheetDocument.WorkbookPart.Workbook = new Workbook();

            _sheets = new Dictionary<string, ExcelSheet>();
        }

        public ExcelSheet AddSheet(string name, params double[] columnWidths)
        {
            foreach (var key in _sheets.Keys)
            {
                _sheets[key].Save();
            }

            var worksheetPart = _spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet();
            worksheetPart.Worksheet.Append(new SheetData());

            if (_spreadsheetDocument.WorkbookPart.Workbook.Sheets == null)
            {
                _spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
            }

            _spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().AppendChild(new Sheet
            {
                Id = _spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = (uint)_sheets.Keys.Count + 1,
                Name = name
            });

            if (_spreadsheetDocument.WorkbookPart.WorkbookStylesPart == null)
            {
                _spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                _spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet = CreateStylesheet();
            }

            _sheets[name] = new ExcelSheet(worksheetPart, columnWidths);
            return _sheets[name];
        }

        /// <summary>
        /// List of predefined NumberFormatId values https://github.com/closedxml/closedxml/wiki/NumberFormatId-Lookup-Table
        /// </summary>
        /// <returns></returns>
        private Stylesheet CreateStylesheet()
        {
            var styleSheet = new Stylesheet();
            styleSheet.CellFormats = new CellFormats
            (
                //default
                new CellFormat
                {
                    ApplyAlignment = true,
                    Alignment = new Alignment { WrapText = true }
                },
                //date
                new CellFormat
                {
                    ApplyNumberFormat = true,
                    NumberFormatId = 14,
                    ApplyAlignment = true,
                    Alignment = new Alignment { WrapText = true }
                },
                //text
                new CellFormat
                {
                    ApplyNumberFormat = true,
                    NumberFormatId = 49,
                    ApplyAlignment = true,
                    Alignment = new Alignment { WrapText = true }
                },
                //decimal
                new CellFormat
                {
                    ApplyNumberFormat = true,
                    NumberFormatId = 2,
                    ApplyAlignment = true,
                    Alignment = new Alignment { WrapText = true }
                },
                //int
                new CellFormat
                {
                    ApplyNumberFormat = true,
                    NumberFormatId = 1,
                    ApplyAlignment = true,
                    Alignment = new Alignment { WrapText = true }
                }
            );
            return styleSheet;
        }

        public void Dispose()
        {
            _spreadsheetDocument.Dispose();
        }
    }
}
