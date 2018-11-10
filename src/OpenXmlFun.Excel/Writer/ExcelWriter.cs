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
            _spreadsheetDocument.WorkbookPart.Workbook.Append(new BookViews(new WorkbookView()));

            _sheets = new Dictionary<string, ExcelSheet>();
        }

        public ExcelSheet AddSheet(string name)
        {
            if (_sheets.ContainsKey(name))
            {
                throw new InvalidOperationException($"[{name}] sheet already exists.");
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
                _spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet = ExcelStylesheetProvider.Stylesheet;
            }

            _sheets[name] = new ExcelSheet(worksheetPart);
            return _sheets[name];
        }

        public void Dispose()
        {
            foreach (var key in _sheets.Keys)
            {
                _sheets[key].Save();
            }

            _spreadsheetDocument.Save();
            _spreadsheetDocument.Dispose();
        }
    }
}
