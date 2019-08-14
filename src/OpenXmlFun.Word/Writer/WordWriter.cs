using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlFun.Word.Writer
{
    public class WordWriter : IDisposable
    {
        private readonly WordprocessingDocument wordDocument;

        public WordWriter(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException(nameof(filePath));
            wordDocument = WordprocessingDocument.Create(filePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
            MainDocumentPart documentPart = wordDocument.AddMainDocumentPart();
            documentPart.Document = new Document
            {
                Body = new Body()
            };
        }

        public void Dispose()
        {
            wordDocument.Save();
            wordDocument.Dispose();
        }
    }
}
