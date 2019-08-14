using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OpenXmlFun.Word.Parser
{
    public sealed class WordParser<T> : IDisposable where T : new()
    {
        private readonly WordprocessingDocument wordDoument;
        private readonly Body body;

        public WordParser(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                throw new FileNotFoundException(nameof(filePath));

            wordDoument = WordprocessingDocument.Open(filePath, true);
            body = wordDoument.MainDocumentPart.Document.Body;
            if (body == null)
                throw new InvalidOperationException("There is no body part in document");

        }


        // TODO: Implement document parsing
        public T Parse()
        {
            throw new NotImplementedException();
        }



        public void Dispose()
        {
            wordDoument.Dispose();
        }
    }
}
