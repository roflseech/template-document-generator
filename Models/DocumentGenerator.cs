using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace TemplateDocumentGenerator.Models
{
    class DocumentGenerator : IDisposable
    {
        private DocX documentTemplate;

        public DocumentGenerator(string fileName)
        {
            documentTemplate = Xceed.Words.NET.DocX.Load(fileName);
        }

        public void GenerateDocument(List<(string, string)> replacements, string outFileName)
        {
            var clonedDocument = documentTemplate.Copy();
            foreach(var r in replacements)
            {
                clonedDocument.ReplaceText(r.Item1, r.Item2);
            }
            clonedDocument.SaveAs(outFileName);
        }
        public void Dispose()
        {
            documentTemplate.Dispose();
        }

    }
}
