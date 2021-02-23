using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace TemplateDocumentGenerator.Models
{
    class DocumentLoader : IDisposable
    {
        public class IncorrectTemplateError : Exception
        {
            public IncorrectTemplateError(string message) : base(message) { }
        }

        private DocX documentTemplate;

        public DocumentLoader(string fileName)
        {
            documentTemplate = DocX.Load(fileName);
        }

        public void GenerateDocument(List<Variable> variables, string outFileName)
        {
            var clonedDocument = documentTemplate.Copy();
            foreach(var r in variables)
            {
                var newValue = r.Value;
                if (newValue == null) newValue = "";
                clonedDocument.ReplaceText(r.Name, newValue);
            }
            clonedDocument.SaveAs(outFileName);
        }

        public List<string> FindVariables()
        {
            var detectedVariables = documentTemplate.FindUniqueByPattern(@"<[\w \=]{1,}>", RegexOptions.IgnoreCase);
            if (detectedVariables.Count == 0)
            {
                throw new IncorrectTemplateError("No variables detected");
            }
            return detectedVariables;
        }

        public void Dispose()
        {
            documentTemplate.Dispose();
        }
    }
}
