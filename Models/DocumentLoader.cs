using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace TemplateDocumentGenerator.Models
{
    class DocumentLoader : IDisposable
    {
        class IncorrectTemplateError : Exception
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
                clonedDocument.ReplaceText(r.Name, r.Value);
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
            /*for(int i = 0; i < detectedVariables.Count;i++)
            {
                string s = detectedVariables[i].ToLower();
                detectedVariables[i] = s.Substring(1,1).ToUpper() + s.Substring(2, s.Length - 3);
            }*/
            return detectedVariables;
        }

        public void Dispose()
        {
            documentTemplate.Dispose();
        }
    }
}
