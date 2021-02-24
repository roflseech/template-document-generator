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
    /// <summary>
    /// Disposable class, which provides functions to work with documents.
    /// </summary>
    class DocumentLoader : IDisposable
    {
        private DocX documentTemplate;
        /// <summary>
        /// Excpetion which informs that template cannot be used
        /// </summary>
        public class IncorrectTemplateError : Exception
        {
            public IncorrectTemplateError(string message) : base(message) { }
        }
        /// <summary>
        /// Loads given file.
        /// </summary>
        /// <param name="fileName"></param>
        public DocumentLoader(string fileName)
        {
            documentTemplate = DocX.Load(fileName);
        }
        /// <summary>
        /// Generates document file based on loaded template, with given variables and file name.
        /// </summary>
        /// <param name="variables"></param>
        /// <param name="outFileName"></param>
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
        /// <summary>
        /// Finds all variables in loaded document.
        /// </summary>
        /// <returns></returns>
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
