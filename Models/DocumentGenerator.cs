using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TemplateDocumentGenerator.Models
{
    static class DocumentGenerator
    {
        public static void Generate(string templateFullName, string outFolder, string namePattern, List<Variable> variables)
        {
            using(var dl = new DocumentLoader(templateFullName))
            {
                string fileName = Path.GetFileName(templateFullName);
                string extension = Path.GetExtension(templateFullName);
                string outFileName = ParseTemplateName(namePattern, fileName, extension, variables);
                dl.GenerateDocument(variables, outFolder + outFileName);
            }
        }
        private static string ParseTemplateName(string namePattern, string fileName, string extension, List<Variable> variables)
        {
            string result = namePattern;
            foreach (var r in variables)
            {
                result = result.Replace(r.Name, r.Value);
            }
            result = result.Replace("<tempname>", fileName);
            result = result + "." + extension;
            return result;
        }
    }
}
