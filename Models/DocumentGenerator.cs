using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TemplateDocumentGenerator.Models
{
    /// <summary>
    /// Static class which simplifies document generation
    /// </summary>
    static class DocumentGenerator
    {
        /// <summary>
        /// Generates one document.
        /// Assumes that input is correct.
        /// </summary>
        /// <param name="templateFullName">Full path to template</param>
        /// <param name="outFolder">Full path to output folder</param>
        /// <param name="namePattern">String with tags which are going to be replaced with corresponding values</param>
        /// <param name="variables"></param>
        public static void Generate(string templateFullName, string outFolder, string namePattern, List<Variable> variables)
        {
            using(var dl = new DocumentLoader(templateFullName))
            {
                string fileName = Path.GetFileName(templateFullName);
                string extension = Path.GetExtension(templateFullName);
                fileName = fileName.Replace(extension, "");
                string outFileName = ParseTemplateName(namePattern, fileName, extension, variables);
                dl.GenerateDocument(variables, outFolder + outFileName);
            }
        }
        /// <summary>
        /// Generates template name based on naming pattern and variables
        /// </summary>
        /// <param name="namePattern">String with tags</param>
        /// <param name="fileName">Short file name</param>
        /// <param name="extension">File extension with in .extension format</param>
        /// <param name="variables"></param>
        /// <returns></returns>
        private static string ParseTemplateName(string namePattern, string fileName, string extension, List<Variable> variables)
        {
            string result = namePattern;
            foreach (var r in variables)
            {
                result = result.Replace(r.Name, r.Value);
            }
            result = result.Replace("<tempname>", fileName);
            result = result + extension;
            return result;
        }
    }
}
