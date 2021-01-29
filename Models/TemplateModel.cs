using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.Serialization;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace TemplateDocumentGenerator.Models
{
    class IncorrectTemplateError : Exception
    {
        public IncorrectTemplateError(string message) : base(message) {}
    }
    class TemplateModel : INotifyPropertyChanged
    {
        private string fileName;
        private string namingPattern;
        private List<string> detectedVariablesList;
        private bool skipMissingData;

        public string ShortFileName
        {
            get { return Path.GetFileName(fileName); }
        }
        public string FileName
        {
            get { return fileName; }
            set
            {
                fileName = value;
                UpdateDetectedVariables();
                OnPropertyChanged("Title");
                OnPropertyChanged("ShortFileName");
            }
        }

        public string NamingPattern
        {
            get { return namingPattern; }
            set
            {
                namingPattern = value;
                OnPropertyChanged("NamingPattern");
            }
        }
        public string DetectedVariables
        {
            get { return string.Join(",", detectedVariablesList); }
        }
        public List<string> DetectedVariablesList
        {
            get { return detectedVariablesList; }
            set
            {
                detectedVariablesList = value;
                OnPropertyChanged("DetectedVariablesList");
                OnPropertyChanged("DetectedVariables");
            }
        }
        public bool SkipMissingData
        {
            get { return skipMissingData; }
            set
            {
                skipMissingData = value;
                OnPropertyChanged("SkipMissingData");
            }
        }
        private void UpdateDetectedVariables()
        {
            using (var document = Xceed.Words.NET.DocX.Load(FileName))
            {
                var detectedVariables = document.FindUniqueByPattern(@"<[\w \=]{1,}>", RegexOptions.IgnoreCase);
                if (detectedVariables.Count > 0)
                {
                    DetectedVariablesList = detectedVariables;
                }
                else throw new IncorrectTemplateError("No variables detected");
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }
    }
}
