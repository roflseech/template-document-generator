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
    class DocumentTemplate : INotifyPropertyChanged
    {
        private string fileName;
        private List<string> variables;
        private bool active;

        public List<string> Variables
        {
            get { return variables; }
            set
            { 
                variables = value;
                OnPropertyChanged("Variables");
            }
        }

        public bool IsActive {
            get { return active; } 
            set
            {
                active = value;
                OnPropertyChanged("IsActive");
            }
        }
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
                OnPropertyChanged("FileName");
                OnPropertyChanged("ShortFileName");
            }
        }
        public void ReloadVaraibles()
        {
            using (var dl = new DocumentLoader(fileName))
            {
                Variables = dl.FindVariables();
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
