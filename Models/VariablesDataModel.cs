using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace TemplateDocumentGenerator.Models
{
    class VariablesDataModel : INotifyPropertyChanged
    {
        private List<string> allVariables;
        public List<string> AllVariables
        {
            get
            {
                return allVariables;
            }
            set
            {
                allVariables = value;
                OnPropertyChanged("AllVariables");
            }
        }
        public void UpdateVariables(List<List<string>> variables)
        {
            List<string> tmp = new List<string>();
            foreach(var list in variables)
            {
                foreach(var variable in list)
                {
                    if (!tmp.Contains(variable)) tmp.Add(variable);
                }
            }
            AllVariables = tmp;
        }
        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }
    }
}
