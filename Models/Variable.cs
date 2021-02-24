using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace TemplateDocumentGenerator.Models
{
    /// <summary>
    /// Variable class which invokes notification whenever it's properties changed
    /// </summary>
    class Variable : INotifyPropertyChanged
    {
        private string name;
        private string value;
        public Variable()
        {

        }
        public Variable(string name, string value)
        {
            this.name = name;
            this.value = value;
        }
        public string Name
        {
            get { return name; }
            set
            {
                name = value;
                OnPropertyChanged("Name");
            }
        }
        public string Value
        {
            get { return value; }
            set
            {
                this.value = value;
                OnPropertyChanged("Value");
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
