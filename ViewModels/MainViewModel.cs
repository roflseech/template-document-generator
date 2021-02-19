using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;

using System.Windows.Data;
using System.Windows.Input;

namespace TemplateDocumentGenerator.ViewModels
{
    class MainViewModel : INotifyPropertyChanged
    {
        public ObservableCollection<Models.DocumentTemplate> TemplatesList { get; set; }
        private Models.DocumentTemplate selectedTemplate;
        public Models.DocumentTemplate SelectedTemplate
        {
            get { return selectedTemplate; }
            set
            {
                selectedTemplate = value;
                OnPropertyChanged("SelectedTemplate");
            }
        }
        private object templatesListLock;

        public ObservableCollection<Models.Variable> VariablesList { get; set; }
        private string namePattern;
        private string outPath;

        public string NamePattern 
        {
            get { return namePattern; }
            set
            {
                namePattern = value;
                OnPropertyChanged("NamePattern");
            }
        }
        public string OutPath
        {
            get { return outPath; }
            set
            {
                outPath = value;
                OnPropertyChanged("OutPath");
            }
        }

        private Models.TemplatesListUpdater templatesListUpdater;
        

        private bool FileCheck(string s)
        {
            return Path.GetExtension(s).CompareTo(".docx") == 0 ||
                Path.GetExtension(s).CompareTo(".doc") == 0;
        }

        public MainViewModel()
        {
            TemplatesList = new ObservableCollection<Models.DocumentTemplate>();
            templatesListLock = new object();
            BindingOperations.EnableCollectionSynchronization(TemplatesList, templatesListLock);

            VariablesList = new ObservableCollection<Models.Variable>();
            templatesListUpdater = new Models.TemplatesListUpdater("Templates", TemplatesList, FileCheck);
            TemplatesList.CollectionChanged += AddChangeNotifications;
            templatesListUpdater.ScanFolder();
            NamePattern = "<tempname>";
            OutPath = Directory.GetCurrentDirectory();
            var a =Directory.Exists("Templates");
        }

        private void AddChangeNotifications(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.OldItems != null)
            {
                foreach (INotifyPropertyChanged item in e.OldItems)
                    item.PropertyChanged -= TemplatePropertyChanged;
            }
            if (e.NewItems != null)
            {
                foreach (INotifyPropertyChanged item in e.NewItems)
                    item.PropertyChanged += TemplatePropertyChanged;
            }
        }
        private void TemplatePropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName != "IsActive") return;

            var changedTempalte = (Models.DocumentTemplate)sender;
            if(changedTempalte.IsActive)
            {
                changedTempalte.ReloadVaraibles();
                var changedTemplateVariables = changedTempalte.Variables;
                foreach(var variable in changedTemplateVariables)
                {
                    bool needsInclusion = true;
                    foreach(var item in VariablesList)
                    {
                        if(item.Name == variable)
                        {
                            needsInclusion = false;
                            break;
                        }
                    }
                    if (needsInclusion) VariablesList.Add(
                         new Models.Variable
                         {
                             Name = variable
                         });
                }
            }
            else
            {
                var detectedVariables = new HashSet<string>();
                foreach(var a in TemplatesList)
                {
                    if(a.IsActive)
                    {
                        foreach (var variable in a.Variables)
                        {
                            detectedVariables.Add(variable);
                        }
                    }
                }

                for (int i = VariablesList.Count - 1; i >= 0; i--)
                {
                    if(!detectedVariables.Contains(VariablesList[i].Name))
                    {
                        VariablesList.RemoveAt(i);
                    }
                }
            }
        }

        public ICommand ChooseOutPath
        {
            get
            {
                return new DelegateCommand((obj) =>
                {
                    var dialog = new CommonOpenFileDialog();
                    dialog.IsFolderPicker = true;
                    CommonFileDialogResult result = dialog.ShowDialog();
                    if(result == CommonFileDialogResult.Ok)
                    {
                        OutPath = dialog.FileName;
                    }
                });
            }
        }
        private bool VerifyInput()
        {
            if (!Directory.Exists(OutPath)) return false;
            string tmp = namePattern;
            foreach(var a in VariablesList)
            {
                tmp = tmp.Replace(a.Name, a.Value);
            }
            if(tmp.LastIndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                return false;
            }
            return true;
        }
        public ICommand GenerateDocuments
        {
            get
            {
                return new DelegateCommand((obj) =>
                {
                    if (!VerifyInput()) return;

                    var outPathFixed = OutPath;
                    if(outPathFixed[outPathFixed.Length-1] != '\\') outPathFixed = outPathFixed + "\\";

                    foreach (var template in TemplatesList)
                    {
                        if(template.IsActive)
                        {
                            Models.DocumentGenerator.Generate(template.FileName,
                                outPathFixed, 
                                NamePattern, 
                                new List<Models.Variable>(VariablesList));
                        }
                    }
                });
            }
        }
        public ICommand AddTemplate
        {
            get
            {
                return new DelegateCommand((obj) =>
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog();
                    openFileDialog.Filter = "Documents|*.docx;*.doc";
                    if (openFileDialog.ShowDialog() == true)
                    {
                        File.Copy(openFileDialog.FileName, "Templates\\" + Path.GetFileName(openFileDialog.FileName));
                    }
                });
            }
        }
        public ICommand RemoveTemplate
        {
            get
            {
                return new DelegateCommand((obj) =>
                {
                    if(selectedTemplate != null)
                    {
                        File.Delete(selectedTemplate.FileName);
                    }
                });
            }
        }
        public ICommand OpenTemplatesFolder
        {
            get
            {
                return new DelegateCommand((obj) =>
                {
                    Process.Start(@"Templates");
                });
            }
        }
        
        /*public ICommand AddTemplateFromFile
{
   get
   {
       return new DelegateCommand((obj) =>
       {
           OpenFileDialog openFileDialog = new OpenFileDialog();
           if (openFileDialog.ShowDialog() == true)
           {
               Models.DocumentTemplate newTemplate;
               for(int i = 0; i < 100; i++)
               {
                   try
                   {
                       //trying to open file, can throw exceptions
                       newTemplate = new Models.DocumentTemplate
                       {
                           FileName = openFileDialog.FileName,
                           NamingPattern = "Document_<count>",
                           SkipMissingData = false
                       };

                       TemplatesList.Add(newTemplate);
                   }
                   catch (Models.IncorrectTemplateError)
                   {
                       MessageBox.Show("No variables detected");
                   }
                   catch
                   {
                       MessageBox.Show("Cannot read doceument");
                   }
               }
           }
       });
   }
}*/
        /*public ICommand RemoveSelectedTemplate
        {
            get
            {
                return new DelegateCommand((obj) =>
                {
                    TemplatesList.Remove(SelectedTemplate);
                });
            }
        }*/
        /*public ICommand AddVariable
        {
            get
            {
                return new DelegateCommand((obj) =>
                {
                    //variablesTable.
                });
            }
        }*/
        /*public ICommand AddRow
        {
            get
            {
                return new DelegateCommand((obj) =>
                {
                    variablesTable.Rows.Add(variablesTable.NewRow());
                });
            }
        }*/
        /*public ICommand DeleteVariable
        {
            get
            {
                return new DelegateCommand((obj) =>
                {
                   // variablesTable.Columns.Remove(SelectedDataGridCell.Column);
                });
            }
        }*/
        /*public ICommand DeleteRow
        {
            get
            {
                return new DelegateCommand((obj) =>
                {
                    System.Console.WriteLine(SelectedDataCell);
                   // variablesTable.Rows.Remove(variablesTable.)
                });
            }
        }*/
        /*public ICommand GenerateDocuments
        {
            get
            {
                return new DelegateCommand((obj) =>
                {
                    var dialog = new CommonOpenFileDialog();
                    dialog.IsFolderPicker = true;
                    CommonFileDialogResult result = dialog.ShowDialog();

                    foreach (var template in TemplatesList)
                    {
                        using (var documentGenerator = new Models.DocumentLoader(template.FileName))
                        {
                            for (int i = 1; i < VariablesTable.Rows.Count; i++)
                            {
                                var replacements = new List<(string, string)>();
                                for (int j = 0; j < VariablesTable.Columns.Count; j++)
                                {
                                    replacements.Add(
                                        (VariablesTable.Rows[0][j] as string, 
                                        VariablesTable.Rows[i][j] as string));
                                }
                                string outName = ParseTemplateName(template, i, replacements);
                                documentGenerator.GenerateDocument(replacements, dialog.FileName + outName);
                            }
                        }
                    }
                });
            }
        }*/
        /*private static string ParseTemplateName(Models.DocumentTemplate template, int index, List<(string, string)> replacements)
        {
            string result = template.NamingPattern;
            foreach (var r in replacements)
            {
                result = result.Replace(r.Item1, r.Item2);
            }
            result = result.Replace("<count>", index.ToString());
            result = result.Replace("<name>", Path.GetFileName(template.FileName));
            result = result + "." + Path.GetExtension(template.FileName);
            return result;
        }*/
        /*private void OnTemplatesListUpdated(object sender, NotifyCollectionChangedEventArgs e)
        {
        }*/

        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }
    }
}
