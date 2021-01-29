using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace TemplateDocumentGenerator.ViewModels
{
    class MainViewModel : INotifyPropertyChanged
    {
        public ObservableCollection<Models.TemplateModel> TemplatesList { get; set; }
        private Models.TemplateModel selectedTemplate;
        //private Models.VariablesDataModel variables;
        private DataTable variablesTable;
        private int selectedDataCell;
        public Visibility SelectedTemplateExist 
        {
            get { return selectedTemplate == null ? Visibility.Collapsed : Visibility.Visible; }
        }
        public DataTable VariablesTable
        {
            get { return variablesTable; }
            set
            {
                variablesTable = value;
                OnPropertyChanged("VariablesTable");
            }
        }

        public int SelectedDataCell
        {
            get { return selectedDataCell; }
            set
            {
                selectedDataCell = value;
                OnPropertyChanged("SelectedDataCell");
            }
        }

        public MainViewModel()
        {
            TemplatesList = new ObservableCollection<Models.TemplateModel>();
            //variables = new Models.VariablesDataModel();
            var newTable = new DataTable();
            for (int i = 0; i < 2; i++)
                newTable.Columns.Add(new DataColumn("column_" + i));

            var headerRow = newTable.NewRow();
            
            for (int i = 0; i < 5; i++)
                newTable.Rows.Add(newTable.NewRow());
           // newTable.Rows.
            VariablesTable = newTable;

            TemplatesList.CollectionChanged += OnTemplatesListUpdated;
        }

        public Models.TemplateModel SelectedTemplate
        {
            get { return selectedTemplate; }
            set
            {
                selectedTemplate = value;
                OnPropertyChanged("SelectedTemplate");
                OnPropertyChanged("SelectedTemplateExist");
            }
        }
        public ICommand AddTemplateFromFile
        {
            get
            {
                return new DelegateCommand((obj) =>
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog();
                    if (openFileDialog.ShowDialog() == true)
                    {
                        Models.TemplateModel newTemplate;

                        try
                        {
                            //trying to open file, can throw exceptions
                            newTemplate = new Models.TemplateModel
                            {
                                FileName = openFileDialog.FileName,
                                NamingPattern = "Document_<count>",
                                SkipMissingData = false
                            };
                            
                            TemplatesList.Add(newTemplate);
                        }
                        catch(Models.IncorrectTemplateError)
                        {
                            MessageBox.Show("No variables detected");
                        }
                        catch
                        {
                            MessageBox.Show("Cannot read doceument");
                        }
                    }
                });
            }
        }
        public ICommand RemoveSelectedTemplate
        {
            get
            {
                return new DelegateCommand((obj) =>
                {
                    TemplatesList.Remove(SelectedTemplate);
                });
            }
        }
        public ICommand AddVariable
        {
            get
            {
                return new DelegateCommand((obj) =>
                {
                    //variablesTable.
                });
            }
        }
        public ICommand AddRow
        {
            get
            {
                return new DelegateCommand((obj) =>
                {
                    variablesTable.Rows.Add(variablesTable.NewRow());
                });
            }
        }
        public ICommand DeleteVariable
        {
            get
            {
                return new DelegateCommand((obj) =>
                {
                   // variablesTable.Columns.Remove(SelectedDataGridCell.Column);
                });
            }
        }
        public ICommand DeleteRow
        {
            get
            {
                return new DelegateCommand((obj) =>
                {
                    System.Console.WriteLine(SelectedDataCell);
                   // variablesTable.Rows.Remove(variablesTable.)
                });
            }
        }
        public ICommand GenerateDocuments
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
                        using (var documentGenerator = new Models.DocumentGenerator(template.FileName))
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
        }
        private static string ParseTemplateName(Models.TemplateModel template, int index, List<(string, string)> replacements)
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
        }
        private void OnTemplatesListUpdated(object sender, NotifyCollectionChangedEventArgs e)
        {
        }

        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }
    }
}
