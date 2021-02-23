using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Resources;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Markup;
using TemplateDocumentGenerator.Properties;

namespace TemplateDocumentGenerator.ViewModels
{
    class MainViewModel : INotifyPropertyChanged
    {
        private string statusText;
        public string StatusText
        {
            get { return statusText; }
            set
            {
                statusText = value;
                OnPropertyChanged("StatusText");
            }
        }
        Dictionary<string, ResourceDictionary> localizations;

        public ObservableCollection<Models.Variable> Languages { get; set; }
        private Models.Variable selectedLanguage;
        private string previousLanguage;

        public Models.Variable SelectedLanguage
        {
            get { return selectedLanguage; }
            set
            {
                if(selectedLanguage != null)
                {
                    previousLanguage = selectedLanguage.Value;
                }
                selectedLanguage = value;
                OnPropertyChanged("SelectedLanguage");
            }
        }
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

            Languages = new ObservableCollection<Models.Variable>();
            localizations = new Dictionary<string, ResourceDictionary>();
            AddLanguage("en", "English", "Properties/Localization_english.xaml");
            AddLanguage("ru", "Русский", "Properties/Localization_russian.xaml");
            Application.Current.Resources.MergedDictionaries.Add(localizations["en"]);
            SelectedLanguage = Languages[0];

            PropertyChanged += LanguageChanging;
        }
        private void AddLanguage(string tag, string name, string uri)
        {
            Languages.Add(new Models.Variable(name, tag));
            localizations[tag] = new ResourceDictionary();
            localizations[tag].Source = new Uri(uri, UriKind.Relative);
        }
        private void LanguageChanging(object sender, PropertyChangedEventArgs e)
        {
            if(e.PropertyName == "SelectedLanguage")
            {
                if(previousLanguage != null)
                {
                    Application.Current.Resources.MergedDictionaries.Remove(localizations[previousLanguage]);
                }
                Application.Current.Resources.MergedDictionaries.Add(localizations[selectedLanguage.Value]);
            }
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
        private bool VerifyOutputDirectory()
        {
            return Directory.Exists(OutPath);
        }
        private bool VerifyNamingPattern()
        {
            string tmp = namePattern.Replace("<tempname>", "");
            foreach (var a in VariablesList)
            {
                tmp = tmp.Replace(a.Name, a.Value);
            }
            if (tmp.LastIndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                return false;
            }
            return true;
        }
        private int ActiveTemplatesCount()
        {
            int count = 0;
            foreach (var a in TemplatesList)
            {
                if (a.IsActive) count++;
            }
            return count;
        }
        private bool VerifyMoreThanOneTemplateSelected()
        {
            int count = ActiveTemplatesCount();
            return count > 0;
        }
        private bool VerifyNamingPatternEmpty()
        {
            return namePattern != "";
        }
        private bool VerifyManyTempaltesWithNoTempname()
        {
            return !((ActiveTemplatesCount() > 1) && 
                (namePattern.IndexOf("<tempname>") < 0));
        }
        private bool VerifyInput()
        {
            var currentLocResources = localizations[selectedLanguage.Value];

            if (!VerifyOutputDirectory())
            {
                StatusText = (string)currentLocResources["wrong_output_directory"];
                return false;
            }
            if (!VerifyNamingPatternEmpty())
            {
                StatusText = (string)currentLocResources["naming_pattern_empty"];
                return false;
            }
            if (!VerifyNamingPattern())
            {
                StatusText = (string)currentLocResources["wrong_naming_pattern"];
                return false;
            }
            if (!VerifyMoreThanOneTemplateSelected())
            {
                StatusText = (string)currentLocResources["no_templates_selected"];
                return false;
            }
            if(!VerifyManyTempaltesWithNoTempname())
            {
                StatusText = (string)currentLocResources["many_templates_with_no_tempname"];
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
                    if (!VerifyInput())
                    {
                        return;
                    }
                    var currentLocResources = localizations[selectedLanguage.Value];
                    StatusText = (string)currentLocResources["generation_started"];

                    var outPathFixed = OutPath;
                    if(outPathFixed[outPathFixed.Length-1] != '\\') outPathFixed = outPathFixed + "\\";

                    foreach (var template in TemplatesList)
                    {
                        if(template.IsActive)
                        {
                            try
                            {
                                Models.DocumentGenerator.Generate(template.FileName,
                                    outPathFixed,
                                    NamePattern,
                                    new List<Models.Variable>(VariablesList));
                            }
                            catch(Exception e)
                            {
                                StatusText = e.Message;
                            }
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
        
        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }
    }
}
