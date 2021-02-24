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
        private Dictionary<string, ResourceDictionary> localizations;
        private Models.TemplatesListUpdater templatesListUpdater;

        private string statusText;
        private string namePattern;
        private string outPath;

        private string selectedLanguage;
        private string previousLanguage;

        private Models.DocumentTemplate selectedTemplate;

        private object templatesListLock;


        public ObservableCollection<string> Languages { get; set; }
        public ObservableCollection<Models.Variable> VariablesList { get; set; }
        public ObservableCollection<Models.DocumentTemplate> TemplatesList { get; set; }
        public string StatusText
        {
            get { return statusText; }
            set
            {
                statusText = value;
                OnPropertyChanged("StatusText");
            }
        }
        public string SelectedLanguage
        {
            get { return selectedLanguage; }
            set
            {
                if(selectedLanguage != null)
                {
                    previousLanguage = selectedLanguage;
                }
                selectedLanguage = value;
                Settings.Default.Language = selectedLanguage;
                Settings.Default.Save();
                OnPropertyChanged("SelectedLanguage");
            }
        }
        public Models.DocumentTemplate SelectedTemplate
        {
            get { return selectedTemplate; }
            set
            {
                selectedTemplate = value;
                OnPropertyChanged("SelectedTemplate");
            }
        }
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
                Settings.Default.OutPath = outPath;
                Settings.Default.Save();
                OnPropertyChanged("OutPath");
            }
        }
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
            
            Languages = new ObservableCollection<string>();
            localizations = new Dictionary<string, ResourceDictionary>();
            AddLanguage("English", "Properties/Localization_english.xaml");
            AddLanguage("Русский", "Properties/Localization_russian.xaml");
            PropertyChanged += LanguageChanging;

            SelectedLanguage = (string)Settings.Default.Language;

            if(Settings.Default.OutPath == "")
            {
                OutPath = Directory.GetCurrentDirectory();
            }
            else
            {
                OutPath = Settings.Default.OutPath;
            }
        }
        private void AddLanguage(string name, string uri)
        {
            Languages.Add(name);
            localizations[name] = new ResourceDictionary();
            localizations[name].Source = new Uri(uri, UriKind.Relative);
        }
        private void LanguageChanging(object sender, PropertyChangedEventArgs e)
        {
            if(e.PropertyName == "SelectedLanguage")
            {
                if(previousLanguage != null)
                {
                    Application.Current.Resources.MergedDictionaries.Remove(localizations[previousLanguage]);
                }
                Application.Current.Resources.MergedDictionaries.Add(localizations[selectedLanguage]);
                StatusText = "";
            }
        }
        /// <summary>
        /// Adds listener to IsActive property changes.
        /// </summary>
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
        /// <summary>
        /// Checks file acessibility and reloads variables whenever template is set as active.
        /// </summary>
        private void TemplatePropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName != "IsActive") return;

            var changedTempalte = (Models.DocumentTemplate)sender;
            if(changedTempalte.IsActive)
            {
                try
                {
                    using(var dl = new Models.DocumentLoader(changedTempalte.FileName))
                    {
                        changedTempalte.Variables = dl.FindVariables();
                    }

                    var changedTemplateVariables = changedTempalte.Variables;
                    foreach (var variable in changedTemplateVariables)
                    {
                        bool needsInclusion = true;
                        foreach (var item in VariablesList)
                        {
                            if (item.Name == variable)
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
                    StatusText = "";
                }
                catch(System.IO.IOException exception)
                {
                    changedTempalte.IsActive = false;
                    var currentLocResources = localizations[selectedLanguage];
                    StatusText = String.Format(
                        (string)currentLocResources["file_open_error"],
                        changedTempalte.ShortFileName);
                }
                catch (System.IO.FileFormatException exception)
                {
                    changedTempalte.IsActive = false;
                    var currentLocResources = localizations[selectedLanguage];
                    StatusText = String.Format(
                        (string)currentLocResources["wrong_file_format"],
                        changedTempalte.ShortFileName);
                }
                catch (Models.DocumentLoader.IncorrectTemplateError exception)
                {
                    changedTempalte.IsActive = false;
                    var currentLocResources = localizations[selectedLanguage];
                    StatusText = String.Format(
                        (string)currentLocResources["no_variables_detected"],
                        changedTempalte.ShortFileName);
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
                    StatusText = "";
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
            var currentLocResources = localizations[selectedLanguage];

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
        private string GenerationStatusReport(int errorsCount)
        {
            var currentLocResources = localizations[selectedLanguage];
            int activeTemplatesCount = ActiveTemplatesCount();
            string result = "";
            if (errorsCount == 0)
            {
                result = (string)currentLocResources["generation_successful"] + 
                    String.Format($" ({activeTemplatesCount}/{activeTemplatesCount})");
            }
            else
            {
                result = (string)currentLocResources["generation_failed"] +
                    String.Format($" ({activeTemplatesCount-errorsCount}/{activeTemplatesCount})");
            }
            return result;
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

                    var outPathFixed = OutPath;
                    if(outPathFixed[outPathFixed.Length-1] != '\\') outPathFixed = outPathFixed + "\\";
                    
                    int errorsCount = 0;
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
                                //Add error log?
                                errorsCount++;
                            }
                        }
                    }
                    StatusText = GenerationStatusReport(errorsCount);
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
                        if(!Directory.Exists("Templates"))
                        {
                            Directory.CreateDirectory("Templates");
                        }
                        File.Copy(openFileDialog.FileName, "Templates\\" + Path.GetFileName(openFileDialog.FileName));
                    }
                    StatusText = "";
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
                    StatusText = "";
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
                    StatusText = "";
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
