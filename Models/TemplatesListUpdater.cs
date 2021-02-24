using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TemplateDocumentGenerator.Models
{
    /// <summary>
    /// Watches given folder for changes and updates templates list based on detected files.
    /// Also watches folder which contains given folder to react on it's deleting or renaming
    /// </summary>
    class TemplatesListUpdater : IDisposable
    {
        public delegate bool FileCheckFuncDelegate(string s);

        private FileSystemWatcher templatesFolderWatcher;
        private FileSystemWatcher templatesWatcher;
        private ObservableCollection<DocumentTemplate> templatesList;
        private FileCheckFuncDelegate fileCheckFunc;
        /// <summary>
        /// Initializes FileSystemWatchers and events callbacks
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="templatesList"></param>
        /// <param name="fileCheckFunc">Function which checks is file appropriate or not</param>
        public TemplatesListUpdater(string folder, 
            ObservableCollection<DocumentTemplate> templatesList,
            FileCheckFuncDelegate fileCheckFunc)
        {
            this.templatesList = templatesList;
            this.fileCheckFunc = fileCheckFunc;

            if (Directory.Exists(folder))
            {
                templatesWatcher = new FileSystemWatcher(folder);
                templatesWatcher.EnableRaisingEvents = true;
            }
            else templatesWatcher = new FileSystemWatcher();
            templatesWatcher.NotifyFilter = NotifyFilters.LastAccess
                                 | NotifyFilters.LastWrite
                                 | NotifyFilters.FileName
                                 | NotifyFilters.DirectoryName;
            
            templatesWatcher.Created += OnCreated;
            templatesWatcher.Deleted += OnDeleted;
            templatesWatcher.Renamed += OnRenamed;

            templatesFolderWatcher = new FileSystemWatcher(Directory.GetCurrentDirectory());
            templatesFolderWatcher.EnableRaisingEvents = true;
            templatesFolderWatcher.Renamed += OnMainFolderRenamed;
        }
        /// <summary>
        /// Force update state of the folder.
        /// </summary>
        public void ScanFolder()
        {
            templatesList.Clear();
            //if templates watcher was disabled, then given folder doesn't exist
            if (!templatesWatcher.EnableRaisingEvents) return;
            string[] files = Directory.GetFiles(templatesWatcher.Path);
            
            foreach (string s in files)
            {
                if(fileCheckFunc(s))
                {
                    var template = new DocumentTemplate
                    {
                        FileName = s,
                        IsActive = false
                    };

                    templatesList.Add(template);
                }
            }
        }
        public void SetTemplatesList(ObservableCollection<DocumentTemplate> templatesList)
        {
            this.templatesList = templatesList;
            ScanFolder();
        }
        public void Dispose()
        {
            templatesWatcher.Dispose();
        }
        private void OnCreated(object source, FileSystemEventArgs e)
        {
            var template = new DocumentTemplate
            {
                FileName = e.FullPath,
                IsActive = false
            };

            templatesList.Add(template);
        }
        private void OnDeleted(object source, FileSystemEventArgs e)
        {
            templatesList.Remove(templatesList.Where(x => x.ShortFileName == e.Name).Single());
        }
        private void OnRenamed(object source, RenamedEventArgs e)
        {
            foreach(var a in templatesList)
            {
                if (a.ShortFileName == e.OldName)
                {
                    a.FileName = e.FullPath;
                    break;
                }
            }
        }
        private void OnMainFolderRenamed(object source, RenamedEventArgs e)
        {
            if (e.OldName == "Templates")
            {
                templatesList.Clear();
                templatesWatcher.EnableRaisingEvents = false;
            }
            else if (e.Name == "Templates")
            {
                templatesWatcher.Path = e.FullPath;
                templatesWatcher.EnableRaisingEvents = true;
                ScanFolder();
            }
        }
        private void OnMainFolderCreated(object source, FileSystemEventArgs e)
        {
            if (e.Name == "Templates")
            {
                templatesWatcher.Path = e.FullPath;
                templatesWatcher.EnableRaisingEvents = true;
                ScanFolder();
            }
        }
    }
}
