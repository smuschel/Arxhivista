using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.ServiceProcess;

namespace Arxhivista
{
    public partial class Arxhivista : ServiceBase
    {
        private const string APP_DATA = "APPDATA";
        private const string USER_PROFILE = "USERPROFILE";
        private const string DOCUMENTS_FOLDER_NAME = "Documents";

        private const string DOCX_EXTENSION = "docx";
        private const string PPTX_EXTENSION = "pptx";
        private const string XLSX_EXTENSION = "xlsx";
        private const string OFFICE_TEMP_PREFIX = "~$";

        private const string ARXHIVISTA_TEMPFILE_POSTFIX = ".arxhivsta-tmp";
        private const string ARXHIVISTA_META_FOLDER = ".arxhivista";
        private const string GITIGNORE = ".gitignore"; 

        public const string SERVICE_NAME = "Arxhivista";
        protected string appDataPath;
        protected string userProfileDocumentsPath;
        protected FileSystemWatcher watcher;
        private List<FileSystemWatcher> definedWatchersList;

        public Arxhivista()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {            
            appDataPath = Environment.GetEnvironmentVariable(APP_DATA);
            CreateRepositoryRoot();
            userProfileDocumentsPath = Environment.GetEnvironmentVariable(USER_PROFILE) + "\\" + DOCUMENTS_FOLDER_NAME;
            definedWatchersList = new List<FileSystemWatcher>();

            LogEvent("Targeting Repo folder " + appDataPath);
            LogEvent("Listening for file changes in " + userProfileDocumentsPath);

            CreateWatcher("*.docx");
            CreateWatcher("*.pptx");
            CreateWatcher("*.xlsx");
            
        }

        private void OnCreated(object sender, FileSystemEventArgs e)
        {
            LogEvent("Create " + e.FullPath);
            HandleRelevantFileSystemEvent(e.FullPath, e.Name);
        }

        private void OnRenamed(object sender, RenamedEventArgs e)
        {
            LogEvent("Rename " + e.FullPath + " " + e.OldName);
            HandleRelevantFileSystemEvent(e.FullPath, e.Name);
        }

        private void HandleRelevantFileSystemEvent(string fullPath, string name)
        {
            try
            {
                if ((name.EndsWith(DOCX_EXTENSION) || name.EndsWith(PPTX_EXTENSION) || name.EndsWith(XLSX_EXTENSION)) && !name.Contains(OFFICE_TEMP_PREFIX))
                {
                    CreateOfficeDocumentRepository(name);
                    CopyDocumentContentToRepository(fullPath, name);
                    UpdateGitRepository(name);
                }
            } catch (Exception e)
            {
                LogEvent("HRFSE: " + e.Message);
            }
        }

        protected override void OnStop()
        {
            foreach (FileSystemWatcher watcher in definedWatchersList)
            {
                watcher.EnableRaisingEvents = false;
                watcher.Dispose();
            }

            definedWatchersList.Clear();
            
        }

        private void LogEvent(string message)
        {
            string eventSource = SERVICE_NAME + " Monitoring Service";
            DateTime dt = new DateTime();
            dt = System.DateTime.UtcNow;
            message = dt.ToLocalTime() + ": " + message;

            EventLog.WriteEntry(eventSource, message);
        }

        private void CreateRepositoryRoot()
        {
            try
            {
                string repositoryPath = appDataPath + "\\" + SERVICE_NAME;
                if (!Directory.Exists(repositoryPath))
                {
                    Directory.CreateDirectory(repositoryPath);
                }
            } catch (Exception e)
            {
                LogEvent("CRR: " + e.Message);
            }
        }

        private void UpdateGitRepository(string name)
        {
            string archivePath = GetRepositoryRootPath() + name + "\\";
            try
            {
                GitRepository.Init(archivePath);
                GitRepository.Add(archivePath);
                GitRepository.Commit(archivePath);
                GitRepository.Tag(archivePath);
            } catch (Exception e)
            {
                LogEvent("UGR: " + e.Message);
            }
        }

        private string GetRepositoryRootPath()
        {
            return appDataPath + "\\" + SERVICE_NAME + "\\";
        }

        private void CreateOfficeDocumentRepository(string name)
        {
            try
            {
                string archivePath = GetRepositoryRootPath() + name;
                if (!Directory.Exists(archivePath))
                {
                    LogEvent("Creating repository for " + name + " in " + archivePath);
                    Directory.CreateDirectory(archivePath);
                }

                string arxhivistaMetaPath = archivePath + "\\" + ARXHIVISTA_META_FOLDER;
                if (!Directory.Exists(arxhivistaMetaPath))
                {
                    Directory.CreateDirectory(arxhivistaMetaPath);
                    CreateGitignore(archivePath);
                }
            } catch (Exception e)
            {
                LogEvent("CODR: " + e.Message);
            }
        }

        private void CreateGitignore(string archivePath)
        {
            File.WriteAllText(archivePath + "\\" + GITIGNORE, ARXHIVISTA_META_FOLDER + "/");
        }

        private void CreateWatcher(string filter)
        {
            watcher = new FileSystemWatcher(userProfileDocumentsPath, filter);
            watcher.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.Attributes;
            watcher.IncludeSubdirectories = true;
            watcher.Renamed += new RenamedEventHandler(OnRenamed);
            watcher.Created += new FileSystemEventHandler(OnCreated);
            watcher.EnableRaisingEvents = true;

            definedWatchersList.Add(watcher);
        }

        private void CopyDocumentContentToRepository(string fullPath, string name)
        {
            string archivePath = GetRepositoryRootPath() + name;
            
            try
            {
                File.Copy(fullPath, fullPath + ARXHIVISTA_TEMPFILE_POSTFIX);
                DirectoryInfo di = new DirectoryInfo(archivePath);

                foreach (FileInfo file in di.GetFiles())
                {
                    file.Delete();
                }
                foreach (DirectoryInfo dir in di.GetDirectories())
                {
                    if (dir.Name.Equals(".git"))
                    {
                        // skip, need Git repo info!
                    } else
                    {
                        dir.Delete(true);
                    }                    
                }
                ZipFile.ExtractToDirectory(fullPath + ARXHIVISTA_TEMPFILE_POSTFIX, archivePath);
                File.Delete(fullPath + ARXHIVISTA_TEMPFILE_POSTFIX);
            } catch (System.IO.IOException ioe)
            {
                LogEvent("CDCTR: " + ioe.Message);
            }            
        }
    }
}
