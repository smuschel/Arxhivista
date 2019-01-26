using LibGit2Sharp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace Arxhivista
{
    public partial class Arxhivista : ServiceBase
    {
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
            appDataPath = Environment.GetEnvironmentVariable("APPDATA");
            CreateRepositoryRoot();
            userProfileDocumentsPath = Environment.GetEnvironmentVariable("USERPROFILE") + "\\Documents";
            definedWatchersList = new List<FileSystemWatcher>();

            LogEvent("Targeting Repo folder " + appDataPath);
            LogEvent("Listening for DOCX Changes in " + userProfileDocumentsPath);

            CreateWatcher("*.docx");
            CreateWatcher("*.pptx");
            CreateWatcher("*.xlsx");
            
        }

        private void OnCreated(object sender, FileSystemEventArgs e)
        {
            LogEvent("Create " + e.FullPath);
            if (e.FullPath.EndsWith(".docx") || e.FullPath.EndsWith(".pptx") || e.FullPath.EndsWith(".xlsx"))
            {
                CreateOfficeDocumentRepository(e.Name);
                CopyDocumentContentToRepository(e.FullPath, e.Name);
                UpdateGitRepository(e.Name);
            }
        }

        private void OnRenamed(object sender, RenamedEventArgs e)
        {
            LogEvent("Rename " + e.FullPath + " " + e.OldName);
            if (e.FullPath.EndsWith(".docx") || e.FullPath.EndsWith(".pptx") || e.FullPath.EndsWith(".xlsx"))
            {
                CreateOfficeDocumentRepository(e.Name);
                CopyDocumentContentToRepository(e.FullPath, e.Name);
                UpdateGitRepository(e.Name);
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
            string eventSource = "Arxhivista Monitor Service";
            DateTime dt = new DateTime();
            dt = System.DateTime.UtcNow;
            message = dt.ToLocalTime() + ": " + message;

            EventLog.WriteEntry(eventSource, message);
        }

        private void CreateRepositoryRoot()
        {
            string repositoryPath = appDataPath + "\\Arxhivista";
            if (!Directory.Exists(repositoryPath))
            {
                Directory.CreateDirectory(repositoryPath);
            }
        }

        private void UpdateGitRepository(string name)
        {
            string archivePath = GetRepositoryRootPath() + name;
            GitInit(archivePath);
            GitAdd(archivePath);
            GitCommit(archivePath);
            GitTag(archivePath);
        }

        private string GetRepositoryRootPath()
        {
            return appDataPath + "\\Arxhivista\\";
        }

        private void CreateOfficeDocumentRepository(string name)
        {
            string archivePath = GetRepositoryRootPath() + name;
            if (!Directory.Exists(archivePath))
            {
                LogEvent("Creating Arxhivista repository for " + name + " in " + archivePath);
                Directory.CreateDirectory(archivePath);
            }
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
            File.Copy(fullPath, fullPath + ".arxhivista-tmp");
            try
            {
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
                ZipFile.ExtractToDirectory(fullPath + ".arxhivista-tmp", archivePath);                
            } catch (System.IO.IOException ioe)
            {
                LogEvent(ioe.Message);
            }
            File.Delete(fullPath + ".arxhivista-tmp");
        }

        private void GitInit(string repositoryPath)
        {
            if (!Directory.Exists(repositoryPath + "\\.git"))
            {
                Repository.Init(repositoryPath);
            }
        }

        private void GitAdd(string repositoryPath)
        {
            using(var repo = new Repository(repositoryPath))
            {
                Commands.Stage(repo, "*");
            }
        }

        private void GitCommit(string repositoryPath)
        {
            using (var repo = new Repository(repositoryPath))
            {
                Signature author = new Signature("Arxhivista", "arxhivista@home.local", DateTime.Now);
                Signature committer = author;
                string message = string.Format("Snapshot created on {0}", DateTime.Now.ToString("yyyy-MM-dd, HH:mm:ss"));
                Commit commit = repo.Commit(message, author, committer);
            }
        }

        private void GitTag(string repositoryPath)
        {
            using (var repo = new Repository(repositoryPath))
            {
                string tag = DateTime.Now.ToString("yyyyMMddHHmmss");
                Tag t = repo.ApplyTag(tag);
            }
        }
    }
}
