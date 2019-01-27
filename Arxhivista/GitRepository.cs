using System;
using System.IO;
using LibGit2Sharp;

namespace Arxhivista
{
    class GitRepository
    {
        private const string GIT_FOLDER_NAME = ".git";

        public static void Init(string repositoryPath)
        {
            if (!Directory.Exists(repositoryPath + GIT_FOLDER_NAME))
            {
                Repository.Init(repositoryPath);
            }
        }

        public static void Add(string repositoryPath)
        {
            using (var repo = new Repository(repositoryPath))
            {
                Commands.Stage(repo, "*");
            }
        }

        public static void Commit(string repositoryPath)
        {
            using (var repo = new Repository(repositoryPath))
            {
                Signature author = new Signature(Arxhivista.SERVICE_NAME, "arxhivista@home.local", DateTime.Now);
                Signature committer = author;
                string message = string.Format("Snapshot created on {0}", DateTime.Now.ToString("yyyy-MM-dd, HH:mm:ss"));
                Commit commit = repo.Commit(message, author, committer);
            }
        }

        public static void Tag(string repositoryPath)
        {
            using (var repo = new Repository(repositoryPath))
            {
                string tag = DateTime.Now.ToString("yyyyMMddHHmmss");
                Tag t = repo.ApplyTag(tag);
            }
        }
    }
}
