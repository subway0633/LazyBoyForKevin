using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace Utility
{
    public class TaskCopyFile : Task, ITask

    {
        public List<string> CopyFiles { get; private set; }

        private List<InfoFile> filesInfo = new List<InfoFile>();

        public TaskCopyFile(string title, InfoFile[] files) : base(title)
        {
            if (files != null) filesInfo.AddRange(files);
            CopyFiles = new List<string>();
        }

        void ITask.Run()
        {
            CopyAllFile();
        }

        public void CopyAllFile()
        {
            {
                foreach (InfoFile f in filesInfo)
                    try
                    {
                        File.Copy(f.FromFile, f.Tofile);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
            };
        }
    }
}