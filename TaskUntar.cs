using SevenZip;
using System;
using System.Collections.Generic;
using System.IO;

namespace Utility
{
    public class TaskUntar : Task, ITask
    {
        public string ArchiveFile { get { return _archiveFile.Value(); } }
        public string ExtractFolder { get { return _extractFolder.Value(); } }
        public List<string> ExtractFiles { get; private set; }
        private string _archiveFile;
        private string _extractFolder;
        private List<InfoFile> filesInfo = new List<InfoFile>();

        public TaskUntar(string title, string archiveFile, string extractFolder, InfoFile[] files) : base(title)
        {
            _archiveFile = archiveFile;
            _extractFolder = extractFolder;
            if (files != null) filesInfo.AddRange(files);
            ExtractFiles = new List<string>();
        }

        void ITask.Run()
        {
            if (filesInfo.Count == 0) ExtractAllBy7z();
            else ExtractFilesBy7z();
        }

        public void ExtractAllBy7z()
        {
            SevenZipExtractor.SetLibraryPath(Helper.Library7zPath);
            using (SevenZipExtractor zip = new SevenZipExtractor(ArchiveFile))
            {
                try
                {
                    zip.ExtractArchive(ExtractFolder);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            };
        }

        public void ExtractFilesBy7z()
        {
            SevenZipExtractor.SetLibraryPath(Helper.Library7zPath);
            using (SevenZipExtractor zip = new SevenZipExtractor(ArchiveFile))
            {
                try
                {
                    foreach (var entry in zip.ArchiveFileData)
                    {
                        foreach (InfoFile f in filesInfo)
                        {
                            InfoFile nf = f.SetDefaultFolder(ExtractFolder);
                            if (entry.FileName.Contains(nf.FromFile))
                            {
                                string file = nf.Tofile.Length > 0 ? nf.Tofile : Path.Combine(ExtractFolder, entry.FileName);
                                Directory.CreateDirectory(Path.GetDirectoryName(file));
                                FileStream fs = File.OpenWrite(file);
                                zip.ExtractFile(entry.FileName, fs);
                                ExtractFiles.Add(file);
                                fs.Close();
                                break;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }
    }
}