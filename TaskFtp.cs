using Renci.SshNet;
using Renci.SshNet.Common;
using Renci.SshNet.Sftp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;

namespace Utility
{
    public class TaskFtp : Task, ITask
    {
        private string hostName;
        private int port;
        private string userName;
        private string password;
        private bool isByWeb;

        private List<InfoFile> filesInfo = new List<InfoFile>();

        public bool IsBinary { private get; set; }
        public bool IsPassive { private get; set; }

        public TaskFtp(string title, string hostName, string userName, string password, InfoFile[] files, bool isByWeb = true, int port = 21) : base(title)
        {
            this.userName = userName;
            this.port = port;
            this.hostName = hostName;
            this.password = password;
            this.isByWeb = isByWeb;
            IsBinary = IsPassive = true;
            filesInfo.AddRange(files);
        }

        void ITask.Run()
        {
            if (port == 22)
            {
                DownloadViaSftp();
                return;
            }
            string address = $"ftp://{hostName}:{port}";
            if (isByWeb) DownloadViaWeb(address);
            else DownloadViaFtp(address);
        }

        private void DownloadViaWeb(string address)
        {
            using (WebClient webReq = new WebClient())
            {
                webReq.Credentials = new NetworkCredential(userName, password);
                foreach (InfoFile f in filesInfo)
                {
                    string remote = $"{address}{f.FromFile}";
                    string local = $"{address}{f.Tofile}";
                    webReq.DownloadFile(remote, local);
                }
            }
        }

        private void DownloadViaFtp(string address)
        {
            foreach (InfoFile f in filesInfo)
            {
                string remote = $"{address}{f.FromFile}";
                string local = $"{address}{f.Tofile}";

                FtpWebRequest ftpReq = WebRequest.Create(remote) as FtpWebRequest;
                ftpReq.Method = WebRequestMethods.Ftp.DownloadFile;
                ftpReq.UseBinary = IsBinary;
                ftpReq.UsePassive = IsPassive;
                ftpReq.Credentials = new NetworkCredential(userName, password);

                using (FtpWebResponse res = (FtpWebResponse)ftpReq.GetResponse())
                {
                    FileStream fs = new FileStream(local, FileMode.Create, FileAccess.Write);
                    int buffer = 1024;
                    byte[] b = new byte[buffer];
                    int i = 0;
                    Stream stream = res.GetResponseStream();
                    while ((i = stream.Read(b, 0, buffer)) > 0)
                    {
                        fs.Write(b, 0, i);
                    }
                }
            }
        }

        private void DownloadViaSftp()
        {
            using (var sftp = new SftpClient(hostName, port, userName, password))
            {
                sftp.Connect();
                foreach (InfoFile f in filesInfo)
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(f.Tofile));
                    using (var file = File.OpenWrite(f.Tofile))
                    {
                        sftp.DownloadFile(f.FromFile, file);
                    }
                }

                sftp.Disconnect();
            }
        }
    }
}