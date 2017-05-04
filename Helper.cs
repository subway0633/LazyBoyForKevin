using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Utility
{
    public static class Helper
    {
        private static string sessionFolder = File.ReadAllLines(Path.Combine(AssemblyFolder, "session_folder.txt"))[0];
        public static string Library7zPath { get { return Path.Combine(AssemblyFolder, "7z.dll"); } }

        public static string AssemblyFolder { get { return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location); } }

        public static string GetIT()
        {
            return string.Join(";", File.ReadAllLines(Path.Combine(AssemblyFolder, "IT.txt")));
        }

        public static string SessionFile()
        {
            return SearchFiile(".zip", sessionFolder);
        }

        public static string LastMonthEndFile()
        {
            string lastMonthEnd = DateTime.Now.AddDays(-DateTime.Now.Day).ToString("yyyy-MM");
            return SearchFiile(lastMonthEnd, sessionFolder);
        }

        public static string SearchFiile(string pattern, string link)
        {
            var directory = new DirectoryInfo(Path.GetDirectoryName(link));
            var file = directory.GetFiles()
                .Where(f => f.Name.Contains(pattern))
                .OrderByDescending(f => f.LastWriteTime)
                .First();
            return file?.FullName;
        }

        public static DateTime SessionDate
        {
            get
            {
                string sessionFile = SessionFile();
                return DateTime.Parse(Path.GetFileNameWithoutExtension(sessionFile));
            }
        }

        public static Config GetConfig(string configName)
        {
            string configPath = Path.Combine(AssemblyFolder, $"config\\{configName}.txt");
            string config = File.ReadAllText(configPath)
                .Replace(@"\n", System.Environment.NewLine)
                .Replace(@"\", @"\\")
                .Replace("$OUTLOOK$", "Utility.TaskOutlook, Utility")
                .Replace("$VBA$", "Utility.TaskVBA, Utility")
                .Replace("$TAR$", "Utility.TaskTar, Utility")
                .Replace("$UNTAR$", "Utility.TaskUntar, Utility")
                .Replace("$FTP$", "Utility.TaskFtp, Utility")
                .Replace("$EXCEL$", "Utility.TaskExcel, Utility")
                .Replace("$COPYFILE$", "Utility.TaskCopyFile, Utility")
                .Replace("$REJECTFILE$", "Utility.TaskRejectFile, Utility");

            List<Task> tasks = JsonConvert.DeserializeObject<List<Task>>(config, new JsonSerializerSettings
            {
                TypeNameHandling = TypeNameHandling.Auto
            });

            return new Config(tasks);
        }
    }
}