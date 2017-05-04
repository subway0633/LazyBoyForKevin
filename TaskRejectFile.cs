using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Utility
{
    public class TaskRejectFile : Task, ITask
    {
        private const string defaultQueryWord = "Algo spec not defined";
        private string[] defaultIgnoreSubtype = new string[] { "CDO", "Structured Note" };

        private TaskUntar untarTask;

        private string queryWord;
        private string[] ignoreSubtype;
        private string to;

        public TaskRejectFile(string title, TaskUntar untarTask, string queryWord, string[] ignoreSubtype, string to) : base(title)
        {
            this.untarTask = untarTask;
            this.queryWord = queryWord ?? defaultQueryWord;
            this.ignoreSubtype = ignoreSubtype ?? defaultIgnoreSubtype;
            this.to = to;
        }

        void ITask.Run()
        {
            List<string> rejectErmName = new List<string>();
            untarTask.Run();
            string rejectFile = untarTask.ExtractFiles[0];
            string[] csvContent = File.ReadAllLines(rejectFile);
            for (int i = 0; i < csvContent.Length; i++)
            {
                string line = csvContent[i];
                if (!line.Contains(queryWord)) continue;
                bool isIgnoreSubtype = false;

                foreach (string subType in ignoreSubtype)
                {
                    if (line.ToUpper().Contains($"SUB TYPE = {subType.ToUpper()}"))
                    {
                        isIgnoreSubtype = true;
                        break;
                    }
                }
                if (isIgnoreSubtype) continue;

                TextFieldParser parser = new TextFieldParser(new StringReader(line));
                parser.HasFieldsEnclosedInQuotes = true;
                parser.SetDelimiters(",");

                string[] fields;
                while (!parser.EndOfData)
                {
                    fields = parser.ReadFields();
                    rejectErmName.Add(fields[1]);
                }
                parser.Close();
            }
            if (rejectErmName.Count == 0) rejectErmName.Add(@"Nothing to do!! \^_^/");
            TaskOutlook outlook = new TaskOutlook(Title, rejectErmName.ToArray(), "big5", null, true, to, null);
            outlook.Run();
        }
    }
}