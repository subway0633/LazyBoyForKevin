using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Utility
{
    public class Task
    {
        public string Title { get { return _title.Value(); } }
        private string _title;

        public bool Issuccess { get; private set; }

        public void Log(Exception ex)
        {
            logMsg.AppendLine($"錯誤：{ex.Message}");
        }

        public void Log(string s)
        {
            logMsg.AppendLine(s);
        }

        private StringBuilder logMsg = new StringBuilder();

        public string Message { get { return logMsg.ToString(); } }

        public Task(string title)
        {
            _title = title ?? string.Empty;
        }

        public void Run()
        {
            try
            {
                if (this is ITask) ((ITask)this).Run();
                Issuccess = true;
            }
            catch (Exception ex)
            {
                Log(ex);
                Issuccess = false;
            }
        }
    }
}