using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Utility;

namespace LazyBoy
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            if (args.Length == 0) return;
            StringBuilder errorMsg = new StringBuilder();
            string configName = args[0];
            try
            {
                Config c = Helper.GetConfig(configName);
                foreach (var t in c.Tasks)
                {
                    t.Run();
                    if (!t.Issuccess)
                        errorMsg.AppendLine($"{t.Title} : {t.Message}");
                }
            }
            catch (Exception ex)
            {
                errorMsg.AppendLine(ex.Message);
            }
            finally
            {
                string errorLogFolder = Directory.GetParent(Helper.AssemblyFolder).FullName + $"\\{configName}\\ERROR_LOG";
                Directory.CreateDirectory(errorLogFolder);
                if (errorMsg.Length > 0)
                {
                    File.WriteAllText($"{errorLogFolder}\\log{DateTime.Now.ToString("yyyyMMdd-HHmmss")}.txt", errorMsg.ToString());
                    try
                    {
                        TaskOutlook t = new TaskOutlook(configName + "排程失敗", new string[] { errorMsg.ToString() }, null, null, true, Helper.GetIT(), null);
                        t.Run();
                    }
                    catch { }
                }
                else
                {
                    File.WriteAllText($"{errorLogFolder}\\log{DateTime.Now.ToString("yyyyMMdd")}-ok.txt", "done");
                }
            }
        }
    }
}