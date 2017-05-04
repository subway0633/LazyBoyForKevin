using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using EXCEL = Microsoft.Office.Interop.Excel;

namespace Utility
{
    public class TaskVBA : Task, ITask
    {
        private string file;
        private InfoMacro[] macros;
        private string workSheet;

        public TaskVBA(string title, string file, InfoMacro[] macros, string workSheet) : base(title)
        {
            this.file = file;
            this.macros = macros;
            this.workSheet = workSheet;
        }

        void ITask.Run()
        {
            RunningMarcro();
        }

        private void RunningMarcro()
        {
            EXCEL.Application appSource = new EXCEL.Application();
            EXCEL.Workbook wbSource = appSource.Workbooks.Open(file);

            try
            {
                var destinationSheet = (EXCEL.Worksheet)wbSource.Sheets[workSheet];
                destinationSheet.Activate();

                bool isRefreshSuccess = false;
                for (int m = 0; m < macros.Length; m++)
                {
                    InfoMacro im = macros[m];

                    for (int i = 0; i < 4; i++)
                    {
                        System.Threading.Thread.Sleep(30000);

                        string v2 = destinationSheet.Range[im.CheckCell].Value.ToString();

                        if (v2.Contains(im.CheckValue))
                        {
                            isRefreshSuccess = true;
                            break;
                        }
                        else appSource.RTD.RefreshData();
                    }

                    if (isRefreshSuccess)
                    {
                        appSource.GetType().InvokeMember("Run", BindingFlags.Default | BindingFlags.InvokeMethod, null, appSource, new string[] { im.Command });
                    }
                    else
                    {
                        Log(new Exception("作業失敗，請手動執行"));
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                wbSource.Close(false);
                appSource.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(appSource);
            }
        }
    }
}