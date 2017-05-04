using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using EXCEL = Microsoft.Office.Interop.Excel;

namespace Utility
{
    public class TaskExcel : Task, ITask
    {
        private InfoFile fileInfo;
        private string template;
        private string sheet;

        public TaskExcel(InfoFile file, string sheet, string template, string title) : base(title)
        {
            this.template = template;
            this.fileInfo = file;
            this.sheet = sheet;
        }

        void ITask.Run()
        {
            ImportCsv();
        }

        private void ImportCsv()
        {
            EXCEL.Application appSource = new EXCEL.Application();
            try
            {
                string[] csvContent = File.ReadAllLines(fileInfo.FromFile);
                EXCEL.Workbook wbSource = appSource.Workbooks.Open(template);
                var destinationSheet = (EXCEL.Worksheet)wbSource.Sheets[sheet];
                destinationSheet.Activate();
                for (int i = 0; i < csvContent.Length; i++)
                {
                    string line = csvContent[i];
                    TextFieldParser parser = new TextFieldParser(new StringReader(line));

                    // You can also read from a file
                    // TextFieldParser parser = new TextFieldParser("mycsvfile.csv");

                    parser.HasFieldsEnclosedInQuotes = true;
                    parser.SetDelimiters(",");

                    string[] fields;

                    while (!parser.EndOfData)
                    {
                        fields = parser.ReadFields();
                        for (int j = 0; j < fields.Length; j++)
                        {
                            destinationSheet.Cells[i + 1, j + 1] = fields[j];
                        }
                    }

                    parser.Close();
                }

                wbSource.SaveCopyAs(fileInfo.Tofile);
                wbSource.Close(false);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                appSource.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(appSource);
            }
        }
    }
}