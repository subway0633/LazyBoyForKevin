using System;
using System.Collections.Generic;
using System.Text;
using EXCEL = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;
using WORD = Microsoft.Office.Interop.Word;

namespace Utility
{
    public class TaskOutlook : Task, ITask
    {
        private string[] body;
        private List<string> _attachmentsPath = new List<string>();
        private bool sendImmediately;
        private string to;
        private string cc;
        private string encode;

        private string test;

        public TaskOutlook(string title, string[] body, string encode, string[] attachmentsPath, bool sendImmediately, string to, string cc) : base(title)
        {
            test = title;
            this.body = body;
            if (attachmentsPath != null) this._attachmentsPath.AddRange(attachmentsPath);
            this.sendImmediately = sendImmediately;
            this.to = to;
            this.encode = (encode ?? "BIG5").ToUpper();
            this.cc = cc;
        }

        void ITask.Run()
        {
            SendByOutlook();
        }

        private void SendByOutlook()
        {
            Outlook.Application app = new Outlook.Application();
            Outlook.MailItem mailItem = app.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.Subject = this.Title;
            mailItem.To = to;

            if (!string.IsNullOrEmpty(cc)) mailItem.CC = cc;
            if (!string.IsNullOrEmpty(encode))
            {
                //https://msdn.microsoft.com/en-us/library/office/ff860730.aspx
                if (encode == "BIG5") mailItem.InternetCodepage = 950;
                else if (encode == "UTF8") mailItem.InternetCodepage = 65001;
            }

            foreach (string _attachment in _attachmentsPath)
                mailItem.Attachments.Add(_attachment.Value());

            SetMailBody(mailItem);

            if (sendImmediately)
            {
                mailItem.Display(false);
                mailItem.Send();
            }
            else
            {
                mailItem.Display(true);
            }
        }

        private void SetMailBody(Outlook.MailItem mi)
        {
            var isp = mi.GetInspector;
            isp.Activate();
            var doc = isp.WordEditor as WORD.Document;
            if (body == null || body.Length == 0) return;
            doc.Range().Text = body[0].Value();
            for (int i = 1; i < body.Length; i++)
            {
                doc.Range().InsertParagraphAfter();
                string b = body[i];
                if (b.Length > 5 && b.Substring(0, 5).ToUpper() == "EXCEL") CopyExcelContent(b.Split('|'), doc);
                else doc.Range(doc.Range().StoryLength - 1).InsertAfter(b.Value());
            }
        }

        private void CopyExcelContent(string[] excelCommand, WORD.Document doc)
        {
            string file = excelCommand[1].Value();
            string sheet = excelCommand[2].Value();
            string startCell = excelCommand[3];
            string endCell = excelCommand[4];

            EXCEL.Application app = new EXCEL.Application();
            try
            {
                EXCEL.Workbook wb = app.Workbooks.Open(file);
                var destinationSheet = (EXCEL.Worksheet)wb.Sheets[sheet];
                destinationSheet.Activate();
                destinationSheet.Range[startCell, endCell].Copy();
                doc.Range(doc.Range().StoryLength - 1).Paste();
                app.DisplayAlerts = false;
                wb.Close(false);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            }
        }
    }
}