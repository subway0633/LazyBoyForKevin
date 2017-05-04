using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Utility
{
    public static class StringExt
    {
        private static readonly string lastMonth = "LASTMONTH";
        private static readonly string lastMonthEndFile = "LASTMONTHENDFILE";
        private static readonly string sessionDate = "SESSIONDATE";
        private static readonly string sessionFile = "SESSIONFILE";
        private static readonly string quarter = "QUARTER";

        public static string Value(this string originalText)
        {
            if (string.IsNullOrEmpty(originalText)) return originalText;
            if (originalText == lastMonthEndFile) return Helper.LastMonthEndFile();
            if (originalText == sessionFile) return Helper.SessionFile();
            string result = ReplaceSpecialDate(originalText);
            result = ReplaceNormalDate(result);
            result = ReplaceQuarter(result);
            return result;
        }

        /// <summary>
        /// 解析特殊日期
        /// </summary>
        /// <param name="originalText"></param>
        /// <returns></returns>
        private static string ReplaceSpecialDate(string originalText)
        {
            Regex regex = new Regex($@"[%]([\w-/]+)(;?)({lastMonth}|{sessionDate})[%]");
            Match m = regex.Match(originalText);

            string keyWord = string.Empty;
            DateTime keyDate = new DateTime();
            string format = null;
            if (m.Success)
            {
                format = m.Groups[1].ToString();
                keyWord = m.Groups[3].ToString();
                if (keyWord == sessionDate) keyDate = Helper.SessionDate;
                else if (keyWord == lastMonth) keyDate = DateTime.Now.AddDays(-DateTime.Now.Day);
                else keyWord = null;
            }
            else return originalText;

            if (keyWord == null) return "尚未實作此種關鍵字檢合法";

            return originalText.Replace(m.Groups[0].ToString(), keyDate.ToString(format));
        }

        /// <summary>
        /// 解析 {日期格式};+-{日期平移數}
        /// </summary>
        /// <param name="originalText">原始文字</param>
        /// <returns></returns>
        private static string ReplaceNormalDate(string originalText)
        {
            string result = originalText;
            Regex regex = new Regex(@"[%]([\w-/]+)(;?)([+-]?)([\d]*)[%]");
            Match m = regex.Match(result);
            while (m.Success)
            {
                double shift;
                double.TryParse(m.Groups[3].ToString() + m.Groups[4].ToString(), out shift);
                DateTime date = DateTime.Today.AddDays(shift);
                result = result.Replace(m.Groups[0].ToString(), date.ToString(m.Groups[1].ToString()));
                m = m.NextMatch();
            }

            return result;
        }

        private static string ReplaceQuarter(string originalText)
        {
            string result = originalText;
            Regex regex = new Regex($@"[%]({quarter})(,(\d+))" + "{4}[%]");
            Match m = regex.Match(result);

            if (m.Success)
            {
                string patten = m.Groups[0].ToString();
                var qArray = patten.Replace("%", string.Empty).Replace(quarter, string.Empty).Split(',').ToList();
                int q;
                try
                {
                    q = qArray.IndexOf(DateTime.Now.Month.ToString());
                }
                catch
                {
                    q = 0;
                }
                result = originalText.Replace(patten, $"{DateTime.Now.Year} Q{q}");
            }
            return result;
        }
    }
}