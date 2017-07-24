using CSSPWebToolsTaskRunner.Services.Resources;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class M21_3FMLogService_Old
    {
        // Variables - Properties
        private int LineNumber = 0;
        private string NotUsed = "";
        private string TheLine = "";
        private string TheValue = "";

        #region Properties
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        public DateTime StartExecutionDate { get; set; }
        public float TotalElapseTimeInSeconds { get; set; }
        public string CompletionTxt { get; set; }
        #endregion Properties

        // Constructor 
        public M21_3FMLogService(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
            StartExecutionDate = new DateTime(1900, 1, 1);
            TotalElapseTimeInSeconds = 0;
            CompletionTxt = "";
        }
        public void Read_M21_3FM_Log(FileInfo fi)
        {
            if (!fi.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", fi.FullName);
                return;
            }

            StreamReader sr = new StreamReader(fi.FullName, Encoding.GetEncoding("iso-8859-1"));
            Read_m21_3fm_DocumentLog(sr);
            sr.Close();
        }
        private string GetTheLineLog(StreamReader srl)
        {
            TheLine = srl.ReadLine();
            LineNumber += 1;
            return TheLine == null ? null : TheLine.Trim();
        }
        private bool CheckNextLineLog(StreamReader srl, string TextToVerify)
        {
            TheLine = srl.ReadLine();
            LineNumber += 1;
            TheLine = TheLine.Trim();
            if (TheLine.Length == 0 && TextToVerify.Length == 0)
            {
                return true;
            }
            else if (TheLine.Length != 0 && TextToVerify.Length == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, TextToVerify);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", TextToVerify);
                return false;
            }
            else if (!TheLine.StartsWith(TextToVerify))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, TextToVerify);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", TextToVerify);
                return false;
            }
            return true;
        }
        private void Read_m21_3fm_DocumentLog(StreamReader sr)
        {
            Read_m21_3fmLog_TopFileInfoLog(sr);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            // skiping many lines
            TheLine = GetTheLineLog(sr);
            if (TheLine == null) return;
            while (!TheLine.Contains("Overall Timings"))
            {
                // reading until we find the line that contains "Overall Timings"
                TheLine = GetTheLineLog(sr);
                if (TheLine == null) return;
            }

            Read_m21_3fmLog_OverallTimings(sr);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            while (!TheLine.Contains("Performance"))
            {
                // reading until we find the line that contains "Performance"
                TheLine = GetTheLineLog(sr);
                if (TheLine == null) return;
            }

            Read_m21_3fmLog_Completion(sr);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;
        }
        private void Read_m21_3fmLog_TopFileInfoLog(StreamReader sr)
        {
            TheLine = GetTheLineLog(sr);
            TheValue = TheLine.Substring(TheLine.IndexOf(":") + 1).Trim();
            int Year = int.Parse(TheValue.Substring(0, 4));
            int Month = int.Parse(TheValue.Substring(4, 2));
            int Day = int.Parse(TheValue.Substring(6, 2));

            TheLine = GetTheLineLog(sr);
            TheValue = TheLine.Substring(TheLine.IndexOf(":") + 1).Trim();
            int Hour = int.Parse(TheValue.Substring(0, 2));
            int Minute = int.Parse(TheValue.Substring(2, 2));
            int Second = int.Parse(TheValue.Substring(4, 2));

            StartExecutionDate = new DateTime(Year, Month, Day, Hour, Minute, Second);
        }
        private void Read_m21_3fmLog_OverallTimings(StreamReader sr)
        {
            TheLine = GetTheLineLog(sr);
            if (TheLine == null) return;
            while (!TheLine.Trim().StartsWith("Total"))
            {
                // reading until we find the line that Starts with "Total"
                TheLine = GetTheLineLog(sr);
                if (TheLine == null) return;
            }

            TheValue = TheLine.Trim();
            TotalElapseTimeInSeconds = (int)float.Parse(TheValue.Substring(TheValue.LastIndexOf(" ")));
        }
        private void Read_m21_3fmLog_Completion(StreamReader sr)
        {
            TheLine = GetTheLineLog(sr);
            if (TheLine == null) return;
            TheLine = GetTheLineLog(sr);
            if (TheLine == null) return;
            TheLine = GetTheLineLog(sr);
            if (TheLine == null) return;
            TheLine = GetTheLineLog(sr);
            if (TheLine == null) return;
            TheValue = TheLine.Trim();
            CompletionTxt = TheValue;
        }
    }
}
