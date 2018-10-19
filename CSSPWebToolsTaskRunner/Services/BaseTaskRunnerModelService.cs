using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;
using CSSPDBDLL.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace CSSPWebToolsTaskRunner.Services
{
    // Class
    public class BWObj
    {
        public BWObj()
        {
            bw = new BackgroundWorker();
            appTaskModel = new AppTaskModel();
            TextLanguageList = new List<TextLanguage>();
        }

        public BackgroundWorker bw { get; set; }
        public int Index { get; set; }
        public AppTaskCommandEnum appTaskCommand { get; set; }
        public AppTaskModel appTaskModel { get; set; }
        public List<TextLanguage> TextLanguageList { get; set; }
    }
    public class OtherFileInfo
    {
        public int TVFileID { get; set; }
        public string ClientFullFileName { get; set; }
        public string ServerFullFileName { get; set; }
        public bool IsOutput { get; set; }
    }
    public class PeakDifference
    {
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public float Value { get; set; }
    }
    public class TextLanguage
    {
        public LanguageEnum Language { get; set; }
        public string Text { get; set; }
    }
    public class UserStateObj
    {
        public string ProgressText { get; set; }
    }
   
}
