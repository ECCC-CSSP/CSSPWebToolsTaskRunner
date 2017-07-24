using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using CSSPWebToolsTaskRunner.Services;
using CSSPWebToolsTaskRunner.Services.Resources;
using CSSPWebToolsTaskRunner;
using System.Transactions;
using System.Text;
using CSSPWebToolsDBDLL.Models;
using CSSPWebToolsDBDLL.Services;
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;
using System.Threading;
using System.Globalization;

namespace CSSPWebToolsTaskRunner.Services
{
    public class TxtService
    {
        #region Variables
        private string r = "OgW2S3EHhQ(6!Z$odV7eAGnim/#YIClk9vF&1@5xDUa)wPLu*BN.t,c8%JRMbK^yqzXpfTj4sr0:d";
        #endregion Variables

        #region Properties
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        public SamplingPlanService _SampingPlanService { get; private set; }
        public SamplingPlanSubsectorService _SamplingPlanSubsectorService { get; private set; }
        public SamplingPlanSubsectorSiteService _SamplingPlanSubsectorSiteService { get; private set; }
        public MWQMSubsectorService _MWQMSubsectorService { get; private set; }
        public AppTaskService _AppTaskService { get; private set; }
        public TVFileService _TVFileService { get; private set; }
        #endregion Properties
        #region Constructors
        public TxtService(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
            _SampingPlanService = new SamplingPlanService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _SamplingPlanSubsectorService = new SamplingPlanSubsectorService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _SamplingPlanSubsectorSiteService = new SamplingPlanSubsectorSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _MWQMSubsectorService = new MWQMSubsectorService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _AppTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _TVFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        }
        #endregion Constructors

        #region Functions public
        private string CreateCode(string textToCode)
        {
            List<int> intList = new List<int>();
            Random rd = new Random();
            string str = textToCode;
            foreach (char c in str)
            {
                int pos = r.IndexOf(c);
                int first = rd.Next(pos + 1, pos + 9);
                int second = rd.Next(2, 9);
                int tot = (first * second) + pos;
                intList.Add(tot);
                intList.Add(first);
            }

            StringBuilder sb = new StringBuilder();
            foreach (int i in intList)
            {
                sb.Append(i.ToString() + ",");
            }

            return sb.ToString();
        }
        public void CreateSamplingPlanSamplingPlan()
        {
            string NotUsed = "";

            int TVItemID = 0;

            TVItemID = int.Parse(_AppTaskService.GetAppTaskParamStr(_TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "TVItemID"));

            if (TVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.Parameter_NotFound, "TVItemID");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Parameter_NotFound", "TVItemID");
                return;
            }

            int SamplingPlanID = 0;

            SamplingPlanID = int.Parse(_AppTaskService.GetAppTaskParamStr(_TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "SamplingPlanID"));

            if (SamplingPlanID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.Parameter_NotFound, "SamplingPlanID");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Parameter_NotFound", "SamplingPlanID");
                return;
            }

            SamplingPlanModel SamplingPlanModel = _SampingPlanService.GetSamplingPlanModelWithSamplingPlanIDDB(SamplingPlanID);
            if (!string.IsNullOrWhiteSpace(SamplingPlanModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.SamplingPlan, TaskRunnerServiceRes.SamplingPlanID, SamplingPlanID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.SamplingPlan, TaskRunnerServiceRes.SamplingPlanID, SamplingPlanID.ToString());
                return;
            }

            string ServerFilePath = _TVFileService.GetServerFilePath(TVItemID);

            FileInfo fi = new FileInfo(_TVFileService.ChoseEDriveOrCDrive(ServerFilePath) + SamplingPlanModel.SamplingPlanName.Replace("C:\\CSSPLabSheets\\", ""));

            TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            TxtServiceSamplingPlanSamplingPlan txtServiceSamplingPlanSamplingPlan = new TxtServiceSamplingPlanSamplingPlan(_TaskRunnerBaseService);
            txtServiceSamplingPlanSamplingPlan.Generate(fi, SamplingPlanID);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.MWQMSamplingPlanAutoGenerate, FilePurposeEnum.Generated);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            SamplingPlanModel.SamplingPlanFileTVItemID = tvItemModelFile.TVItemID;

            SamplingPlanModel = _SampingPlanService.PostUpdateSamplingPlanDB(SamplingPlanModel);
            if (!string.IsNullOrWhiteSpace(SamplingPlanModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.SamplingPlan, SamplingPlanModel.Error);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.SamplingPlan, SamplingPlanModel.Error);
                return;
            }

        }
        public string GetCodeString(string code)
        {
            string retStr = "";
            List<int> intList = new List<int>();
            List<string> strList = code.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            for (int i = 0, count = strList.Count(); i < count; i = i + 2)
            {
                retStr = retStr + r.Substring((int.Parse(strList[i]) % int.Parse(strList[i + 1])), 1);
            }

            return retStr;
        }
        #endregion Functions public

        #region Functions private
        #endregion Functions private
    }
}
