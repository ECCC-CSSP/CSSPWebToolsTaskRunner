using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using CSSPWebToolsTaskRunner.Services.Resources;
using System.Transactions;
using System.Text;
using CSSPDBDLL.Models;
using CSSPDBDLL.Services;
using System.Threading;
using System.Globalization;
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;

namespace CSSPWebToolsTaskRunner.Services
{
    public class TxtServiceSamplingPlanSamplingPlan
    {
        #region Variables
        private string r = "OgW2S3EHhQ(6!Z$odV7eAGnim/#YIClk9vF&1@5xDUa)wPLu*BN.t,c8%JRMbK^yqzXpfTj4sr0:d";

        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        public SamplingPlanService _SamplingPlanService { get; private set; }
        public SamplingPlanSubsectorService _SamplingPlanSubsectorService { get; private set; }
        public SamplingPlanSubsectorSiteService _SamplingPlanSubsectorSiteService { get; private set; }
        public MWQMSubsectorService _MWQMSubsectorService { get; private set; }
        #endregion Variables

        #region Constructors
        public TxtServiceSamplingPlanSamplingPlan(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
            _SamplingPlanService = new SamplingPlanService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _SamplingPlanSubsectorService = new SamplingPlanSubsectorService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _SamplingPlanSubsectorSiteService = new SamplingPlanSubsectorSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _MWQMSubsectorService = new MWQMSubsectorService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
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
        public void Generate(FileInfo fi, int SamplingPlanID)
        {
            string NotUsed = "";

            if (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr)
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-CA");
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("fr-CA");
            }
            else
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-CA");
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-CA");
            }

            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);

            DirectoryInfo di = new DirectoryInfo(ServerFilePath);
            if (!di.Exists)
                di.Create();

            if (fi.Exists)
                fi.Delete();

            SamplingPlanModel SamplingPlanModel = _SamplingPlanService.GetSamplingPlanModelWithSamplingPlanIDDB(SamplingPlanID);
            if (!string.IsNullOrWhiteSpace(SamplingPlanModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.SamplingPlan, TaskRunnerServiceRes.SamplingPlanID, SamplingPlanID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.SamplingPlan, TaskRunnerServiceRes.SamplingPlanID, SamplingPlanID.ToString());
                return;
            }

            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemModel tvItemModelProvince = tvItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);

            StringBuilder sb = new StringBuilder();

            sb.AppendLine("Version\t1\t");
            sb.AppendLine("Sampling Plan Type\t" + SamplingPlanModel.SamplingPlanType.ToString() + "\t");
            sb.AppendLine("Sample Type\t" + SamplingPlanModel.SampleType.ToString() + "\t");
            sb.AppendLine("Lab Sheet Type\t" + SamplingPlanModel.LabSheetType.ToString() + "\t");

            List<SamplingPlanSubsectorModel> SamplingPlanSubsectorModelList = _SamplingPlanSubsectorService.GetSamplingPlanSubsectorModelListWithSamplingPlanIDDB(SamplingPlanID);

            foreach (SamplingPlanSubsectorModel SamplingPlanSubsectorModel in SamplingPlanSubsectorModelList)
            {
                MWQMSubsectorModel mwqmSubsectorModel = _MWQMSubsectorService.GetMWQMSubsectorModelWithMWQMSubsectorTVItemIDDB(SamplingPlanSubsectorModel.SubsectorTVItemID);
                if (!string.IsNullOrWhiteSpace(mwqmSubsectorModel.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MWQMSubsector, TaskRunnerServiceRes.SubsectorTVItemID, SamplingPlanSubsectorModel.SubsectorTVItemID.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MWQMSubsector, TaskRunnerServiceRes.SubsectorTVItemID, SamplingPlanSubsectorModel.SubsectorTVItemID.ToString());
                }

                List<SamplingPlanSubsectorSiteModel> SamplingPlanSubsectorSiteModelList = _SamplingPlanSubsectorSiteService.GetSamplingPlanSubsectorSiteModelListWithSamplingPlanSubsectorIDDB(SamplingPlanSubsectorModel.SamplingPlanSubsectorID);

                if (SamplingPlanSubsectorSiteModelList.Count > 0)
                {
                    sb.AppendLine("Subsector\t" + SamplingPlanSubsectorModel.SubsectorTVText + "\t" + SamplingPlanSubsectorModel.SubsectorTVItemID + "\t" + mwqmSubsectorModel.TideLocationSIDText);


                    sb.Append("MWQM Sites\t");
                    foreach (SamplingPlanSubsectorSiteModel SamplingPlanSubsectorSiteModel in SamplingPlanSubsectorSiteModelList)
                    {
                        sb.Append(SamplingPlanSubsectorSiteModel.MWQMSiteText + ",");
                    }
                    sb.Append("\t");
                    foreach (SamplingPlanSubsectorSiteModel SamplingPlanSubsectorSiteModel in SamplingPlanSubsectorSiteModelList)
                    {
                        sb.Append(SamplingPlanSubsectorSiteModel.MWQMSiteTVItemID + ",");
                    }
                    sb.AppendLine("");
                }

                if (SamplingPlanSubsectorSiteModelList.Where(c => c.IsDuplicate == true).Count() > 0)
                {
                    sb.Append("Daily Duplicate\t");
                    foreach (SamplingPlanSubsectorSiteModel SamplingPlanSubsectorSiteModel in SamplingPlanSubsectorSiteModelList.Where(c => c.IsDuplicate == true))
                    {
                        sb.Append(SamplingPlanSubsectorSiteModel.MWQMSiteText + ",");
                    }
                    sb.Append("\t");
                    foreach (SamplingPlanSubsectorSiteModel SamplingPlanSubsectorSiteModel in SamplingPlanSubsectorSiteModelList.Where(c => c.IsDuplicate == true))
                    {
                        sb.Append(SamplingPlanSubsectorSiteModel.MWQMSiteTVItemID + ",");
                    }
                    sb.AppendLine("");
                }
            }

            sb.AppendLine("App\t" + CreateCode(SamplingPlanModel.AccessCode) + "\t" + CreateCode(SamplingPlanModel.Year.ToString()));
            sb.AppendLine("Precision Criteria\t" + CreateCode(SamplingPlanModel.DailyDuplicatePrecisionCriteria.ToString()) + "\t" + CreateCode(SamplingPlanModel.IntertechDuplicatePrecisionCriteria.ToString()));
            sb.AppendLine("Include Laboratory QA/QC\t" + SamplingPlanModel.IncludeLaboratoryQAQC.ToString() + "\t" + CreateCode(SamplingPlanModel.ApprovalCode));
            sb.AppendLine("Backup Directory\t" + SamplingPlanModel.BackupDirectory);

            StreamWriter sw = fi.CreateText();
            sw.Write(sb.ToString());
            sw.Close();
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

    }

}
