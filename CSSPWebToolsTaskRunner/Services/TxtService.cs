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
using CSSPWebToolsDBDLL;

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
        public TVItemService _TVItemService { get; private set; }
        public PolSourceSiteService _PolSourceSiteService { get; private set; }
        public AddressService _AddressService { get; private set; }
        public PolSourceObservationService _PolSourceObservationService { get; private set; }
        public PolSourceObservationIssueService _PolSourceObservationIssueService { get; private set; }
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
            _TVItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _PolSourceSiteService = new PolSourceSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _AddressService = new AddressService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _PolSourceObservationService = new PolSourceObservationService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _PolSourceObservationIssueService = new PolSourceObservationIssueService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
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
        public void CreateCSVOfMWQMSites()
        {
            string NotUsed = "";
            string ProvInit = "";
            List<string> ProvInitList = new List<string>()
            {
                "BC", "ME", "NB", "NL", "NS", "PE", "QC",
            };
            List<string> ProvList = new List<string>()
            {
                "British Columbia", "Maine", "New Brunswick", "Newfoundland and Labrador", "Nova Scotia", "Prince Edward Island", "Québec",
            };

            int TVItemID = _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID;

            if (TVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.Parameter_NotFound, "TVItemID");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Parameter_NotFound", "TVItemID");
                return;
            }

            TVItemModel tvItemModel = _TVItemService.GetTVItemModelWithTVItemIDDB(TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotDeleteFile_Error_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                return;
            }

            for (int i = 0, countProv = ProvList.Count; i < countProv; i++)
            {
                if (ProvList[i] == tvItemModel.TVText)
                {
                    ProvInit = ProvInitList[i];
                    break;
                }
            }

            string ServerFilePath = _TVFileService.GetServerFilePath(TVItemID);

            FileInfo fi = new FileInfo(_TVFileService.ChoseEDriveOrCDrive(ServerFilePath) + $"Sites_{ProvInit}.csv");

            TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            // loop through all the MWQMSites etc...

            StringBuilder sb = new StringBuilder();

            sb.AppendLine("Site_ID,Lat,Long");

            if (fi.Exists)
            {
                try
                {
                    fi.Delete();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDeleteFile_Error_, fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDeleteFile_Error_", fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    return;
                }
            }

            using (CSSPWebToolsDBEntities db = new CSSPWebToolsDBEntities())
            {
                var tvItemProv = (from c in db.TVItems
                                  from cl in db.TVItemLanguages
                                  where c.TVItemID == cl.TVItemID
                                  && c.TVItemID == TVItemID
                                  && cl.Language == (int)LanguageEnum.en
                                  && c.TVType == (int)TVTypeEnum.Province
                                  select new { c, cl }).FirstOrDefault();

                if (tvItemProv == null)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                    return;
                }

                for (int i = 0, countProv = ProvList.Count; i < countProv; i++)
                {
                    if (ProvList[i] == tvItemProv.cl.TVText)
                    {
                        ProvInit = ProvInitList[i];
                        break;
                    }
                }

                NotUsed = string.Format(TaskRunnerServiceRes.Creating_, fi.Name);
                List<TextLanguage> TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", fi.Name);

                _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);

                var tvItemSSList = (from t in db.TVItems
                                    from tl in db.TVItemLanguages
                                    where t.TVItemID == tl.TVItemID
                                    && tl.Language == (int)LanguageEnum.en
                                    && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                    && t.TVType == (int)TVTypeEnum.Subsector
                                    orderby tl.TVText
                                    select new { t, tl }).ToList();

                var MonitoringSiteList = (from t in db.TVItems
                                          from mi in db.MapInfos
                                          from mip in db.MapInfoPoints
                                          let hasSample = (from c in db.MWQMSamples
                                                           where c.MWQMSiteTVItemID == t.TVItemID
                                                           && c.UseForOpenData == true
                                                           select c).Any()
                                          where mi.TVItemID == t.TVItemID
                                          && mip.MapInfoID == mi.MapInfoID
                                          && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                          && t.TVType == (int)TVTypeEnum.MWQMSite
                                          && mi.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Point
                                          && mi.TVType == (int)TVTypeEnum.MWQMSite
                                          && hasSample == true
                                          select new { t, mip, hasSample }).ToList();


                int TotalCount2 = tvItemSSList.Count;
                int Count2 = 0;
                foreach (var tvItemSS in tvItemSSList)
                {
                    if (Count2 % 20 == 0)
                    {
                        _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)(100.0f * ((float)Count2 / (float)TotalCount2)));

                        NotUsed = string.Format(TaskRunnerServiceRes.Creating_, fi.Name + " --- doing " + tvItemSS.tl.TVText + "");
                        TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", fi.Name + " --- doing " + tvItemSS.tl.TVText + "");

                        _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);
                    }

                    Count2 += 1;

                    foreach (var mwqmSite in MonitoringSiteList.Where(c => c.t.ParentID == tvItemSS.t.TVItemID))
                    {
                        string MS = mwqmSite.t.TVItemID.ToString();
                        string Lat = (mwqmSite.mip != null ? mwqmSite.mip.Lat.ToString("F6") : "");
                        string Lng = (mwqmSite.mip != null ? mwqmSite.mip.Lng.ToString("F6") : "");
                        sb.AppendLine($"{MS},{Lat.Replace(",", ".")},{Lng.Replace(",", ".")}");
                    }
                }
            }

            UnicodeEncoding encoding = new UnicodeEncoding();

            FileStream fs = fi.Create();
            byte[] bytes = encoding.GetBytes(sb.ToString());
            fs.Write(bytes, 0, bytes.Length);
            fs.Close();

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.CSVOfMWQMSites, FilePurposeEnum.OpenData);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;
        }
        public void CreateCSVNationalOfMWQMSites()
        {
            string NotUsed = "";
            string ProvInit = "";
            List<string> ProvInitList = new List<string>()
            {
                "BC", "ME", "NB", "NL", "NS", "PE", "QC",
            };
            List<string> ProvList = new List<string>()
            {
                "British Columbia", "Maine", "New Brunswick", "Newfoundland and Labrador", "Nova Scotia", "Prince Edward Island", "Québec",
            };

            int TVItemID = _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID;

            if (TVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.Parameter_NotFound, "TVItemID");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Parameter_NotFound", "TVItemID");
                return;
            }

            TVItemModel tvItemModelCountry = _TVItemService.GetTVItemModelWithTVItemIDDB(TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelCountry.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotDeleteFile_Error_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                return;
            }

            List<TVItemModel> tvItemModelProvList = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelCountry.TVItemID, TVTypeEnum.Province);

            StringBuilder sb = new StringBuilder();

            string ServerFilePath = _TVFileService.GetServerFilePath(TVItemID);

            FileInfo fi = new FileInfo(_TVFileService.ChoseEDriveOrCDrive(ServerFilePath) + $"Sites_National.csv");

            TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            int CountProv = 0;
            foreach (TVItemModel tvItemModelProv in tvItemModelProvList)
            {
                CountProv += 1;
                for (int i = 0, countProv = ProvList.Count; i < countProv; i++)
                {
                    if (ProvList[i] == tvItemModelProv.TVText)
                    {
                        ProvInit = ProvInitList[i];
                        break;
                    }
                }

                sb.AppendLine("Prov,Site_ID,Lat,Long");

                if (fi.Exists)
                {
                    try
                    {
                        fi.Delete();
                    }
                    catch (Exception ex)
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDeleteFile_Error_, fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDeleteFile_Error_", fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                        return;
                    }
                }

                using (CSSPWebToolsDBEntities db = new CSSPWebToolsDBEntities())
                {
                    var tvItemProv = (from c in db.TVItems
                                      from cl in db.TVItemLanguages
                                      where c.TVItemID == cl.TVItemID
                                      && c.TVItemID == tvItemModelProv.TVItemID
                                      && cl.Language == (int)LanguageEnum.en
                                      && c.TVType == (int)TVTypeEnum.Province
                                      select new { c, cl }).FirstOrDefault();

                    if (tvItemProv == null)
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, tvItemModelProv.TVItemID.ToString());
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, tvItemModelProv.TVItemID.ToString());
                        return;
                    }

                    for (int i = 0, countProv = ProvList.Count; i < countProv; i++)
                    {
                        if (ProvList[i] == tvItemProv.cl.TVText)
                        {
                            ProvInit = ProvInitList[i];
                            break;
                        }
                    }

                    NotUsed = string.Format(TaskRunnerServiceRes.Creating_, fi.Name);
                    List<TextLanguage> TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", fi.Name);

                    _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);

                    var tvItemSSList = (from t in db.TVItems
                                        from tl in db.TVItemLanguages
                                        where t.TVItemID == tl.TVItemID
                                        && tl.Language == (int)LanguageEnum.en
                                        && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                        && t.TVType == (int)TVTypeEnum.Subsector
                                        orderby tl.TVText
                                        select new { t, tl }).ToList();

                    var MonitoringSiteList = (from t in db.TVItems
                                              from mi in db.MapInfos
                                              from mip in db.MapInfoPoints
                                              let hasSample = (from c in db.MWQMSamples
                                                               where c.MWQMSiteTVItemID == t.TVItemID
                                                               && c.UseForOpenData == true
                                                               select c).Any()
                                              where mi.TVItemID == t.TVItemID
                                              && mip.MapInfoID == mi.MapInfoID
                                              && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                              && t.TVType == (int)TVTypeEnum.MWQMSite
                                              && mi.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Point
                                              && mi.TVType == (int)TVTypeEnum.MWQMSite
                                              && hasSample == true
                                              select new { t, mip, hasSample }).ToList();


                    int TotalCount2 = tvItemSSList.Count;
                    int Count2 = 0;
                    foreach (var tvItemSS in tvItemSSList)
                    {
                        if (Count2 % 100 == 0)
                        {
                            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)((CountProv - 1) / 6.0f * 100.0f) +  (int)(100.0f / 6.0f * ((float)Count2 / (float)TotalCount2)));

                            NotUsed = string.Format(TaskRunnerServiceRes.Creating_, fi.Name + " --- doing " + tvItemSS.tl.TVText + "");
                            TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", fi.Name + " --- doing " + tvItemSS.tl.TVText + "");

                            _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);
                        }

                        Count2 += 1;

                        foreach (var mwqmSite in MonitoringSiteList.Where(c => c.t.ParentID == tvItemSS.t.TVItemID))
                        {
                            string MS = mwqmSite.t.TVItemID.ToString();
                            string Lat = (mwqmSite.mip != null ? mwqmSite.mip.Lat.ToString("F6") : "");
                            string Lng = (mwqmSite.mip != null ? mwqmSite.mip.Lng.ToString("F6") : "");
                            sb.AppendLine($"{ProvInit},{MS},{Lat.Replace(",", ".")},{Lng.Replace(",", ".")}");
                        }
                    }
                }
            }

            UnicodeEncoding encoding = new UnicodeEncoding();

            FileStream fs = fi.Create();
            byte[] bytes = encoding.GetBytes(sb.ToString());
            fs.Write(bytes, 0, bytes.Length);
            fs.Close();

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.CSVOfMWQMSites, FilePurposeEnum.OpenData);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;
        }
        public void CreateCSVOfMWQMSamples()
        {
            string NotUsed = "";
            string ProvInit = "";
            List<string> ProvInitList = new List<string>()
            {
                "BC", "ME", "NB", "NL", "NS", "PE", "QC",
            };
            List<string> ProvList = new List<string>()
            {
                "British Columbia", "Maine", "New Brunswick", "Newfoundland and Labrador", "Nova Scotia", "Prince Edward Island", "Québec",
            };
            int TVItemID = _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID;

            if (TVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.Parameter_NotFound, "TVItemID");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Parameter_NotFound", "TVItemID");
                return;
            }

            TVItemModel tvItemModel = _TVItemService.GetTVItemModelWithTVItemIDDB(TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotDeleteFile_Error_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                return;
            }

            for (int i = 0, countProv = ProvList.Count; i < countProv; i++)
            {
                if (ProvList[i] == tvItemModel.TVText)
                {
                    ProvInit = ProvInitList[i];
                    break;
                }
            }

            string ServerFilePath = _TVFileService.GetServerFilePath(TVItemID);

            FileInfo fi = new FileInfo(_TVFileService.ChoseEDriveOrCDrive(ServerFilePath) + $"Samples_Echantillons_{ProvInit}.csv");

            TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            // loop through all the MWQMSites etc...

            StringBuilder sb = new StringBuilder();

            sb.AppendLine("Site_ID,Date,FC_MPN_CF_NPP,Temp_°C,Sal_0/00,pH"); //,Depth/Profondeur");

            if (fi.Exists)
            {
                try
                {
                    fi.Delete();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDeleteFile_Error_, fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDeleteFile_Error_", fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    return;
                }
            }

            using (CSSPWebToolsDBEntities db = new CSSPWebToolsDBEntities())
            {
                var tvItemProv = (from c in db.TVItems
                                  from cl in db.TVItemLanguages
                                  where c.TVItemID == cl.TVItemID
                                  && c.TVItemID == TVItemID
                                  && cl.Language == (int)LanguageEnum.en
                                  && c.TVType == (int)TVTypeEnum.Province
                                  select new { c, cl }).FirstOrDefault();

                if (tvItemProv == null)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                    return;
                }

                for (int i = 0, countProv = ProvList.Count; i < countProv; i++)
                {
                    if (ProvList[i] == tvItemProv.cl.TVText)
                    {
                        ProvInit = ProvInitList[i];
                        break;
                    }
                }

                NotUsed = string.Format(TaskRunnerServiceRes.Creating_, fi.Name);
                List<TextLanguage> TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", fi.Name);

                _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);

                var tvItemSSList = (from t in db.TVItems
                                    from tl in db.TVItemLanguages
                                    where t.TVItemID == tl.TVItemID
                                    && tl.Language == (int)LanguageEnum.en
                                    && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                    && t.TVType == (int)TVTypeEnum.Subsector
                                    orderby tl.TVText
                                    select new { t, tl }).ToList();

                var MonitoringSiteList = (from t in db.TVItems
                                          from mi in db.MapInfos
                                          from mip in db.MapInfoPoints
                                          let hasSample = (from c in db.MWQMSamples
                                                           where c.MWQMSiteTVItemID == t.TVItemID
                                                           && c.UseForOpenData == true
                                                           select c).Any()
                                          where mi.TVItemID == t.TVItemID
                                          && mip.MapInfoID == mi.MapInfoID
                                          && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                          && t.TVType == (int)TVTypeEnum.MWQMSite
                                          && mi.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Point
                                          && mi.TVType == (int)TVTypeEnum.MWQMSite
                                          && hasSample == true
                                          select new { t, mip, hasSample }).ToList();

                int TotalCount2 = tvItemSSList.Count;
                int Count2 = 0;
                foreach (var tvItemSS in tvItemSSList)
                {
                    if (Count2 % 2 == 0)
                    {
                        _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)(100.0f * ((float)Count2 / (float)TotalCount2)));

                        NotUsed = string.Format(TaskRunnerServiceRes.Creating_, fi.Name + " --- doing " + tvItemSS.tl.TVText + "");
                        TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", fi.Name + " --- doing " + tvItemSS.tl.TVText + "");

                        _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);
                    }

                    Count2 += 1;

                    foreach (var mwqmSite in MonitoringSiteList.Where(c => c.t.ParentID == tvItemSS.t.TVItemID))
                    {
                        using (CSSPWebToolsDBEntities db2 = new CSSPWebToolsDBEntities())
                        {
                            List<MWQMSample> sampleList = (from c in db2.MWQMSamples
                                                           where c.MWQMSiteTVItemID == mwqmSite.t.TVItemID
                                                           && c.UseForOpenData == true
                                                           orderby c.SampleDateTime_Local ascending
                                                           select c).ToList();

                            foreach (MWQMSample mwqmSample in sampleList)
                            {
                                string MS = mwqmSite.t.TVItemID.ToString();
                                string D = mwqmSample.SampleDateTime_Local.ToString("yyyy-MM-dd");
                                string FC = (mwqmSample.FecCol_MPN_100ml < 2 ? "< 2" : (mwqmSample.FecCol_MPN_100ml > 1600 ? "> 1600" : mwqmSample.FecCol_MPN_100ml.ToString().Replace(",",".")));
                                string Temp = (mwqmSample.WaterTemp_C != null ? ((double)mwqmSample.WaterTemp_C).ToString("F1").Replace(",",".") : "");
                                string Sal = (mwqmSample.Salinity_PPT != null ? ((double)mwqmSample.Salinity_PPT).ToString("F1").Replace(",",".") : "");
                                string pH = (mwqmSample.PH != null ? ((double)mwqmSample.PH).ToString("F1").Replace(",",".") : "");
                                string Depth = (mwqmSample.Depth_m != null ? ((double)mwqmSample.Depth_m).ToString("F1").Replace(",",".") : "");
                                sb.AppendLine($"{MS},{D},{FC},{Temp},{Sal},{pH}"); //,{Depth}");
                            }
                        }
                    }
                }
            }

            UnicodeEncoding encoding = new UnicodeEncoding();

            FileStream fs = fi.Create();
            byte[] bytes = encoding.GetBytes(sb.ToString());
            fs.Write(bytes, 0, bytes.Length);
            fs.Close();

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.MWQMSamplingPlanAutoGenerate, FilePurposeEnum.OpenData);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;
        }
        public void CreateCSVNationalOfMWQMSamples()
        {
            string NotUsed = "";
            string ProvInit = "";
            List<string> ProvInitList = new List<string>()
            {
                "BC", "ME", "NB", "NL", "NS", "PE", "QC",
            };
            List<string> ProvList = new List<string>()
            {
                "British Columbia", "Maine", "New Brunswick", "Newfoundland and Labrador", "Nova Scotia", "Prince Edward Island", "Québec",
            };
            int TVItemID = _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID;

            if (TVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.Parameter_NotFound, "TVItemID");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Parameter_NotFound", "TVItemID");
                return;
            }

            TVItemModel tvItemModelCountry = _TVItemService.GetTVItemModelWithTVItemIDDB(TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelCountry.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotDeleteFile_Error_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                return;
            }

            List<TVItemModel> tvItemModelProvList = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelCountry.TVItemID, TVTypeEnum.Province);

            StringBuilder sb = new StringBuilder();

            string ServerFilePath = _TVFileService.GetServerFilePath(TVItemID);

            FileInfo fi = new FileInfo(_TVFileService.ChoseEDriveOrCDrive(ServerFilePath) + $"Samples_Echantillons_National.csv");

            TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            int CountProv = 0;
            foreach (TVItemModel tvItemModelProv in tvItemModelProvList)
            {
                CountProv += 1;

                for (int i = 0, countProv = ProvList.Count; i < countProv; i++)
                {
                    if (ProvList[i] == tvItemModelProv.TVText)
                    {
                        ProvInit = ProvInitList[i];
                        break;
                    }
                }

                sb.AppendLine("Prov,Site_ID,Date,FC_MPN_CF_NPP,Temp_°C,Sal_0/00,pH"); //,Depth/Profondeur");

                if (fi.Exists)
                {
                    try
                    {
                        fi.Delete();
                    }
                    catch (Exception ex)
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDeleteFile_Error_, fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDeleteFile_Error_", fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                        return;
                    }
                }

                using (CSSPWebToolsDBEntities db = new CSSPWebToolsDBEntities())
                {
                    var tvItemProv = (from c in db.TVItems
                                      from cl in db.TVItemLanguages
                                      where c.TVItemID == cl.TVItemID
                                      && c.TVItemID == tvItemModelProv.TVItemID
                                      && cl.Language == (int)LanguageEnum.en
                                      && c.TVType == (int)TVTypeEnum.Province
                                      select new { c, cl }).FirstOrDefault();

                    if (tvItemProv == null)
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, tvItemModelProv.TVItemID.ToString());
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, tvItemModelProv.TVItemID.ToString());
                        return;
                    }

                    for (int i = 0, countProv = ProvList.Count; i < countProv; i++)
                    {
                        if (ProvList[i] == tvItemProv.cl.TVText)
                        {
                            ProvInit = ProvInitList[i];
                            break;
                        }
                    }

                    NotUsed = string.Format(TaskRunnerServiceRes.Creating_, fi.Name);
                    List<TextLanguage> TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", fi.Name);

                    _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);

                    var tvItemSSList = (from t in db.TVItems
                                        from tl in db.TVItemLanguages
                                        where t.TVItemID == tl.TVItemID
                                        && tl.Language == (int)LanguageEnum.en
                                        && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                        && t.TVType == (int)TVTypeEnum.Subsector
                                        orderby tl.TVText
                                        select new { t, tl }).ToList();

                    var MonitoringSiteList = (from t in db.TVItems
                                              from mi in db.MapInfos
                                              from mip in db.MapInfoPoints
                                              let hasSample = (from c in db.MWQMSamples
                                                               where c.MWQMSiteTVItemID == t.TVItemID
                                                               && c.UseForOpenData == true
                                                               select c).Any()
                                              where mi.TVItemID == t.TVItemID
                                              && mip.MapInfoID == mi.MapInfoID
                                              && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                              && t.TVType == (int)TVTypeEnum.MWQMSite
                                              && mi.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Point
                                              && mi.TVType == (int)TVTypeEnum.MWQMSite
                                              && hasSample == true
                                              select new { t, mip, hasSample }).ToList();

                    int TotalCount2 = tvItemSSList.Count;
                    int Count2 = 0;
                    foreach (var tvItemSS in tvItemSSList)
                    {
                        if (Count2 % 100 == 0)
                        {
                            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)((CountProv - 1) / 6.0f * 100.0f) + (int)(100.0f / 6.0f * ((float)Count2 / (float)TotalCount2)));

                            NotUsed = string.Format(TaskRunnerServiceRes.Creating_, fi.Name + " --- doing " + tvItemSS.tl.TVText + "");
                            TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", fi.Name + " --- doing " + tvItemSS.tl.TVText + "");

                            _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);
                        }

                        Count2 += 1;

                        foreach (var mwqmSite in MonitoringSiteList.Where(c => c.t.ParentID == tvItemSS.t.TVItemID))
                        {
                            using (CSSPWebToolsDBEntities db2 = new CSSPWebToolsDBEntities())
                            {
                                List<MWQMSample> sampleList = (from c in db2.MWQMSamples
                                                               where c.MWQMSiteTVItemID == mwqmSite.t.TVItemID
                                                               && c.UseForOpenData == true
                                                               orderby c.SampleDateTime_Local ascending
                                                               select c).ToList();

                                foreach (MWQMSample mwqmSample in sampleList)
                                {
                                    string MS = mwqmSite.t.TVItemID.ToString();
                                    string D = mwqmSample.SampleDateTime_Local.ToString("yyyy-MM-dd");
                                    string FC = (mwqmSample.FecCol_MPN_100ml < 2 ? "< 2" : (mwqmSample.FecCol_MPN_100ml > 1600 ? "> 1600" : mwqmSample.FecCol_MPN_100ml.ToString().Replace(",", ".")));
                                    string Temp = (mwqmSample.WaterTemp_C != null ? ((double)mwqmSample.WaterTemp_C).ToString("F1").Replace(",", ".") : "");
                                    string Sal = (mwqmSample.Salinity_PPT != null ? ((double)mwqmSample.Salinity_PPT).ToString("F1").Replace(",", ".") : "");
                                    string pH = (mwqmSample.PH != null ? ((double)mwqmSample.PH).ToString("F1").Replace(",", ".") : "");
                                    string Depth = (mwqmSample.Depth_m != null ? ((double)mwqmSample.Depth_m).ToString("F1").Replace(",", ".") : "");
                                    sb.AppendLine($"{ProvInit},{MS},{D},{FC},{Temp},{Sal},{pH}"); //,{Depth}");
                                }
                            }
                        }
                    }
                }
            }

            UnicodeEncoding encoding = new UnicodeEncoding();

            FileStream fs = fi.Create();
            byte[] bytes = encoding.GetBytes(sb.ToString());
            fs.Write(bytes, 0, bytes.Length);
            fs.Close();

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.MWQMSamplingPlanAutoGenerate, FilePurposeEnum.OpenData);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;
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

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.MWQMSamplingPlanAutoGenerate, FilePurposeEnum.TemplateGenerated);
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
