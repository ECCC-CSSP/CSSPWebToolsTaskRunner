using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CSSPWebToolsTaskRunner.Services.Resources;
using CSSPWebToolsTaskRunner.Services;
using System.IO;
using System.Transactions;
using System.Windows.Forms;
using System.Threading;
using System.ComponentModel;
using CSSPWebToolsDBDLL.Models;
using CSSPWebToolsDBDLL.Services;
using System.Net;
using CSSPModelsDLL.Models;
using CSSPEnumsDLL.Enums;
using CSSPWebToolsDBDLL;
using System.Data.OleDb;

namespace CSSPWebToolsTaskRunner.Services
{
    public class HydrometricService
    {
        #region Variables
        public List<string> ProvList = new List<string>() { "BC", "NB", "NL", "NS", "PE", "QC" };
        public List<string> ProvNameListEN = new List<string>() { "British Columbia", "New Brunswick", "Newfoundland and Labrador", "Nova Scotia", "Prince Edward Island", "Québec" };
        public List<string> ProvNameListFR = new List<string>() { "Colombie-Britannique", "Nouveau-Brunswick", "Terre-Neuve-et-Labrador", "Nouvelle-Écosse", "Île-du-Prince-Édouard", "Québec" };
        public string URLUpdateHydrometricSiteInfo = "https://wateroffice.ec.gc.ca/station_metadata/station_result_e.html?search_type=province&province={0}";
        public string URLUpdateHydrometricSiteInfoLatLngWMOElevTC = "http://Hydrometric.weather.gc.ca/HydrometricData/dailydata_e.html?timeframe=2&StationID={0}&Year=1800&Month=1&Day=01";
        public string UpdateHydrometricSiteDailyFromStartDateToEndDate = "http://Hydrometric.weather.gc.ca/HydrometricData/bulk_data_e.html?format=csv&stationID={0}&Year={1}&Month=1&Day=1&timeframe=2&submit=Download+Data";
        //public string GetHydrometricSiteDataForRun = "http://Hydrometric.weather.gc.ca/HydrometricData/bulk_data_e.html?format=csv&stationID={0}&Year={1}&Month=1&Day=1&timeframe=2&submit=Download+Data";
        public string UrlToGetHydrometricSiteDataForRunsOfYear = "http://Hydrometric.weather.gc.ca/Hydrometric_data/bulk_data_e.html?format=csv&stationID={0}&Year={1}&Month=1&Day=1&timeframe=2&submit=Download+Data";
        public string UpdateHydrometricSiteHourlyFromStartDateToEndDate = "http://Hydrometric.weather.gc.ca/Hydrometric_data/bulk_data_e.html?format=csv&stationID={0}&Year={1}&Month={2}&Day=1&timeframe=1&submit=Download+Data";
        public bool WebBrowserContentAnalysed = true;
        public bool AllDone = false;
        #endregion Variables

        #region Properties
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        #endregion Properties

        #region Constructors
        public HydrometricService(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
        }
        #endregion Constructors

        #region Events
        //public void browser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        //{
        //    string NotUsed = "";

        //    if (e.Url.AbsolutePath != (sender as WebBrowser).Url.AbsolutePath)
        //        return;

        //    if ((sender as WebBrowser).Url == e.Url)
        //    {
        //        if ((sender as WebBrowser).Url.ToString().Contains("HydrometricData/dailydata_e.html"))
        //        {
        //            UpdateHydrometricSiteInfoLatLngWMOElevTCParse(sender, e);
        //        }
        //        else if ((sender as WebBrowser).Url.ToString().Contains("HydrometricData/bulkdata_e.html"))
        //        {
        //            if ((sender as WebBrowser).Url.ToString().Contains("timeframe=2"))
        //            {
        //                UpdateHydrometricSiteInfoLatLngWMOElevTCParse(sender, e);
        //            }
        //            else if ((sender as WebBrowser).Url.ToString().Contains("timeframe=1"))
        //            {
        //                UpdateHydrometricSiteInfoLatLngWMOElevTCParse(sender, e);
        //            }
        //            else
        //            {
        //            }
        //        }
        //        else
        //        {

        //        }
        //    }
        //    else
        //    {
        //        NotUsed = TaskRunnerServiceRes.CurrentURLNotTheSameInDocumentCompletedEvent;
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("CurrentURLNotTheSameInDocumentCompletedEvent");
        //        return;
        //    }

        //    Application.ExitThread();   // Stops the thread
        //}
        #endregion Events

        #region Functions public
        public bool FindHtmlElement(HtmlElement htmlElement, int count, string tagName, string message)
        {
            string NotUsed = "";

            if (htmlElement.Children.Count < count)
            {
                AllDone = true;
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, message);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", message);
                return false;
            }

            HtmlElement htmlElementChild = htmlElement.Children[count - 1];
            if (htmlElementChild.TagName.ToUpper() != tagName.ToUpper())
            {
                AllDone = true;
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, message);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", message);
                return false;
            }

            return true;
        }
        //public void RunBrowserThread(string URL)
        //{
        //    var thread = new Thread(() =>
        //    {
        //        var br = new WebBrowser();
        //        br.DocumentCompleted += browser_DocumentCompleted;
        //        br.Navigate(URL);
        //        Application.Run();
        //    });
        //    thread.SetApartmentState(ApartmentState.STA);
        //    thread.Start();

        //    return;
        //}
        //public void GetAllDischargesForYear()
        //{
        //    string NotUsed = "";

        //    TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    MWQMSubsectorService mwqmSubsectorService = new MWQMSubsectorService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

        //    AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //    if (_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID == 0)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes._Required, TaskRunnerServiceRes.TVItemID);
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_Required", TaskRunnerServiceRes.TVItemID);
        //        return;
        //    }
        //    if (_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID2 == 0)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes._Required, TaskRunnerServiceRes.TVItemID2);
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_Required", TaskRunnerServiceRes.TVItemID2);
        //        return;
        //    }

        //    TVItemModel tvItemModelProvince = tvItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //    if (!string.IsNullOrWhiteSpace(tvItemModelProvince.Error))
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        return;
        //    }

        //    if (tvItemModelProvince.TVType != TVTypeEnum.Province)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.TVTypeShouldBe_, TVTypeEnum.Province.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("TVTypeShouldBe_", TVTypeEnum.Province.ToString());
        //        return;
        //    }

        //    string Parameters = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters;
        //    string[] ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
        //    int ProvinceTVItemID = 0;
        //    int Year = 0;
        //    foreach (string s in ParamValueList)
        //    {
        //        string[] ParamValue = s.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

        //        if (ParamValue.Length != 2)
        //        {
        //            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, TaskRunnerServiceRes.Parameters);
        //            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", TaskRunnerServiceRes.Parameters);
        //            return;
        //        }

        //        if (ParamValue[0] == "ProvinceTVItemID")
        //        {
        //            ProvinceTVItemID = int.Parse(ParamValue[1]);
        //        }
        //        else if (ParamValue[0] == "Year")
        //        {
        //            Year = int.Parse(ParamValue[1]);
        //        }
        //        else
        //        {
        //            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, ParamValue[0]);
        //            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", ParamValue[0].ToString());
        //            return;
        //        }
        //    }

        //    if (tvItemModelProvince.TVItemID != ProvinceTVItemID)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes._NotEqualTo_, "tvItemModelProvince.TVItemID[" + tvItemModelProvince.TVItemID.ToString() + "]", "ProvinceTVItemID[" + ProvinceTVItemID.ToString() + "]");
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("_NotEqualTo_", "tvItemModelProvince.TVItemID[" + tvItemModelProvince.TVItemID.ToString() + "]", "ProvinceTVItemID[" + ProvinceTVItemID.ToString() + "]");
        //        return;
        //    }

        //    List<TVItemModel> tvItemModelSubsectorList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(ProvinceTVItemID, TVTypeEnum.Subsector).OrderBy(c => c.TVText).ToList();
        //    if (tvItemModelSubsectorList.Count == 0)
        //    {
        //        return;
        //    }

        //    int CountSS = 0;
        //    string Status = appTaskModel.StatusText;
        //    foreach (TVItemModel tvItemModel in tvItemModelSubsectorList)
        //    {
        //        CountSS += 1;
        //        if (CountSS % 1 == 0)
        //        {
        //            appTaskModel.PercentCompleted = (100 * CountSS) / tvItemModelSubsectorList.Count;
        //            appTaskModel.StatusText = Status + " --- " + tvItemModel.TVText;
        //            appTaskService.PostUpdateAppTask(appTaskModel);
        //        }

        //        List<MWQMRun> mwqmRunList = new List<MWQMRun>();
        //        using (CSSPWebToolsDBEntities db = new CSSPWebToolsDBEntities())
        //        {
        //            mwqmRunList = (from c in db.MWQMRuns
        //                           where c.SubsectorTVItemID == tvItemModel.TVItemID
        //                           select c).ToList();
        //        }

        //        if (mwqmRunList.Count == 0)
        //        {
        //            continue;
        //        }

        //        GetHydrometricSitesDataForSubsectorRunsOfYear(tvItemModel.TVItemID, Year);
        //        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        {
        //            return;
        //        }
        //    }

        //    appTaskModel.PercentCompleted = 100;
        //    appTaskService.PostUpdateAppTask(appTaskModel);
        //}
        //public void GetHydrometricSitesDataForRunsOfYear()
        //{
        //    string NotUsed = "";

        //    TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

        //    AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //    if (_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID == 0)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes._Required, TaskRunnerServiceRes.TVItemID);
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_Required", TaskRunnerServiceRes.TVItemID);
        //        return;
        //    }

        //    if (_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID2 == 0)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes._Required, TaskRunnerServiceRes.TVItemID2);
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_Required", TaskRunnerServiceRes.TVItemID2);
        //        return;
        //    }

        //    TVItemModel tvItemModelSubsector = tvItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //    if (!string.IsNullOrWhiteSpace(tvItemModelSubsector.Error))
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        return;
        //    }

        //    if (tvItemModelSubsector.TVType != TVTypeEnum.Subsector)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.TVTypeShouldBe_, TVTypeEnum.Subsector.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("TVTypeShouldBe_", TVTypeEnum.Subsector.ToString());
        //        return;
        //    }

        //    string Parameters = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters;
        //    string[] ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
        //    int SubsectorTVItemID = 0;
        //    int Year = 0;
        //    foreach (string s in ParamValueList)
        //    {
        //        string[] ParamValue = s.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

        //        if (ParamValue.Length != 2)
        //        {
        //            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, TaskRunnerServiceRes.Parameters);
        //            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", TaskRunnerServiceRes.Parameters);
        //            return;
        //        }

        //        if (ParamValue[0] == "SubsectorTVItemID")
        //        {
        //            SubsectorTVItemID = int.Parse(ParamValue[1]);
        //        }
        //        else if (ParamValue[0] == "Year")
        //        {
        //            Year = int.Parse(ParamValue[1]);
        //        }
        //        else
        //        {
        //            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, ParamValue[0]);
        //            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", ParamValue[0].ToString());
        //            return;
        //        }
        //    }

        //    if (tvItemModelSubsector.TVItemID != SubsectorTVItemID)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes._NotEqualTo_, "tvItemModelSubsector.TVItemID[" + tvItemModelSubsector.TVItemID.ToString() + "]", "SubsectorTVItemID[" + SubsectorTVItemID.ToString() + "]");
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("_NotEqualTo_", "tvItemModelSubsector.TVItemID[" + tvItemModelSubsector.TVItemID.ToString() + "]", "SubsectorTVItemID[" + SubsectorTVItemID.ToString() + "]");
        //        return;
        //    }

        //    GetHydrometricSitesDataForSubsectorRunsOfYear(SubsectorTVItemID, Year);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //    {
        //        return;
        //    }

        //    // should get all the runs for the particular year and subsector

        //    appTaskModel.PercentCompleted = 100;
        //    appTaskService.PostUpdateAppTask(appTaskModel);
        //}
        //public void GetHydrometricSitesDataForSubsectorRunsOfYear(int SubsectorTVItemID, int Year)
        //{
        //    string NotUsed = "";
        //    int CurrentYear = DateTime.Now.Year;

        //    HydrometricDataValueService hydrometricDataValueService = new HydrometricDataValueService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    HydrometricSiteService hydrometricSiteService = new HydrometricSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    MWQMRunService mwqmRunService = new MWQMRunService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    UseOfSiteService useOfSiteService = new UseOfSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

        //    AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //    List<MWQMRunModel> mwqmRunModelList = mwqmRunService.GetMWQMRunModelListWithSubsectorTVItemIDDB(SubsectorTVItemID).Where(c => c.DateTime_Local.Year == Year).OrderBy(c => c.DateTime_Local).ToList();

        //    // need to get all Hydrometric sites for this particular subsector and run
        //    List<UseOfSiteModel> useOfSiteModelList = useOfSiteService.GetUseOfSiteModelListWithSiteTypeAndSubsectorTVItemIDDB(SiteTypeEnum.Hydrometric, SubsectorTVItemID);
        //    List<int> HydrometricSiteTVItemID = new List<int>();

        //    //appTaskModel.PercentCompleted = 5;
        //    //appTaskService.PostUpdateAppTask(appTaskModel);

        //    int Count = 0;
        //    int TotalCount = mwqmRunModelList.Count() * useOfSiteModelList.Count();
        //    foreach (UseOfSiteModel useOfSiteModel in useOfSiteModelList)
        //    {

        //        int EndYear = (useOfSiteModel.EndYear == null ? CurrentYear : (int)useOfSiteModel.EndYear);
        //        if (Year >= useOfSiteModel.StartYear && Year <= EndYear)
        //        {
        //            HydrometricSiteModel HydrometricSiteModel = hydrometricSiteService.GetHydrometricSiteModelWithHydrometricSiteTVItemIDDB(useOfSiteModel.SiteTVItemID);
        //            if (!string.IsNullOrWhiteSpace(HydrometricSiteModel.Error))
        //            {
        //                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.HydrometricSite, TaskRunnerServiceRes.HydrometricSiteTVItemID, useOfSiteModel.SiteTVItemID.ToString());
        //                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.HydrometricSite, TaskRunnerServiceRes.HydrometricSiteTVItemID, useOfSiteModel.SiteTVItemID.ToString());
        //                return;
        //            }

        //            string httpStrDaily = "";

        //            using (WebClient webClient = new WebClient())
        //            {
        //                WebProxy webProxy = new WebProxy();
        //                webClient.Proxy = webProxy;
        //                string url = string.Format(UrlToGetHydrometricSiteDataForRunsOfYear, HydrometricSiteModel.ECDBID, Year);
        //                httpStrDaily = webClient.DownloadString(new Uri(url));
        //                if (httpStrDaily.Length > 0)
        //                {
        //                    if (httpStrDaily.Substring(0, "\"Station Name".Length) == "\"Station Name")
        //                    {
        //                        httpStrDaily = httpStrDaily.Replace("\"", "").Replace("\n", "\r\n");
        //                    }
        //                }
        //                else
        //                {
        //                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotReadFile_, url);
        //                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotReadFile_", url);
        //                    return;
        //                }
        //            }

        //            foreach (MWQMRunModel mwqmRunModel in mwqmRunModelList)
        //            {
        //                DateTime RunDate = new DateTime(mwqmRunModel.DateTime_Local.Year, mwqmRunModel.DateTime_Local.Month, mwqmRunModel.DateTime_Local.Day);
        //                DateTime RunDateMinus10 = RunDate.AddDays(-10);

        //                DateTime HydrometricStartDate = new DateTime(HydrometricSiteModel.StartDate_Local.Value.Year, HydrometricSiteModel.StartDate_Local.Value.Month, HydrometricSiteModel.StartDate_Local.Value.Day);
        //                DateTime HydrometricEndDate = new DateTime(HydrometricSiteModel.EndDate_Local.Value.Year, HydrometricSiteModel.EndDate_Local.Value.Month, HydrometricSiteModel.EndDate_Local.Value.Day);

        //                Count += 1;

        //                //appTaskModel.PercentCompleted = 100 * Count / TotalCount;
        //                //appTaskService.PostUpdateAppTask(appTaskModel);

        //                if (HydrometricStartDate <= RunDate && HydrometricEndDate >= RunDate)
        //                {
        //                    UpdateDailyValuesForHydrometricSiteTVItemID(HydrometricSiteModel, httpStrDaily, RunDateMinus10, RunDate, new List<DateTime>() { RunDate });
        //                    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    {
        //                        return;
        //                    }
        //                }

        //                if (RunDate.Year != RunDateMinus10.Year)
        //                {
        //                    string httpStrDaily2 = "";

        //                    using (WebClient webClient2 = new WebClient())
        //                    {
        //                        WebProxy webProxy2 = new WebProxy();
        //                        webClient2.Proxy = webProxy2;
        //                        string url2 = string.Format(UrlToGetHydrometricSiteDataForRunsOfYear, HydrometricSiteModel.ECDBID, RunDateMinus10.Year);
        //                        httpStrDaily2 = webClient2.DownloadString(new Uri(url2));
        //                        if (httpStrDaily2.Length > 0)
        //                        {
        //                            if (httpStrDaily2.Substring(0, "\"Station Name".Length) == "\"Station Name")
        //                            {
        //                                httpStrDaily2 = httpStrDaily2.Replace("\"", "").Replace("\n", "\r\n");
        //                            }
        //                        }
        //                        else
        //                        {
        //                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotReadFile_, url2);
        //                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotReadFile_", url2);
        //                            return;
        //                        }
        //                    }

        //                    UpdateDailyValuesForHydrometricSiteTVItemID(HydrometricSiteModel, httpStrDaily2, RunDateMinus10, RunDate, new List<DateTime>() { RunDateMinus10 });
        //                    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    {
        //                        return;
        //                    }

        //                }
        //            }
        //        }
        //    }
        //}
        public void UpdateHydrometricSitesInformationForProvinceTVItemID()
        {
            string NotUsed = "";

            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            HydrometricSiteService hydrometricSiteService = new HydrometricSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            if (_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._Required, TaskRunnerServiceRes.TVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_Required", TaskRunnerServiceRes.TVItemID);
                return;
            }

            TVItemModel tvItemModelProvince = tvItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelProvince.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                return;
            }

            if (tvItemModelProvince.TVType != TVTypeEnum.Province)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.TVTypeShouldBe_, TVTypeEnum.Province.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("TVTypeShouldBe_", TVTypeEnum.Province.ToString());
                return;
            }

            string Prov = "";
            for (int i = 0; i < ProvList.Count; i++)
            {
                Application.DoEvents();

                if (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr)
                {
                    if (ProvNameListFR[i] == tvItemModelProvince.TVText)
                    {
                        Prov = ProvList[i];
                        break;
                    }
                }
                else
                {
                    if (ProvNameListEN[i] == tvItemModelProvince.TVText)
                    {
                        Prov = ProvList[i];
                        break;
                    }
                }
            }

            if (string.IsNullOrWhiteSpace(Prov))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.Province_NotFound, tvItemModelProvince.TVText);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Province_NotFound", tvItemModelProvince.TVText);
                return;
            }

            string ProvincePath = tvFileService.GetServerFilePath(tvItemModelProvince.TVItemID);

            FileInfo fi = new FileInfo(ProvincePath + $"HydrometricSiteInfo_{ Prov }.csv");

            FileInfo fiXlsx = new FileInfo(fi.FullName.Replace(".csv", ".xlsx"));

            if (fiXlsx.Exists)
            {
                try
                {
                    fiXlsx.Delete();
                }
                catch (Exception ex)
                {
                    string InnerException = ex.InnerException != null ? $" InnerException: { ex.InnerException.Message }" : "";
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDeleteFile_Error_, fiXlsx.FullName, $"{ ex.Message } { InnerException }");
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDeleteFile_Error_", fiXlsx.FullName, $"{ ex.Message } { InnerException }");
                    return;
                }
            }

            if (fi.Exists)
            {
                try
                {
                    fi.Delete();
                }
                catch (Exception ex)
                {
                    string InnerException = ex.InnerException != null ? $" InnerException: { ex.InnerException.Message }" : "";
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDeleteFile_Error_, fi.FullName, $"{ ex.Message } { InnerException }");
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDeleteFile_Error_", fi.FullName, $"{ ex.Message } { InnerException }");
                    return;
                }
            }

            using (WebClient webClient = new WebClient())
            {
                WebProxy webProxy = new WebProxy();
                webClient.Proxy = webProxy;
                string url = string.Format(URLUpdateHydrometricSiteInfo, Prov);
                webClient.DownloadFile(new Uri(url), fi.FullName);

                fi = new FileInfo(ProvincePath + $"HydrometricSiteInfo_{ Prov }.csv");

                if (!fi.Exists)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fi.FullName);
                    return;
                }
            }

            Microsoft.Office.Interop.Excel.Application appExcel = new Microsoft.Office.Interop.Excel.Application();

            if (appExcel == null)
            {
                NotUsed = TaskRunnerServiceRes.CouldNotStartExcel;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("CouldNotStartExcel");
                return;
            }
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = appExcel.Workbooks.Open(fi.FullName);
            appExcel.ActiveWorkbook.SaveAs(fiXlsx.FullName, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
            xlWorkbook.Close();
            appExcel.Quit();

            fi = new FileInfo(fiXlsx.FullName);

            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fiXlsx.FullName + ";Extended Properties=Excel 12.0";

            OleDbConnection conn = new OleDbConnection(connectionString);

            conn.Open();
            OleDbDataReader reader;
            OleDbCommand comm = new OleDbCommand("Select * from [HydrometricSiteInfo_" + Prov + "$];");

            comm.Connection = conn;
            reader = comm.ExecuteReader();

            int countTotal = 0;
            while (reader.Read())
            {
                countTotal += 1;
            }

            conn.Close();

            conn.Open();
            comm.Connection = conn;
            reader = comm.ExecuteReader();

            List<string> ShouldBeList = new List<string>()
            {
                "Station Number", "Station Name", "Province", "Status", "Latitude", "Longitude", "Year From", "Year To",
                "Gross Drainage Area (km2)", "Effective Drainage Area (km2)", "Regulation", "Data Type", "Operation Schedule",
                "Sediment", "RHBN", "Real-Time", "Datum Name", "Publishing Office", "Operating Agency", "Contributed"
            };

            reader.Read();

            for (int i = 0, count = ShouldBeList.Count; i < count; i++)
            {
                string varName = reader.GetName(i);
                if (varName != ShouldBeList[i])
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.Variable_ShouldBe_ItIs_, (i + 1).ToString(), ShouldBeList[i], varName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("Variable_ShouldBe_ItIs_", (i + 1).ToString(), ShouldBeList[i], varName);
                    return;
                }
            }

            int countRead = 0;
            while (reader.Read())
            {
                countRead += 1;

                if (countRead % 20 == 0)
                {
                    if (_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID != 0) // 0 ==> testing
                    {
                        _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)(100.0f * countRead / countTotal));
                    }
                }

                string StationNumber = "";
                string StationName = "";
                string Province = "";
                string Status = "";
                string Latitude = "";
                string Longitude = "";
                string YearFrom = "";
                string YearTo = "";
                string GrossDrainageArea_km2 = "";
                string EffectiveDrainageArea_km2 = "";
                string Regulation = "";
                string DataType = "";
                string OperationSchedule = "";
                string Sediment = "";
                string RHBN = "";
                string RealTime = "";
                string DatumName = "";
                string PublishingOffice = "";
                string OperatingAgency = "";
                string Contributed = "";

                // StationNumber
                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                {
                    StationNumber = "";
                }
                else
                {
                    StationNumber = reader.GetValue(0).ToString().Trim();
                    StationNumber = StationNumber.Replace(@"\", "").Replace(@"""", "");
                }

                if (StationNumber == "01AD015")
                {
                    int sliefj = 34;
                }
                // StationName
                if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                {
                    StationName = "";
                }
                else
                {
                    StationName = reader.GetValue(1).ToString().Trim();
                    StationName = StationName.Replace(@"\", "").Replace(@"""", "");
                }

                // Province
                if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                {
                    Province = "";
                }
                else
                {
                    Province = reader.GetValue(2).ToString().Trim();
                    Province = Province.Replace(@"\", "").Replace(@"""", "");
                }

                // Status
                if (reader.GetValue(3).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(3).ToString()))
                {
                    Status = "";
                }
                else
                {
                    Status = reader.GetValue(3).ToString().Trim();
                    Status = Status.Replace(@"\", "").Replace(@"""", "");
                }

                // Latitude
                if (reader.GetValue(4).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(4).ToString()))
                {
                    Latitude = "";
                }
                else
                {
                    Latitude = reader.GetValue(4).ToString().Trim();
                    Latitude = Latitude.Replace(@"\", "").Replace(@"""", "");
                }

                // Longitude
                if (reader.GetValue(5).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(5).ToString()))
                {
                    Longitude = "";
                }
                else
                {
                    Longitude = reader.GetValue(5).ToString().Trim();
                    Longitude = Longitude.Replace(@"\", "").Replace(@"""", "");
                }

                // YearFrom
                if (reader.GetValue(6).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(6).ToString()))
                {
                    YearFrom = "";
                }
                else
                {
                    YearFrom = reader.GetValue(6).ToString().Trim();
                    YearFrom = YearFrom.Replace(@"\", "").Replace(@"""", "");
                }

                // YearTo
                if (reader.GetValue(7).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(7).ToString()))
                {
                    YearTo = "";
                }
                else
                {
                    YearTo = reader.GetValue(7).ToString().Trim();
                    YearTo = YearTo.Replace(@"\", "").Replace(@"""", "");
                }

                // GrossDrainageArea_km2
                if (reader.GetValue(8).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(8).ToString()))
                {
                    GrossDrainageArea_km2 = "";
                }
                else
                {
                    GrossDrainageArea_km2 = reader.GetValue(8).ToString().Trim();
                    GrossDrainageArea_km2 = GrossDrainageArea_km2.Replace(@"\", "").Replace(@"""", "");
                }

                // EffectiveDrainageArea_km2
                if (reader.GetValue(9).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(9).ToString()))
                {
                    EffectiveDrainageArea_km2 = "";
                }
                else
                {
                    EffectiveDrainageArea_km2 = reader.GetValue(9).ToString().Trim();
                    EffectiveDrainageArea_km2 = EffectiveDrainageArea_km2.Replace(@"\", "").Replace(@"""", "");
                }

                // Regulation
                if (reader.GetValue(10).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(10).ToString()))
                {
                    Regulation = "";
                }
                else
                {
                    Regulation = reader.GetValue(10).ToString().Trim();
                    Regulation = Regulation.Replace(@"\", "").Replace(@"""", "");
                }

                // DataType
                if (reader.GetValue(11).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(11).ToString()))
                {
                    DataType = "";
                }
                else
                {
                    DataType = reader.GetValue(11).ToString().Trim();
                    DataType = DataType.Replace(@"\", "").Replace(@"""", "");
                }

                // OperationSchedule
                if (reader.GetValue(12).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(12).ToString()))
                {
                    OperationSchedule = "";
                }
                else
                {
                    OperationSchedule = reader.GetValue(12).ToString().Trim();
                    OperationSchedule = OperationSchedule.Replace(@"\", "").Replace(@"""", "");
                }

                // Sediment
                if (reader.GetValue(13).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(13).ToString()))
                {
                    Sediment = "";
                }
                else
                {
                    Sediment = reader.GetValue(13).ToString().Trim();
                    Sediment = Sediment.Replace(@"\", "").Replace(@"""", "");
                }

                // RHBN
                if (reader.GetValue(14).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(14).ToString()))
                {
                    RHBN = "";
                }
                else
                {
                    RHBN = reader.GetValue(14).ToString().Trim();
                    RHBN = RHBN.Replace(@"\", "").Replace(@"""", "");
                }

                // RealTime
                if (reader.GetValue(15).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(15).ToString()))
                {
                    RealTime = "";
                }
                else
                {
                    RealTime = reader.GetValue(15).ToString().Trim();
                    RealTime = RealTime.Replace(@"\", "").Replace(@"""", "");
                }

                // DatumName
                if (reader.GetValue(16).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(16).ToString()))
                {
                    DatumName = "";
                }
                else
                {
                    DatumName = reader.GetValue(16).ToString().Trim();
                    DatumName = DatumName.Replace(@"\", "").Replace(@"""", "");
                }

                // PublishingOffice
                if (reader.GetValue(17).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(17).ToString()))
                {
                    PublishingOffice = "";
                }
                else
                {
                    PublishingOffice = reader.GetValue(17).ToString().Trim();
                    PublishingOffice = PublishingOffice.Replace(@"\", "").Replace(@"""", "");
                }

                // OperatingAgency
                if (reader.GetValue(18).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(18).ToString()))
                {
                    OperatingAgency = "";
                }
                else
                {
                    OperatingAgency = reader.GetValue(18).ToString().Trim();
                    OperatingAgency = OperatingAgency.Replace(@"\", "").Replace(@"""", "");
                }

                // Contributed
                if (reader.GetValue(19).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(19).ToString()))
                {
                    Contributed = "";
                }
                else
                {
                    Contributed = reader.GetValue(19).ToString().Trim();
                    Contributed = Contributed.Replace(@"\", "").Replace(@"""", "");
                }

                if (string.IsNullOrWhiteSpace(DataType))
                {
                    // the hydrometric site is not an actual hydrometric site but a WQ or Buoy site.
                    continue;
                }

                if (!int.TryParse(YearFrom, out int StartYear))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_As_, "YearFrom", "int");
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotParse_As_", "YearFrom", "int");
                    return;
                }

                if (!int.TryParse(YearTo, out int EndYear))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_As_, "YearTo", "int");
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotParse_As_", "YearTo", "int");
                    return;
                }

                bool NoDrainageArea = false;
                float DrainageArea = 0.0f;
                if (!string.IsNullOrWhiteSpace(GrossDrainageArea_km2.Trim()))
                {
                    if (!float.TryParse(GrossDrainageArea_km2, out DrainageArea))
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_As_, "GrossDrainageArea_km2", "float");
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotParse_As_", "GrossDrainageArea_km2", "float");
                        return;
                    }
                }
                else
                {
                    NoDrainageArea = true;
                }

                bool? IsNatural = null;
                if (Regulation.Trim() == "N")
                {
                    IsNatural = true;
                }
                if (Regulation.Trim() == "R")
                {
                    IsNatural = false;
                }

                bool? HasDischarge = null;
                bool? HasLevel = null;

                if (string.IsNullOrWhiteSpace(DataType))
                {
                }

                if (DataType.Contains("Flow"))
                {
                    HasDischarge = true;
                }

                if (DataType.Contains("Level"))
                {
                    HasLevel = true;
                }

                if (HasDischarge != true)
                {
                    continue;
                }

                if (EndYear == DateTime.Now.Year)
                {
                    EndYear = EndYear + 10;
                }

                HydrometricSiteModel hydrometricSiteModelRet = new HydrometricSiteModel();

                HydrometricSiteModel hydrometricSiteModel = new HydrometricSiteModel()
                {
                    FedSiteNumber = StationNumber,
                    HydrometricSiteName = StationName,
                    Province = Province,
                    IsActive = Status == "Active" ? true : false,
                    StartDate_Local = new DateTime(StartYear, 1, 1),
                    EndDate_Local = new DateTime(EndYear, 12, 31),
                    DrainageArea_km2 = NoDrainageArea == true ? 0.0f : DrainageArea,
                    IsNatural = IsNatural,
                    Sediment = Sediment == "Y" ? true : false,
                    RHBN = RHBN == "Y" ? true : false,
                    RealTime = RealTime == "Y" ? true : false,
                    HasDischarge = HasDischarge,
                    HasLevel = HasLevel,
                };

                HydrometricSiteModel hydrometricSiteModelExist = hydrometricSiteService.GetHydrometricSiteModelExistDB(hydrometricSiteModel);
                if (string.IsNullOrWhiteSpace(hydrometricSiteModelExist.Error))
                {
                    if (hydrometricSiteModelExist.IsActive != hydrometricSiteModel.IsActive
                        || hydrometricSiteModelExist.StartDate_Local != hydrometricSiteModel.StartDate_Local
                        || hydrometricSiteModelExist.EndDate_Local != hydrometricSiteModel.EndDate_Local
                        || hydrometricSiteModelExist.DrainageArea_km2 != hydrometricSiteModel.DrainageArea_km2
                        || hydrometricSiteModelExist.IsNatural != hydrometricSiteModel.IsNatural
                        || hydrometricSiteModelExist.RHBN != hydrometricSiteModel.RHBN
                        || hydrometricSiteModelExist.RealTime != hydrometricSiteModel.RealTime
                        || hydrometricSiteModelExist.HasDischarge != hydrometricSiteModel.HasDischarge
                        || hydrometricSiteModelExist.HasLevel != hydrometricSiteModel.HasLevel)
                    {
                        hydrometricSiteModelExist.IsActive = hydrometricSiteModel.IsActive;
                        hydrometricSiteModelExist.StartDate_Local = hydrometricSiteModel.StartDate_Local;
                        hydrometricSiteModelExist.EndDate_Local = hydrometricSiteModel.EndDate_Local;
                        if (hydrometricSiteModel.DrainageArea_km2 == 0.0f)
                        {
                            hydrometricSiteModelExist.DrainageArea_km2 = null;
                        }
                        else
                        {
                            hydrometricSiteModelExist.DrainageArea_km2 = hydrometricSiteModel.DrainageArea_km2;
                        }
                        hydrometricSiteModelExist.IsNatural = hydrometricSiteModel.IsNatural;
                        hydrometricSiteModelExist.RHBN = hydrometricSiteModel.RHBN;
                        hydrometricSiteModelExist.RealTime = hydrometricSiteModel.RealTime;
                        hydrometricSiteModelExist.HasDischarge = hydrometricSiteModel.HasDischarge;
                        hydrometricSiteModelExist.HasLevel = hydrometricSiteModel.HasLevel;

                        hydrometricSiteModelRet = hydrometricSiteService.PostUpdateHydrometricSiteDB(hydrometricSiteModelExist);
                        if (!string.IsNullOrWhiteSpace(hydrometricSiteModelRet.Error))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.HydrometricSite, hydrometricSiteModelRet.Error);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.HydrometricSite, hydrometricSiteModelRet.Error);
                            return;
                        }
                    }
                    else
                    {
                        hydrometricSiteModelRet = hydrometricSiteModelExist;
                    }
                }
                else
                {
                    string TVText = hydrometricSiteModel.HydrometricSiteName;
                    TVText = TVText + (string.IsNullOrWhiteSpace(hydrometricSiteModel.FedSiteNumber) == true ? "" : " F(" + hydrometricSiteModel.FedSiteNumber + ")");

                    TVItemModel tvItemModelHydrometric = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelProvince.TVItemID, TVText, TVTypeEnum.HydrometricSite);
                    if (!string.IsNullOrWhiteSpace(tvItemModelHydrometric.Error))
                    {
                        tvItemModelHydrometric = tvItemService.PostAddChildTVItemDB(tvItemModelProvince.TVItemID, TVText, TVTypeEnum.HydrometricSite);
                        if (!string.IsNullOrWhiteSpace(tvItemModelHydrometric.Error))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVItem, tvItemModelHydrometric.Error);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVItem, tvItemModelHydrometric.Error);
                            return;
                        }
                    }

                    hydrometricSiteModel.HydrometricSiteTVItemID = tvItemModelHydrometric.TVItemID;
                    if (hydrometricSiteModel.DrainageArea_km2 == 0.0f)
                    {
                        hydrometricSiteModelExist.DrainageArea_km2 = null;
                    }
                    else
                    {
                        hydrometricSiteModelExist.DrainageArea_km2 = hydrometricSiteModel.DrainageArea_km2;
                    }

                    hydrometricSiteModelRet = hydrometricSiteService.PostAddHydrometricSiteDB(hydrometricSiteModel);
                    if (!string.IsNullOrWhiteSpace(hydrometricSiteModelRet.Error))
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.HydrometricSite, hydrometricSiteModelRet.Error);
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.HydrometricSite, hydrometricSiteModelRet.Error);
                        return;
                    }
                }

                if (!float.TryParse(Latitude, out float Lat))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_As_, "Latitude", "float");
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotParse_As_", "Latitude", "float");
                    return;
                }

                if (!float.TryParse(Longitude, out float Lng))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_As_, "Longitude", "float");
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotParse_As_", "Longitude", "float");
                    return;
                }

                List<MapInfoPointModel> mapInfoPointModelList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(hydrometricSiteModelRet.HydrometricSiteTVItemID, TVTypeEnum.HydrometricSite, MapInfoDrawTypeEnum.Point);

                if (mapInfoPointModelList.Count > 0)
                {
                    if (mapInfoPointModelList[0].Lat != Lat || mapInfoPointModelList[0].Lng != Lng || mapInfoPointModelList[0].Ordinal != 0)
                    {
                        mapInfoPointModelList[0].Lat = Lat;
                        mapInfoPointModelList[0].Lng = Lng;
                        mapInfoPointModelList[0].Ordinal = 0;

                        MapInfoPointModel mapInfoPointModel = mapInfoService._MapInfoPointService.PostUpdateMapInfoPointDB(mapInfoPointModelList[0]);
                        if (!string.IsNullOrWhiteSpace(mapInfoPointModel.Error))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.MapInfoPoint, mapInfoPointModel.Error);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.MapInfoPoint, mapInfoPointModel.Error);
                            return;
                        }
                    }
                }
                else
                {
                    List<Coord> coordList = new List<Coord>()
                        {
                            new Coord() { Lat = Lat, Lng = Lng, Ordinal = 0 }
                        };

                    MapInfoModel mapInfoModelRet = mapInfoService.CreateMapInfoObjectDB(coordList, MapInfoDrawTypeEnum.Point, TVTypeEnum.HydrometricSite, hydrometricSiteModelRet.HydrometricSiteTVItemID);
                    if (!string.IsNullOrWhiteSpace(mapInfoModelRet.Error))
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.MapInfo, mapInfoModelRet.Error);
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.MapInfo, mapInfoModelRet.Error);
                        return;
                    }
                }
            }

            return;
        }
        ////public void UpdateHydrometricSiteDailyAndHourlyFromStartDateToEndDate()
        ////{
        ////    string NotUsed = "";

        ////    string[] ParamValueList = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
        ////    int HydrometricSiteTVItemID = 0;
        ////    int StartYear = 0;
        ////    int StartMonth = 0;
        ////    int StartDay = 0;
        ////    int EndYear = 0;
        ////    int EndMonth = 0;
        ////    int EndDay = 0;

        ////    if (_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID == 0)
        ////    {
        ////        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.TVItemID);
        ////        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.TVItemID);
        ////    }

        ////    if (ParamValueList.Count() != 9)
        ////    {
        ////        NotUsed = string.Format(TaskRunnerServiceRes.ParameterCount_NotEqual_, ParamValueList.Count().ToString(), 9);
        ////        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("ParameterCount_NotEqual_", ParamValueList.Count().ToString(), "9");
        ////    }

        ////    foreach (string s in ParamValueList)
        ////    {
        ////        string[] ParamValue = s.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

        ////        if (ParamValue.Length != 2)
        ////        {
        ////            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, TaskRunnerServiceRes.Parameters);
        ////            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", TaskRunnerServiceRes.Parameters);
        ////            return;
        ////        }

        ////        if (ParamValue[0] == "HydrometricSiteTVItemID")
        ////        {
        ////            HydrometricSiteTVItemID = int.Parse(ParamValue[1]);
        ////        }
        ////        else if (ParamValue[0] == "StartYear")
        ////        {
        ////            StartYear = int.Parse(ParamValue[1]);
        ////        }
        ////        else if (ParamValue[0] == "StartMonth")
        ////        {
        ////            StartMonth = int.Parse(ParamValue[1]);
        ////        }
        ////        else if (ParamValue[0] == "StartDay")
        ////        {
        ////            StartDay = int.Parse(ParamValue[1]);
        ////        }
        ////        else if (ParamValue[0] == "EndYear")
        ////        {
        ////            EndYear = int.Parse(ParamValue[1]);
        ////        }
        ////        else if (ParamValue[0] == "EndMonth")
        ////        {
        ////            EndMonth = int.Parse(ParamValue[1]);
        ////        }
        ////        else if (ParamValue[0] == "EndDay")
        ////        {
        ////            EndDay = int.Parse(ParamValue[1]);
        ////        }
        ////        else if (ParamValue[0] == "Generate")
        ////        {
        ////            // nothing
        ////        }
        ////        else if (ParamValue[0] == "Command")
        ////        {
        ////            // nothing
        ////        }
        ////        else
        ////        {
        ////            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, ParamValue[0]);
        ////            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", ParamValue[0].ToString());
        ////            return;
        ////        }
        ////    }

        ////    if (HydrometricSiteTVItemID == 0)
        ////    {
        ////        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, HydrometricSiteTVItemID.ToString());
        ////        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.HydrometricSiteTVItemID);
        ////        return;
        ////    }

        ////    if (StartYear == 0)
        ////    {
        ////        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, StartYear.ToString());
        ////        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.StartYear);
        ////        return;
        ////    }

        ////    if (StartMonth == 0)
        ////    {
        ////        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, StartMonth.ToString());
        ////        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.StartMonth);
        ////        return;
        ////    }

        ////    if (StartDay == 0)
        ////    {
        ////        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, StartDay.ToString());
        ////        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.StartDay);
        ////        return;
        ////    }

        ////    if (EndYear == 0)
        ////    {
        ////        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, EndYear.ToString());
        ////        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.EndYear);
        ////        return;
        ////    }

        ////    if (EndMonth == 0)
        ////    {
        ////        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, EndMonth.ToString());
        ////        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.EndMonth);
        ////        return;
        ////    }

        ////    if (EndDay == 0)
        ////    {
        ////        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, EndDay.ToString());
        ////        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.EndDay);
        ////        return;
        ////    }

        ////    HydrometricSiteService HydrometricSiteService = new HydrometricSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        ////    HydrometricSiteModel HydrometricSiteModel = HydrometricSiteService.GetHydrometricSiteModelWithHydrometricSiteTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        ////    if (!string.IsNullOrWhiteSpace(HydrometricSiteModel.Error))
        ////    {
        ////        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.HydrometricSite, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        ////        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.HydrometricSite, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        ////        return;
        ////    }

        ////    DateTime StartDate = new DateTime(StartYear, StartMonth, StartDay).AddDays(-10);
        ////    DateTime EndDate = new DateTime(EndYear, EndMonth, EndDay).AddDays(10);

        ////    for (int year = StartDate.Year; year <= EndDate.Year; year++)
        ////    {
        ////        string httpStrDaily = "";

        ////        using (WebClient webClient = new WebClient())
        ////        {
        ////            WebProxy webProxy = new WebProxy();
        ////            webClient.Proxy = webProxy;
        ////            string url = string.Format(UpdateHydrometricSiteDailyFromStartDateToEndDate, HydrometricSiteModel.ECDBID, year);
        ////            httpStrDaily = webClient.DownloadString(new Uri(url));
        ////            if (httpStrDaily.Length > 0)
        ////            {
        ////                if (httpStrDaily.Substring(0, "\"Station Name".Length) == "\"Station Name")
        ////                {
        ////                    httpStrDaily = httpStrDaily.Replace("\"", "").Replace("\n", "\r\n");
        ////                }
        ////            }
        ////            else
        ////            {
        ////                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotReadFile_, url);
        ////                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotReadFile_", url);
        ////                return;
        ////            }
        ////        }
        ////        UpdateDailyValuesForHydrometricSiteTVItemID(HydrometricSiteModel, httpStrDaily, StartDate, EndDate, new List<DateTime>());
        ////        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        ////            return;

        ////    }
        ////    return;
        ////}
        ////public void UpdateHydrometricSiteDailyAndHourlyForSubsectorFromStartDateToEndDate()
        ////{
        ////    string NotUsed = "";

        ////    HydrometricDataValueService HydrometricDataValueService = new HydrometricDataValueService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        ////    HydrometricSiteService HydrometricSiteService = new HydrometricSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        ////    TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        ////    MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        ////    UseOfSiteService useOfSiteService = new UseOfSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

        ////    string[] ParamValueList = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
        ////    int SubsectorTVItemID = 0;
        ////    int StartYear = 0;
        ////    int StartMonth = 0;
        ////    int StartDay = 0;
        ////    int EndYear = 0;
        ////    int EndMonth = 0;
        ////    int EndDay = 0;

        ////    if (_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID == 0)
        ////    {
        ////        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.TVItemID);
        ////        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.TVItemID);
        ////        return;
        ////    }


        ////    if (ParamValueList.Count() != 9)
        ////    {
        ////        NotUsed = string.Format(TaskRunnerServiceRes.ParameterCount_NotEqual_, ParamValueList.Count().ToString(), 9);
        ////        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("ParameterCount_NotEqual_", ParamValueList.Count().ToString(), "9");
        ////        return;
        ////    }

        ////    foreach (string s in ParamValueList)
        ////    {
        ////        string[] ParamValue = s.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

        ////        if (ParamValue.Length != 2)
        ////        {
        ////            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, TaskRunnerServiceRes.Parameters);
        ////            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", TaskRunnerServiceRes.Parameters);
        ////            return;
        ////        }

        ////        if (ParamValue[0] == "SubsectorTVItemID")
        ////        {
        ////            SubsectorTVItemID = int.Parse(ParamValue[1]);
        ////        }
        ////        else if (ParamValue[0] == "StartYear")
        ////        {
        ////            StartYear = int.Parse(ParamValue[1]);
        ////        }
        ////        else if (ParamValue[0] == "StartMonth")
        ////        {
        ////            StartMonth = int.Parse(ParamValue[1]);
        ////        }
        ////        else if (ParamValue[0] == "StartDay")
        ////        {
        ////            StartDay = int.Parse(ParamValue[1]);
        ////        }
        ////        else if (ParamValue[0] == "EndYear")
        ////        {
        ////            EndYear = int.Parse(ParamValue[1]);
        ////        }
        ////        else if (ParamValue[0] == "EndMonth")
        ////        {
        ////            EndMonth = int.Parse(ParamValue[1]);
        ////        }
        ////        else if (ParamValue[0] == "EndDay")
        ////        {
        ////            EndDay = int.Parse(ParamValue[1]);
        ////        }
        ////        else if (ParamValue[0] == "Generate")
        ////        {
        ////            // nothing
        ////        }
        ////        else if (ParamValue[0] == "Command")
        ////        {
        ////            // nothing
        ////        }
        ////        else
        ////        {
        ////            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, ParamValue[0]);
        ////            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", ParamValue[0].ToString());
        ////            return;
        ////        }
        ////    }

        ////    if (SubsectorTVItemID == 0)
        ////    {
        ////        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, SubsectorTVItemID.ToString());
        ////        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.SubsectorTVItemID);
        ////        return;
        ////    }

        ////    if (StartYear == 0)
        ////    {
        ////        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, StartYear.ToString());
        ////        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.StartYear);
        ////        return;
        ////    }

        ////    if (StartMonth == 0)
        ////    {
        ////        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, StartMonth.ToString());
        ////        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.StartMonth);
        ////        return;
        ////    }

        ////    if (StartDay == 0)
        ////    {
        ////        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, StartDay.ToString());
        ////        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.StartDay);
        ////        return;
        ////    }

        ////    if (EndYear == 0)
        ////    {
        ////        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, EndYear.ToString());
        ////        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.EndYear);
        ////        return;
        ////    }

        ////    if (EndMonth == 0)
        ////    {
        ////        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, EndMonth.ToString());
        ////        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.EndMonth);
        ////        return;
        ////    }

        ////    if (EndDay == 0)
        ////    {
        ////        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, EndDay.ToString());
        ////        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.EndDay);
        ////        return;
        ////    }

        ////    DateTime StartDate = new DateTime(StartYear, StartMonth, StartDay).AddDays(-10);
        ////    DateTime EndDate = new DateTime(EndYear, EndMonth, EndDay).AddDays(10);

        ////    TVItemModel tvItemModelSubsector = tvItemService.GetTVItemModelWithTVItemIDDB(SubsectorTVItemID);
        ////    if (!string.IsNullOrWhiteSpace(tvItemModelSubsector.Error))
        ////    {
        ////        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, SubsectorTVItemID.ToString());
        ////        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, SubsectorTVItemID.ToString());
        ////        return;

        ////    }

        ////    List<MapInfoPointModel> mapInfoPointListSubsector = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.Subsector, MapInfoDrawTypeEnum.Point);
        ////    if (mapInfoPointListSubsector.Count == 0)
        ////    {
        ////        NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBeNullOrEmpty, TaskRunnerServiceRes.mapInfoPointListSubsector);
        ////        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_ShouldNotBeNullOrEmpty", TaskRunnerServiceRes.mapInfoPointListSubsector);
        ////        return;
        ////    }

        ////    float LookupRadius = 100000.0f; // 100 km radius

        ////    List<MapInfoModel> mapInfoModelList = new List<MapInfoModel>();
        ////    List<HydrometricSiteModel> HydrometricSiteModelListWithHourly = new List<HydrometricSiteModel>();
        ////    List<HydrometricSiteModel> HydrometricSiteModelListWithDaily = new List<HydrometricSiteModel>();

        ////    while (mapInfoModelList.Count == 0)
        ////    {
        ////        mapInfoModelList = mapInfoService.GetMapInfoModelWithinCircleWithTVTypeAndMapInfoDrawTypeDB((float)mapInfoPointListSubsector[0].Lat, (float)mapInfoPointListSubsector[0].Lng, LookupRadius, TVTypeEnum.HydrometricSite, MapInfoDrawTypeEnum.Point);

        ////        foreach (MapInfoModel mapInfoModel in mapInfoModelList)
        ////        {
        ////            HydrometricSiteModel HydrometricSiteModel = HydrometricSiteService.GetHydrometricSiteModelWithHydrometricSiteTVItemIDDB(mapInfoModel.TVItemID);
        ////            if (!string.IsNullOrWhiteSpace(HydrometricSiteModel.Error))
        ////            {
        ////                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, mapInfoModel.TVItemID.ToString());
        ////                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, mapInfoModel.TVItemID.ToString());
        ////                return;
        ////            }
        ////            if (HydrometricSiteModel.HourlyStartDate_Local != null)
        ////            {
        ////                HydrometricSiteModelListWithHourly.Add(HydrometricSiteModel);
        ////            }
        ////            if (HydrometricSiteModel.DailyStartDate_Local != null)
        ////            {
        ////                HydrometricSiteModelListWithDaily.Add(HydrometricSiteModel);
        ////            }
        ////        }

        ////        LookupRadius += 10000;
        ////    }

        ////    for (int year = StartDate.Year; year <= EndDate.Year; year++)
        ////    {

        ////        List<HydrometricSiteModel> HydrometricSiteForDailyModelListUsed = new List<HydrometricSiteModel>();
        ////        //List<HydrometricSiteModel> HydrometricSiteForHourlyModelListUsed = new List<HydrometricSiteModel>();
        ////        foreach (HydrometricSiteModel HydrometricSiteModel in HydrometricSiteModelListWithDaily)
        ////        {
        ////            if (HydrometricSiteForDailyModelListUsed.Count > 2)
        ////                break;

        ////            if (HydrometricSiteModel.DailyStartDate_Local.Value <= StartDate && HydrometricSiteModel.DailyEndDate_Local.Value >= EndDate)
        ////            {
        ////                HydrometricDataValueModel HydrometricDataValueStartDone = HydrometricDataValueService.GetHydrometricDataValueModelExitDB(new HydrometricDataValueModel() { HydrometricSiteID = HydrometricSiteModel.HydrometricSiteID, DateTime_Local = StartDate });
        ////                HydrometricDataValueModel HydrometricDataValueEndDone = HydrometricDataValueService.GetHydrometricDataValueModelExitDB(new HydrometricDataValueModel() { HydrometricSiteID = HydrometricSiteModel.HydrometricSiteID, DateTime_Local = EndDate });
        ////                if (string.IsNullOrWhiteSpace(HydrometricDataValueStartDone.Error) && string.IsNullOrWhiteSpace(HydrometricDataValueEndDone.Error))
        ////                {
        ////                    HydrometricSiteForDailyModelListUsed.Add(HydrometricSiteModel);
        ////                    continue;
        ////                }
        ////                string httpStr = "";
        ////                using (WebClient webClient = new WebClient())
        ////                {
        ////                    WebProxy webProxy = new WebProxy();
        ////                    webClient.Proxy = webProxy;
        ////                    string url = string.Format(UpdateHydrometricSiteDailyFromStartDateToEndDate, HydrometricSiteModel.ECDBID, year);
        ////                    httpStr = webClient.DownloadString(new Uri(url));
        ////                    if (httpStr.Length > 0)
        ////                    {
        ////                        if (httpStr.Substring(0, "\"Station Name".Length) == "\"Station Name")
        ////                        {
        ////                            httpStr = httpStr.Replace("\"", "").Replace("\n", "\r\n");
        ////                        }
        ////                    }
        ////                    else
        ////                    {
        ////                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotReadFile_, url);
        ////                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotReadFile_", url);
        ////                        return;
        ////                    }
        ////                }
        ////                HydrometricDataValueModel HydrometricDataValueModel = UpdateDailyValuesForHydrometricSiteTVItemID(HydrometricSiteModel, httpStr, StartDate, EndDate, new List<DateTime>());
        ////                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        ////                    return;

        ////                if (HydrometricDataValueModel.Rainfall_mm != null || HydrometricDataValueModel.TotalPrecip_mm_cm != null)
        ////                {
        ////                    HydrometricSiteForDailyModelListUsed.Add(HydrometricSiteModel);
        ////                }

        ////                UseOfSiteModel useOfSiteModelNew = new UseOfSiteModel()
        ////                {
        ////                    SiteTVItemID = HydrometricSiteModel.HydrometricSiteTVItemID,
        ////                    SubsectorTVItemID = tvItemModelSubsector.TVItemID,
        ////                    SiteType = SiteTypeEnum.Hydrometric,
        ////                    Ordinal = 0, // will be replaced 
        ////                    StartYear = year,
        ////                    EndYear = year,
        ////                    UseWeight = false,
        ////                    Weight_perc = 1.0f,
        ////                    UseEquation = false,
        ////                };

        ////                List<UseOfSiteModel> useOfSiteModelList = useOfSiteService.GetUseOfSiteModelListWithSiteTypeAndSubsectorTVItemIDDB(SiteTypeEnum.Hydrometric, tvItemModelSubsector.TVItemID);
        ////                bool hasUpdated = false;
        ////                foreach (UseOfSiteModel useOfSiteModel in useOfSiteModelList)
        ////                {
        ////                    if (useOfSiteModel.EndYear == year - 1 || useOfSiteModel.EndYear == year)
        ////                    {
        ////                        useOfSiteModel.EndYear = year;
        ////                        UseOfSiteModel useOfSiteModelRet = useOfSiteService.PostUpdateUseOfSiteDB(useOfSiteModel);
        ////                        if (!string.IsNullOrWhiteSpace(useOfSiteModelRet.Error))
        ////                        {
        ////                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.UseOfSite, useOfSiteModelRet.Error);
        ////                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.UseOfSite, useOfSiteModelRet.Error);
        ////                            return;
        ////                        }
        ////                        hasUpdated = true;
        ////                        break;
        ////                    }
        ////                }
        ////                if (!hasUpdated)
        ////                {
        ////                    if (useOfSiteModelList.Count > 0)
        ////                    {
        ////                        useOfSiteModelNew.Ordinal = useOfSiteModelList.OrderBy(c => c.Ordinal).Last().Ordinal + 1;
        ////                    }
        ////                    UseOfSiteModel useOfSiteModelRet = useOfSiteService.PostAddUseOfSiteDB(useOfSiteModelNew);
        ////                    if (!string.IsNullOrWhiteSpace(useOfSiteModelRet.Error))
        ////                    {
        ////                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.UseOfSite, useOfSiteModelRet.Error);
        ////                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(useOfSiteModelRet.Error);
        ////                        return;
        ////                    }
        ////                }

        ////                //if (!string.IsNullOrWhiteSpace(HydrometricDataValueModel.HourlyValues))
        ////                //{
        ////                //    HydrometricSiteForHourlyModelListUsed.Add(HydrometricSiteModel);
        ////                //}
        ////            }
        ////        }

        ////        //    if (HydrometricSiteForHourlyModelListUsed.Count == 0)
        ////        //    {
        ////        //        foreach (HydrometricSiteModel HydrometricSiteModel in HydrometricSiteModelListWithHourly)
        ////        //        {
        ////        //            if (HydrometricSiteForHourlyModelListUsed.Count > 0)
        ////        //                break;

        ////        //            if (HydrometricSiteModel.DailyStartDate_Local.Value <= StartDate && HydrometricSiteModel.DailyEndDate_Local.Value >= EndDate)
        ////        //            {
        ////        //                string httpStr = "";
        ////        //                using (WebClient webClient = new WebClient())
        ////        //                {
        ////        //                    WebProxy webProxy = new WebProxy();
        ////        //                    webClient.Proxy = webProxy;
        ////        //                    string url = string.Format(UpdateHydrometricSiteDailyFromStartDateToEndDate, HydrometricSiteModel.ECDBID, year);
        ////        //                    httpStr = webClient.DownloadString(new Uri(url));
        ////        //                    if (httpStr.Length > 0)
        ////        //                    {
        ////        //                        if (httpStr.Substring(0, "\"Station Name".Length) == "\"Station Name")
        ////        //                        {
        ////        //                            httpStr = httpStr.Replace("\"", "").Replace("\n", "\r\n");
        ////        //                        }
        ////        //                    }
        ////        //                    else
        ////        //                    {
        ////        //                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotReadFile_, url);
        ////        //                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotReadFile_", url);
        ////        //                        return;
        ////        //                    }
        ////        //                }
        ////        //                HydrometricDataValueModel HydrometricDataValueModel = UpdateDailyValuesForHydrometricSiteTVItemID(HydrometricSiteModel.HydrometricSiteTVItemID, httpStr, StartDate, EndDate);
        ////        //                if (!string.IsNullOrWhiteSpace(HydrometricDataValueModel.Error))
        ////        //                {
        ////        //                    NotUsed = string.Format(TaskRunnerServiceRes._Required, TaskRunnerServiceRes.TVItemID);
        ////        //                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.TVItemID));
        ////        //                    return;
        ////        //                }
        ////        //                if (HydrometricDataValueModel.Rainfall_mm != null)
        ////        //                {
        ////        //                    HydrometricSiteForDailyModelListUsed.Add(HydrometricSiteModel);
        ////        //                }
        ////        //                //if (HydrometricDataValueModel.HourlyValues != null)
        ////        //                //{
        ////        //                //    HydrometricSiteForHourlyModelListUsed.Add(HydrometricSiteModel);
        ////        //                //    break;
        ////        //                //}
        ////        //            }
        ////        //        }
        ////        //    }

        ////        //    foreach (HydrometricSiteModel HydrometricSiteModelHourly in HydrometricSiteForHourlyModelListUsed)
        ////        //    {
        ////        //        HydrometricSiteForDailyModelListUsed.Add(HydrometricSiteModelHourly);
        ////        //    }
        ////    }
        ////}
        //public void UpdateDailyValuesForHydrometricSiteTVItemID(HydrometricSiteModel HydrometricSiteModel, string httpStrDaily, DateTime StartDate, DateTime EndDate, List<DateTime> HourlyDateListToLoad)
        //{
        //    string NotUsed = "";

        //    //string httpStrHourly = "";
        //    //HydrometricDataValueModel HydrometricDataValueModel = new HydrometricDataValueModel();
        //    List<string> FullProvList = new List<string>() { "BRITISH COLUMBIA", "NEW BRUNSWICK", "NEWFOUNDLAND", "NOVA SCOTIA", "PRINCE EDWARD ISLAND", "QUEBEC" };
        //    List<string> ShortProvList = new List<string>() { "BC", "NB", "NL", "NS", "PE", "QC" };

        //    StringBuilder hourlyValues = new StringBuilder();

        //    hourlyValues.Clear();

        //    HydrometricSiteService HydrometricSiteService = new HydrometricSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    HydrometricDataValueService HydrometricDataValueService = new HydrometricDataValueService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

        //    using (TextReader tr = new StringReader(httpStrDaily))
        //    {

        //        int countLine = 0;
        //        string LookupTxt = "";
        //        while (true)
        //        {
        //            countLine += 1;
        //            string lineStr = tr.ReadLine();
        //            if (lineStr == null)
        //            {
        //                break;
        //            }

        //            List<string> lineValueArr = lineStr.Split(",".ToCharArray(), StringSplitOptions.None).ToList();

        //            if (countLine == 1)
        //            {
        //                LookupTxt = "Station Name";
        //                if (lineValueArr[0] != LookupTxt)
        //                {
        //                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_AtLine_, LookupTxt, countLine.ToString());
        //                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
        //                    return;
        //                }
        //                LookupTxt = HydrometricSiteModel.HydrometricSiteName;
        //                if (lineValueArr[1] != LookupTxt)
        //                {
        //                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_AtLine_, LookupTxt, countLine.ToString());
        //                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
        //                    return;
        //                }
        //            }
        //            if (countLine == 2)
        //            {
        //                LookupTxt = "Province";
        //                if (lineValueArr[0] != LookupTxt)
        //                {
        //                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_AtLine_, LookupTxt, countLine.ToString());
        //                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
        //                    return;
        //                }
        //                LookupTxt = lineValueArr[1];
        //                if (!FullProvList.Contains(LookupTxt))
        //                {
        //                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_AtLine_, LookupTxt, countLine.ToString());
        //                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
        //                    return;
        //                }
        //                for (int i = 0; i < 6; i++)
        //                {
        //                    if (lineValueArr[1] == FullProvList[i])
        //                    {
        //                        if (HydrometricSiteModel.Province != ShortProvList[i])
        //                        {
        //                            NotUsed = string.Format(TaskRunnerServiceRes.Province_NotEqualTo_AtLine_, ShortProvList[i], HydrometricSiteModel.Province, countLine.ToString());
        //                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
        //                            return;
        //                        }
        //                    }
        //                }
        //            }
        //            if (countLine == 6)
        //            {
        //                LookupTxt = "Hydrometric Identifier";
        //                if (lineValueArr[0] != LookupTxt)
        //                {
        //                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_AtLine_, LookupTxt, countLine.ToString());
        //                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
        //                    return;
        //                }
        //                if (lineValueArr[1].Length > 0)
        //                {
        //                    LookupTxt = HydrometricSiteModel.HydrometricID;
        //                    if (lineValueArr[1] != HydrometricSiteModel.HydrometricID)
        //                    {
        //                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_AtLine_, LookupTxt, countLine.ToString());
        //                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
        //                        return;
        //                    }
        //                }
        //            }
        //            if (lineValueArr[0].Contains("-"))
        //            {
        //                if (lineValueArr[0].Substring(4, 1) == "-")
        //                {
        //                    if (lineValueArr.Count != 27)
        //                    {
        //                        NotUsed = string.Format(TaskRunnerServiceRes.CountOfValuesInLine_ShouldBe_, countLine, "27");
        //                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CountOfValuesInLine_ShouldBe_", countLine.ToString(), "27");
        //                        return;
        //                    }

        //                    int Year = int.Parse(lineValueArr[1]);
        //                    int Month = int.Parse(lineValueArr[2]);
        //                    int Day = int.Parse(lineValueArr[3]);

        //                    if (Year == 0 || Month == 0 || Day == 0)
        //                    {
        //                        NotUsed = string.Format(TaskRunnerServiceRes.YearNotCorrect_AtLine_, Year.ToString() + " - " + Month.ToString() + " - " + Day.ToString(), countLine.ToString());
        //                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("YearNotCorrect_AtLine_", Year.ToString() + " - " + Month.ToString() + " - " + Day.ToString(), countLine.ToString());
        //                        return;
        //                    }

        //                    DateTime LineDate = new DateTime(Year, Month, Day);

        //                    if (!(StartDate <= LineDate && EndDate >= LineDate))
        //                    {
        //                        continue;
        //                    }

        //                    string httpStrHourly = "";
        //                    if (HourlyDateListToLoad.Contains(LineDate))
        //                    {
        //                        using (WebClient webClient = new WebClient())
        //                        {
        //                            WebProxy webProxy = new WebProxy();
        //                            webClient.Proxy = webProxy;
        //                            string url = string.Format(UpdateHydrometricSiteHourlyFromStartDateToEndDate, HydrometricSiteModel.ECDBID, Year, LineDate.Month);
        //                            httpStrHourly = webClient.DownloadString(new Uri(url));
        //                            if (httpStrHourly.Length > 0)
        //                            {
        //                                if (httpStrHourly.Substring(0, "\"Station Name".Length) == "\"Station Name")
        //                                {
        //                                    httpStrHourly = httpStrHourly.Replace("\"", "").Replace("\n", "\r\n");
        //                                }
        //                            }
        //                            else
        //                            {
        //                                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotReadFile_, url);
        //                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotReadFile_", url);
        //                                return;
        //                            }
        //                        }

        //                        if (HydrometricSiteModel.HourlyStartDate_Local <= LineDate && HydrometricSiteModel.HourlyEndDate_Local >= LineDate)
        //                        {
        //                            UpdateHourlyValuesForHydrometricSiteAndDate(HydrometricSiteModel, httpStrHourly, LineDate, hourlyValues);
        //                            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                                return;

        //                            if (!string.IsNullOrWhiteSpace(hourlyValues.ToString()))
        //                            {
        //                                string Title = "Year,Month,Day,Hour,Temp_C,DewPoint_C,RelHum_Perc,WindDir_10deg,WindSpd_km_h,Visibility_km,StnPress_kPa,Humidx,WindChill\r\n";
        //                                hourlyValues.Append(Title + hourlyValues.ToString());
        //                            }
        //                        }
        //                    }


        //                    float? MaxTemp = null;
        //                    if (lineValueArr[5].Length > 0)
        //                    {
        //                        MaxTemp = float.Parse(lineValueArr[5]);
        //                    }
        //                    float? MinTemp = null;
        //                    if (lineValueArr[7].Length > 0)
        //                    {
        //                        MinTemp = float.Parse(lineValueArr[7]);
        //                    }
        //                    float? MeanTemp = null;
        //                    if (lineValueArr[9].Length > 0)
        //                    {
        //                        MeanTemp = float.Parse(lineValueArr[9]);
        //                    }
        //                    float? HeatDegDays = null;
        //                    if (lineValueArr[11].Length > 0)
        //                    {
        //                        HeatDegDays = float.Parse(lineValueArr[11]);
        //                    }
        //                    float? CoolDegDays = null;
        //                    if (lineValueArr[13].Length > 0)
        //                    {
        //                        CoolDegDays = float.Parse(lineValueArr[13]);
        //                    }
        //                    float? TotalRain = null;
        //                    if (lineValueArr[15].Length > 0)
        //                    {
        //                        if (lineValueArr[15].Trim() == "T")
        //                        {
        //                            TotalRain = 0.0f;
        //                        }
        //                        else
        //                        {
        //                            TotalRain = float.Parse(lineValueArr[15]);
        //                        }
        //                    }
        //                    float? TotalSnow = null;
        //                    if (lineValueArr[17].Length > 0)
        //                    {
        //                        if (lineValueArr[17].Trim() == "T")
        //                        {
        //                            TotalSnow = 0.0f;
        //                        }
        //                        else
        //                        {
        //                            TotalSnow = float.Parse(lineValueArr[17]);
        //                        }
        //                    }
        //                    float? TotalPrecip = null;
        //                    if (lineValueArr[19].Length > 0)
        //                    {
        //                        if (lineValueArr[19].Trim() == "T")
        //                        {
        //                            TotalPrecip = 0.0f;
        //                        }
        //                        else
        //                        {
        //                            TotalPrecip = float.Parse(lineValueArr[19]);
        //                        }
        //                    }
        //                    float? SnowOnGround = null;
        //                    if (lineValueArr[21].Length > 0)
        //                    {
        //                        if (lineValueArr[21].Trim() == "T")
        //                        {
        //                            SnowOnGround = 0.0f;
        //                        }
        //                        else
        //                        {
        //                            SnowOnGround = float.Parse(lineValueArr[21]);
        //                        }
        //                    }
        //                    float? DirMaxGust = null;
        //                    if (lineValueArr[23].Length > 0)
        //                    {
        //                        DirMaxGust = float.Parse(lineValueArr[23]);
        //                    }
        //                    float? SpdMaxGust = null;
        //                    if (lineValueArr[25].Length > 0)
        //                    {
        //                        if (lineValueArr[25].Substring(0, 1) == "<")
        //                        {
        //                            SpdMaxGust = float.Parse(lineValueArr[25].Substring(1));
        //                        }
        //                        else
        //                        {
        //                            SpdMaxGust = float.Parse(lineValueArr[25]);
        //                        }
        //                    }

        //                    HydrometricDataValueModel HydrometricDataValueModelNew = new HydrometricDataValueModel()
        //                    {
        //                        HydrometricSiteID = HydrometricSiteModel.HydrometricSiteID,
        //                        HasBeenRead = true,
        //                        CoolDegDays_C = CoolDegDays,
        //                        DateTime_Local = LineDate,
        //                        DirMaxGust_0North = DirMaxGust,
        //                        HeatDegDays_C = HeatDegDays,
        //                        HourlyValues = hourlyValues.ToString(),
        //                        Keep = true,
        //                        MaxTemp_C = MaxTemp,
        //                        MinTemp_C = MinTemp,
        //                        Rainfall_mm = TotalRain,
        //                        RainfallEntered_mm = null,
        //                        Snow_cm = TotalSnow,
        //                        SnowOnGround_cm = SnowOnGround,
        //                        SpdMaxGust_kmh = SpdMaxGust,
        //                        StorageDataType = StorageDataTypeEnum.Archived,
        //                        TotalPrecip_mm_cm = TotalPrecip,
        //                    };
        //                    HydrometricDataValueModel HydrometricDataValueModelExist = HydrometricDataValueService.GetHydrometricDataValueModelExitDB(HydrometricDataValueModelNew);
        //                    if (!string.IsNullOrWhiteSpace(HydrometricDataValueModelExist.Error))
        //                    {
        //                        HydrometricDataValueModel HydrometricDataValueModelRet = HydrometricDataValueService.PostAddHydrometricDataValueDB(HydrometricDataValueModelNew);
        //                        if (!string.IsNullOrWhiteSpace(HydrometricDataValueModelRet.Error))
        //                        {
        //                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.HydrometricDataValue, HydrometricDataValueModelRet.Error);
        //                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.HydrometricDataValue, HydrometricDataValueModelRet.Error);
        //                            return;
        //                        }
        //                    }
        //                    else
        //                    {

        //                        if (HydrometricDataValueModelExist.HasBeenRead == false)
        //                        {
        //                            HydrometricDataValueModelNew.HydrometricDataValueID = HydrometricDataValueModelExist.HydrometricDataValueID;
        //                            HydrometricDataValueModelNew.RainfallEntered_mm = HydrometricDataValueModelExist.RainfallEntered_mm;
        //                            HydrometricDataValueModelNew.HasBeenRead = true;
        //                            HydrometricDataValueModelExist = HydrometricDataValueService.PostUpdateHydrometricDataValueDB(HydrometricDataValueModelNew);
        //                            if (!string.IsNullOrWhiteSpace(HydrometricDataValueModelExist.Error))
        //                            {
        //                                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.HydrometricDataValue, HydrometricDataValueModelExist.Error); ;
        //                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.HydrometricDataValue, HydrometricDataValueModelExist.Error);
        //                                return;
        //                            }
        //                        }
        //                    }

        //                }
        //            }
        //        }
        //    }
        //}
        //public void UpdateHourlyValuesForHydrometricSiteAndDate(HydrometricSiteModel HydrometricSiteModel, string httpStrHourly, DateTime CurrentDate, StringBuilder hourlyValues)
        //{
        //    string NotUsed = "";
        //    bool HasValue = false;

        //    //HydrometricDataValueModel HydrometricDataValueModel = new HydrometricDataValueModel();
        //    List<string> FullProvList = new List<string>() { "BRITISH COLUMBIA", "NEW BRUNSWICK", "NEWFOUNDLAND", "NOVA SCOTIA", "PRINCE EDWARD ISLAND", "QUEBEC" };
        //    List<string> ShortProvList = new List<string>() { "BC", "NB", "NL", "NS", "PE", "QC" };

        //    hourlyValues.Clear();

        //    using (TextReader tr = new StringReader(httpStrHourly))
        //    {
        //        string LookupTxt = "";
        //        int countLine = 0;
        //        while (true)
        //        {
        //            countLine += 1;
        //            string lineStr = tr.ReadLine();
        //            if (lineStr == null)
        //                break;

        //            List<string> lineValueArr = lineStr.Split(",".ToCharArray(), StringSplitOptions.None).ToList();
        //            if (countLine == 1)
        //            {
        //                LookupTxt = "Station Name";
        //                if (lineValueArr[0] != LookupTxt)
        //                {
        //                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_AtLine_, LookupTxt, countLine.ToString());
        //                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
        //                    return;
        //                }
        //                LookupTxt = HydrometricSiteModel.HydrometricSiteName;
        //                if (lineValueArr[1] != LookupTxt)
        //                {
        //                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_AtLine_, LookupTxt, countLine.ToString());
        //                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
        //                    return;
        //                }
        //            }
        //            if (countLine == 2)
        //            {
        //                LookupTxt = "Province";
        //                if (lineValueArr[0] != LookupTxt)
        //                {
        //                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_AtLine_, LookupTxt, countLine.ToString());
        //                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
        //                    return;
        //                }
        //                LookupTxt = lineValueArr[1];
        //                if (!FullProvList.Contains(LookupTxt))
        //                {
        //                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_AtLine_, LookupTxt, countLine.ToString());
        //                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
        //                    return;
        //                }
        //                for (int i = 0; i < 6; i++)
        //                {
        //                    if (lineValueArr[1] == FullProvList[i])
        //                    {
        //                        if (HydrometricSiteModel.Province != ShortProvList[i])
        //                        {
        //                            NotUsed = string.Format(TaskRunnerServiceRes.Province_NotEqualTo_AtLine_, ShortProvList[i], HydrometricSiteModel.Province, countLine.ToString());
        //                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
        //                            return;
        //                        }
        //                    }
        //                }
        //            }
        //            if (countLine == 6)
        //            {
        //                LookupTxt = "Hydrometric Identifier";
        //                if (lineValueArr[0] != LookupTxt)
        //                {
        //                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_AtLine_, LookupTxt, countLine.ToString());
        //                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
        //                    return;
        //                }
        //                if (lineValueArr[1].Length > 0)
        //                {
        //                    LookupTxt = HydrometricSiteModel.HydrometricID;
        //                    if (lineValueArr[1] != HydrometricSiteModel.HydrometricID)
        //                    {
        //                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_AtLine_, LookupTxt, countLine.ToString());
        //                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
        //                        return;
        //                    }
        //                }
        //            }
        //            if (lineValueArr[0].Contains("-"))
        //            {
        //                if (lineValueArr[0].Substring(4, 1) == "-")
        //                {
        //                    if (lineValueArr.Count != 25 && lineValueArr.Count != 26 && lineValueArr.Count != 27 && lineValueArr.Count != 28 && lineValueArr.Count != 5)
        //                    {
        //                        NotUsed = string.Format(TaskRunnerServiceRes.CountOfValuesInLine_ShouldBe_, countLine, "25");
        //                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CountOfValuesInLine_ShouldBe_", countLine.ToString(), "25");
        //                        return;
        //                    }

        //                    int Year = int.Parse(lineValueArr[1]);
        //                    int Month = int.Parse(lineValueArr[2]);
        //                    int Day = int.Parse(lineValueArr[3]);

        //                    if (Year == 0 || Month == 0 || Day == 0)
        //                    {
        //                        NotUsed = string.Format(TaskRunnerServiceRes.YearNotCorrect_AtLine_, Year.ToString() + " - " + Month.ToString() + " - " + Day.ToString(), countLine.ToString());
        //                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("YearNotCorrect_AtLine_", Year.ToString() + " - " + Month.ToString() + " - " + Day.ToString(), countLine.ToString());
        //                        return;
        //                    }


        //                    DateTime LineDate = new DateTime(Year, Month, Day);

        //                    if (CurrentDate != LineDate)
        //                    {
        //                        continue;
        //                    }

        //                    int? Hour = null;
        //                    if (lineValueArr[4].Length > 2)
        //                    {
        //                        Hour = int.Parse(lineValueArr[4].Substring(0, 2));
        //                    }

        //                    float? Temp_C = null;
        //                    float? DewPoint_C = null;
        //                    float? RelHum_Perc = null;
        //                    float? WindDir_10deg = null;
        //                    float? WindSpd_km_h = null;
        //                    float? Visibility_km = null;
        //                    float? StnPress_kPa = null;
        //                    float? Humidx = null;
        //                    float? WindChill = null;

        //                    if (lineValueArr.Count > 5)
        //                    {
        //                        HasValue = true;
        //                        if (lineValueArr[6].Length > 0)
        //                        {
        //                            Temp_C = float.Parse(lineValueArr[6]);
        //                        }
        //                        if (lineValueArr[8].Length > 0)
        //                        {
        //                            DewPoint_C = float.Parse(lineValueArr[8]);
        //                        }
        //                        if (lineValueArr[10].Length > 0)
        //                        {
        //                            RelHum_Perc = float.Parse(lineValueArr[10]);
        //                        }
        //                        if (lineValueArr[12].Length > 0)
        //                        {
        //                            WindDir_10deg = float.Parse(lineValueArr[12]);
        //                        }
        //                        if (lineValueArr[14].Length > 0)
        //                        {
        //                            WindSpd_km_h = float.Parse(lineValueArr[14]);
        //                        }
        //                        if (lineValueArr[16].Length > 0)
        //                        {
        //                            Visibility_km = float.Parse(lineValueArr[16]);
        //                        }
        //                        if (lineValueArr[18].Length > 0)
        //                        {
        //                            StnPress_kPa = float.Parse(lineValueArr[18]);
        //                        }
        //                        if (lineValueArr[20].Length > 0)
        //                        {
        //                            Humidx = float.Parse(lineValueArr[20]);
        //                        }
        //                        if (lineValueArr[22].Length > 0)
        //                        {
        //                            WindChill = float.Parse(lineValueArr[22]);
        //                        }

        //                        hourlyValues.AppendLine(
        //                            Year.ToString() + "," +
        //                            Month.ToString() + "," +
        //                            Day.ToString() + "," +
        //                            Hour.ToString() + "," +
        //                            (Temp_C == null ? "" : Temp_C.ToString()) + "," +
        //                            (DewPoint_C == null ? "" : DewPoint_C.ToString()) + "," +
        //                            (RelHum_Perc == null ? "" : RelHum_Perc.ToString()) + "," +
        //                            (WindDir_10deg == null ? "" : WindDir_10deg.ToString()) + "," +
        //                            (WindSpd_km_h == null ? "" : WindSpd_km_h.ToString()) + "," +
        //                            (Visibility_km == null ? "" : Visibility_km.ToString()) + "," +
        //                            (StnPress_kPa == null ? "" : StnPress_kPa.ToString()) + "," +
        //                            (Humidx == null ? "" : Humidx.ToString()) + "," +
        //                            (WindChill == null ? "" : WindChill.ToString())
        //                            );
        //                    }
        //                }
        //            }
        //        }
        //    }

        //    if (!HasValue)
        //    {
        //        hourlyValues.Clear();
        //    }
        //}
        #endregion Functions public

        #region Functions private

        #endregion Functions private

    }
}
