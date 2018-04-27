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

namespace CSSPWebToolsTaskRunner.Services
{
    public class ClimateService
    {
        #region Variables
        public List<string> ProvList = new List<string>() { "BC", "NB", "NL", "NS", "PE", "QC" };
        public List<string> ProvNameListEN = new List<string>() { "British Columbia", "New Brunswick", "Newfoundland and Labrador", "Nova Scotia", "Prince Edward Island", "Québec" };
        public List<string> ProvNameListFR = new List<string>() { "Colombie-Britannique", "Nouveau-Brunswick", "Terre-Neuve-et-Labrador", "Nouvelle-Écosse", "Île-du-Prince-Édouard", "Québec" };
        public string URLUpdateClimateSiteInfoHourlyDailyMonthly = "http://climate.weather.gc.ca/advanceSearch/searchHistoricDataStations_e.html?searchType=stnProv&timeframe=1&lstProvince={0}&optLimit=yearRange&StartYear=1840&EndYear=2015&Year=2050&Month=1&Day=1&selRowPerPage=100&cmdProvSubmit=Search&startRow={1}";
        public string URLUpdateClimateSiteInfoLatLngWMOElevTC = "http://climate.weather.gc.ca/climateData/dailydata_e.html?timeframe=2&StationID={0}&Year=1800&Month=1&Day=01";
        public string UpdateClimateSiteDailyFromStartDateToEndDate = "http://climate.weather.gc.ca/climateData/bulk_data_e.html?format=csv&stationID={0}&Year={1}&Month=1&Day=1&timeframe=2&submit=Download+Data";
        //public string GetClimateSiteDataForRun = "http://climate.weather.gc.ca/climateData/bulk_data_e.html?format=csv&stationID={0}&Year={1}&Month=1&Day=1&timeframe=2&submit=Download+Data";
        public string UrlToGetClimateSiteDataForRunsOfYear = "http://climate.weather.gc.ca/climate_data/bulk_data_e.html?format=csv&stationID={0}&Year={1}&Month=1&Day=1&timeframe=2&submit=Download+Data";
        public string UpdateClimateSiteHourlyFromStartDateToEndDate = "http://climate.weather.gc.ca/climate_data/bulk_data_e.html?format=csv&stationID={0}&Year={1}&Month={2}&Day=1&timeframe=1&submit=Download+Data";
        public bool WebBrowserContentAnalysed = true;
        public bool AllDone = false;
        #endregion Variables

        #region Properties
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        #endregion Properties

        #region Constructors
        public ClimateService(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
        }
        #endregion Constructors

        #region Events
        public void browser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            string NotUsed = "";

            if (e.Url.AbsolutePath != (sender as WebBrowser).Url.AbsolutePath)
                return;

            if ((sender as WebBrowser).Url == e.Url)
            {
                if ((sender as WebBrowser).Url.ToString().Contains("advanceSearch/searchHistoricDataStations_e.html"))
                {
                    UpdateClimateSitesInformationForProvinceTVItemIDParse(sender, e);
                }
                else if ((sender as WebBrowser).Url.ToString().Contains("climateData/dailydata_e.html"))
                {
                    UpdateClimateSiteInfoLatLngWMOElevTCParse(sender, e);
                }
                else if ((sender as WebBrowser).Url.ToString().Contains("climateData/bulkdata_e.html"))
                {
                    if ((sender as WebBrowser).Url.ToString().Contains("timeframe=2"))
                    {
                        UpdateClimateSiteInfoLatLngWMOElevTCParse(sender, e);
                    }
                    else if ((sender as WebBrowser).Url.ToString().Contains("timeframe=1"))
                    {
                        UpdateClimateSiteInfoLatLngWMOElevTCParse(sender, e);
                    }
                    else
                    {
                    }
                }
                else
                {

                }
            }
            else
            {
                NotUsed = TaskRunnerServiceRes.CurrentURLNotTheSameInDocumentCompletedEvent;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("CurrentURLNotTheSameInDocumentCompletedEvent");
                return;
            }

            Application.ExitThread();   // Stops the thread
        }
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
        public void RunBrowserThread(string URL)
        {
            var thread = new Thread(() =>
            {
                var br = new WebBrowser();
                br.DocumentCompleted += browser_DocumentCompleted;
                br.Navigate(URL);
                Application.Run();
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();

            return;
        }
        public void FindMissingPrecipForProvince()
        {
            string NotUsed = "";

            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MWQMRunService mwqmRunService = new MWQMRunService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

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

            string Parameters = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters;
            string[] ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            int ProvinceTVItemID = 0;
            foreach (string s in ParamValueList)
            {
                string[] ParamValue = s.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                if (ParamValue.Length != 2)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, TaskRunnerServiceRes.Parameters);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", TaskRunnerServiceRes.Parameters);
                    return;
                }

                if (ParamValue[0] == "ProvinceTVItemID")
                {
                    ProvinceTVItemID = int.Parse(ParamValue[1]);
                }
                else
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, ParamValue[0]);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", ParamValue[0].ToString());
                    return;
                }
            }

            if (tvItemModelProvince.TVItemID != ProvinceTVItemID)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._NotEqualTo_, "tvItemModelProvince.TVItemID[" + tvItemModelProvince.TVItemID.ToString() + "]", "ProvinceTVItemID[" + ProvinceTVItemID.ToString() + "]");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("_NotEqualTo_", "tvItemModelProvince.TVItemID[" + tvItemModelProvince.TVItemID.ToString() + "]", "ProvinceTVItemID[" + ProvinceTVItemID.ToString() + "]");
                return;
            }

            List<TVItemModel> tvItemModelSubsectorList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelProvince.TVItemID, TVTypeEnum.Subsector).OrderBy(c => c.TVText).ToList();
            if (tvItemModelSubsectorList.Count == 0)
            {
                return;
            }

            StringBuilder sb = new StringBuilder();

            int CountSS = 0;
            foreach (TVItemModel tvItemModel in tvItemModelSubsectorList)
            {
                CountSS += 1;
                if (CountSS % 10 == 0)
                {
                    appTaskModel.PercentCompleted = (100 * CountSS) / tvItemModelSubsectorList.Count;
                    appTaskService.PostUpdateAppTask(appTaskModel);
                }

                List<MWQMRunModel> mwqmRunModelList = mwqmRunService.GetMWQMRunModelListWithSubsectorTVItemIDDB(tvItemModel.TVItemID);

                foreach (MWQMRunModel mwqmRunModel in mwqmRunModelList)
                {
                    string RainMissing = "";
                    if (mwqmRunModel.RainDay0_mm == null)
                    {
                        RainMissing = RainMissing + " RainDay0,";
                    }
                    if (mwqmRunModel.RainDay1_mm == null)
                    {
                        RainMissing = RainMissing + " RainDay1,";
                    }
                    if (mwqmRunModel.RainDay2_mm == null)
                    {
                        RainMissing = RainMissing + " RainDay2,";
                    }
                    if (mwqmRunModel.RainDay3_mm == null)
                    {
                        RainMissing = RainMissing + " RainDay3,";
                    }
                    if (mwqmRunModel.RainDay4_mm == null)
                    {
                        RainMissing = RainMissing + " RainDay4,";
                    }
                    if (mwqmRunModel.RainDay5_mm == null)
                    {
                        RainMissing = RainMissing + " RainDay5,";
                    }
                    if (mwqmRunModel.RainDay6_mm == null)
                    {
                        RainMissing = RainMissing + " RainDay6,";
                    }
                    if (mwqmRunModel.RainDay7_mm == null)
                    {
                        RainMissing = RainMissing + " RainDay7,";
                    }
                    if (mwqmRunModel.RainDay8_mm == null)
                    {
                        RainMissing = RainMissing + " RainDay8,";
                    }
                    if (mwqmRunModel.RainDay9_mm == null)
                    {
                        RainMissing = RainMissing + " RainDay9,";
                    }
                    if (mwqmRunModel.RainDay10_mm == null)
                    {
                        RainMissing = RainMissing + " RainDay10,";
                    }

                    if (!string.IsNullOrWhiteSpace(RainMissing))
                    {
                        sb.AppendLine(tvItemModel.TVText.Replace(",", "_") + "," + mwqmRunModel.MWQMRunTVText.Replace(",", "_") + "," + RainMissing);
                    }
                }
            }

            string ServerPath = tvFileService.GetServerFilePath(ProvinceTVItemID);
            if (!ServerPath.EndsWith(@"\"))
            {
                ServerPath = ServerPath + @"\";
            }

            string DateText = "_" + DateTime.Now.Year.ToString() + "_" + DateTime.Now.Month.ToString() + "_" + DateTime.Now.Day.ToString() + "_" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString();

            FileInfo fi = new FileInfo(ServerPath + $@"MissingPrecip{DateText}.csv");

            DirectoryInfo di = new DirectoryInfo(fi.DirectoryName);

            if (!di.Exists)
            {
                try
                {
                    di.Create();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateDirectory__, di.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateDirectory__", di.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return;
                }
            }

            StreamWriter sw = fi.CreateText();
            sw.Write(sb.ToString());
            sw.Flush();
            sw.Close();

            fi = new FileInfo(fi.FullName);

            if (!fi.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fi.FullName);
                return;
            }

            TVItemModel tvItemModelRet = tvItemService.PostAddChildTVItemDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi.Name.Replace(fi.Extension, ""), TVTypeEnum.File);
            if (!string.IsNullOrWhiteSpace(tvItemModelRet.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVItem, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVItem, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                return;
            }

            TVFileModel tvFileModelNew = new TVFileModel();
            tvFileModelNew.TVFileTVItemID = tvItemModelRet.TVItemID;
            tvFileModelNew.TemplateTVType = 0;
            tvFileModelNew.ReportTypeID = null;
            tvFileModelNew.Parameters = Parameters;
            tvFileModelNew.ServerFileName = fi.Name;
            tvFileModelNew.FilePurpose = FilePurposeEnum.Information;
            tvFileModelNew.Language = _TaskRunnerBaseService._BWObj.appTaskModel.Language;
            tvFileModelNew.Year = DateTime.Now.Year;
            tvFileModelNew.FileDescription = "Missing Pecipitation";
            tvFileModelNew.FileType = tvFileService.GetFileType(fi.Extension);
            tvFileModelNew.FileSize_kb = (((int)fi.Length / 1024) == 0 ? 1 : (int)fi.Length / 1024);
            tvFileModelNew.FileInfo = TaskRunnerServiceRes.FileName + "[" + fi.Name + "]\r\n" + TaskRunnerServiceRes.FileType + "[" + fi.Extension + "]\r\n";
            tvFileModelNew.FileCreatedDate_UTC = fi.LastWriteTimeUtc;
            tvFileModelNew.ServerFilePath = (fi.DirectoryName + @"\").Replace(@"C:\", @"E:\");
            tvFileModelNew.LastUpdateDate_UTC = DateTime.UtcNow;
            tvFileModelNew.LastUpdateContactTVItemID = _TaskRunnerBaseService._BWObj.appTaskModel.LastUpdateContactTVItemID;

            TVFileModel tvFileModelRet = tvFileService.PostAddTVFileDB(tvFileModelNew);
            if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                return;
            }

            appTaskModel.PercentCompleted = 100;
            appTaskService.PostUpdateAppTask(appTaskModel);
        }
        public void FillRunPrecipByClimateSitePriorityForYear()
        {
            string NotUsed = "";

            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MWQMSubsectorService mwqmSubsectorService = new MWQMSubsectorService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

            if (_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._Required, TaskRunnerServiceRes.TVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_Required", TaskRunnerServiceRes.TVItemID);
                return;
            }
            if (_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID2 == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._Required, TaskRunnerServiceRes.TVItemID2);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_Required", TaskRunnerServiceRes.TVItemID2);
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

            string Parameters = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters;
            string[] ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            int ProvinceTVItemID = 0;
            int Year = 0;
            foreach (string s in ParamValueList)
            {
                string[] ParamValue = s.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                if (ParamValue.Length != 2)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, TaskRunnerServiceRes.Parameters);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", TaskRunnerServiceRes.Parameters);
                    return;
                }

                if (ParamValue[0] == "ProvinceTVItemID")
                {
                    ProvinceTVItemID = int.Parse(ParamValue[1]);
                }
                else if (ParamValue[0] == "Year")
                {
                    Year = int.Parse(ParamValue[1]);
                }
                else
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, ParamValue[0]);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", ParamValue[0].ToString());
                    return;
                }
            }

            if (tvItemModelProvince.TVItemID != ProvinceTVItemID)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._NotEqualTo_, "tvItemModelProvince.TVItemID[" + tvItemModelProvince.TVItemID.ToString() + "]", "ProvinceTVItemID[" + ProvinceTVItemID.ToString() + "]");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("_NotEqualTo_", "tvItemModelProvince.TVItemID[" + tvItemModelProvince.TVItemID.ToString() + "]", "ProvinceTVItemID[" + ProvinceTVItemID.ToString() + "]");
                return;
            }

            List<TVItemModel> tvItemModelSubsectorList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(ProvinceTVItemID, TVTypeEnum.Subsector).OrderBy(c => c.TVText).ToList();
            if (tvItemModelSubsectorList.Count == 0)
            {
                return;
            }

            string status = appTaskModel.StatusText;

            int CountSS = 0;
            foreach (TVItemModel tvItemModel in tvItemModelSubsectorList)
            {
                CountSS += 1;
                if (CountSS % 1 == 0)
                {
                    appTaskModel.StatusText = $"{status} --- doing {tvItemModel.TVText} for year {Year}";
                    appTaskModel.PercentCompleted = (100 * CountSS) / tvItemModelSubsectorList.Count;
                    appTaskService.PostUpdateAppTask(appTaskModel);
                }

                MWQMSubsectorModel mwqmSubsectorModelRet = mwqmSubsectorService.ClimateSiteSetDataToUseByAverageOrPriorityDB(tvItemModel.TVItemID, Year, "Priority");
                if (!string.IsNullOrWhiteSpace(mwqmSubsectorModelRet.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotRun_Error_, "ClimateSiteSetDataToUseByAverageOrPriorityDB", mwqmSubsectorModelRet.Error);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotRun_Error_", "ClimateSiteSetDataToUseByAverageOrPriorityDB", mwqmSubsectorModelRet.Error);
                    return;
                }
            }

            appTaskModel.PercentCompleted = 100;
            appTaskService.PostUpdateAppTask(appTaskModel);
        }
        public void GetAllPrecipitationForYear()
        {
            string NotUsed = "";

            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MWQMSubsectorService mwqmSubsectorService = new MWQMSubsectorService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

            if (_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._Required, TaskRunnerServiceRes.TVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_Required", TaskRunnerServiceRes.TVItemID);
                return;
            }
            if (_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID2 == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._Required, TaskRunnerServiceRes.TVItemID2);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_Required", TaskRunnerServiceRes.TVItemID2);
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

            string Parameters = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters;
            string[] ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            int ProvinceTVItemID = 0;
            int Year = 0;
            foreach (string s in ParamValueList)
            {
                string[] ParamValue = s.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                if (ParamValue.Length != 2)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, TaskRunnerServiceRes.Parameters);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", TaskRunnerServiceRes.Parameters);
                    return;
                }

                if (ParamValue[0] == "ProvinceTVItemID")
                {
                    ProvinceTVItemID = int.Parse(ParamValue[1]);
                }
                else if (ParamValue[0] == "Year")
                {
                    Year = int.Parse(ParamValue[1]);
                }
                else
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, ParamValue[0]);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", ParamValue[0].ToString());
                    return;
                }
            }

            if (tvItemModelProvince.TVItemID != ProvinceTVItemID)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._NotEqualTo_, "tvItemModelProvince.TVItemID[" + tvItemModelProvince.TVItemID.ToString() + "]", "ProvinceTVItemID[" + ProvinceTVItemID.ToString() + "]");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("_NotEqualTo_", "tvItemModelProvince.TVItemID[" + tvItemModelProvince.TVItemID.ToString() + "]", "ProvinceTVItemID[" + ProvinceTVItemID.ToString() + "]");
                return;
            }

            List<TVItemModel> tvItemModelSubsectorList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(ProvinceTVItemID, TVTypeEnum.Subsector).OrderBy(c => c.TVText).ToList();
            if (tvItemModelSubsectorList.Count == 0)
            {
                return;
            }

            int CountSS = 0;
            string Status = appTaskModel.StatusText;
            foreach (TVItemModel tvItemModel in tvItemModelSubsectorList)
            {
                CountSS += 1;
                if (CountSS % 1 == 0)
                {
                    appTaskModel.PercentCompleted = (100 * CountSS) / tvItemModelSubsectorList.Count;
                    appTaskModel.StatusText = Status + " --- " + tvItemModel.TVText;
                    appTaskService.PostUpdateAppTask(appTaskModel);
                }

                List<MWQMRun> mwqmRunList = new List<MWQMRun>();
                using (CSSPWebToolsDBEntities db = new CSSPWebToolsDBEntities())
                {
                    mwqmRunList = (from c in db.MWQMRuns
                                   where c.SubsectorTVItemID == tvItemModel.TVItemID
                                   select c).ToList();
                }

                if (mwqmRunList.Count == 0)
                {
                    continue;
                }

                GetClimateSitesDataForSubsectorRunsOfYear(tvItemModel.TVItemID, Year);
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                {
                    return;
                }
            }

            appTaskModel.PercentCompleted = 100;
            appTaskService.PostUpdateAppTask(appTaskModel);
        }
        public void GetClimateSitesDataForRunsOfYear()
        {
            string NotUsed = "";

            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

            if (_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._Required, TaskRunnerServiceRes.TVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_Required", TaskRunnerServiceRes.TVItemID);
                return;
            }

            if (_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID2 == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._Required, TaskRunnerServiceRes.TVItemID2);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_Required", TaskRunnerServiceRes.TVItemID2);
                return;
            }

            TVItemModel tvItemModelSubsector = tvItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelSubsector.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                return;
            }

            if (tvItemModelSubsector.TVType != TVTypeEnum.Subsector)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.TVTypeShouldBe_, TVTypeEnum.Subsector.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("TVTypeShouldBe_", TVTypeEnum.Subsector.ToString());
                return;
            }

            string Parameters = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters;
            string[] ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            int SubsectorTVItemID = 0;
            int Year = 0;
            foreach (string s in ParamValueList)
            {
                string[] ParamValue = s.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                if (ParamValue.Length != 2)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, TaskRunnerServiceRes.Parameters);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", TaskRunnerServiceRes.Parameters);
                    return;
                }

                if (ParamValue[0] == "SubsectorTVItemID")
                {
                    SubsectorTVItemID = int.Parse(ParamValue[1]);
                }
                else if (ParamValue[0] == "Year")
                {
                    Year = int.Parse(ParamValue[1]);
                }
                else
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, ParamValue[0]);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", ParamValue[0].ToString());
                    return;
                }
            }

            if (tvItemModelSubsector.TVItemID != SubsectorTVItemID)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._NotEqualTo_, "tvItemModelSubsector.TVItemID[" + tvItemModelSubsector.TVItemID.ToString() + "]", "SubsectorTVItemID[" + SubsectorTVItemID.ToString() + "]");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("_NotEqualTo_", "tvItemModelSubsector.TVItemID[" + tvItemModelSubsector.TVItemID.ToString() + "]", "SubsectorTVItemID[" + SubsectorTVItemID.ToString() + "]");
                return;
            }
     
            GetClimateSitesDataForSubsectorRunsOfYear(SubsectorTVItemID, Year);

            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
            {
                return;
            }

            // should get all the runs for the particular year and subsector

            appTaskModel.PercentCompleted = 100;
            appTaskService.PostUpdateAppTask(appTaskModel);
        }
        public void GetClimateSitesDataForSubsectorRunsOfYear(int SubsectorTVItemID, int Year)
        {
            string NotUsed = "";
            int CurrentYear = DateTime.Now.Year;

            ClimateDataValueService climateDataValueService = new ClimateDataValueService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            ClimateSiteService climateSiteService = new ClimateSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MWQMRunService mwqmRunService = new MWQMRunService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            UseOfSiteService useOfSiteService = new UseOfSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

            List<MWQMRunModel> mwqmRunModelList = mwqmRunService.GetMWQMRunModelListWithSubsectorTVItemIDDB(SubsectorTVItemID).Where(c => c.DateTime_Local.Year == Year).OrderBy(c => c.DateTime_Local).ToList();

            // need to get all climate sites for this particular subsector and run
            List<UseOfSiteModel> useOfSiteModelList = useOfSiteService.GetUseOfSiteModelListWithSiteTypeAndSubsectorTVItemIDDB(SiteTypeEnum.Climate, SubsectorTVItemID);
            List<int> ClimateSiteTVItemID = new List<int>();

            //appTaskModel.PercentCompleted = 5;
            //appTaskService.PostUpdateAppTask(appTaskModel);

            int Count = 0;
            int TotalCount = mwqmRunModelList.Count() * useOfSiteModelList.Count();
            foreach (UseOfSiteModel useOfSiteModel in useOfSiteModelList)
            {

                int EndYear = (useOfSiteModel.EndYear == null ? CurrentYear : (int)useOfSiteModel.EndYear);
                if (Year >= useOfSiteModel.StartYear && Year <= EndYear)
                {
                    ClimateSiteModel climateSiteModel = climateSiteService.GetClimateSiteModelWithClimateSiteTVItemIDDB(useOfSiteModel.SiteTVItemID);
                    if (!string.IsNullOrWhiteSpace(climateSiteModel.Error))
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.ClimateSite, TaskRunnerServiceRes.ClimateSiteTVItemID, useOfSiteModel.SiteTVItemID.ToString());
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.ClimateSite, TaskRunnerServiceRes.ClimateSiteTVItemID, useOfSiteModel.SiteTVItemID.ToString());
                        return;
                    }

                    string httpStrDaily = "";

                    using (WebClient webClient = new WebClient())
                    {
                        WebProxy webProxy = new WebProxy();
                        webClient.Proxy = webProxy;
                        string url = string.Format(UrlToGetClimateSiteDataForRunsOfYear, climateSiteModel.ECDBID, Year);
                        httpStrDaily = webClient.DownloadString(new Uri(url));
                        if (httpStrDaily.Length > 0)
                        {
                            if (httpStrDaily.Substring(0, "\"Station Name".Length) == "\"Station Name")
                            {
                                httpStrDaily = httpStrDaily.Replace("\"", "").Replace("\n", "\r\n");
                            }
                        }
                        else
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotReadFile_, url);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotReadFile_", url);
                            return;
                        }
                    }

                    foreach (MWQMRunModel mwqmRunModel in mwqmRunModelList)
                    {
                        DateTime RunDate = new DateTime(mwqmRunModel.DateTime_Local.Year, mwqmRunModel.DateTime_Local.Month, mwqmRunModel.DateTime_Local.Day);
                        DateTime RunDateMinus10 = RunDate.AddDays(-10);

                        DateTime ClimateStartDate = new DateTime(climateSiteModel.DailyStartDate_Local.Value.Year, climateSiteModel.DailyStartDate_Local.Value.Month, climateSiteModel.DailyStartDate_Local.Value.Day);
                        DateTime ClimateEndDate = new DateTime(climateSiteModel.DailyEndDate_Local.Value.Year, climateSiteModel.DailyEndDate_Local.Value.Month, climateSiteModel.DailyEndDate_Local.Value.Day);

                        Count += 1;

                        //appTaskModel.PercentCompleted = 100 * Count / TotalCount;
                        //appTaskService.PostUpdateAppTask(appTaskModel);

                        if (ClimateStartDate <= RunDate && ClimateEndDate >= RunDate)
                        {
                            UpdateDailyValuesForClimateSiteTVItemID(climateSiteModel, httpStrDaily, RunDateMinus10, RunDate, new List<DateTime>() { RunDate });
                            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                            {
                                return;
                            }
                        }

                        if (RunDate.Year != RunDateMinus10.Year)
                        {
                            string httpStrDaily2 = "";

                            using (WebClient webClient2 = new WebClient())
                            {
                                WebProxy webProxy2 = new WebProxy();
                                webClient2.Proxy = webProxy2;
                                string url2 = string.Format(UrlToGetClimateSiteDataForRunsOfYear, climateSiteModel.ECDBID, RunDateMinus10.Year);
                                httpStrDaily2 = webClient2.DownloadString(new Uri(url2));
                                if (httpStrDaily2.Length > 0)
                                {
                                    if (httpStrDaily2.Substring(0, "\"Station Name".Length) == "\"Station Name")
                                    {
                                        httpStrDaily2 = httpStrDaily2.Replace("\"", "").Replace("\n", "\r\n");
                                    }
                                }
                                else
                                {
                                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotReadFile_, url2);
                                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotReadFile_", url2);
                                    return;
                                }
                            }

                            UpdateDailyValuesForClimateSiteTVItemID(climateSiteModel, httpStrDaily2, RunDateMinus10, RunDate, new List<DateTime>() { RunDateMinus10 });
                            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                            {
                                return;
                            }

                        }
                    }
                }
            }
        }
        public void UpdateClimateSitesInformationForProvinceTVItemID()
        {
            string NotUsed = "";

            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            ClimateSiteService climateSiteService = new ClimateSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

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

            string Province = "";
            for (int i = 0; i < ProvList.Count; i++)
            {
                Application.DoEvents();

                if (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr)
                {
                    if (ProvNameListFR[i] == tvItemModelProvince.TVText)
                    {
                        Province = ProvList[i];
                        break;
                    }
                }
                else
                {
                    if (ProvNameListEN[i] == tvItemModelProvince.TVText)
                    {
                        Province = ProvList[i];
                        break;
                    }
                }
            }

            if (string.IsNullOrWhiteSpace(Province))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.Province_NotFound, tvItemModelProvince.TVText);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Province_NotFound", tvItemModelProvince.TVText);
                return;
            }

            for (int i = 1; i < 10000; i += 100)
            {
                while (!WebBrowserContentAnalysed)
                {
                    Application.DoEvents();
                }
                WebBrowserContentAnalysed = false;
                RunBrowserThread(string.Format(URLUpdateClimateSiteInfoHourlyDailyMonthly, Province, i));
                while (!WebBrowserContentAnalysed)
                {
                    Application.DoEvents();
                }

                if (AllDone)
                {
                    break;
                }
            }

            TVItemModel tvItemModelProv = tvItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelProv.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString()));
                return;
            }

            List<TVItemModel> tvItemModelClimateSiteList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, TVTypeEnum.ClimateSite);

            foreach (TVItemModel tvItemModelClimateSite in tvItemModelClimateSiteList)
            {
                //if (tvItemModelClimateSite.TVItemID < 231576)
                //    continue;

                while (!WebBrowserContentAnalysed)
                {
                    Application.DoEvents();
                }
                ClimateSiteModel climateSiteModel = climateSiteService.GetClimateSiteModelWithClimateSiteTVItemIDDB(tvItemModelClimateSite.TVItemID);
                if (!string.IsNullOrWhiteSpace(tvItemModelProv.Error))
                {
                    NotUsed = tvItemModelProv.Error;
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(tvItemModelProv.Error);
                    return;
                }

                WebBrowserContentAnalysed = false;
                RunBrowserThread(string.Format(URLUpdateClimateSiteInfoLatLngWMOElevTC, climateSiteModel.ECDBID));
                while (!WebBrowserContentAnalysed)
                {
                    Application.DoEvents();
                }
            }
            return;
        }
        public void UpdateClimateSitesInformationForProvinceTVItemIDParse(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            string NotUsed = "";

            HtmlDocument doc = (sender as WebBrowser).Document;
            HtmlElementCollection htmlElementCollection = doc.GetElementsByTagName("FORM");
            if (htmlElementCollection.Count == 0)
                AllDone = true;

            int CountInterForm = 0;
            foreach (HtmlElement htmlElementForm in htmlElementCollection)
            {
                if (htmlElementForm.GetAttribute("ACTION") == "/lib/climateData/Interform.php")
                {
                    CountInterForm += 1;

                    HtmlElement htmlElementDiv = htmlElementForm.Children[0];
                    HtmlElementCollection htmlElementCollectionDiv = htmlElementDiv.Children;

                    DateTime MarkDate = new DateTime(1800, 1, 1);
                    string Prov = "";
                    string StationName = "";
                    string StationID = "";
                    bool HasHourly = false;
                    DateTime HourlyStart = MarkDate;
                    DateTime HourlyEnd = MarkDate;
                    bool HasDaily = false;
                    DateTime DailyStart = MarkDate;
                    DateTime DailyEnd = MarkDate;
                    bool HasMonthly = false;
                    DateTime MonthlyStart = MarkDate;
                    DateTime MonthlyEnd = MarkDate;
                    foreach (HtmlElement htmlElementChild1 in htmlElementCollectionDiv)
                    {
                        if (htmlElementChild1.TagName == "INPUT")
                        {
                            if (htmlElementChild1.GetAttribute("NAME") == "Prov")
                            {
                                Prov = htmlElementChild1.GetAttribute("VALUE");
                            }
                            if (htmlElementChild1.GetAttribute("NAME") == "StationID")
                            {
                                StationID = htmlElementChild1.GetAttribute("VALUE");
                            }
                            if (htmlElementChild1.GetAttribute("NAME") == "hlyRange")
                            {
                                if (htmlElementChild1.GetAttribute("VALUE").Length > 1)
                                {
                                    string Value = htmlElementChild1.GetAttribute("VALUE");
                                    HasHourly = true;
                                    HourlyStart = new DateTime(int.Parse(Value.Substring(0, 4)), int.Parse(Value.Substring(5, 2)), int.Parse(Value.Substring(8, 2)));
                                    HourlyEnd = new DateTime(int.Parse(Value.Substring(11, 4)), int.Parse(Value.Substring(16, 2)), int.Parse(Value.Substring(19, 2)));
                                }
                            }
                            if (htmlElementChild1.GetAttribute("NAME") == "dlyRange")
                            {
                                if (htmlElementChild1.GetAttribute("VALUE").Length > 1)
                                {
                                    string Value = htmlElementChild1.GetAttribute("VALUE");
                                    HasDaily = true;
                                    DailyStart = new DateTime(int.Parse(Value.Substring(0, 4)), int.Parse(Value.Substring(5, 2)), int.Parse(Value.Substring(8, 2)));
                                    DailyEnd = new DateTime(int.Parse(Value.Substring(11, 4)), int.Parse(Value.Substring(16, 2)), int.Parse(Value.Substring(19, 2)));
                                }
                            }
                            if (htmlElementChild1.GetAttribute("NAME") == "mlyRange")
                            {
                                if (htmlElementChild1.GetAttribute("VALUE").Length > 1)
                                {
                                    string Value = htmlElementChild1.GetAttribute("VALUE");
                                    HasMonthly = true;
                                    MonthlyStart = new DateTime(int.Parse(Value.Substring(0, 4)), int.Parse(Value.Substring(5, 2)), int.Parse(Value.Substring(8, 2)));
                                    MonthlyEnd = new DateTime(int.Parse(Value.Substring(11, 4)), int.Parse(Value.Substring(16, 2)), int.Parse(Value.Substring(19, 2)));
                                }
                            }
                        }
                        else if (htmlElementChild1.TagName == "DIV")
                        {
                            HtmlElement htmlElementStationName = htmlElementChild1.Children[0].Children[0];
                            StationName = htmlElementStationName.InnerText.Trim();
                        }
                        else
                        {
                        }
                    }

                    if (string.IsNullOrWhiteSpace(StationName))
                    {
                        AllDone = true;
                        NotUsed = string.Format(TaskRunnerServiceRes._WasNotFound, TaskRunnerServiceRes.StationName);
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_WasNotFound", TaskRunnerServiceRes.StationName);
                        return;
                    }

                    if (string.IsNullOrWhiteSpace(Prov))
                    {
                        AllDone = true;
                        NotUsed = string.Format(TaskRunnerServiceRes._WasNotFound, TaskRunnerServiceRes.Prov);
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_WasNotFound", TaskRunnerServiceRes.Prov);
                        return;
                    }


                    if (string.IsNullOrWhiteSpace(StationID))
                    {
                        AllDone = true;
                        NotUsed = string.Format(TaskRunnerServiceRes._WasNotFound, TaskRunnerServiceRes.StationID);
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_WasNotFound", TaskRunnerServiceRes.StationID);
                        return;
                    }

                    if (!HasHourly && !HasDaily && !HasMonthly)
                    {
                        AllDone = true;
                        NotUsed = TaskRunnerServiceRes.HasHourlyOrHasDailyOrHasMonthlyHasToBeTrue;
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("HasHourlyOrHasDailyOrHasMonthlyHasToBeTrue");
                        return;
                    }

                    // everything OK
                    // update DB
                    ClimateSiteService climateSiteService = new ClimateSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

                    ClimateSiteModel climateSiteModelNew = new ClimateSiteModel()
                    {
                        ClimateSiteName = StationName,
                        ECDBID = int.Parse(StationID),
                        Province = Prov,
                    };

                    if ((new List<string>() { "NB", "NS", "PE" }).Contains(climateSiteModelNew.Province))
                    {
                        climateSiteModelNew.TimeOffset_hour = -4;
                    }
                    else if ((new List<string>() { "NL" }).Contains(climateSiteModelNew.Province))
                    {
                        climateSiteModelNew.TimeOffset_hour = -3.5;
                    }
                    else if ((new List<string>() { "QC" }).Contains(climateSiteModelNew.Province))
                    {
                        climateSiteModelNew.TimeOffset_hour = -5;
                    }
                    else if ((new List<string>() { "BC" }).Contains(climateSiteModelNew.Province))
                    {
                        climateSiteModelNew.TimeOffset_hour = -8;
                    }
                    else
                    {
                        climateSiteModelNew.TimeOffset_hour = -99;
                    }

                    bool IsNew = false;
                    ClimateSiteModel climateSiteModel = climateSiteService.GetClimateSiteModelExistDB(climateSiteModelNew);
                    if (!string.IsNullOrWhiteSpace(climateSiteModel.Error))
                    {
                        IsNew = true;
                        climateSiteModel = new ClimateSiteModel();
                    }

                    climateSiteModel.ClimateSiteName = climateSiteModelNew.ClimateSiteName;
                    climateSiteModel.ECDBID = climateSiteModelNew.ECDBID;
                    climateSiteModel.Province = climateSiteModelNew.Province;
                    climateSiteModel.TimeOffset_hour = climateSiteModelNew.TimeOffset_hour;

                    // hourly
                    climateSiteModel.HourlyNow = (HourlyEnd > DateTime.Today.AddDays(-5) ? true : false);
                    if (HourlyStart == MarkDate)
                        climateSiteModel.HourlyStartDate_Local = null;
                    else
                        climateSiteModel.HourlyStartDate_Local = HourlyStart;

                    if (HourlyEnd == MarkDate)
                        climateSiteModel.HourlyEndDate_Local = null;
                    else
                        climateSiteModel.HourlyEndDate_Local = HourlyEnd;

                    // daily
                    climateSiteModel.DailyNow = (DailyEnd > DateTime.Today.AddDays(-5) ? true : false);
                    if (DailyStart == MarkDate)
                        climateSiteModel.DailyStartDate_Local = null;
                    else
                        climateSiteModel.DailyStartDate_Local = DailyStart;

                    if (DailyEnd == MarkDate)
                        climateSiteModel.DailyEndDate_Local = null;
                    else
                        climateSiteModel.DailyEndDate_Local = DailyEnd;

                    // monthly
                    climateSiteModel.MonthlyNow = (MonthlyEnd > DateTime.Today.AddDays(-5) ? true : false);
                    if (MonthlyStart == MarkDate)
                        climateSiteModel.MonthlyStartDate_Local = null;
                    else
                        climateSiteModel.MonthlyStartDate_Local = MonthlyStart;

                    if (MonthlyEnd == MarkDate)
                        climateSiteModel.MonthlyEndDate_Local = null;
                    else
                        climateSiteModel.MonthlyEndDate_Local = MonthlyEnd;

                    if (IsNew)
                    {
                        string FullProvName = "";
                        for (int i = 0; i < ProvList.Count; i++)
                        {
                            if (ProvList[i] == Prov)
                            {
                                FullProvName = ProvNameListEN[i];
                                break;
                            }
                        }

                        TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

                        TVItemModel tvItemModelRoot = tvItemService.GetRootTVItemModelDB();
                        if (!string.IsNullOrWhiteSpace(tvItemModelRoot.Error))
                        {
                            AllDone = true;
                            NotUsed = TaskRunnerServiceRes.CouldNotFindTVItemRoot;
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("CouldNotFindTVItemRoot");
                            return;
                        }

                        TVItemModel tvItemModelProv = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, FullProvName, TVTypeEnum.Province);
                        if (!string.IsNullOrWhiteSpace(tvItemModelProv.Error))
                        {
                            AllDone = true;
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_StartingWith_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVText, FullProvName);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_StartingWith_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVText, FullProvName);
                            return;
                        }

                        if (climateSiteModel.DailyStartDate_Local != null && climateSiteModel.DailyEndDate_Local != null)
                        {
                            if (climateSiteModel.DailyEndDate_Local.Value.Year > 1979)
                            {
                                int CountTVItem = 0;
                                string TVText = StationName;
                                TVItemModel tvItemModelClimate = new TVItemModel() { Error = "Empty" };
                                while (!string.IsNullOrWhiteSpace(tvItemModelClimate.Error))
                                {
                                    tvItemModelClimate = tvItemService.PostAddChildTVItemDB(tvItemModelProv.TVItemID, TVText, TVTypeEnum.ClimateSite);
                                    CountTVItem += 1;
                                    TVText = StationName + "(" + CountTVItem.ToString() + ")";
                                }

                                climateSiteModel.ClimateSiteTVItemID = tvItemModelClimate.TVItemID;
                                climateSiteModel.ClimateSiteName = climateSiteModelNew.ClimateSiteName;
                                climateSiteModel.ECDBID = climateSiteModelNew.ECDBID;
                                climateSiteModel.Province = climateSiteModelNew.Province;

                                ClimateSiteModel climateSiteModelRet = climateSiteService.PostAddClimateSiteDB(climateSiteModel);
                                if (!string.IsNullOrWhiteSpace(climateSiteModel.Error))
                                {
                                    AllDone = true;
                                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.ClimateSite, climateSiteModel.Error);
                                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.ClimateSite, climateSiteModel.Error);
                                    return;
                                }

                            }
                            else
                            {
                                ClimateSiteModel climateSiteModelRet = climateSiteService.PostUpdateClimateSiteDB(climateSiteModel);
                                if (!string.IsNullOrWhiteSpace(climateSiteModel.Error))
                                {
                                    AllDone = true;
                                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.ClimateSite, climateSiteModel.Error);
                                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.ClimateSite, climateSiteModel.Error);
                                    return;
                                }
                            }
                        }
                    }
                }

            }

            if (CountInterForm == 0)
                AllDone = true;

            WebBrowserContentAnalysed = true;
        }
        public void UpdateClimateSiteInfoLatLngWMOElevTCParse(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            string NotUsed = "";
            string url = (sender as WebBrowser).Url.ToString();

            HtmlDocument doc = (sender as WebBrowser).Document;
            HtmlElement htmlElementMainIn = doc.GetElementById("wb-main-in");
            if (htmlElementMainIn == null)
            {
                AllDone = true;
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, "wb-main-in");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", "wb-main-in");
                return;
            }

            string ClimateSiteNameStr = "";
            string ClimateSiteName = "";
            string LatStr = "";
            float? Lat = 0.0f;
            string LngStr = "";
            float? Lng = 0.0f;
            string ElevStr = "";
            float? Elev = 0.0f;
            string ClimateIDStr = "";
            string ClimateID = "";
            string WMOIDStr = "";
            int? WMOID = null;
            string TCIDStr = "";
            string TCID = "";

            if (!FindHtmlElement(htmlElementMainIn, 3, "DIV", "wb-main-in child [3]"))
                return;

            if (!FindHtmlElement(htmlElementMainIn.Children[2], 1, "TABLE", "wb-main-in child [3, 1]"))
                return;

            if (!FindHtmlElement(htmlElementMainIn.Children[2].Children[0], 2, "TBODY", "wb-main-in child [3, 1, 2]"))
                return;

            // Doing first row
            if (!FindHtmlElement(htmlElementMainIn.Children[2].Children[0].Children[1], 1, "TR", "wb-main-in child [3, 1, 2, 1]"))
                return;

            // Getting TH
            if (!FindHtmlElement(htmlElementMainIn.Children[2].Children[0].Children[1].Children[0], 1, "TH", "wb-main-in child [3, 1, 2, 1, 1]"))
                return;

            // Getting B
            if (!FindHtmlElement(htmlElementMainIn.Children[2].Children[0].Children[1].Children[0].Children[0], 1, "B", "wb-main-in child [3, 1, 2, 1, 1, 1]"))
                return;

            ClimateSiteNameStr = htmlElementMainIn.Children[2].Children[0].Children[1].Children[0].Children[0].Children[0].InnerHtml;
            //  ALBANY<br>PRINCE EDWARD ISLAND
            if (ClimateSiteNameStr != null)
            {
                ClimateSiteNameStr = ClimateSiteNameStr.Trim();
                ClimateSiteName = ClimateSiteNameStr.Substring(0, ClimateSiteNameStr.IndexOf("<"));
            }
            else
            {
                ClimateSiteName = null;
            }

            // Doing second row
            if (!FindHtmlElement(htmlElementMainIn.Children[2].Children[0].Children[1], 2, "TR", "wb-main-in child [3, 1, 2, 2]"))
                return;


            // Getting Lat
            if (!FindHtmlElement(htmlElementMainIn.Children[2].Children[0].Children[1].Children[1], 1, "TH", "wb-main-in child [3, 1, 2, 2, 1]"))
                return;

            if (!FindHtmlElement(htmlElementMainIn.Children[2].Children[0].Children[1].Children[1].Children[0], 1, "A", "wb-main-in child [3, 1, 2, 2, 1, 1]"))
                return;

            if (htmlElementMainIn.Children[2].Children[0].Children[1].Children[1].Children[0].Children[0].GetAttribute("HREF") == "http://climate.weather.gc.ca/glossary_e.html#latitude")
            {
                if (!FindHtmlElement(htmlElementMainIn.Children[2].Children[0].Children[1].Children[1], 2, "TD", "wb-main-in child [3, 1, 2, 2, 2]"))
                    return;

                LatStr = htmlElementMainIn.Children[2].Children[0].Children[1].Children[1].Children[1].InnerText;
                // 46°16'00.000" N
                if (LatStr != null)
                {
                    float Deg = float.Parse(LatStr.Substring(0, LatStr.IndexOf("°")));
                    float Min = float.Parse(LatStr.Substring(LatStr.IndexOf("°") + 1, LatStr.IndexOf("'") - LatStr.IndexOf("°") - 1));
                    float Sec = float.Parse(LatStr.Substring(LatStr.IndexOf("'") + 1, LatStr.IndexOf("\"") - LatStr.IndexOf("'") - 1));
                    Lat = Deg + Min / 60 + Sec / 3600;
                    Lat = Lat * -1;
                }
                else
                {
                    Lat = null;
                }
            }


            // Getting Lng
            if (!FindHtmlElement(htmlElementMainIn.Children[2].Children[0].Children[1].Children[1], 3, "TH", "wb-main-in child [3, 1, 2, 2, 3]"))
                return;

            if (!FindHtmlElement(htmlElementMainIn.Children[2].Children[0].Children[1].Children[1].Children[2], 1, "A", "wb-main-in child [3, 1, 2, 2, 3, 1]"))
                return;

            if (htmlElementMainIn.Children[2].Children[0].Children[1].Children[1].Children[2].Children[0].GetAttribute("HREF") == "http://climate.weather.gc.ca/glossary_e.html#longitude")
            {
                if (!FindHtmlElement(htmlElementMainIn.Children[2].Children[0].Children[1].Children[1], 4, "TD", "wb-main-in child [3, 1, 2, 2, 4]"))
                    return;

                LngStr = htmlElementMainIn.Children[2].Children[0].Children[1].Children[1].Children[3].InnerText;
                //63°35'00.000" W
                if (LngStr != null)
                {
                    float Deg = float.Parse(LngStr.Substring(0, LngStr.IndexOf("°")));
                    float Min = float.Parse(LngStr.Substring(LngStr.IndexOf("°") + 1, LngStr.IndexOf("'") - LngStr.IndexOf("°") - 1));
                    float Sec = float.Parse(LngStr.Substring(LngStr.IndexOf("'") + 1, LngStr.IndexOf("\"") - LngStr.IndexOf("'") - 1));
                    Lng = Deg + Min / 60 + Sec / 3600;
                    Lng = Lng * -1;
                }
                else
                {
                    Lng = null;
                }
            }

            // Getting Elevation
            if (!FindHtmlElement(htmlElementMainIn.Children[2].Children[0].Children[1].Children[1], 5, "TH", "wb-main-in child [3, 1, 2, 2, 5]"))
                return;

            if (!FindHtmlElement(htmlElementMainIn.Children[2].Children[0].Children[1].Children[1].Children[4], 1, "A", "wb-main-in child [3, 1, 2, 2, 3, 1]"))
                return;

            if (htmlElementMainIn.Children[2].Children[0].Children[1].Children[1].Children[4].Children[0].GetAttribute("HREF") == "http://climate.weather.gc.ca/glossary_e.html#elevation")
            {
                if (!FindHtmlElement(htmlElementMainIn.Children[2].Children[0].Children[1].Children[1], 6, "TD", "wb-main-in child [3, 1, 2, 2, 6]"))
                    return;

                ElevStr = htmlElementMainIn.Children[2].Children[0].Children[1].Children[1].Children[5].InnerText;
                if (ElevStr != null)
                {
                    if (!string.IsNullOrWhiteSpace(ElevStr))
                    {
                        Elev = float.Parse(ElevStr.Trim().Substring(0, ElevStr.Trim().IndexOf(" ")));
                    }
                    else
                    {
                        Elev = null;
                    }
                }
                else
                {
                    Elev = null;
                }
            }

            // Doing third row
            if (!FindHtmlElement(htmlElementMainIn.Children[2].Children[0].Children[1], 3, "TR", "wb-main-in child [3, 1, 2, 3]"))
                return;


            // Getting ClimateID
            if (!FindHtmlElement(htmlElementMainIn.Children[2].Children[0].Children[1].Children[2], 1, "TH", "wb-main-in child [3, 1, 2, 3, 1]"))
                return;

            if (!FindHtmlElement(htmlElementMainIn.Children[2].Children[0].Children[1].Children[2].Children[0], 1, "A", "wb-main-in child [3, 1, 2, 3, 1, 1]"))
                return;

            if (htmlElementMainIn.Children[2].Children[0].Children[1].Children[2].Children[0].Children[0].GetAttribute("HREF") == "http://climate.weather.gc.ca/glossary_e.html#climate_ID")
            {
                if (!FindHtmlElement(htmlElementMainIn.Children[2].Children[0].Children[1].Children[2], 2, "TD", "wb-main-in child [3, 1, 2, 3, 2]"))
                    return;

                ClimateIDStr = htmlElementMainIn.Children[2].Children[0].Children[1].Children[2].Children[1].InnerText;
                if (ClimateIDStr != null)
                {
                    ClimateID = ClimateIDStr.Trim();
                }
                else
                {
                    ClimateID = null;
                }
            }


            // Getting WMOID
            if (!FindHtmlElement(htmlElementMainIn.Children[2].Children[0].Children[1].Children[2], 3, "TH", "wb-main-in child [3, 1, 2, 3, 3]"))
                return;

            if (!FindHtmlElement(htmlElementMainIn.Children[2].Children[0].Children[1].Children[2].Children[2], 1, "A", "wb-main-in child [3, 1, 2, 3, 3, 1]"))
                return;

            if (htmlElementMainIn.Children[2].Children[0].Children[1].Children[2].Children[2].Children[0].GetAttribute("HREF") == "http://climate.weather.gc.ca/glossary_e.html#wmo_id")
            {
                if (!FindHtmlElement(htmlElementMainIn.Children[2].Children[0].Children[1].Children[2], 4, "TD", "wb-main-in child [3, 1, 2, 3, 4]"))
                    return;

                WMOIDStr = htmlElementMainIn.Children[2].Children[0].Children[1].Children[2].Children[3].InnerText;
                if (WMOIDStr != null)
                {
                    WMOID = int.Parse(WMOIDStr.Trim());
                    if (WMOID == 0)
                        WMOID = null;
                }
                else
                {
                    WMOID = null;
                }
            }

            // Getting TCID
            if (!FindHtmlElement(htmlElementMainIn.Children[2].Children[0].Children[1].Children[2], 5, "TH", "wb-main-in child [3, 1, 2, 3, 5]"))
                return;

            if (!FindHtmlElement(htmlElementMainIn.Children[2].Children[0].Children[1].Children[2].Children[4], 1, "A", "wb-main-in child [3, 1, 2, 3, 3, 1]"))
                return;

            if (htmlElementMainIn.Children[2].Children[0].Children[1].Children[2].Children[4].Children[0].GetAttribute("HREF") == "http://climate.weather.gc.ca/glossary_e.html#tc_ID")
            {
                if (!FindHtmlElement(htmlElementMainIn.Children[2].Children[0].Children[1].Children[2], 6, "TD", "wb-main-in child [3, 1, 2, 3, 6]"))
                    return;

                TCIDStr = htmlElementMainIn.Children[2].Children[0].Children[1].Children[2].Children[5].InnerText;
                if (TCIDStr != null)
                {
                    TCID = TCIDStr.Trim();
                }
                else
                {
                    TCID = null;
                }
            }

            ClimateSiteService climateSiteService = new ClimateSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            string ECDBIDStr = url.Substring(url.IndexOf("StationID=") + 10);
            ECDBIDStr = ECDBIDStr.Substring(0, ECDBIDStr.IndexOf("&"));
            int ECDBID = int.Parse(ECDBIDStr);

            ClimateSiteModel climateSiteModel = climateSiteService.GetClimateSiteModelWithECDBIDDB(ECDBID);
            if (!string.IsNullOrWhiteSpace(climateSiteModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.ClimateSite, TaskRunnerServiceRes.ECDBID, ECDBID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.ClimateSite, TaskRunnerServiceRes.ECDBID, ECDBID.ToString());
                return;
            }

            MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            List<MapInfoPointModel> mapInfoPointModelList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(climateSiteModel.ClimateSiteTVItemID, TVTypeEnum.ClimateSite, MapInfoDrawTypeEnum.Point);

            if (mapInfoPointModelList.Count == 0)
            {
                if (Lat == null || Lng == null)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.MapInfo, TaskRunnerServiceRes.LatOrLngIsNull);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.MapInfo, TaskRunnerServiceRes.LatOrLngIsNull);
                    return;
                }

                List<Coord> coordList = new List<Coord>() { new Coord() { Lat = (float)Lat, Lng = (float)Lng, Ordinal = 0 } };

                MapInfoModel mapInfoModel = mapInfoService.CreateMapInfoObjectDB(coordList, MapInfoDrawTypeEnum.Point, TVTypeEnum.ClimateSite, climateSiteModel.ClimateSiteTVItemID);
                if (!string.IsNullOrWhiteSpace(mapInfoModel.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateMapInfoObjectForDrawType_TVType_TVItemID_, MapInfoDrawTypeEnum.Point.ToString(), TVTypeEnum.ClimateSite.ToString(), climateSiteModel.ClimateSiteTVItemID.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotCreateMapInfoObjectForDrawType_TVType_TVItemID_", MapInfoDrawTypeEnum.Point.ToString(), TVTypeEnum.ClimateSite.ToString(), climateSiteModel.ClimateSiteTVItemID.ToString());
                    return;
                }
            }

            mapInfoPointModelList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(climateSiteModel.ClimateSiteTVItemID, TVTypeEnum.ClimateSite, MapInfoDrawTypeEnum.Point);
            if (mapInfoPointModelList.Count == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBeNullOrEmpty, TaskRunnerServiceRes.MapInfoPointModelList);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_ShouldNotBeNullOrEmpty", TaskRunnerServiceRes.MapInfoPointModelList);
                return;
            }

            if (mapInfoPointModelList[0].Lat != Lat || mapInfoPointModelList[0].Lng != Lng)
            {
                mapInfoPointModelList[0].Lat = (double)Lat;
                mapInfoPointModelList[0].Lng = (double)Lng;

                MapInfoPointModel mapInfoPointModel = mapInfoService._MapInfoPointService.PostUpdateMapInfoPointDB(mapInfoPointModelList[0]);
                if (!string.IsNullOrWhiteSpace(mapInfoPointModel.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.MapInfoPoint, mapInfoPointModel.Error);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.MapInfoPoint, mapInfoPointModel.Error);
                    return;
                }
            }


            climateSiteModel.ClimateSiteName = ClimateSiteName;
            climateSiteModel.Elevation_m = Elev;
            climateSiteModel.ClimateID = ClimateID;
            climateSiteModel.WMOID = WMOID;
            climateSiteModel.TCID = TCID;

            ClimateSiteModel climateSiteModelRet = climateSiteService.PostUpdateClimateSiteDB(climateSiteModel);
            if (!string.IsNullOrWhiteSpace(climateSiteModelRet.Error))
            {
                AllDone = true;
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.ClimateSite, climateSiteModelRet.Error);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.ClimateSite, climateSiteModelRet.Error);
                return;
            }

            WebBrowserContentAnalysed = true;
        }
        //public void UpdateClimateSiteDailyAndHourlyFromStartDateToEndDate()
        //{
        //    string NotUsed = "";

        //    string[] ParamValueList = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
        //    int ClimateSiteTVItemID = 0;
        //    int StartYear = 0;
        //    int StartMonth = 0;
        //    int StartDay = 0;
        //    int EndYear = 0;
        //    int EndMonth = 0;
        //    int EndDay = 0;

        //    if (_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID == 0)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.TVItemID);
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.TVItemID);
        //    }

        //    if (ParamValueList.Count() != 9)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.ParameterCount_NotEqual_, ParamValueList.Count().ToString(), 9);
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("ParameterCount_NotEqual_", ParamValueList.Count().ToString(), "9");
        //    }

        //    foreach (string s in ParamValueList)
        //    {
        //        string[] ParamValue = s.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

        //        if (ParamValue.Length != 2)
        //        {
        //            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, TaskRunnerServiceRes.Parameters);
        //            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", TaskRunnerServiceRes.Parameters);
        //            return;
        //        }

        //        if (ParamValue[0] == "ClimateSiteTVItemID")
        //        {
        //            ClimateSiteTVItemID = int.Parse(ParamValue[1]);
        //        }
        //        else if (ParamValue[0] == "StartYear")
        //        {
        //            StartYear = int.Parse(ParamValue[1]);
        //        }
        //        else if (ParamValue[0] == "StartMonth")
        //        {
        //            StartMonth = int.Parse(ParamValue[1]);
        //        }
        //        else if (ParamValue[0] == "StartDay")
        //        {
        //            StartDay = int.Parse(ParamValue[1]);
        //        }
        //        else if (ParamValue[0] == "EndYear")
        //        {
        //            EndYear = int.Parse(ParamValue[1]);
        //        }
        //        else if (ParamValue[0] == "EndMonth")
        //        {
        //            EndMonth = int.Parse(ParamValue[1]);
        //        }
        //        else if (ParamValue[0] == "EndDay")
        //        {
        //            EndDay = int.Parse(ParamValue[1]);
        //        }
        //        else if (ParamValue[0] == "Generate")
        //        {
        //            // nothing
        //        }
        //        else if (ParamValue[0] == "Command")
        //        {
        //            // nothing
        //        }
        //        else
        //        {
        //            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, ParamValue[0]);
        //            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", ParamValue[0].ToString());
        //            return;
        //        }
        //    }

        //    if (ClimateSiteTVItemID == 0)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, ClimateSiteTVItemID.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.ClimateSiteTVItemID);
        //        return;
        //    }

        //    if (StartYear == 0)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, StartYear.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.StartYear);
        //        return;
        //    }

        //    if (StartMonth == 0)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, StartMonth.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.StartMonth);
        //        return;
        //    }

        //    if (StartDay == 0)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, StartDay.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.StartDay);
        //        return;
        //    }

        //    if (EndYear == 0)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, EndYear.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.EndYear);
        //        return;
        //    }

        //    if (EndMonth == 0)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, EndMonth.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.EndMonth);
        //        return;
        //    }

        //    if (EndDay == 0)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, EndDay.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.EndDay);
        //        return;
        //    }

        //    ClimateSiteService climateSiteService = new ClimateSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    ClimateSiteModel climateSiteModel = climateSiteService.GetClimateSiteModelWithClimateSiteTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //    if (!string.IsNullOrWhiteSpace(climateSiteModel.Error))
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.ClimateSite, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.ClimateSite, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        return;
        //    }

        //    DateTime StartDate = new DateTime(StartYear, StartMonth, StartDay).AddDays(-10);
        //    DateTime EndDate = new DateTime(EndYear, EndMonth, EndDay).AddDays(10);

        //    for (int year = StartDate.Year; year <= EndDate.Year; year++)
        //    {
        //        string httpStrDaily = "";

        //        using (WebClient webClient = new WebClient())
        //        {
        //            WebProxy webProxy = new WebProxy();
        //            webClient.Proxy = webProxy;
        //            string url = string.Format(UpdateClimateSiteDailyFromStartDateToEndDate, climateSiteModel.ECDBID, year);
        //            httpStrDaily = webClient.DownloadString(new Uri(url));
        //            if (httpStrDaily.Length > 0)
        //            {
        //                if (httpStrDaily.Substring(0, "\"Station Name".Length) == "\"Station Name")
        //                {
        //                    httpStrDaily = httpStrDaily.Replace("\"", "").Replace("\n", "\r\n");
        //                }
        //            }
        //            else
        //            {
        //                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotReadFile_, url);
        //                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotReadFile_", url);
        //                return;
        //            }
        //        }
        //        UpdateDailyValuesForClimateSiteTVItemID(climateSiteModel, httpStrDaily, StartDate, EndDate, new List<DateTime>());
        //        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //            return;

        //    }
        //    return;
        //}
        //public void UpdateClimateSiteDailyAndHourlyForSubsectorFromStartDateToEndDate()
        //{
        //    string NotUsed = "";

        //    ClimateDataValueService climateDataValueService = new ClimateDataValueService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    ClimateSiteService climateSiteService = new ClimateSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    UseOfSiteService useOfSiteService = new UseOfSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

        //    string[] ParamValueList = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
        //    int SubsectorTVItemID = 0;
        //    int StartYear = 0;
        //    int StartMonth = 0;
        //    int StartDay = 0;
        //    int EndYear = 0;
        //    int EndMonth = 0;
        //    int EndDay = 0;

        //    if (_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID == 0)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.TVItemID);
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.TVItemID);
        //        return;
        //    }


        //    if (ParamValueList.Count() != 9)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.ParameterCount_NotEqual_, ParamValueList.Count().ToString(), 9);
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("ParameterCount_NotEqual_", ParamValueList.Count().ToString(), "9");
        //        return;
        //    }

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
        //        else if (ParamValue[0] == "StartYear")
        //        {
        //            StartYear = int.Parse(ParamValue[1]);
        //        }
        //        else if (ParamValue[0] == "StartMonth")
        //        {
        //            StartMonth = int.Parse(ParamValue[1]);
        //        }
        //        else if (ParamValue[0] == "StartDay")
        //        {
        //            StartDay = int.Parse(ParamValue[1]);
        //        }
        //        else if (ParamValue[0] == "EndYear")
        //        {
        //            EndYear = int.Parse(ParamValue[1]);
        //        }
        //        else if (ParamValue[0] == "EndMonth")
        //        {
        //            EndMonth = int.Parse(ParamValue[1]);
        //        }
        //        else if (ParamValue[0] == "EndDay")
        //        {
        //            EndDay = int.Parse(ParamValue[1]);
        //        }
        //        else if (ParamValue[0] == "Generate")
        //        {
        //            // nothing
        //        }
        //        else if (ParamValue[0] == "Command")
        //        {
        //            // nothing
        //        }
        //        else
        //        {
        //            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, ParamValue[0]);
        //            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", ParamValue[0].ToString());
        //            return;
        //        }
        //    }

        //    if (SubsectorTVItemID == 0)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, SubsectorTVItemID.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.SubsectorTVItemID);
        //        return;
        //    }

        //    if (StartYear == 0)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, StartYear.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.StartYear);
        //        return;
        //    }

        //    if (StartMonth == 0)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, StartMonth.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.StartMonth);
        //        return;
        //    }

        //    if (StartDay == 0)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, StartDay.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.StartDay);
        //        return;
        //    }

        //    if (EndYear == 0)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, EndYear.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.EndYear);
        //        return;
        //    }

        //    if (EndMonth == 0)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, EndMonth.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.EndMonth);
        //        return;
        //    }

        //    if (EndDay == 0)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, EndDay.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.EndDay);
        //        return;
        //    }

        //    DateTime StartDate = new DateTime(StartYear, StartMonth, StartDay).AddDays(-10);
        //    DateTime EndDate = new DateTime(EndYear, EndMonth, EndDay).AddDays(10);

        //    TVItemModel tvItemModelSubsector = tvItemService.GetTVItemModelWithTVItemIDDB(SubsectorTVItemID);
        //    if (!string.IsNullOrWhiteSpace(tvItemModelSubsector.Error))
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, SubsectorTVItemID.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, SubsectorTVItemID.ToString());
        //        return;

        //    }

        //    List<MapInfoPointModel> mapInfoPointListSubsector = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.Subsector, MapInfoDrawTypeEnum.Point);
        //    if (mapInfoPointListSubsector.Count == 0)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBeNullOrEmpty, TaskRunnerServiceRes.mapInfoPointListSubsector);
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_ShouldNotBeNullOrEmpty", TaskRunnerServiceRes.mapInfoPointListSubsector);
        //        return;
        //    }

        //    float LookupRadius = 100000.0f; // 100 km radius

        //    List<MapInfoModel> mapInfoModelList = new List<MapInfoModel>();
        //    List<ClimateSiteModel> climateSiteModelListWithHourly = new List<ClimateSiteModel>();
        //    List<ClimateSiteModel> climateSiteModelListWithDaily = new List<ClimateSiteModel>();

        //    while (mapInfoModelList.Count == 0)
        //    {
        //        mapInfoModelList = mapInfoService.GetMapInfoModelWithinCircleWithTVTypeAndMapInfoDrawTypeDB((float)mapInfoPointListSubsector[0].Lat, (float)mapInfoPointListSubsector[0].Lng, LookupRadius, TVTypeEnum.ClimateSite, MapInfoDrawTypeEnum.Point);

        //        foreach (MapInfoModel mapInfoModel in mapInfoModelList)
        //        {
        //            ClimateSiteModel climateSiteModel = climateSiteService.GetClimateSiteModelWithClimateSiteTVItemIDDB(mapInfoModel.TVItemID);
        //            if (!string.IsNullOrWhiteSpace(climateSiteModel.Error))
        //            {
        //                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, mapInfoModel.TVItemID.ToString());
        //                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, mapInfoModel.TVItemID.ToString());
        //                return;
        //            }
        //            if (climateSiteModel.HourlyStartDate_Local != null)
        //            {
        //                climateSiteModelListWithHourly.Add(climateSiteModel);
        //            }
        //            if (climateSiteModel.DailyStartDate_Local != null)
        //            {
        //                climateSiteModelListWithDaily.Add(climateSiteModel);
        //            }
        //        }

        //        LookupRadius += 10000;
        //    }

        //    for (int year = StartDate.Year; year <= EndDate.Year; year++)
        //    {

        //        List<ClimateSiteModel> climateSiteForDailyModelListUsed = new List<ClimateSiteModel>();
        //        //List<ClimateSiteModel> climateSiteForHourlyModelListUsed = new List<ClimateSiteModel>();
        //        foreach (ClimateSiteModel climateSiteModel in climateSiteModelListWithDaily)
        //        {
        //            if (climateSiteForDailyModelListUsed.Count > 2)
        //                break;

        //            if (climateSiteModel.DailyStartDate_Local.Value <= StartDate && climateSiteModel.DailyEndDate_Local.Value >= EndDate)
        //            {
        //                ClimateDataValueModel climateDataValueStartDone = climateDataValueService.GetClimateDataValueModelExitDB(new ClimateDataValueModel() { ClimateSiteID = climateSiteModel.ClimateSiteID, DateTime_Local = StartDate });
        //                ClimateDataValueModel climateDataValueEndDone = climateDataValueService.GetClimateDataValueModelExitDB(new ClimateDataValueModel() { ClimateSiteID = climateSiteModel.ClimateSiteID, DateTime_Local = EndDate });
        //                if (string.IsNullOrWhiteSpace(climateDataValueStartDone.Error) && string.IsNullOrWhiteSpace(climateDataValueEndDone.Error))
        //                {
        //                    climateSiteForDailyModelListUsed.Add(climateSiteModel);
        //                    continue;
        //                }
        //                string httpStr = "";
        //                using (WebClient webClient = new WebClient())
        //                {
        //                    WebProxy webProxy = new WebProxy();
        //                    webClient.Proxy = webProxy;
        //                    string url = string.Format(UpdateClimateSiteDailyFromStartDateToEndDate, climateSiteModel.ECDBID, year);
        //                    httpStr = webClient.DownloadString(new Uri(url));
        //                    if (httpStr.Length > 0)
        //                    {
        //                        if (httpStr.Substring(0, "\"Station Name".Length) == "\"Station Name")
        //                        {
        //                            httpStr = httpStr.Replace("\"", "").Replace("\n", "\r\n");
        //                        }
        //                    }
        //                    else
        //                    {
        //                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotReadFile_, url);
        //                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotReadFile_", url);
        //                        return;
        //                    }
        //                }
        //                ClimateDataValueModel climateDataValueModel = UpdateDailyValuesForClimateSiteTVItemID(climateSiteModel, httpStr, StartDate, EndDate, new List<DateTime>());
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                if (climateDataValueModel.Rainfall_mm != null || climateDataValueModel.TotalPrecip_mm_cm != null)
        //                {
        //                    climateSiteForDailyModelListUsed.Add(climateSiteModel);
        //                }

        //                UseOfSiteModel useOfSiteModelNew = new UseOfSiteModel()
        //                {
        //                    SiteTVItemID = climateSiteModel.ClimateSiteTVItemID,
        //                    SubsectorTVItemID = tvItemModelSubsector.TVItemID,
        //                    SiteType = SiteTypeEnum.Climate,
        //                    Ordinal = 0, // will be replaced 
        //                    StartYear = year,
        //                    EndYear = year,
        //                    UseWeight = false,
        //                    Weight_perc = 1.0f,
        //                    UseEquation = false,
        //                };

        //                List<UseOfSiteModel> useOfSiteModelList = useOfSiteService.GetUseOfSiteModelListWithSiteTypeAndSubsectorTVItemIDDB(SiteTypeEnum.Climate, tvItemModelSubsector.TVItemID);
        //                bool hasUpdated = false;
        //                foreach (UseOfSiteModel useOfSiteModel in useOfSiteModelList)
        //                {
        //                    if (useOfSiteModel.EndYear == year - 1 || useOfSiteModel.EndYear == year)
        //                    {
        //                        useOfSiteModel.EndYear = year;
        //                        UseOfSiteModel useOfSiteModelRet = useOfSiteService.PostUpdateUseOfSiteDB(useOfSiteModel);
        //                        if (!string.IsNullOrWhiteSpace(useOfSiteModelRet.Error))
        //                        {
        //                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.UseOfSite, useOfSiteModelRet.Error);
        //                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.UseOfSite, useOfSiteModelRet.Error);
        //                            return;
        //                        }
        //                        hasUpdated = true;
        //                        break;
        //                    }
        //                }
        //                if (!hasUpdated)
        //                {
        //                    if (useOfSiteModelList.Count > 0)
        //                    {
        //                        useOfSiteModelNew.Ordinal = useOfSiteModelList.OrderBy(c => c.Ordinal).Last().Ordinal + 1;
        //                    }
        //                    UseOfSiteModel useOfSiteModelRet = useOfSiteService.PostAddUseOfSiteDB(useOfSiteModelNew);
        //                    if (!string.IsNullOrWhiteSpace(useOfSiteModelRet.Error))
        //                    {
        //                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.UseOfSite, useOfSiteModelRet.Error);
        //                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(useOfSiteModelRet.Error);
        //                        return;
        //                    }
        //                }

        //                //if (!string.IsNullOrWhiteSpace(climateDataValueModel.HourlyValues))
        //                //{
        //                //    climateSiteForHourlyModelListUsed.Add(climateSiteModel);
        //                //}
        //            }
        //        }

        //        //    if (climateSiteForHourlyModelListUsed.Count == 0)
        //        //    {
        //        //        foreach (ClimateSiteModel climateSiteModel in climateSiteModelListWithHourly)
        //        //        {
        //        //            if (climateSiteForHourlyModelListUsed.Count > 0)
        //        //                break;

        //        //            if (climateSiteModel.DailyStartDate_Local.Value <= StartDate && climateSiteModel.DailyEndDate_Local.Value >= EndDate)
        //        //            {
        //        //                string httpStr = "";
        //        //                using (WebClient webClient = new WebClient())
        //        //                {
        //        //                    WebProxy webProxy = new WebProxy();
        //        //                    webClient.Proxy = webProxy;
        //        //                    string url = string.Format(UpdateClimateSiteDailyFromStartDateToEndDate, climateSiteModel.ECDBID, year);
        //        //                    httpStr = webClient.DownloadString(new Uri(url));
        //        //                    if (httpStr.Length > 0)
        //        //                    {
        //        //                        if (httpStr.Substring(0, "\"Station Name".Length) == "\"Station Name")
        //        //                        {
        //        //                            httpStr = httpStr.Replace("\"", "").Replace("\n", "\r\n");
        //        //                        }
        //        //                    }
        //        //                    else
        //        //                    {
        //        //                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotReadFile_, url);
        //        //                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotReadFile_", url);
        //        //                        return;
        //        //                    }
        //        //                }
        //        //                ClimateDataValueModel climateDataValueModel = UpdateDailyValuesForClimateSiteTVItemID(climateSiteModel.ClimateSiteTVItemID, httpStr, StartDate, EndDate);
        //        //                if (!string.IsNullOrWhiteSpace(climateDataValueModel.Error))
        //        //                {
        //        //                    NotUsed = string.Format(TaskRunnerServiceRes._Required, TaskRunnerServiceRes.TVItemID);
        //        //                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.TVItemID));
        //        //                    return;
        //        //                }
        //        //                if (climateDataValueModel.Rainfall_mm != null)
        //        //                {
        //        //                    climateSiteForDailyModelListUsed.Add(climateSiteModel);
        //        //                }
        //        //                //if (climateDataValueModel.HourlyValues != null)
        //        //                //{
        //        //                //    climateSiteForHourlyModelListUsed.Add(climateSiteModel);
        //        //                //    break;
        //        //                //}
        //        //            }
        //        //        }
        //        //    }

        //        //    foreach (ClimateSiteModel climateSiteModelHourly in climateSiteForHourlyModelListUsed)
        //        //    {
        //        //        climateSiteForDailyModelListUsed.Add(climateSiteModelHourly);
        //        //    }
        //    }
        //}
        public void UpdateDailyValuesForClimateSiteTVItemID(ClimateSiteModel climateSiteModel, string httpStrDaily, DateTime StartDate, DateTime EndDate, List<DateTime> HourlyDateListToLoad)
        {
            string NotUsed = "";

            //string httpStrHourly = "";
            //ClimateDataValueModel climateDataValueModel = new ClimateDataValueModel();
            List<string> FullProvList = new List<string>() { "BRITISH COLUMBIA", "NEW BRUNSWICK", "NEWFOUNDLAND", "NOVA SCOTIA", "PRINCE EDWARD ISLAND", "QUEBEC" };
            List<string> ShortProvList = new List<string>() { "BC", "NB", "NL", "NS", "PE", "QC" };

            StringBuilder hourlyValues = new StringBuilder();

            hourlyValues.Clear();

            ClimateSiteService climateSiteService = new ClimateSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            ClimateDataValueService climateDataValueService = new ClimateDataValueService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            using (TextReader tr = new StringReader(httpStrDaily))
            {

                int countLine = 0;
                string LookupTxt = "";
                while (true)
                {
                    countLine += 1;
                    string lineStr = tr.ReadLine();
                    if (lineStr == null)
                    {
                        break;
                    }

                    List<string> lineValueArr = lineStr.Split(",".ToCharArray(), StringSplitOptions.None).ToList();

                    if (countLine == 1)
                    {
                        LookupTxt = "Station Name";
                        if (lineValueArr[0] != LookupTxt)
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_AtLine_, LookupTxt, countLine.ToString());
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
                            return;
                        }
                        LookupTxt = climateSiteModel.ClimateSiteName;
                        if (lineValueArr[1] != LookupTxt)
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_AtLine_, LookupTxt, countLine.ToString());
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
                            return;
                        }
                    }
                    if (countLine == 2)
                    {
                        LookupTxt = "Province";
                        if (lineValueArr[0] != LookupTxt)
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_AtLine_, LookupTxt, countLine.ToString());
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
                            return;
                        }
                        LookupTxt = lineValueArr[1];
                        if (!FullProvList.Contains(LookupTxt))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_AtLine_, LookupTxt, countLine.ToString());
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
                            return;
                        }
                        for (int i = 0; i < 6; i++)
                        {
                            if (lineValueArr[1] == FullProvList[i])
                            {
                                if (climateSiteModel.Province != ShortProvList[i])
                                {
                                    NotUsed = string.Format(TaskRunnerServiceRes.Province_NotEqualTo_AtLine_, ShortProvList[i], climateSiteModel.Province, countLine.ToString());
                                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
                                    return;
                                }
                            }
                        }
                    }
                    if (countLine == 6)
                    {
                        LookupTxt = "Climate Identifier";
                        if (lineValueArr[0] != LookupTxt)
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_AtLine_, LookupTxt, countLine.ToString());
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
                            return;
                        }
                        if (lineValueArr[1].Length > 0)
                        {
                            LookupTxt = climateSiteModel.ClimateID;
                            if (lineValueArr[1] != climateSiteModel.ClimateID)
                            {
                                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_AtLine_, LookupTxt, countLine.ToString());
                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
                                return;
                            }
                        }
                    }
                    if (lineValueArr[0].Contains("-"))
                    {
                        if (lineValueArr[0].Substring(4, 1) == "-")
                        {
                            if (lineValueArr.Count != 27)
                            {
                                NotUsed = string.Format(TaskRunnerServiceRes.CountOfValuesInLine_ShouldBe_, countLine, "27");
                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CountOfValuesInLine_ShouldBe_", countLine.ToString(), "27");
                                return;
                            }

                            int Year = int.Parse(lineValueArr[1]);
                            int Month = int.Parse(lineValueArr[2]);
                            int Day = int.Parse(lineValueArr[3]);

                            if (Year == 0 || Month == 0 || Day == 0)
                            {
                                NotUsed = string.Format(TaskRunnerServiceRes.YearNotCorrect_AtLine_, Year.ToString() + " - " + Month.ToString() + " - " + Day.ToString(), countLine.ToString());
                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("YearNotCorrect_AtLine_", Year.ToString() + " - " + Month.ToString() + " - " + Day.ToString(), countLine.ToString());
                                return;
                            }

                            DateTime LineDate = new DateTime(Year, Month, Day);

                            if (!(StartDate <= LineDate && EndDate >= LineDate))
                            {
                                continue;
                            }

                            string httpStrHourly = "";
                            if (HourlyDateListToLoad.Contains(LineDate))
                            {
                                using (WebClient webClient = new WebClient())
                                {
                                    WebProxy webProxy = new WebProxy();
                                    webClient.Proxy = webProxy;
                                    string url = string.Format(UpdateClimateSiteHourlyFromStartDateToEndDate, climateSiteModel.ECDBID, Year, LineDate.Month);
                                    httpStrHourly = webClient.DownloadString(new Uri(url));
                                    if (httpStrHourly.Length > 0)
                                    {
                                        if (httpStrHourly.Substring(0, "\"Station Name".Length) == "\"Station Name")
                                        {
                                            httpStrHourly = httpStrHourly.Replace("\"", "").Replace("\n", "\r\n");
                                        }
                                    }
                                    else
                                    {
                                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotReadFile_, url);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotReadFile_", url);
                                        return;
                                    }
                                }

                                if (climateSiteModel.HourlyStartDate_Local <= LineDate && climateSiteModel.HourlyEndDate_Local >= LineDate)
                                {
                                    UpdateHourlyValuesForClimateSiteAndDate(climateSiteModel, httpStrHourly, LineDate, hourlyValues);
                                    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                                        return;

                                    if (!string.IsNullOrWhiteSpace(hourlyValues.ToString()))
                                    {
                                        string Title = "Year,Month,Day,Hour,Temp_C,DewPoint_C,RelHum_Perc,WindDir_10deg,WindSpd_km_h,Visibility_km,StnPress_kPa,Humidx,WindChill\r\n";
                                        hourlyValues.Append(Title + hourlyValues.ToString());
                                    }
                                }
                            }


                            float? MaxTemp = null;
                            if (lineValueArr[5].Length > 0)
                            {
                                MaxTemp = float.Parse(lineValueArr[5]);
                            }
                            float? MinTemp = null;
                            if (lineValueArr[7].Length > 0)
                            {
                                MinTemp = float.Parse(lineValueArr[7]);
                            }
                            float? MeanTemp = null;
                            if (lineValueArr[9].Length > 0)
                            {
                                MeanTemp = float.Parse(lineValueArr[9]);
                            }
                            float? HeatDegDays = null;
                            if (lineValueArr[11].Length > 0)
                            {
                                HeatDegDays = float.Parse(lineValueArr[11]);
                            }
                            float? CoolDegDays = null;
                            if (lineValueArr[13].Length > 0)
                            {
                                CoolDegDays = float.Parse(lineValueArr[13]);
                            }
                            float? TotalRain = null;
                            if (lineValueArr[15].Length > 0)
                            {
                                if (lineValueArr[15].Trim() == "T")
                                {
                                    TotalRain = 0.0f;
                                }
                                else
                                {
                                    TotalRain = float.Parse(lineValueArr[15]);
                                }
                            }
                            float? TotalSnow = null;
                            if (lineValueArr[17].Length > 0)
                            {
                                if (lineValueArr[17].Trim() == "T")
                                {
                                    TotalSnow = 0.0f;
                                }
                                else
                                {
                                    TotalSnow = float.Parse(lineValueArr[17]);
                                }
                            }
                            float? TotalPrecip = null;
                            if (lineValueArr[19].Length > 0)
                            {
                                if (lineValueArr[19].Trim() == "T")
                                {
                                    TotalPrecip = 0.0f;
                                }
                                else
                                {
                                    TotalPrecip = float.Parse(lineValueArr[19]);
                                }
                            }
                            float? SnowOnGround = null;
                            if (lineValueArr[21].Length > 0)
                            {
                                if (lineValueArr[21].Trim() == "T")
                                {
                                    SnowOnGround = 0.0f;
                                }
                                else
                                {
                                    SnowOnGround = float.Parse(lineValueArr[21]);
                                }
                            }
                            float? DirMaxGust = null;
                            if (lineValueArr[23].Length > 0)
                            {
                                DirMaxGust = float.Parse(lineValueArr[23]);
                            }
                            float? SpdMaxGust = null;
                            if (lineValueArr[25].Length > 0)
                            {
                                if (lineValueArr[25].Substring(0, 1) == "<")
                                {
                                    SpdMaxGust = float.Parse(lineValueArr[25].Substring(1));
                                }
                                else
                                {
                                    SpdMaxGust = float.Parse(lineValueArr[25]);
                                }
                            }

                            ClimateDataValueModel climateDataValueModelNew = new ClimateDataValueModel()
                            {
                                ClimateSiteID = climateSiteModel.ClimateSiteID,
                                HasBeenRead = true,
                                CoolDegDays_C = CoolDegDays,
                                DateTime_Local = LineDate,
                                DirMaxGust_0North = DirMaxGust,
                                HeatDegDays_C = HeatDegDays,
                                HourlyValues = hourlyValues.ToString(),
                                Keep = true,
                                MaxTemp_C = MaxTemp,
                                MinTemp_C = MinTemp,
                                Rainfall_mm = TotalRain,
                                RainfallEntered_mm = null,
                                Snow_cm = TotalSnow,
                                SnowOnGround_cm = SnowOnGround,
                                SpdMaxGust_kmh = SpdMaxGust,
                                StorageDataType = StorageDataTypeEnum.Archived,
                                TotalPrecip_mm_cm = TotalPrecip,
                            };
                            ClimateDataValueModel climateDataValueModelExist = climateDataValueService.GetClimateDataValueModelExitDB(climateDataValueModelNew);
                            if (!string.IsNullOrWhiteSpace(climateDataValueModelExist.Error))
                            {
                                ClimateDataValueModel climateDataValueModelRet = climateDataValueService.PostAddClimateDataValueDB(climateDataValueModelNew);
                                if (!string.IsNullOrWhiteSpace(climateDataValueModelRet.Error))
                                {
                                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.ClimateDataValue, climateDataValueModelRet.Error);
                                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.ClimateDataValue, climateDataValueModelRet.Error);
                                    return;
                                }
                            }
                            else
                            {

                                if (climateDataValueModelExist.HasBeenRead == false)
                                {
                                    climateDataValueModelNew.ClimateDataValueID = climateDataValueModelExist.ClimateDataValueID;
                                    climateDataValueModelNew.RainfallEntered_mm = climateDataValueModelExist.RainfallEntered_mm;
                                    climateDataValueModelNew.HasBeenRead = true;
                                    climateDataValueModelExist = climateDataValueService.PostUpdateClimateDataValueDB(climateDataValueModelNew);
                                    if (!string.IsNullOrWhiteSpace(climateDataValueModelExist.Error))
                                    {
                                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.ClimateDataValue, climateDataValueModelExist.Error); ;
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.ClimateDataValue, climateDataValueModelExist.Error);
                                        return;
                                    }
                                }
                            }

                        }
                    }
                }
            }
        }
        public void UpdateHourlyValuesForClimateSiteAndDate(ClimateSiteModel climateSiteModel, string httpStrHourly, DateTime CurrentDate, StringBuilder hourlyValues)
        {
            string NotUsed = "";
            bool HasValue = false;

            //ClimateDataValueModel climateDataValueModel = new ClimateDataValueModel();
            List<string> FullProvList = new List<string>() { "BRITISH COLUMBIA", "NEW BRUNSWICK", "NEWFOUNDLAND", "NOVA SCOTIA", "PRINCE EDWARD ISLAND", "QUEBEC" };
            List<string> ShortProvList = new List<string>() { "BC", "NB", "NL", "NS", "PE", "QC" };

            hourlyValues.Clear();

            using (TextReader tr = new StringReader(httpStrHourly))
            {
                string LookupTxt = "";
                int countLine = 0;
                while (true)
                {
                    countLine += 1;
                    string lineStr = tr.ReadLine();
                    if (lineStr == null)
                        break;

                    List<string> lineValueArr = lineStr.Split(",".ToCharArray(), StringSplitOptions.None).ToList();
                    if (countLine == 1)
                    {
                        LookupTxt = "Station Name";
                        if (lineValueArr[0] != LookupTxt)
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_AtLine_, LookupTxt, countLine.ToString());
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
                            return;
                        }
                        LookupTxt = climateSiteModel.ClimateSiteName;
                        if (lineValueArr[1] != LookupTxt)
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_AtLine_, LookupTxt, countLine.ToString());
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
                            return;
                        }
                    }
                    if (countLine == 2)
                    {
                        LookupTxt = "Province";
                        if (lineValueArr[0] != LookupTxt)
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_AtLine_, LookupTxt, countLine.ToString());
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
                            return;
                        }
                        LookupTxt = lineValueArr[1];
                        if (!FullProvList.Contains(LookupTxt))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_AtLine_, LookupTxt, countLine.ToString());
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
                            return;
                        }
                        for (int i = 0; i < 6; i++)
                        {
                            if (lineValueArr[1] == FullProvList[i])
                            {
                                if (climateSiteModel.Province != ShortProvList[i])
                                {
                                    NotUsed = string.Format(TaskRunnerServiceRes.Province_NotEqualTo_AtLine_, ShortProvList[i], climateSiteModel.Province, countLine.ToString());
                                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
                                    return;
                                }
                            }
                        }
                    }
                    if (countLine == 6)
                    {
                        LookupTxt = "Climate Identifier";
                        if (lineValueArr[0] != LookupTxt)
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_AtLine_, LookupTxt, countLine.ToString());
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
                            return;
                        }
                        if (lineValueArr[1].Length > 0)
                        {
                            LookupTxt = climateSiteModel.ClimateID;
                            if (lineValueArr[1] != climateSiteModel.ClimateID)
                            {
                                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_AtLine_, LookupTxt, countLine.ToString());
                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind_AtLine_", LookupTxt, countLine.ToString());
                                return;
                            }
                        }
                    }
                    if (lineValueArr[0].Contains("-"))
                    {
                        if (lineValueArr[0].Substring(4, 1) == "-")
                        {
                            if (lineValueArr.Count != 25 && lineValueArr.Count != 26 && lineValueArr.Count != 27 && lineValueArr.Count != 28 && lineValueArr.Count != 5)
                            {
                                NotUsed = string.Format(TaskRunnerServiceRes.CountOfValuesInLine_ShouldBe_, countLine, "25");
                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CountOfValuesInLine_ShouldBe_", countLine.ToString(), "25");
                                return;
                            }

                            int Year = int.Parse(lineValueArr[1]);
                            int Month = int.Parse(lineValueArr[2]);
                            int Day = int.Parse(lineValueArr[3]);

                            if (Year == 0 || Month == 0 || Day == 0)
                            {
                                NotUsed = string.Format(TaskRunnerServiceRes.YearNotCorrect_AtLine_, Year.ToString() + " - " + Month.ToString() + " - " + Day.ToString(), countLine.ToString());
                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("YearNotCorrect_AtLine_", Year.ToString() + " - " + Month.ToString() + " - " + Day.ToString(), countLine.ToString());
                                return;
                            }


                            DateTime LineDate = new DateTime(Year, Month, Day);

                            if (CurrentDate != LineDate)
                            {
                                continue;
                            }

                            int? Hour = null;
                            if (lineValueArr[4].Length > 2)
                            {
                                Hour = int.Parse(lineValueArr[4].Substring(0, 2));
                            }

                            float? Temp_C = null;
                            float? DewPoint_C = null;
                            float? RelHum_Perc = null;
                            float? WindDir_10deg = null;
                            float? WindSpd_km_h = null;
                            float? Visibility_km = null;
                            float? StnPress_kPa = null;
                            float? Humidx = null;
                            float? WindChill = null;

                            if (lineValueArr.Count > 5)
                            {
                                HasValue = true;
                                if (lineValueArr[6].Length > 0)
                                {
                                    Temp_C = float.Parse(lineValueArr[6]);
                                }
                                if (lineValueArr[8].Length > 0)
                                {
                                    DewPoint_C = float.Parse(lineValueArr[8]);
                                }
                                if (lineValueArr[10].Length > 0)
                                {
                                    RelHum_Perc = float.Parse(lineValueArr[10]);
                                }
                                if (lineValueArr[12].Length > 0)
                                {
                                    WindDir_10deg = float.Parse(lineValueArr[12]);
                                }
                                if (lineValueArr[14].Length > 0)
                                {
                                    WindSpd_km_h = float.Parse(lineValueArr[14]);
                                }
                                if (lineValueArr[16].Length > 0)
                                {
                                    Visibility_km = float.Parse(lineValueArr[16]);
                                }
                                if (lineValueArr[18].Length > 0)
                                {
                                    StnPress_kPa = float.Parse(lineValueArr[18]);
                                }
                                if (lineValueArr[20].Length > 0)
                                {
                                    Humidx = float.Parse(lineValueArr[20]);
                                }
                                if (lineValueArr[22].Length > 0)
                                {
                                    WindChill = float.Parse(lineValueArr[22]);
                                }

                                hourlyValues.AppendLine(
                                    Year.ToString() + "," +
                                    Month.ToString() + "," +
                                    Day.ToString() + "," +
                                    Hour.ToString() + "," +
                                    (Temp_C == null ? "" : Temp_C.ToString()) + "," +
                                    (DewPoint_C == null ? "" : DewPoint_C.ToString()) + "," +
                                    (RelHum_Perc == null ? "" : RelHum_Perc.ToString()) + "," +
                                    (WindDir_10deg == null ? "" : WindDir_10deg.ToString()) + "," +
                                    (WindSpd_km_h == null ? "" : WindSpd_km_h.ToString()) + "," +
                                    (Visibility_km == null ? "" : Visibility_km.ToString()) + "," +
                                    (StnPress_kPa == null ? "" : StnPress_kPa.ToString()) + "," +
                                    (Humidx == null ? "" : Humidx.ToString()) + "," +
                                    (WindChill == null ? "" : WindChill.ToString())
                                    );
                            }
                        }
                    }
                }
            }

            if (!HasValue)
            {
                hourlyValues.Clear();
            }
        }
        #endregion Functions public

        #region Functions private

        #endregion Functions private

    }
}
