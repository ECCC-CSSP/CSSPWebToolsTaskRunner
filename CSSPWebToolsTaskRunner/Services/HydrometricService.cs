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
using CSSPDBDLL.Models;
using CSSPDBDLL.Services;
using System.Net;
using CSSPModelsDLL.Models;
using CSSPEnumsDLL.Enums;
using CSSPDBDLL;
using System.Data.OleDb;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Collections.ObjectModel;

namespace CSSPWebToolsTaskRunner.Services
{
    public class HydrometricService
    {
        #region Variables
        public List<string> ProvList = new List<string>() { "BC", "NB", "NL", "NS", "PE", "QC" };
        public List<string> ProvNameListEN = new List<string>() { "British Columbia", "New Brunswick", "Newfoundland and Labrador", "Nova Scotia", "Prince Edward Island", "Québec" };
        public List<string> ProvNameListFR = new List<string>() { "Colombie-Britannique", "Nouveau-Brunswick", "Terre-Neuve-et-Labrador", "Nouvelle-Écosse", "Île-du-Prince-Édouard", "Québec" };
        public string URLUpdateHydrometricSiteInfo = "https://wateroffice.ec.gc.ca/station_metadata/station_result_e.html?search_type=province&province={0}";
        public string UrlToGetHydrometricSiteDataForYear = "https://wateroffice.ec.gc.ca/report/historical_e.html?stn={0}&mode=Table&type=h2oArc&results_type=historical&dataType=Daily&parameterType={1}&year={2}";
        public bool WebBrowserContentAnalysed = true;
        public bool AllDone = false;
        public string FedStationNameToDo = "";
        public string DataTypeToDo = "";
        public int YearToDo = 0;
        public int HydrometricSiteIDToDo = 0;
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
        #endregion Events

        #region Functions public
        public void LoadHydrometricDataValueDB()
        {
            TVItemService _TVItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MikeSourceService _MikeSourceService = new MikeSourceService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MikeScenarioService _MikeScenarioService = new MikeScenarioService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MWQMRunService _MWQMRunService = new MWQMRunService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            HydrometricSiteService _HydrometricSiteService = new HydrometricSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            HydrometricDataValueService _HydrometricDataValueService = new HydrometricDataValueService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            int HydrometricSiteTVItemID = 0;
            int Year = 0;

            string NotUsed = "";

            string Parameters = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters;
            string[] ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            int MikeSourceTVItemID = 0;
            int MikeScenarioTVItemID = 0;
            foreach (string s in ParamValueList)
            {
                string[] ParamValue = s.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                if (ParamValue.Length != 2)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, TaskRunnerServiceRes.Parameters);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", TaskRunnerServiceRes.Parameters);
                    return;
                }

                if (ParamValue[0] == "MikeSourceTVItemID")
                {
                    MikeSourceTVItemID = int.Parse(ParamValue[1]);
                }
                else if (ParamValue[0] == "MikeScenarioTVItemID")
                {
                    MikeScenarioTVItemID = int.Parse(ParamValue[1]);
                }
                else
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, ParamValue[0]);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", ParamValue[0].ToString());
                    return;
                }
            }

            if (MikeSourceTVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBeNullOrEmpty, TaskRunnerServiceRes.MikeScenarioTVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_ShouldNotBeNullOrEmpty", TaskRunnerServiceRes.MikeScenarioTVItemID);
                return;
            }

            TVItemModel tvItemModelMikeSource = _TVItemService.GetTVItemModelWithTVItemIDDB(MikeSourceTVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelMikeSource.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, MikeSourceTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, MikeSourceTVItemID.ToString());
                return;
            }

            MikeSourceModel mikeSourceModel = _MikeSourceService.GetMikeSourceModelWithMikeSourceTVItemIDDB(MikeSourceTVItemID);
            if (!string.IsNullOrWhiteSpace(mikeSourceModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeSource, TaskRunnerServiceRes.MikeSourceTVItemID, MikeSourceTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeSource, TaskRunnerServiceRes.MikeSourceTVItemID, MikeSourceTVItemID.ToString());
                return;
            }

            if (mikeSourceModel.UseHydrometric == false)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._ShouldBeTrue, "UseHydrometric");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_ShouldBeTrue", "UseHydrometric");
                return;
            }

            if (mikeSourceModel.HydrometricTVItemID == null)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBeNull, "HydrometricTVItemID");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_ShouldNotBeNull", "HydrometricTVItemID");
                return;
            }

            HydrometricSiteTVItemID = (int)mikeSourceModel.HydrometricTVItemID;

            if (MikeScenarioTVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBeNullOrEmpty, TaskRunnerServiceRes.MikeScenarioTVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_ShouldNotBeNullOrEmpty", TaskRunnerServiceRes.MikeScenarioTVItemID);
                return;
            }

            HydrometricSiteModel hydrometricSiteModel = _HydrometricSiteService.GetHydrometricSiteModelWithHydrometricSiteTVItemIDDB(HydrometricSiteTVItemID);
            if (!string.IsNullOrWhiteSpace(hydrometricSiteModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.HydrometricSiteModel, TaskRunnerServiceRes.HydrometricSiteTVItemID, HydrometricSiteTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.HydrometricSiteModel, TaskRunnerServiceRes.HydrometricSiteTVItemID, HydrometricSiteTVItemID.ToString());
                return;
            }

            TVItemModel tvItemModelMikeScenario = _TVItemService.GetTVItemModelWithTVItemIDDB(MikeScenarioTVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelMikeScenario.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, MikeSourceTVItemID.ToString());
                return;
            }

            MikeScenarioModel mikeScenarioModel = _MikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(MikeScenarioTVItemID);
            if (!string.IsNullOrWhiteSpace(mikeScenarioModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.MikeScenarioTVItemID, MikeScenarioTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.MikeScenarioTVItemID, MikeSourceTVItemID.ToString());
                return;
            }

            if (mikeScenarioModel.ForSimulatingMWQMRunTVItemID == null)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBeNull, "ForSimulatingMWQMRunTVItemID");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_ShouldNotBeNull", "ForSimulatingMWQMRunTVItemID");
                return;
            }

            MWQMRunModel mwqmRunModel = _MWQMRunService.GetMWQMRunModelWithMWQMRunTVItemIDDB((int)mikeScenarioModel.ForSimulatingMWQMRunTVItemID);
            if (!string.IsNullOrWhiteSpace(mwqmRunModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MWQMRun, TaskRunnerServiceRes.MWQMRunTVItemID, ((int)mikeScenarioModel.ForSimulatingMWQMRunTVItemID).ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MWQMRun, TaskRunnerServiceRes.MWQMRunTVItemID, ((int)mikeScenarioModel.ForSimulatingMWQMRunTVItemID).ToString());
                return;
            }

            Year = mwqmRunModel.DateTime_Local.Year;

            List<HydrometricDataValueModel> hydrometricDataValueModel10ValueList = _HydrometricDataValueService.GetHydrometricDataValueModelListWithHydrometricSiteIDAroundRunDateDB(hydrometricSiteModel.HydrometricSiteID, mwqmRunModel.DateTime_Local);

            if (hydrometricDataValueModel10ValueList.Count != 10)
            {
                List<string> FlowLevelList = new List<string>() { "Flow", "Level" };

                foreach (string FlowLevel in FlowLevelList)
                {

                    if (hydrometricSiteModel.HasDischarge != true && FlowLevel == "Flow")
                    {
                        continue;
                    }

                    if (hydrometricSiteModel.HasDischarge != true && FlowLevel == "Level")
                    {
                        continue;
                    }

                    FedStationNameToDo = hydrometricSiteModel.FedSiteNumber;
                    DataTypeToDo = FlowLevel;
                    YearToDo = Year;
                    HydrometricSiteIDToDo = hydrometricSiteModel.HydrometricSiteID;

                    string url = string.Format(UrlToGetHydrometricSiteDataForYear, hydrometricSiteModel.FedSiteNumber, FlowLevel, Year);

                    using (IWebDriver driver = new ChromeDriver())
                    {
                        GetHydrometricSiteDataForYear(driver, hydrometricSiteModel, FlowLevel, Year);
                    }

                    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                    {
                        return;
                    }

                    if (Year != mwqmRunModel.DateTime_Local.AddDays(-7).Year)
                    {
                        Year = Year - 1;

                        FedStationNameToDo = hydrometricSiteModel.FedSiteNumber;
                        DataTypeToDo = FlowLevel;
                        YearToDo = Year;
                        HydrometricSiteIDToDo = hydrometricSiteModel.HydrometricSiteID;

                        url = string.Format(UrlToGetHydrometricSiteDataForYear, hydrometricSiteModel.FedSiteNumber, FlowLevel, Year);

                        using (IWebDriver driver = new ChromeDriver())
                        {
                            GetHydrometricSiteDataForYear(driver, hydrometricSiteModel, FlowLevel, Year);
                        }

                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                        {
                            return;
                        }
                    }
                }
            }

            return;
        }

        public void GetHydrometricSiteDataForYear(IWebDriver driver, HydrometricSiteModel hydrometricSiteModel, string DataType, int year)
        {
            int RowCount = 37;
            int CellCount = 13;
            IWebElement webElement = null;
            IWebElement webElementTable = null;
            ReadOnlyCollection<IWebElement> webElementRowList = null;
            ReadOnlyCollection<IWebElement> webElementCellList = null;

            if (DataType == "Level")
            {
                RowCount = 35;
            }

            bool GoodURL = false;
            string NotUsed = "";
            bool IsDischarge = true;

            string url = string.Format(UrlToGetHydrometricSiteDataForYear, hydrometricSiteModel.FedSiteNumber, DataType, year);

            if (url.Contains("&parameterType=Flow"))
            {
                IsDischarge = true;
            }
            else if (url.Contains("&parameterType=Level"))
            {
                IsDischarge = false;
            }
            else
            {
                NotUsed = string.Format(TaskRunnerServiceRes.InvalidURL_CouldNotFind_, url, "&parameterType=Flow or &parameterType=Level");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("InvalidURL_CouldNotFind_", url, "&parameterType=Flow or &parameterType=Level");
                return;
            }

            driver.Navigate().GoToUrl(url);

            if (driver.Title.StartsWith("Disclaimer for Hydrometric Information"))
            {
                // need to go through disclaimer first
                IReadOnlyCollection<IWebElement> webElementToClickList = driver.FindElements(By.Name("disclaimer_action"));
                foreach (IWebElement webElementDisclaimer in webElementToClickList)
                {
                    if (webElementDisclaimer.GetAttribute("value") == "I Agree")
                    {
                        webElementDisclaimer.Click();

                        if (driver.Title.StartsWith("Daily Discharge Data") && driver.Title.Contains(FedStationNameToDo))
                        {
                            // strong possibility we are at the right url
                            GoodURL = true;
                        }
                        else if (driver.Title.StartsWith("Daily Water Level Data") && driver.Title.Contains(FedStationNameToDo))
                        {
                            // strong possibility we are at the right url
                            GoodURL = true;
                        }
                        else
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.HTMLTitleShouldStartWith_, "Disclaimer for Hydrometric Information or Daily Discharge Data");
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("HTMLTitleShouldStartWith_", "Disclaimer or Daily Discharge Data");
                            return;
                        }

                        break;
                    }
                }
            }
            else if (driver.Title.StartsWith("Daily Discharge Data") && driver.Title.Contains(FedStationNameToDo))
            {
                // strong possibility we are at the right url
                GoodURL = true;
            }
            else if (driver.Title.StartsWith("Daily Water Level Data") && driver.Title.Contains(FedStationNameToDo))
            {
                // strong possibility we are at the right url
                GoodURL = true;
            }
            else
            {
                NotUsed = string.Format(TaskRunnerServiceRes.HTMLTitleShouldStartWith_, "Disclaimer for Hydrometric Information or Daily Discharge Data or Daily Water Level Data");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("HTMLTitleShouldStartWith_", "Disclaimer or Daily Discharge Data or Daily Water Level Data");
                return;
            }

            if (!GoodURL)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.HTMLTitleShouldStartWith_, "Disclaimer for Hydrometric Information or Daily Discharge Data or Daily Water Level Data");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("HTMLTitleShouldStartWith_", "Disclaimer or Daily Discharge Data or Daily Water Level Data");
                return;
            }

            // now let's parse the page and store the data in the DB

            List<string> idListToVerify = new List<string>() { "selectStn", "selectReportType", "selectDataType", "selectYear" };
            List<string> ShouldBeValueList = new List<string>() { FedStationNameToDo, "Daily", DataTypeToDo, YearToDo.ToString() };

            for (int i = 0, count = idListToVerify.Count; i < count; i++)
            {
                try
                {
                    webElement = driver.FindElement(By.Id(idListToVerify[i]));
                    SelectElement selectedValue = new SelectElement(webElement);
                    string selectedText = selectedValue.SelectedOption.Text;

                    if (ShouldBeValueList[i] != selectedText.Trim())
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindElementID_WithinURL_, idListToVerify[i], url);
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFindElementID_WithinURL_", idListToVerify[i], url);
                        return;
                    }
                }
                catch (Exception)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindElementID_WithinURL_, idListToVerify[i], url);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFindElementID_WithinURL_", idListToVerify[i], url);
                    return;
                }
            }

            // all verification done to really know we are at the right URL

            try
            {
                webElementTable = driver.FindElement(By.Id("wb-auto-1"));

                if (webElementTable.TagName != "table")
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindElementID_WithinURL_, "wb-auto-1", url);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFindElementID_WithinURL_", "wb-auto-1", url);
                    return;
                }
            }
            catch (Exception)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindElementID_WithinURL_, "wb-auto-1", url);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFindElementID_WithinURL_", "wb-auto-1", url);
                return;
            }

            webElementRowList = webElementTable.FindElements(By.TagName("tr"));
            if (webElementRowList.Count() != RowCount)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.TableDoesNotContain_RowsWithinURL_, RowCount.ToString(), url);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("TableDoesNotContain_RowsWithinURL_", RowCount.ToString(), url);
                return;
            }

            webElementCellList = webElementRowList[0].FindElements(By.TagName("th"));
            if (webElementCellList.Count() != CellCount)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.TableDoesNotContain_CellsOnFirstRowWithinURL_, CellCount.ToString(), url);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("TableDoesNotContain_CellsOnFirstRowWithinURL_", CellCount.ToString(), url);
                return;
            }

            List<string> HeaderTextList = new List<string>()
            {
                "Day", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
            };

            for (int i = 0, count = HeaderTextList.Count; i < count; i++)
            {
                string CellText = webElementCellList[i].Text.Trim();
                if (HeaderTextList[i] != CellText)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.TableDoesNotContainProperHeader_OnFirstRowWithinURL_, HeaderTextList[i], url);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("TableDoesNotContainProperHeader_OnFirstRowWithinURL_", HeaderTextList[i], url);
                    return;
                }
            }

            List<HydrometricDataValueModel> hydrometricDataValueModelList = new List<HydrometricDataValueModel>();

            for (int row = 1, count = webElementRowList.Count; row < count; row++) // starting after header
            {
                // getting Day
                webElementCellList = webElementRowList[row].FindElements(By.TagName("th"));
                if (webElementCellList.Count != 1)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.FirstCellOfRow_WasSupposeToBeATHWithinURL_, row.ToString(), url);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("FirstCellOfRow_WasSupposeToBeATHWithinURL_", row.ToString(), url);
                    return;
                }

                string dayText = webElementCellList[0].Text.Trim();

                if (dayText == "Mean")
                {
                    break;
                }

                int day = int.Parse(webElementCellList[0].Text.Trim());

                webElementCellList = webElementRowList[row].FindElements(By.TagName("td"));
                if (webElementCellList.Count != 12)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.ShouldHaveFound_CellsOfTypeTDWithinURL_, 12.ToString(), url);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("ShouldHaveFound_CellsOfTypeTDWithinURL_", 12.ToString(), url);
                    return;
                }

                for (int col = 0, countCol = webElementCellList.Count; col < countCol; col++)
                {
                    int month = col + 1;

                    DateTime ValueDate = DateTime.Now;
                    try
                    {
                        ValueDate = new DateTime(YearToDo, month, day);
                    }
                    catch (Exception)
                    {
                        continue;
                    }

                    string colText = webElementCellList[col].Text;

                    double? value = null;

                    if (!string.IsNullOrWhiteSpace(colText))
                    {
                        colText = colText.Trim();

                        if (colText.Contains(" "))
                        {
                            if (double.TryParse(colText.Substring(0, colText.IndexOf(" ")), out double TempVal))
                            {
                                value = TempVal;
                            }
                        }
                        else
                        {
                            if (double.TryParse(colText, out double TempVal))
                            {
                                value = TempVal;
                            }
                        }
                    }

                    hydrometricDataValueModelList.Add(new HydrometricDataValueModel()
                    {
                        DateTime_Local = ValueDate,
                        DischargeEntered_m3_s = null,
                        Discharge_m3_s = (IsDischarge ? value : null),
                        HasBeenRead = true,
                        HourlyValues = "",
                        HydrometricSiteID = HydrometricSiteIDToDo,
                        Keep = true,
                        Level_m = (!IsDischarge ? value : null),
                        StorageDataType = StorageDataTypeEnum.Archived,
                    });

                }
            }

            HydrometricDataValueService _HydrometricDataValueService = new HydrometricDataValueService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            List<HydrometricDataValueModel> hydrometricDataValueModelInDBList = _HydrometricDataValueService.GetHydrometricDataValueModelListWithHydrometricSiteIDAndYearDB(hydrometricSiteModel.HydrometricSiteID, year);

            foreach (HydrometricDataValueModel hydrometricDataValueModel in hydrometricDataValueModelList.OrderBy(c => c.DateTime_Local))
            {
                HydrometricDataValueModel hydrometricDataValueModelInDB = (from c in hydrometricDataValueModelInDBList
                                                                           where c.DateTime_Local == hydrometricDataValueModel.DateTime_Local
                                                                 select c).FirstOrDefault();

                if (hydrometricDataValueModelInDB == null)
                {
                    HydrometricDataValueModel hydrometricDataValueModelRet = _HydrometricDataValueService.PostAddHydrometricDataValueDB(hydrometricDataValueModel);
                    if (!string.IsNullOrWhiteSpace(hydrometricDataValueModelRet.Error))
                    {
                        NotUsed = hydrometricDataValueModelRet.Error;
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(hydrometricDataValueModelRet.Error);
                        return;
                    }
                }
                else
                {
                    bool NeedChange = false;

                    if (IsDischarge)
                    {
                        if (hydrometricDataValueModelInDB.Discharge_m3_s != hydrometricDataValueModel.Discharge_m3_s)
                        {
                            hydrometricDataValueModelInDB.Discharge_m3_s = hydrometricDataValueModel.Discharge_m3_s;
                            NeedChange = true;
                        }
                    }
                    else
                    {
                        if (hydrometricDataValueModelInDB.Level_m != hydrometricDataValueModel.Level_m)
                        {
                            hydrometricDataValueModelInDB.Level_m = hydrometricDataValueModel.Level_m;
                            NeedChange = true;
                        }
                    }

                    if (NeedChange)
                    {
                        HydrometricDataValueModel hydrometricDataValueModelRet = _HydrometricDataValueService.PostUpdateHydrometricDataValueDB(hydrometricDataValueModelInDB);
                        if (!string.IsNullOrWhiteSpace(hydrometricDataValueModelRet.Error))
                        {
                            NotUsed = hydrometricDataValueModelRet.Error;
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(hydrometricDataValueModelRet.Error);
                            return;
                        }
                    }
                }
            }
        }

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
        #endregion Functions public

        #region Functions private

        #endregion Functions private

    }
}
