using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;
using CSSPWebToolsTaskRunner.Services;
using System.Linq;
using CSSPDBDLL.Models;
using System.ComponentModel;
using System.Transactions;
using Microsoft.QualityTools.Testing.Fakes;
using CSSPDBDLL.Services;
using System.IO;
using System.Windows.Forms;
using CSSPDBDLL;
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;

namespace CSSPWebToolsTaskRunner.Test.Services
{
    /// <summary>
    /// Summary description for BaseMOdelServiceTest
    /// </summary>
    [TestClass]
    public class ClimateServiceTest
    {
        #region Variables
        private TestContext testContextInstance { get; set; }
        #endregion Variables

        #region Properties
        private CSSPWebToolsTaskRunner csspWebToolsTaskRunner { get; set; }
        private BWObj bwObj { get; set; }
        private ClimateService climateService { get; set; }
        private AppTaskService appTaskService { get; set; }
        private ClimateSiteService climateSiteService { get; set; }
        private TVItemService tvItemService { get; set; }
        private MWQMSiteService mwqmSiteService { get; set; }
        private MWQMSampleService mwqmSampleService { get; set; }
        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }
        #endregion Properties

        #region Constructors
        public ClimateServiceTest()
        {
        }
        #endregion Constructors

        #region Initialize and Cleanup
        //
        // You can use the following additional attributes as you write your tests:
        //
        // Use ClassInitialize to run code before running the first test in the class
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // Use ClassCleanup to run code after all tests in a class have run
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Use TestInitialize to run code before running each test 
        //[TestInitialize()]
        //public void MyTestInitialize() {}
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion Initialize and Cleanup

        #region Functions public
        [TestMethod]
        public void ClimateService_Constructor_Test()
        {
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                int CountryTVItemID = 5;

                AppTaskModel appTaskModel = new AppTaskModel()
                {
                    AppTaskID = 0,
                    TVItemID = CountryTVItemID,
                    TVItemID2 = CountryTVItemID,
                    AppTaskCommand = AppTaskCommandEnum.UpdateClimateSiteInformation,
                    AppTaskStatus = AppTaskStatusEnum.Created,
                    PercentCompleted = 1,
                    Parameters = "",
                    Language = LanguageRequest,
                    StartDateTime_UTC = DateTime.Now,
                    EndDateTime_UTC = null,
                    EstimatedLength_second = null,
                    RemainingTime_second = null,
                    LastUpdateDate_UTC = DateTime.Now,
                    LastUpdateContactTVItemID = 2, // Charles LeBlanc
                };

                appTaskModel.AppTaskStatus = AppTaskStatusEnum.Running;

                BWObj bwObj = new BWObj()
                {
                    Index = 1,
                    appTaskModel = appTaskModel,
                    appTaskCommand = appTaskModel.AppTaskCommand,
                    TextLanguageList = new List<TextLanguage>(),
                    bw = new BackgroundWorker(),
                };

                TaskRunnerBaseService taskRunnerBaseService = new TaskRunnerBaseService(new List<BWObj>()
                {
                    bwObj
                });

                taskRunnerBaseService._BWObj = bwObj;

                ClimateService _ClimateService = new ClimateService(taskRunnerBaseService);
                _ClimateService.UpdateClimateSitesInformationForCountryTVItemID();
                Assert.AreEqual(0, taskRunnerBaseService._BWObj.TextLanguageList.Count);

                break;
            }
        }
        [TestMethod]
        public void ClimateService_GetAllPrecipitationForYear_Test()
        {
            // AppTaskID	TVItemID	TVItemID2	AppTaskCommand	AppTaskStatus	PercentCompleted	Parameters	Language	StartDateTime_UTC	EndDateTime_UTC	EstimatedLength_second	RemainingTime_second	LastUpdateDate_UTC	LastUpdateContactTVItemID
            // 14922	7	2018	26	1	1	|||ProvinceTVItemID,7|||Year,2018|||	1	2018-10-23 12:13:31.377	NULL	NULL	NULL	2018-10-23 12:13:31.377	2            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                int ProvinceTVItemID = 7;
                int Year = 2018;

                AppTaskModel appTaskModel = new AppTaskModel()
                {
                    AppTaskID = 0,
                    TVItemID = ProvinceTVItemID,
                    TVItemID2 = ProvinceTVItemID,
                    AppTaskCommand = AppTaskCommandEnum.GetAllPrecipitationForYear,
                    AppTaskStatus = AppTaskStatusEnum.Created,
                    PercentCompleted = 1,
                    Parameters = $"|||ProvinceTVItemID,{ ProvinceTVItemID }|||Year,{ Year }|||",
                    Language = LanguageRequest,
                    StartDateTime_UTC = DateTime.Now,
                    EndDateTime_UTC = null,
                    EstimatedLength_second = null,
                    RemainingTime_second = null,
                    LastUpdateDate_UTC = DateTime.Now,
                    LastUpdateContactTVItemID = 2, // Charles LeBlanc
                };

                appTaskModel.AppTaskStatus = AppTaskStatusEnum.Running;

                BWObj bwObj = new BWObj()
                {
                    Index = 1,
                    appTaskModel = appTaskModel,
                    appTaskCommand = appTaskModel.AppTaskCommand,
                    TextLanguageList = new List<TextLanguage>(),
                    bw = new BackgroundWorker(),
                };

                TaskRunnerBaseService taskRunnerBaseService = new TaskRunnerBaseService(new List<BWObj>()
                {
                    bwObj
                });

                taskRunnerBaseService._BWObj = bwObj;

                ClimateService _ClimateService = new ClimateService(taskRunnerBaseService);
                _ClimateService.GetAllPrecipitationForYear();
                Assert.AreEqual(0, taskRunnerBaseService._BWObj.TextLanguageList.Count);

                break;
            }
        }
        [TestMethod]
        public void ClimateService_FillRunPrecipByClimateSitePriorityForYear_Test()
        {
            // AppTaskID	TVItemID	TVItemID2	AppTaskCommand	AppTaskStatus	PercentCompleted	Parameters	Language	StartDateTime_UTC	EndDateTime_UTC	EstimatedLength_second	RemainingTime_second	LastUpdateDate_UTC	LastUpdateContactTVItemID
            // 14922	7	2018	26	1	1	|||ProvinceTVItemID,7|||Year,2018|||	1	2018-10-23 12:13:31.377	NULL	NULL	NULL	2018-10-23 12:13:31.377	2            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                int ProvinceTVItemID = 9;
                int Year = 2018;

                AppTaskModel appTaskModel = new AppTaskModel()
                {
                    AppTaskID = 0,
                    TVItemID = ProvinceTVItemID,
                    TVItemID2 = ProvinceTVItemID,
                    AppTaskCommand = AppTaskCommandEnum.FillRunPrecipByClimateSitePriorityForYear,
                    AppTaskStatus = AppTaskStatusEnum.Created,
                    PercentCompleted = 1,
                    Parameters = $"|||ProvinceTVItemID,{ ProvinceTVItemID }|||Year,{ Year }|||",
                    Language = LanguageRequest,
                    StartDateTime_UTC = DateTime.Now,
                    EndDateTime_UTC = null,
                    EstimatedLength_second = null,
                    RemainingTime_second = null,
                    LastUpdateDate_UTC = DateTime.Now,
                    LastUpdateContactTVItemID = 2, // Charles LeBlanc
                };

                appTaskModel.AppTaskStatus = AppTaskStatusEnum.Running;

                BWObj bwObj = new BWObj()
                {
                    Index = 1,
                    appTaskModel = appTaskModel,
                    appTaskCommand = appTaskModel.AppTaskCommand,
                    TextLanguageList = new List<TextLanguage>(),
                    bw = new BackgroundWorker(),
                };

                TaskRunnerBaseService taskRunnerBaseService = new TaskRunnerBaseService(new List<BWObj>()
                {
                    bwObj
                });

                taskRunnerBaseService._BWObj = bwObj;

                ClimateService _ClimateService = new ClimateService(taskRunnerBaseService);
                _ClimateService.FillRunPrecipByClimateSitePriorityForYear();
                Assert.AreEqual(0, taskRunnerBaseService._BWObj.TextLanguageList.Count);

                break;
            }
        }
        //    [TestMethod]
        //    public void ClimateService_UpdateClimateSiteInformation_Test()
        //    {
        //        foreach (string LanguageRequest in new List<string>() { "en" /*, "fr" */ })
        //        {
        //            // Arrange 
        //            SetupTest(LanguageRequest);
        //            Assert.IsNotNull(csspWebToolsTaskRunner);

        //            csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";

        //            csspWebToolsTaskRunner.StopTimer();

        //            //using (TransactionScope ts = new TransactionScope())
        //            //{
        //            RemoveAllTask();

        //            TVItemModel tvItemModelRoot = tvItemService.GetRootTVItemModelForLocationDB();
        //            Assert.AreEqual("", tvItemModelRoot.Error);

        //            TVItemModel tvItemModelProvince = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID,
        //"Prince Edward Island", TVTypeEnum.Province);
        //            Assert.AreEqual("", tvItemModelProvince.Error);

        //            climateSiteService.LanguageRequest = LanguageRequest;
        //            string retStr = climateSiteService.UpdateClimateSitesInformationForProvinceTVItemIDDB(tvItemModelProvince.TVItemID);
        //            Assert.AreEqual("", retStr);

        //            SetupBWObj(tvItemModelProvince.TVItemID, LanguageRequest, AppTaskCommandEnum.UpdateClimateSiteInformation, DateTime.Now, DateTime.Now);

        //            climateService = new ClimateService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(climateService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, climateService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            climateService.UpdateClimateSitesInformationForProvinceTVItemID();
        //            Assert.IsTrue(climateService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

        //            RemoveAllTask();

        //            //}
        //        }
        //    }
        //    [TestMethod]
        //    public void ClimateService_UpdateClimateSiteDailyAndHourlyFromStartDateToEndDate_Test()
        //    {
        //        foreach (string LanguageRequest in new List<string>() { "en" /*, "fr" */ })
        //        {
        //            // Arrange 
        //            SetupTest(LanguageRequest);
        //            Assert.IsNotNull(csspWebToolsTaskRunner);

        //            csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";

        //            csspWebToolsTaskRunner.StopTimer();

        //            //using (TransactionScope ts = new TransactionScope())
        //            //{
        //            RemoveAllTask();

        //            TVItemModel tvItemModelRoot = tvItemService.GetRootTVItemModelForLocationDB();
        //            Assert.AreEqual("", tvItemModelRoot.Error);

        //            TVItemModel tvItemModelProvince = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID,
        //(LanguageRequest == "fr" ? "Île-du-Prince-Édouard" : "Prince Edward Island"), TVTypeEnum.Province);
        //            Assert.AreEqual("", tvItemModelProvince.Error);

        //            TVItemModel tvItemModelClimateSite = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelProvince.TVItemID,
        //(LanguageRequest == "fr" ? "NORTH CAPE" : "NORTH CAPE"), TVTypeEnum.ClimateSite);
        //            Assert.AreEqual("", tvItemModelClimateSite.Error);

        //            DateTime StartDate = new DateTime(2015, 1, 1);
        //            DateTime EndDate = new DateTime(2015, 2, 1);

        //            climateSiteService.LanguageRequest = LanguageRequest;
        //            string retStr = climateSiteService.UpdateClimateSiteDailyAndHourlyFromStartDateToEndDateDB(tvItemModelClimateSite.TVItemID, StartDate, EndDate);
        //            Assert.AreEqual("", retStr);

        //            SetupBWObj(tvItemModelClimateSite.TVItemID, LanguageRequest, AppTaskCommandEnum.UpdateClimateSiteDailyAndHourlyFromStartDateToEndDate, StartDate, EndDate);

        //            climateService = new ClimateService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(climateService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, climateService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            climateService.UpdateClimateSiteDailyAndHourlyFromStartDateToEndDate();
        //            Assert.IsTrue(climateService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);
        //            //}
        //        }
        //    }
        //    [TestMethod]
        //    public void ClimateService_UpdateClimateSiteDailyAndHourlyForSubsectorFromStartDateToEndDate_Test()
        //    {
        //        foreach (string LanguageRequest in new List<string>() { "en" /*, "fr" */ })
        //        {
        //            // Arrange 
        //            SetupTest(LanguageRequest);
        //            Assert.IsNotNull(csspWebToolsTaskRunner);

        //            csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";

        //            csspWebToolsTaskRunner.StopTimer();

        //            //using (TransactionScope ts = new TransactionScope())
        //            //{
        //            RemoveAllTask();

        //            TVItemModel tvItemModelRoot = tvItemService.GetRootTVItemModelForLocationDB();
        //            Assert.AreEqual("", tvItemModelRoot.Error);

        //            List<string> provList = new List<string>() { "Prince Edward Island", "New Brunswick", "Nova Scotia", "Newfoundland and Labrador", "Québec", "British Columbia" };

        //            foreach (string Prov in provList)
        //            {
        //                TVItemModel tvItemModelProvince = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, Prov, TVTypeEnum.Province);

        //                Assert.AreEqual("", tvItemModelProvince.Error);

        //                List<TVItemModel> tvItemModelSubsectorList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelProvince.TVItemID, TVTypeEnum.Subsector);

        //                Assert.IsTrue(tvItemModelSubsectorList.Count > 0);

        //                //bool proceed = false;
        //                foreach (TVItemModel tvItemModelSubsector in tvItemModelSubsectorList)
        //                {
        //                    //if (tvItemModelSubsector.TVItemID == 1483)
        //                    //    proceed = true;

        //                    //if (!proceed)
        //                    //    continue;

        //                    List<YearMinMaxDate> yearMinMaxDateList = new List<YearMinMaxDate>();

        //                    Application.DoEvents();

        //                    List<int> mwqmSampleModelYearList = (from c in mwqmSampleService.db.TVItems
        //                                                         from si in mwqmSampleService.db.MWQMSites
        //                                                         from sa in mwqmSampleService.db.MWQMSamples
        //                                                         where c.ParentID == tvItemModelSubsector.TVItemID
        //                                                         && c.TVItemID == si.MWQMSiteTVItemID
        //                                                         && si.MWQMSiteTVItemID == sa.MWQMSiteTVItemID
        //                                                         orderby sa.SampleDateTime_Local
        //                                                         select sa.SampleDateTime_Local.Year).Distinct().ToList();

        //                    foreach (int year in mwqmSampleModelYearList)
        //                    {
        //                        DateTime Earliest = (from c in mwqmSampleService.db.TVItems
        //                                             from si in mwqmSampleService.db.MWQMSites
        //                                             from sa in mwqmSampleService.db.MWQMSamples
        //                                             where c.ParentID == tvItemModelSubsector.TVItemID
        //                                             && c.TVItemID == si.MWQMSiteTVItemID
        //                                             && si.MWQMSiteTVItemID == sa.MWQMSiteTVItemID
        //                                             && sa.SampleDateTime_Local.Year == year
        //                                             orderby sa.SampleDateTime_Local
        //                                             select sa.SampleDateTime_Local).FirstOrDefault();

        //                        DateTime Latest = (from c in mwqmSampleService.db.TVItems
        //                                           from si in mwqmSampleService.db.MWQMSites
        //                                           from sa in mwqmSampleService.db.MWQMSamples
        //                                           where c.ParentID == tvItemModelSubsector.TVItemID
        //                                           && c.TVItemID == si.MWQMSiteTVItemID
        //                                           && si.MWQMSiteTVItemID == sa.MWQMSiteTVItemID
        //                                           && sa.SampleDateTime_Local.Year == year
        //                                           orderby sa.SampleDateTime_Local descending
        //                                           select sa.SampleDateTime_Local).FirstOrDefault();

        //                        YearMinMaxDate yearMinMaxDate = (from c in yearMinMaxDateList
        //                                                         where c.Year == year
        //                                                         select c).FirstOrDefault();

        //                        if (yearMinMaxDate == null)
        //                        {
        //                            YearMinMaxDate yearMinMaxDateNew = new YearMinMaxDate()
        //                            {
        //                                Year = year,
        //                                MinDateTime = Earliest,
        //                                MaxDateTime = Latest,
        //                            };

        //                            yearMinMaxDateList.Add(yearMinMaxDateNew);
        //                        }
        //                        else
        //                        {
        //                            if (yearMinMaxDate.MinDateTime > Earliest)
        //                            {
        //                                yearMinMaxDate.MinDateTime = Earliest;
        //                            }
        //                            if (yearMinMaxDate.MaxDateTime < Latest)
        //                            {
        //                                yearMinMaxDate.MaxDateTime = Latest;
        //                            }
        //                        }
        //                    }
        //                    yearMinMaxDateList = (from c in yearMinMaxDateList orderby c.Year select c).ToList();

        //                    foreach (YearMinMaxDate yearMinMaxDate in yearMinMaxDateList)
        //                    {
        //                        RemoveAllTask();

        //                        climateSiteService.LanguageRequest = LanguageRequest;
        //                        string retStr = climateSiteService.UpdateClimateSiteDailyAndHourlyForSubsectorFromStartDateToEndDateDB(tvItemModelSubsector.TVItemID, yearMinMaxDate.MinDateTime, yearMinMaxDate.MaxDateTime);
        //                        Assert.AreEqual("", retStr);

        //                        SetupBWObj(tvItemModelSubsector.TVItemID, LanguageRequest, AppTaskCommandEnum.UpdateClimateSiteDailyAndHourlyForSubsectorFromStartDateToEndDate, yearMinMaxDate.MinDateTime, yearMinMaxDate.MaxDateTime);

        //                        climateService = new ClimateService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //                        Assert.IsNotNull(climateService);
        //                        Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, climateService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //                        climateService.UpdateClimateSiteDailyAndHourlyForSubsectorFromStartDateToEndDate();
        //                        Assert.IsTrue(climateService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

        //                        RemoveAllTask();

        //                    }
        //                }
        //            }
        //            //}
        //        }
        //    }
        #endregion Functions public

        #region Functions private
        //private void RemoveAllTask()
        //{
        //    List<AppTaskModel> appTaskModelList = appTaskService.GetAppTaskModelListDB();
        //    foreach (AppTaskModel appTaskModel in appTaskModelList)
        //    {
        //        AppTaskModel appTaskModelRet = appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
        //        Assert.AreEqual("", appTaskModelRet.Error);
        //    }
        //}
        //private void SetupBWObj(int TVItemID, string LanguageRequest, AppTaskCommandEnum appTaskCommand, DateTime StartDate, DateTime EndDate)
        //{
        //    using (ShimsContext.Create())
        //    {
        //        SetupShim();

        //        shimTaskRunnerBaseService.DoCommand = () =>
        //        {
        //            return; // just so it does not start to Generate the documents
        //        };
        //        csspWebToolsTaskRunner.GetNextTask();

        //        //
        //        Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel);
        //        Assert.AreEqual(TVItemID, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //        Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.bw);
        //        Assert.IsFalse(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskCommand.ToString()));
        //        Assert.IsTrue(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.Index > 0);
        //        Assert.AreEqual(TVItemID, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //        Assert.IsTrue(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.ErrorText));
        //        Assert.IsTrue(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.StatusText));
        //        if (appTaskCommand == AppTaskCommandEnum.UpdateClimateSiteInformation)
        //        {
        //            string FirstPart = "|||ProvinceTVItemID," + TVItemID.ToString() +
        //              "|||Generate,0" +
        //              "|||Command,1|||";
        //            Assert.AreEqual(FirstPart, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters);
        //            Assert.AreEqual(AppTaskCommandEnum.UpdateClimateSiteInformation, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskCommand);
        //        }
        //        else if (appTaskCommand == AppTaskCommandEnum.UpdateClimateSiteDailyAndHourlyFromStartDateToEndDate)
        //        {
        //            string FirstPart = "|||ClimateSiteTVItemID," + TVItemID.ToString() +
        //              "|||StartYear," + StartDate.Year.ToString() +
        //              "|||StartMonth," + StartDate.Month.ToString() +
        //              "|||StartDay," + StartDate.Day.ToString() +
        //              "|||EndYear," + EndDate.Year.ToString() +
        //              "|||EndMonth," + EndDate.Month.ToString() +
        //              "|||EndDay," + EndDate.Day.ToString() +
        //              "|||Generate,0" +
        //              "|||Command,1|||";
        //            Assert.AreEqual(FirstPart, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters);
        //            Assert.AreEqual(AppTaskCommandEnum.UpdateClimateSiteDailyAndHourlyFromStartDateToEndDate, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskCommand);
        //        }
        //        else if (appTaskCommand == AppTaskCommandEnum.UpdateClimateSiteDailyAndHourlyForSubsectorFromStartDateToEndDate)
        //        {
        //            string FirstPart = "|||SubsectorTVItemID," + TVItemID.ToString() +
        //              "|||StartYear," + StartDate.Year.ToString() +
        //              "|||StartMonth," + StartDate.Month.ToString() +
        //              "|||StartDay," + StartDate.Day.ToString() +
        //              "|||EndYear," + EndDate.Year.ToString() +
        //              "|||EndMonth," + EndDate.Month.ToString() +
        //              "|||EndDay," + EndDate.Day.ToString() +
        //              "|||Generate,0" +
        //              "|||Command,1|||";
        //            Assert.AreEqual(FirstPart, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters);
        //            Assert.AreEqual(AppTaskCommandEnum.UpdateClimateSiteDailyAndHourlyForSubsectorFromStartDateToEndDate, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskCommand);
        //        }
        //        else
        //        {
        //            Assert.IsTrue(false);
        //        }
        //        //    Assert.AreEqual(true, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.bw.IsBusy);
        //        //    Assert.AreEqual(false, csspWebToolsTaskRunner._BWList[0 + 1].bw.IsBusy);
        //    }

        //}
        //public void SetupTest(string LanguageRequest)
        //{
        //    csspWebToolsTaskRunner = new CSSPWebToolsTaskRunner();
        //    appTaskService = new AppTaskService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        //    tvItemService = new TVItemService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        //    climateSiteService = new ClimateSiteService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        //    mwqmSiteService = new MWQMSiteService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        //    mwqmSampleService = new MWQMSampleService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        //}
        //private void SetupShim()
        //{
        //    shimTaskRunnerBaseService = new ShimTaskRunnerBaseService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //}
        #endregion Functions private

    }

    public class YearMinMaxDate
    {
        public int Year { get; set; }
        public DateTime MinDateTime { get; set; }
        public DateTime MaxDateTime { get; set; }
    }
}

