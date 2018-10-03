using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;
using CSSPWebToolsTaskRunner.Services;
using System.Linq;
using CSSPWebToolsDBDLL.Models;
using System.ComponentModel;
using System.Transactions;
using Microsoft.QualityTools.Testing.Fakes;
using CSSPWebToolsDBDLL.Services;
using System.IO;
using System.Windows.Forms;
using CSSPWebToolsDBDLL;
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace CSSPWebToolsTaskRunner.Test.Services
{
    /// <summary>
    /// Summary description for BaseMOdelServiceTest
    /// </summary>
    [TestClass]
    public class HydrometricServiceTest
    {
        #region Variables
        private TestContext testContextInstance { get; set; }
        #endregion Variables

        #region Properties
        private CSSPWebToolsTaskRunner csspWebToolsTaskRunner { get; set; }
        private BWObj bwObj { get; set; }
        private HydrometricService HydrometricService { get; set; }
        private AppTaskService appTaskService { get; set; }
        private HydrometricSiteService hydrometricSiteService { get; set; }
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
        public HydrometricServiceTest()
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
        //[TestMethod]
        //public void HydrometricService_RunBrowserThread_Test()
        //{
        //    // AppTaskID	TVItemID	TVItemID2	AppTaskCommand	AppTaskStatus	PercentCompleted	Parameters	Language	StartDateTime_UTC	EndDateTime_UTC	EstimatedLength_second	RemainingTime_second	LastUpdateDate_UTC	LastUpdateContactTVItemID
        //    // 14775	361210	361210	44	2	1	|||MikeSourceTVItemID,361210|||MikeScenarioTVItemID,357139|||	1	2018-10-02 14:15:24.317	NULL	NULL	NULL	2018-10-02 14:15:25.307	2
        //    foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
        //    {
        //        int MikeSourceTVItemID = 361210;
        //        int MikeScenarioTVItemID = 357139;

        //        string Parameters = $"|||MikeSourceTVItemID,{ MikeSourceTVItemID }|||MikeScenarioTVItemID,{ MikeScenarioTVItemID }|||";

        //        AppTaskModel appTaskModel = new AppTaskModel()
        //        {
        //            AppTaskID = 100000,
        //            TVItemID = MikeSourceTVItemID,
        //            TVItemID2 = MikeSourceTVItemID,
        //            AppTaskCommand = AppTaskCommandEnum.LoadHydrometricDataValue,
        //            AppTaskStatus = AppTaskStatusEnum.Created,
        //            PercentCompleted = 1,
        //            Parameters = Parameters,
        //            Language = LanguageRequest,
        //            StartDateTime_UTC = DateTime.Now,
        //            EndDateTime_UTC = null,
        //            EstimatedLength_second = null,
        //            RemainingTime_second = null,
        //            LastUpdateDate_UTC = DateTime.Now,
        //            LastUpdateContactTVItemID = 2, // Charles LeBlanc
        //        };

        //        appTaskModel.AppTaskStatus = AppTaskStatusEnum.Running;

        //        BWObj bwObj = new BWObj()
        //        {
        //            Index = 1,
        //            appTaskModel = appTaskModel,
        //            appTaskCommand = appTaskModel.AppTaskCommand,
        //            TextLanguageList = new List<TextLanguage>(),
        //            bw = new BackgroundWorker(),
        //        };

        //        TaskRunnerBaseService taskRunnerBaseService = new TaskRunnerBaseService(new List<BWObj>()
        //        {
        //            bwObj
        //        });

        //        taskRunnerBaseService._BWObj = bwObj;

        //        HydrometricService _HydrometricService = new HydrometricService(taskRunnerBaseService);

        //        string url = "https://wateroffice.ec.gc.ca/report/historical_e.html?stn=02ZJ001&mode=Table&type=h2oArc&results_type=historical&dataType=Daily&parameterType=Flow&year=2008";
        //        _HydrometricService.RunBrowserThread(url);
        //        Assert.AreEqual(0, taskRunnerBaseService._BWObj.TextLanguageList.Count);

        //        break;
        //    }

        //}
        [TestMethod]
        public void HydrometricService_GetHydrometricSiteDataForYear_Test()
        {
            // AppTaskID	TVItemID	TVItemID2	AppTaskCommand	AppTaskStatus	PercentCompleted	Parameters	Language	StartDateTime_UTC	EndDateTime_UTC	EstimatedLength_second	RemainingTime_second	LastUpdateDate_UTC	LastUpdateContactTVItemID
            // 14775	361210	361210	44	2	1	|||MikeSourceTVItemID,361210|||MikeScenarioTVItemID,357139|||	1	2018-10-02 14:15:24.317	NULL	NULL	NULL	2018-10-02 14:15:25.307	2
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                SetupTest(LanguageRequest);

                int MikeSourceTVItemID = 361210;
                int MikeScenarioTVItemID = 357139;

                string Parameters = $"|||MikeSourceTVItemID,{ MikeSourceTVItemID }|||MikeScenarioTVItemID,{ MikeScenarioTVItemID }|||";

                AppTaskModel appTaskModel = new AppTaskModel()
                {
                    AppTaskID = 100000,
                    TVItemID = MikeSourceTVItemID,
                    TVItemID2 = MikeSourceTVItemID,
                    AppTaskCommand = AppTaskCommandEnum.LoadHydrometricDataValue,
                    AppTaskStatus = AppTaskStatusEnum.Created,
                    PercentCompleted = 1,
                    Parameters = Parameters,
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

                HydrometricService _HydrometricService = new HydrometricService(taskRunnerBaseService);

                _HydrometricService.FedStationNameToDo = "02ZJ001"; // HydrometricTVItemID = 56367;
                _HydrometricService.DataTypeToDo = "Flow";
                _HydrometricService.YearToDo = 2015;
                _HydrometricService.HydrometricSiteIDToDo = 288;

                HydrometricSiteModel hydrometricSiteModel = hydrometricSiteService.GetHydrometricSiteModelWithHydrometricSiteTVItemIDDB(56367);

                string url = string.Format(_HydrometricService.UrlToGetHydrometricSiteDataForYear, _HydrometricService.FedStationNameToDo, _HydrometricService.DataTypeToDo, _HydrometricService.YearToDo);

                ChromeOptions chromeOptions = new ChromeOptions();
                chromeOptions.AddArgument("headless");
                using (IWebDriver driver = new ChromeDriver()) // chromeOptions))
                {
                    _HydrometricService.GetHydrometricSiteDataForYear(driver, hydrometricSiteModel, _HydrometricService.DataTypeToDo, _HydrometricService.YearToDo);
                    Assert.AreEqual(0, taskRunnerBaseService._BWObj.TextLanguageList.Count);
                }

                break;
            }

        }
        [TestMethod]
        public void HydrometricService_LoadHydrometricDataValueDB_Test()
        {
            // AppTaskID	TVItemID	TVItemID2	AppTaskCommand	AppTaskStatus	PercentCompleted	Parameters	Language	StartDateTime_UTC	EndDateTime_UTC	EstimatedLength_second	RemainingTime_second	LastUpdateDate_UTC	LastUpdateContactTVItemID
            // 14775	361210	361210	44	2	1	|||MikeSourceTVItemID,361210|||MikeScenarioTVItemID,357139|||	1	2018-10-02 14:15:24.317	NULL	NULL	NULL	2018-10-02 14:15:25.307	2
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                int MikeSourceTVItemID = 361210;
                int MikeScenarioTVItemID = 357139;

                string Parameters = $"|||MikeSourceTVItemID,{ MikeSourceTVItemID }|||MikeScenarioTVItemID,{ MikeScenarioTVItemID }|||";

                AppTaskModel appTaskModel = new AppTaskModel()
                {
                    AppTaskID = 100000,
                    TVItemID = MikeSourceTVItemID,
                    TVItemID2 = MikeSourceTVItemID,
                    AppTaskCommand = AppTaskCommandEnum.LoadHydrometricDataValue,
                    AppTaskStatus = AppTaskStatusEnum.Created,
                    PercentCompleted = 1,
                    Parameters = Parameters,
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

                HydrometricService _HydrometricService = new HydrometricService(taskRunnerBaseService);

                _HydrometricService.LoadHydrometricDataValueDB();
                Assert.AreEqual(0, taskRunnerBaseService._BWObj.TextLanguageList.Count);

                break;
            }

        }
        [TestMethod]
        public void HydrometricService_UpdateHydrometricSitesInformationForProvinceTVItemID_Test()
        {
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                int ProvinceTVItemID = 7;

                AppTaskModel appTaskModel = new AppTaskModel()
                {
                    AppTaskID = 0,
                    TVItemID = ProvinceTVItemID,
                    TVItemID2 = ProvinceTVItemID,
                    AppTaskCommand = AppTaskCommandEnum.UpdateHydrometricSiteInformation,
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

                HydrometricService _HydrometricService = new HydrometricService(taskRunnerBaseService);
                _HydrometricService.UpdateHydrometricSitesInformationForProvinceTVItemID();
                Assert.AreEqual(0, taskRunnerBaseService._BWObj.TextLanguageList.Count);

                break;
            }

        }
        #endregion Functions public

        #region Functions private
        public void SetupTest(LanguageEnum LanguageRequest)
        {
            csspWebToolsTaskRunner = new CSSPWebToolsTaskRunner();
            //appTaskService = new AppTaskService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
            //tvItemService = new TVItemService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
            hydrometricSiteService = new HydrometricSiteService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
            //mwqmSiteService = new MWQMSiteService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
            //mwqmSampleService = new MWQMSampleService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        }
        #endregion Functions private

    }


}

