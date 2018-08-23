using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;
using CSSPWebToolsTaskRunner.Services;
using System.Linq;
using CSSPModelsDLL.Models;
using System.ComponentModel;
using System.Transactions;
using CSSPModelsDLL.Services;
using System.IO;
using CSSPWebToolsDBDLL.Services;
using CSSPEnumsDLL.Enums;
using DHI.PFS;

namespace CSSPWebToolsTaskRunner.Test.Services
{
    /// <summary>
    /// Summary description for BaseMOdelServiceTest
    /// </summary>
    [TestClass]
    public class GroupingMWQMSitesAndPolSourceSitesGenerateServiceTest
    {
        #region Variables
        private TestContext testContextInstance { get; set; }
        #endregion Variables

        #region Properties
        private CSSPWebToolsTaskRunner csspWebToolsTaskRunner { get; set; }
        private TVItemService _TVItemService { get; set; }
        private AppTaskService _AppTaskService { get; set; }
        private TVFileService _TVFileService { get; set; }
        private MapInfoService _MapInfoService { get; set; }
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
        public GroupingMWQMSitesAndPolSourceSitesGenerateServiceTest()
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
        public void GroupingMWQMSitesAndPolSourceSitesGenerateService_GenerateLinksBetweenMWQMSitesAndPolSourceSitesForCSSPWebToolsVisualization_Test()
        {
            // AppTaskID	TVItemID	TVItemID2	AppTaskCommand	AppTaskStatus	PercentCompleted	Parameters	Language	StartDateTime_UTC	EndDateTime_UTC	EstimatedLength_second	RemainingTime_second	LastUpdateDate_UTC	LastUpdateContactTVItemID
            // 12643   344316  344316  2   1   1 ||| MikeScenarioTVItemID,344316 ||| 1   2018 - 03 - 20 14:49:58.440 NULL NULL    NULL    2018 - 03 - 20 14:49:58.443 2
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                SetupTest(LanguageRequest);

                int ProvinceTVItemID = 9;

                StringBuilder sbKML = new StringBuilder();
                string Parameters = $"|||ProvinceTVItemID,{ ProvinceTVItemID }|||";

                AppTaskModel appTaskModel = new AppTaskModel()
                {
                    AppTaskID = 100000,
                    TVItemID = ProvinceTVItemID,
                    TVItemID2 = ProvinceTVItemID,
                    AppTaskCommand = AppTaskCommandEnum.GenerateLinksBetweenMWQMSitesAndPolSourceSitesForCSSPWebToolsVisualization,
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

                GroupingMWQMSitesAndPolSourceSitesGenerateService groupingMWQMSitesAndPolSourceSitesGenerateService = new GroupingMWQMSitesAndPolSourceSitesGenerateService(taskRunnerBaseService);

                groupingMWQMSitesAndPolSourceSitesGenerateService.GenerateLinksBetweenMWQMSitesAndPolSourceSitesForCSSPWebToolsVisualization();
                Assert.AreEqual(0, taskRunnerBaseService._BWObj.TextLanguageList.Count);

                break;
            }

        }
        #endregion Functions public

        #region Functions private
        public void SetupTest(LanguageEnum LanguageRequest)
        {
            csspWebToolsTaskRunner = new CSSPWebToolsTaskRunner();
            _TVItemService = new TVItemService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
            _AppTaskService = new AppTaskService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
            _TVFileService = new TVFileService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
            _MapInfoService = new MapInfoService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        }
        #endregion Functions private
    }
}

