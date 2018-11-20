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
using CSSPDBDLL.Services;
using CSSPEnumsDLL.Enums;
using DHI.PFS;

namespace CSSPWebToolsTaskRunner.Test.Services
{
    /// <summary>
    /// Summary description for BaseMOdelServiceTest
    /// </summary>
    [TestClass]
    public class MikeScenarioFileServiceTest
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
        private ReportTypeService _ReportTypeService { get; set; }
        private MikeScenarioService _MikeScenarioService { get; set; }
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
        public MikeScenarioFileServiceTest()
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
        public void MikeScenarioFileService_MikeScenarioPrepareResults_Test()
        {
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                SetupTest(LanguageRequest);

                int MikeScenarioTVItemID = 357139;

                string Parameters = $"|||MikeScenarioTVItemID,{ MikeScenarioTVItemID }|||";

                AppTaskModel appTaskModel = new AppTaskModel()
                {
                    AppTaskID = 100000,
                    TVItemID = MikeScenarioTVItemID,
                    TVItemID2 = MikeScenarioTVItemID,
                    AppTaskCommand = AppTaskCommandEnum.MikeScenarioPrepareResults,
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

                MikeScenarioFileService _MikeScenarioFileService = new MikeScenarioFileService(taskRunnerBaseService);
                _MikeScenarioFileService.MikeScenarioPrepareResults();
                Assert.AreEqual(0, taskRunnerBaseService._BWObj.TextLanguageList.Count);

                break;
            }

        }
        [TestMethod]
        public void MikeScenarioFileService_MikeScenarioAskToRun_Test()
        {
            // AppTaskID TVItemID    TVItemID2 AppTaskCommand  AppTaskStatus PercentCompleted    Parameters Language    StartDateTime_UTC EndDateTime_UTC EstimatedLength_second RemainingTime_second    LastUpdateDate_UTC LastUpdateContactTVItemID
            // 14895   357139  357139  2   1   1 ||| MikeScenarioTVItemID,357139 ||| 1   2018 - 10 - 17 12:04:21.687 NULL NULL    NULL    2018 - 10 - 17 12:04:21.687 2
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                SetupTest(LanguageRequest);

                int MikeScenarioTVItemID = 357139;

                string Parameters = $"|||MikeScenarioTVItemID,{ MikeScenarioTVItemID }|||";

                AppTaskModel appTaskModel = new AppTaskModel()
                {
                    AppTaskID = 100000,
                    TVItemID = MikeScenarioTVItemID,
                    TVItemID2 = MikeScenarioTVItemID,
                    AppTaskCommand = AppTaskCommandEnum.MikeScenarioAskToRun,
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

                MikeScenarioFileService _MikeScenarioFileService = new MikeScenarioFileService(taskRunnerBaseService);
                _MikeScenarioFileService.MikeScenarioAskToRun();
                Assert.AreEqual(0, taskRunnerBaseService._BWObj.TextLanguageList.Count);

                break;
            }

        }

        [TestMethod]
        public void MikeScenarioFileService_CreateInititalConditionFileSalAndTempFromTVFileItemID_Test()
        {
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                SetupTest(LanguageRequest);

                int MikeScenarioTVItemID = 357139;

                string Parameters = $"|||not important|||";

                AppTaskModel appTaskModel = new AppTaskModel()
                {
                    AppTaskID = 100000,
                    TVItemID = MikeScenarioTVItemID,
                    TVItemID2 = MikeScenarioTVItemID,
                    AppTaskCommand = AppTaskCommandEnum.Error,
                    AppTaskStatus = AppTaskStatusEnum.Error,
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

                MikeScenarioFileService _MikeScenarioFileService = new MikeScenarioFileService(taskRunnerBaseService);

                TVItemModel tvItemModelMikeScenario = _TVItemService.GetTVItemModelWithTVItemIDDB(MikeScenarioTVItemID);
                Assert.AreEqual("", tvItemModelMikeScenario.Error);

                MikeScenarioModel mikeScenarioModel = _MikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(MikeScenarioTVItemID);
                Assert.AreEqual("", mikeScenarioModel.Error);

                TVFileModel tvFileModelM21_3FM = _TVFileService.GetTVFileModelWithTVItemIDAndTVFileTypeM21FMOrM3FMDB(MikeScenarioTVItemID);
                Assert.AreEqual("", tvFileModelM21_3FM.Error);

                string ServerPath = _TVFileService.GetServerFilePath(MikeScenarioTVItemID);

                FileInfo fiM21_M3 = new FileInfo(ServerPath + tvFileModelM21_3FM.ServerFileName);
                Assert.AreEqual(true, fiM21_M3.Exists);

                _MikeScenarioFileService.CreateInititalConditionFileSalAndTempFromTVFileItemID(MikeScenarioTVItemID, (int)mikeScenarioModel.UseSalinityAndTemperatureInitialConditionFromTVFileTVItemID, fiM21_M3);
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
            _ReportTypeService = new ReportTypeService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
            _MikeScenarioService = new MikeScenarioService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        }
        #endregion Functions private
    }
}

