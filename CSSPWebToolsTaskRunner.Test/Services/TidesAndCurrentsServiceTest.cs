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
using Microsoft.QualityTools.Testing.Fakes;
using CSSPWebToolsDBDLL.Services;
using System.IO;
using CSSPWebToolsDBDLL.Services;
using CSSPEnumsDLL.Enums;

namespace CSSPWebToolsTaskRunner.Test.Services
{
    /// <summary>
    /// Summary description for BaseMOdelServiceTest
    /// </summary>
    [TestClass]
    public class TidesAndCurrentsServiceTest
    {
        #region Variables
        private TestContext testContextInstance { get; set; }
        #endregion Variables

        #region Properties
        private CSSPWebToolsTaskRunner csspWebToolsTaskRunner { get; set; }
        private BWObj bwObj { get; set; }
        private TidesAndCurrentsService tidesAndCurrentsService { get; set; }
        private AppTaskService appTaskService { get; set; }
        private TVItemService tvItemService { get; set; }
        private TVFileService tvFileService { get; set; }
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
        public TidesAndCurrentsServiceTest()
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
        public void TidesAndCurrentsService_Constructor_Test()
        {
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                // Arrange 
                SetupTest(LanguageRequest);
                Assert.IsNotNull(csspWebToolsTaskRunner);

                AppTaskModel appTaskModel = new AppTaskModel()
                {
                    AppTaskID = 100000,
                    TVItemID = 123456,
                    TVItemID2 = 123456,
                    AppTaskCommand = AppTaskCommandEnum.SetupWebTide,
                    AppTaskStatus = AppTaskStatusEnum.Created,
                    PercentCompleted = 1,
                    Parameters = "|||MikeScenarioTVItemID,336888|||WebTideDataSet,11|||",
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

                using (TransactionScope ts = new TransactionScope())
                {
                    tidesAndCurrentsService = new TidesAndCurrentsService(taskRunnerBaseService);
                    Assert.IsNotNull(tidesAndCurrentsService);
                    Assert.IsNotNull(tidesAndCurrentsService._TaskRunnerBaseService);
                    Assert.AreEqual(bwObj.appTaskModel.AppTaskID, tidesAndCurrentsService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                }
            }
        }
        [TestMethod]
        public void TidesAndCurrentsService_SetWebTide_Test()
        {
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                SetupTest(LanguageRequest);

                int MikeScenarioTVItemID = 336888;
                string Parameters = "|||MikeScenarioTVItemID,336888|||WebTideDataSet,11|||";

                //AppTaskID	TVItemID	TVItemID2	AppTaskCommand	AppTaskStatus	PercentCompleted	Parameters	Language	StartDateTime_UTC	EndDateTime_UTC	EstimatedLength_second	RemainingTime_second	LastUpdateDate_UTC	LastUpdateContactTVItemID
                //10346	336888	336888	8	2	10	|||MikeScenarioTVItemID,336888|||WebTideDataSet,11|||	1	2018-02-12 13:16:35.400	NULL NULL    NULL	2018-02-12 13:16:36.477	2
                AppTaskModel appTaskModel = new AppTaskModel()
                {
                    AppTaskID = 100000,
                    TVItemID = MikeScenarioTVItemID,
                    TVItemID2 = MikeScenarioTVItemID,
                    AppTaskCommand = AppTaskCommandEnum.SetupWebTide,
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
                TidesAndCurrentsService tidesAndCurrentsService = new TidesAndCurrentsService(taskRunnerBaseService);

                tidesAndCurrentsService.SetupWebTide();

                break;
            }

        }

        //[TestMethod]
        //public void TidesAndCurrentsService_CreateHighAndLowTide_Test()
        //{
        //    foreach (string LanguageRequest in new List<string>() { "en", "fr" })
        //    {
        //        // Arrange 
        //        SetupTest(LanguageRequest);
        //        Assert.IsNotNull(csspWebToolsTaskRunner);

        //        csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";

        //        csspWebToolsTaskRunner.StopTimer();

        //        using (TransactionScope ts = new TransactionScope())
        //        {
        //            RemoveAllTask();

        //            TVItemModel tvItemModelRoot = tvItemService.GetRootTVItemModelForLocationDB();

        //            Assert.AreEqual("", tvItemModelRoot.Error);

        //            TVItemModel tvItemModelSubsector = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, "NB-06-020-002", TVTypeEnum.Subsector);
        //            Assert.AreEqual("", tvItemModelSubsector.Error);

        //            List<TVItemModel> tvItemModelTideSiteList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.TideSite);
        //            Assert.IsTrue(tvItemModelTideSiteList.Count > 0);

        //            TVItemModel tvItemModelTideSite = tvItemModelTideSiteList[0];


        //        }

        //        tidesAndCurrentsService = new TidesAndCurrentsService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //        Assert.IsNotNull(tidesAndCurrentsService);
        //        Assert.IsNotNull(tidesAndCurrentsService._TaskRunnerBaseService);

        //        TideModel tideModel = new TideModel(tvFileService.ChoseEDriveOrCDrive(tvFileService.BasePath), WebTideDataSetEnum.nwatl)
        //        {
        //            StartDate = new DateTime(2015, 1, 1),
        //            EndDate = new DateTime(2015, 1, 2),
        //            Lng = (double)-64.39D,
        //            Lat = (double)47.31D,
        //            Steps_min = 60,
        //            DoWaterLevels = false,
        //        };
        //        int Days = 5;

        //        List<PeakDifference> PeakDifferenceList = tidesAndCurrentsService.CreateHighAndLowTide(tideModel, Days);
        //        Assert.IsTrue(PeakDifferenceList.Count > 0);
        //    }
        //}
        //[TestMethod]
        //public void TidesAndCurrentsService_GenerateWebTideNodes_Test()
        //{
        //    foreach (string LanguageRequest in new List<string>() { "en", "fr" })
        //    {
        //        // Arrange 
        //        SetupTest(LanguageRequest);
        //        Assert.IsNotNull(csspWebToolsTaskRunner);

        //        csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";

        //        csspWebToolsTaskRunner.StopTimer();

        //        using (TransactionScope ts = new TransactionScope())
        //        {
        //            RemoveAllTask();

        //            TVItemModel tvItemModelRoot = tvItemService.GetRootTVItemModelForLocationDB();
        //            Assert.AreEqual("", tvItemModelRoot.Error);

        //            SetupBWObj(LanguageRequest);

        //            tidesAndCurrentsService = new TidesAndCurrentsService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(tidesAndCurrentsService);
        //            Assert.IsNotNull(tidesAndCurrentsService._TaskRunnerBaseService);
        //            Assert.AreEqual(bwObj.appTaskModel.AppTaskID, tidesAndCurrentsService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
        //        }
        //    }
        //}
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
        //private void SetupBWObj(string LanguageRequest)
        //{
        //    RemoveAllTask();

        //    tvFileService.LanguageRequest = LanguageRequest;

        //    using (ShimsContext.Create())
        //    {
        //        SetupShim();

        //        shimTaskRunnerBaseService.GenerateDoc = () =>
        //        {
        //            return; // just so it does not start to Generate the documents
        //        };
        //        csspWebToolsTaskRunner.GetNextTask();

        //        //
        //        //Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel);
        //        //Assert.AreEqual(tvItemModel.TVItemID, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //        //Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.bw);
        //        //Assert.IsFalse(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskCommand.ToString()));
        //        //Assert.IsTrue(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.Index > 0);
        //        //Assert.AreEqual(tvItemModel.TVItemID, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //        //Assert.IsTrue(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.ErrorText));
        //        //Assert.IsTrue(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.StatusText));
        //        //string FirstPart = "|||TVItemID," + tvItemModel.TVItemID.ToString() +
        //        //    "|||FileGenerator," + ((int)fileGenerator).ToString() +
        //        //    "|||FileGeneratorType," + ((int)fileGeneratorType).ToString() +
        //        //    "|||Generate,1" +
        //        //    "|||Command,0" +
        //        //    "|||FileName," + FileName;
        //        //Assert.AreEqual(FirstPart, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters.Substring(0, FirstPart.Length));
        //        //string LastPart = "_" + csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Language + ext.ToLower() + "|||";
        //        //int ParamLength = csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters.Length;
        //        //Assert.AreEqual(LastPart, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters.Substring(ParamLength - LastPart.Length));
        //        //Assert.AreEqual(fileGenerator, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.fileGenerator);
        //        //Assert.AreEqual(fileGeneratorType, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.fileGeneratorType);
        //        //Assert.AreEqual(true, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.Generate);
        //        //Assert.AreEqual(false, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.Command);
        //        //Assert.AreEqual(FileName, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.FileName);
        //        //Assert.AreEqual(true, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.bw.IsBusy);
        //        //Assert.AreEqual(false, csspWebToolsTaskRunner._BWList[0 + 1].bw.IsBusy);
        //    }

        //}
        //public void SetupTest(string LanguageRequest)
        //{
        //    csspWebToolsTaskRunner = new CSSPWebToolsTaskRunner();
        //    appTaskService = new AppTaskService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        //    tvItemService = new TVItemService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        //    tvFileService = new TVFileService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        //}
        //private void SetupShim()
        //{
        //    shimTaskRunnerBaseService = new ShimTaskRunnerBaseService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //}
        public void SetupTest(LanguageEnum LanguageRequest)
        {
            csspWebToolsTaskRunner = new CSSPWebToolsTaskRunner();
        }
        #endregion Functions private

    }
}

