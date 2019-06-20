//using System;
//using System.Text;
//using System.Collections.Generic;
//using Microsoft.VisualStudio.TestTools.UnitTesting;
//using System.Reflection;
//using CSSPWebToolsTaskRunner.Services;
//using System.Linq;
//using CSSPWebToolsDB.Models;
//using System.ComponentModel;
//using System.Transactions;
//using CSSPWebToolsTaskRunner.Services.Fakes;
//using Microsoft.QualityTools.Testing.Fakes;
//using CSSPWebToolsDB.Services;
//using System.IO;
//using CSSPWebToolsDBDLL.Services;
//using System.Web;
//using CSSPWebToolsTaskRunner.Services.Resources;
//using CSSPWebToolsTaskRunner;
//using CSSPDBDLL.Models;
//using CSSPDBDLL.Services;
//using CSSPEnumsDLL.Enums;
//using CSSPModelsDLL.Models;
//using System.Threading;
//using System.Globalization;
//using CSSPDBDLL;

//namespace CSSPWebToolsTaskRunner.Test.Services
//{
//    /// <summary>
//    /// Summary description for BaseMOdelServiceTest
//    /// </summary>
//    [TestClass]
//    public class TxtServiceTest
//    {
//        #region Variables
//        private TestContext testContextInstance { get; set; }
//        #endregion Variables

//        #region Properties
//        private CSSPWebToolsTaskRunner csspWebToolsTaskRunner { get; set; }
//        private ShimTaskRunnerBaseService shimTaskRunnerBaseService { get; set; }
//        private TVItemService tvItemService { get; set; }
//        private AppTaskService appTaskService { get; set; }
//        private TVFileService tvFileService { get; set; }
//        private MapInfoService mapInfoService { get; set; }
//        private TxtService txtService { get; set; }
//        private RandomService randomService { get; set; }
//        private MWQMPlanService mwqmPlanService { get; set; }
//        private MWQMPlanSubsectorService mwqmPlanSubsectorService { get; set; }
//        private MWQMPlanSubsectorSiteService mwqmPlanSubsectorSiteService { get; set; }
//        /// <summary>
//        ///Gets or sets the test context which provides
//        ///information about and functionality for the current test run.
//        ///</summary>
//        public TestContext TestContext
//        {
//            get
//            {
//                return testContextInstance;
//            }
//            set
//            {
//                testContextInstance = value;
//            }
//        }
//        #endregion Properties

//        #region Constructors
//        public TxtServiceTest()
//        {
//        }
//        #endregion Constructors

//        #region Initialize and Cleanup
//        //
//        // You can use the following additional attributes as you write your tests:
//        //
//        // Use ClassInitialize to run code before running the first test in the class
//        // [ClassInitialize()]
//        // public static void MyClassInitialize(TestContext testContext) { }
//        //
//        // Use ClassCleanup to run code after all tests in a class have run
//        // [ClassCleanup()]
//        // public static void MyClassCleanup() { }
//        //
//        // Use TestInitialize to run code before running each test 
//        //[TestInitialize()]
//        //public void MyTestInitialize() {}
//        //
//        // Use TestCleanup to run code after each test has run
//        // [TestCleanup()]
//        // public void MyTestCleanup() { }
//        //
//        #endregion Initialize and Cleanup

//        //#region Functions public
//        //[TestMethod]
//        //public void TxtService_Constructor_Test()
//        //{
//        //    foreach (string LanguageRequest in new List<string>() { "en", "fr" })
//        //    {
//        //        SetupTest(LanguageRequest);
//        //        Assert.IsNotNull(csspWebToolsTaskRunner);

//        //        csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";

//        //        csspWebToolsTaskRunner.StopTimer();

//        //        using (TransactionScope ts = new TransactionScope())
//        //        {
//        //            RemoveAllTask();

//        //            MWQMPlanModel mwqmPlanModel = AddMWQMPlanModel();
//        //            Assert.AreEqual("", mwqmPlanModel.Error);

//        //            SetupBWObjForMWQMPlan(mwqmPlanModel, FileGeneratorEnum.MWQMPlan, FileGeneratorTypeEnum.TXT, LanguageRequest);

//        //            // Act                   
//        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
//        //            Assert.AreEqual(1, countAppTask);

//        //            txtService = new TxtService(csspWebToolsTaskRunner._TaskRunnerBaseService);
//        //            Assert.IsNotNull(txtService);
//        //            Assert.IsNotNull(txtService._TaskRunnerBaseService);
//        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, txtService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
//        //        }
//        //    }
//        //}
//        //[TestMethod]
//        //public void TxtService_GenerateMWQMPlanConfigFileTXT_Test()
//        //{
//        //    foreach (string LanguageRequest in new List<string>() { "en" /*, "fr" */ })
//        //    {
//        //        SetupTest(LanguageRequest);
//        //        Assert.IsNotNull(csspWebToolsTaskRunner);

//        //        csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";

//        //        csspWebToolsTaskRunner.StopTimer();

//        //        using (TransactionScope ts = new TransactionScope())
//        //        {
//        //            RemoveAllTask();

//        //            MWQMPlanModel mwqmPlanModel = AddMWQMPlanModel();
//        //            Assert.AreEqual("", mwqmPlanModel.Error);

//        //            SetupBWObjForMWQMPlan(mwqmPlanModel, FileGeneratorEnum.MWQMPlan, FileGeneratorTypeEnum.TXT, LanguageRequest);

//        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
//        //            Assert.AreEqual(1, countAppTask);

//        //            txtService = new TxtService(csspWebToolsTaskRunner._TaskRunnerBaseService);
//        //            Assert.IsNotNull(txtService);
//        //            Assert.IsNotNull(txtService._TaskRunnerBaseService);
//        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, txtService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

//        //            txtService.GenerateMWQMPlanConfigFileTXT();
//        //            Assert.IsTrue(txtService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

//        //            string ServerFilePath = tvFileService.GetServerFilePath(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
//        //            string ServerFileName = csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.FileName;

//        //            TVItemModel tvItemModelFile = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(mwqmPlanModel.ProvinceTVItemID, ServerFileName, TVTypeEnum.File);
//        //            Assert.AreEqual("", tvItemModelFile.Error);

//        //            TVFileModel tvFileModel = tvFileService.GetTVFileModelWithTVFileTVItemIDDB(tvItemModelFile.TVItemID);
//        //            Assert.AreEqual("", tvFileModel.Error);

//        //            mwqmPlanModel.ConfigFileTxtTVItemID = tvFileModel.TVFileTVItemID;

//        //            MWQMPlanModel mwqmPlanModelRet = mwqmPlanService.PostUpdateMWQMPlanDB(mwqmPlanModel);
//        //            Assert.AreEqual("", mwqmPlanModelRet.Error);
//        //            Assert.AreEqual(tvFileModel.TVFileTVItemID, mwqmPlanModelRet.ConfigFileTxtTVItemID);

//        //            FileInfo fi = new FileInfo(ServerFilePath + ServerFileName);
//        //            Assert.IsTrue(fi.Exists);

//        //            if (fi.Exists)
//        //            {
//        //                fi.Delete();
//        //            }
//        //        }
//        //    }
//        //}
//        //#endregion Functions public

//        #region Functions private
//        //private MWQMPlanModel AddMWQMPlanModel()
//        //{
//        //    TVItemModel tvItemModelProvince = randomService.RandomTVItem(TVTypeEnum.Province);
//        //    Assert.AreEqual("", tvItemModelProvince.Error);

//        //    MWQMPlanModel mwqmPlanModelNew = new MWQMPlanModel()
//        //    {
//        //        ProvinceTVItemID = tvItemModelProvince.TVItemID,
//        //        ConfigFileName = randomService.RandomString("", 20),
//        //        ForGroupName = randomService.RandomString("", 20),
//        //        Year = randomService.RandomInt(2015, 2020),
//        //        SecretCode = randomService.RandomString("", 8),
//        //        CreatorTVItemID = 2,
//        //    };

//        //    MWQMPlanModel mwqmPlanModelRet = mwqmPlanService.PostAddMWQMPlanDB(mwqmPlanModelNew);
//        //    Assert.AreEqual("", mwqmPlanModelRet.Error);

//        //    List<TVItemModel> tvItemModelSubsectorList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelProvince.TVItemID, TVTypeEnum.Subsector);

//        //    int countSubsector = 0;
//        //    foreach (TVItemModel tvItemModelSubsector in tvItemModelSubsectorList)
//        //    {
//        //        List<TVItemModel> tvItemModelSubsectorSiteList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.MWQMSite);
//        //        if (tvItemModelSubsectorSiteList.Count > 0)
//        //        {
//        //            countSubsector += 1;
//        //            if (countSubsector > 3)
//        //            {
//        //                break;
//        //            }

//        //            MWQMPlanSubsectorModel mwqmPlanSubsectorModelNew = new MWQMPlanSubsectorModel()
//        //            {
//        //                MWQMPlanID = mwqmPlanModelRet.MWQMPlanID,
//        //                SubsectorTVItemID = tvItemModelSubsector.TVItemID,
//        //            };

//        //            MWQMPlanSubsectorModel mwqmPlanSubsectorModelRet = mwqmPlanSubsectorService.PostAddMWQMPlanSubsectorDB(mwqmPlanSubsectorModelNew);
//        //            Assert.AreEqual("", mwqmPlanSubsectorModelRet.Error);

//        //            int countSite = 0;
//        //            foreach (TVItemModel tvItemModelSusectorSite in tvItemModelSubsectorSiteList)
//        //            {
//        //                countSite += 1;
//        //                if (countSite > 3)
//        //                {
//        //                    break;
//        //                }
//        //                MWQMPlanSubsectorSiteModel mwqmPlanSubsectorSiteModelNew = new MWQMPlanSubsectorSiteModel()
//        //                {
//        //                    IsDuplicate = (countSite == 1 ? true : false),
//        //                    MWQMPlanSubsectorID = mwqmPlanSubsectorModelRet.MWQMPlanSubsectorID,
//        //                    MWQMSiteTVItemID = tvItemModelSusectorSite.TVItemID,
//        //                };

//        //                MWQMPlanSubsectorSiteModel mwqmPlanSubsectorSiteModelRet = mwqmPlanSubsectorSiteService.PostAddMWQMPlanSubsectorSiteDB(mwqmPlanSubsectorSiteModelNew);

//        //                Assert.AreEqual("", mwqmPlanSubsectorSiteModelRet.Error);

//        //            }
//        //        }
//        //    }

//        //    return mwqmPlanModelRet;
//        //}
//        //private void RemoveAllTask()
//        //{
//        //    List<AppTaskModel> appTaskModelList = appTaskService.GetAppTaskModelListDB();
//        //    foreach (AppTaskModel appTaskModel in appTaskModelList)
//        //    {
//        //        AppTaskModel appTaskModelRet = appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
//        //        Assert.AreEqual("", appTaskModelRet.Error);
//        //    }
//        //}
//        //private void SetupBWObjForMWQMPlan(MWQMPlanModel mwqmPlanModel, FileGeneratorEnum fileGenerator, FileGeneratorTypeEnum fileGeneratorType, string LanguageRequest)
//        //{
//        //    RemoveAllTask();

//        //    string ext = tvFileService.GetFileExtension(fileGeneratorType);
//        //    Assert.AreEqual(".txt", ext);

//        //    string ServerFilePath = tvFileService.GetServerFilePath(mwqmPlanModel.ProvinceTVItemID);
//        //    Assert.IsTrue(!string.IsNullOrWhiteSpace(ServerFilePath));

//        //    TVItemModel tvItemModel = tvItemService.GetTVItemModelWithTVItemIDForLocationDB(mwqmPlanModel.ProvinceTVItemID);
//        //    Assert.AreEqual("", tvItemModel.Error);

//        //    string FileName = "config_" + mwqmPlanModel.ConfigFileName + tvFileService.GetFileExtension(fileGeneratorType);
//        //    Assert.IsTrue(!string.IsNullOrWhiteSpace(FileName));

//        //    FileInfo fi = new FileInfo(ServerFilePath + FileName);
//        //    if (tvFileService.GetFileExist(fi))
//        //    {
//        //        fi.Delete();
//        //    }

//        //    fi = new FileInfo(ServerFilePath + FileName);
//        //    Assert.IsFalse(tvFileService.GetFileExist(fi));

//        //    mwqmPlanService.LanguageRequest = LanguageRequest;
//        //    string retStr = mwqmPlanService.MWQMPlanFileGenerateDB(mwqmPlanModel.MWQMPlanID, fileGenerator, fileGeneratorType);
//        //    Assert.AreEqual("", retStr);

//        //    using (ShimsContext.Create())
//        //    {
//        //        SetupShim();

//        //        shimTaskRunnerBaseService.GenerateDoc = () =>
//        //        {
//        //            return; // just so it does not start to Generate the documents
//        //        };
//        //        csspWebToolsTaskRunner.GetNextTask();

//        //        Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel);
//        //        Assert.AreEqual(mwqmPlanModel.ProvinceTVItemID, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
//        //        Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.bw);
//        //        Assert.IsFalse(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskCommand.ToString()));
//        //        Assert.IsTrue(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.Index > 0);
//        //        Assert.IsTrue(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.ErrorText));
//        //        Assert.IsTrue(!string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.StatusText));
//        //        Assert.AreEqual(mwqmPlanModel.ProvinceTVItemID, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
//        //        Assert.AreEqual(mwqmPlanModel.MWQMPlanID, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.MWQMPlanID);
//        //        Assert.AreEqual(fileGenerator, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.fileGenerator);
//        //        Assert.AreEqual(fileGeneratorType, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.fileGeneratorType);
//        //        Assert.AreEqual(true, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.Generate);
//        //        Assert.AreEqual(false, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.Command);
//        //        Assert.AreEqual(FileName, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.FileName);
//        //        Assert.AreEqual(true, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.bw.IsBusy);
//        //        Assert.AreEqual(false, csspWebToolsTaskRunner._BWList[0 + 1].bw.IsBusy);
//        //    }
//        //}
//        //public void SetupTest(string LanguageRequest)
//        //{
//        //    csspWebToolsTaskRunner = new CSSPWebToolsTaskRunner();
//        //    tvItemService = new TVItemService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
//        //    appTaskService = new AppTaskService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
//        //    tvFileService = new TVFileService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
//        //    mapInfoService = new MapInfoService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
//        //    randomService = new RandomService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
//        //    mwqmPlanService = new MWQMPlanService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
//        //    mwqmPlanSubsectorService = new MWQMPlanSubsectorService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
//        //    mwqmPlanSubsectorSiteService = new MWQMPlanSubsectorSiteService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
//        //}
//        //private void SetupShim()
//        //{
//        //    shimTaskRunnerBaseService = new ShimTaskRunnerBaseService(csspWebToolsTaskRunner._TaskRunnerBaseService);
//        //}
//        #endregion Functions private

//    }
//}

