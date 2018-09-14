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
        private HydrometricSiteService HydrometricSiteService { get; set; }
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
        [TestMethod]
        public void HydrometricService_Constructor_Test()
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
        //        if (appTaskCommand == AppTaskCommandEnum.UpdateHydrometricSiteInformation)
        //        {
        //            string FirstPart = "|||ProvinceTVItemID," + TVItemID.ToString() +
        //              "|||Generate,0" +
        //              "|||Command,1|||";
        //            Assert.AreEqual(FirstPart, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters);
        //            Assert.AreEqual(AppTaskCommandEnum.UpdateHydrometricSiteInformation, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskCommand);
        //        }
        //        else if (appTaskCommand == AppTaskCommandEnum.UpdateHydrometricSiteDailyAndHourlyFromStartDateToEndDate)
        //        {
        //            string FirstPart = "|||HydrometricSiteTVItemID," + TVItemID.ToString() +
        //              "|||StartYear," + StartDate.Year.ToString() +
        //              "|||StartMonth," + StartDate.Month.ToString() +
        //              "|||StartDay," + StartDate.Day.ToString() +
        //              "|||EndYear," + EndDate.Year.ToString() +
        //              "|||EndMonth," + EndDate.Month.ToString() +
        //              "|||EndDay," + EndDate.Day.ToString() +
        //              "|||Generate,0" +
        //              "|||Command,1|||";
        //            Assert.AreEqual(FirstPart, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters);
        //            Assert.AreEqual(AppTaskCommandEnum.UpdateHydrometricSiteDailyAndHourlyFromStartDateToEndDate, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskCommand);
        //        }
        //        else if (appTaskCommand == AppTaskCommandEnum.UpdateHydrometricSiteDailyAndHourlyForSubsectorFromStartDateToEndDate)
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
        //            Assert.AreEqual(AppTaskCommandEnum.UpdateHydrometricSiteDailyAndHourlyForSubsectorFromStartDateToEndDate, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskCommand);
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
        //    HydrometricSiteService = new HydrometricSiteService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        //    mwqmSiteService = new MWQMSiteService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        //    mwqmSampleService = new MWQMSampleService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        //}
        //private void SetupShim()
        //{
        //    shimTaskRunnerBaseService = new ShimTaskRunnerBaseService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //}
        #endregion Functions private

    }


}

