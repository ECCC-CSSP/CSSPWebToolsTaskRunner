using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using CSSPWebToolsDB.Models;
using System.Security.Principal;
using CSSPWebToolsDB.Services;
using CSSPWebToolsDB.Services.Resources;
using System.Transactions;
using CSSPWebToolsDB.Services.Fakes;
using Microsoft.QualityTools.Testing.Fakes;
using System.Threading;
using System.Globalization;
using CSSPWebToolsTaskRunner.Services;
using System.Linq;
using CSSPWebToolsTaskRunner.Fakes;
using CSSPWebToolsTaskRunner.Services.Fakes;
using System.IO;

namespace CSSPWebToolsTaskRunner.Test.App
{
    /// <summary>
    /// Summary description for AddressServiceTest
    /// </summary>
    [TestClass]
    public class CSSPWebToolsTaskRunnerTest
    {
        #region Variables
        private TestContext testContextInstance { get; set; }
        #endregion Variables

        #region Properties
        private CSSPWebToolsTaskRunner csspWebToolsTaskRunner { get; set; }
        private ShimTaskRunnerBaseService shimTaskRunnerBaseService { get; set; }
        private ShimCSSPWebToolsTaskRunner shimCSSPWebToolsTaskRunner { get; set; }
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
        public CSSPWebToolsTaskRunnerTest()
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

        #region Testing Methods Public
        [TestMethod]
        public void CSSPWebToolsTaskRunner_Constructor_Test()
        {
            // Arrange 
            SetupTest();

            // Assert
            Assert.IsNotNull(csspWebToolsTaskRunner);
            Assert.AreEqual(200, csspWebToolsTaskRunner._BWCount);
            Assert.IsNotNull(csspWebToolsTaskRunner._BWList);
            Assert.AreEqual(200, csspWebToolsTaskRunner._BWList.Count);
            Assert.IsNotNull(csspWebToolsTaskRunner._BWList[0].bw);
            Assert.IsNotNull(csspWebToolsTaskRunner._BWList[0].appTaskModel);
            Assert.AreEqual(0, csspWebToolsTaskRunner._BWList[0].Index);
            Assert.AreEqual(AppTaskCommandEnum.Error, csspWebToolsTaskRunner._BWList[0].appTaskCommand);
            Assert.AreEqual(FileGeneratorEnum.Error, csspWebToolsTaskRunner._BWList[0].fileGenerator);
            Assert.AreEqual(FileGeneratorTypeEnum.Error, csspWebToolsTaskRunner._BWList[0].fileGeneratorType);
            Assert.IsFalse(csspWebToolsTaskRunner._BWList[0].Generate);
            Assert.IsFalse(csspWebToolsTaskRunner._BWList[0].Command);
            Assert.AreEqual("", csspWebToolsTaskRunner._BWList[0].FileName);
            Assert.IsNotNull(csspWebToolsTaskRunner._RichTextBoxStatus);
            Assert.IsTrue(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._RichTextBoxStatus.Text));
            Assert.AreEqual(0, csspWebToolsTaskRunner._SkipTimerCount);
            Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService);
            Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._AppTaskLanguageService);
            Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._AppTaskService);
            Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._User);
            Assert.AreEqual("Charles.LeBlanc@ec.gc.ca", csspWebToolsTaskRunner._TaskRunnerBaseService._User.Identity.Name);
            Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._MapInfoService);
            Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._MikeBoundaryConditionService);
            Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._MikeScenarioService);
            Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._MikeSourceService);
            Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._MikeSourceStartEndService);
            Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._TVFileService);
            Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._TVItemLanguageService);
            Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._TVItemService);
            Assert.IsNull(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj);
            Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObjList);
            Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._TaskRunnerBaseService);
            Assert.IsNotNull(csspWebToolsTaskRunner._TimerCheckTask);
            Assert.IsTrue(csspWebToolsTaskRunner._TimerCheckTask.Enabled);
            Assert.AreEqual(1000, csspWebToolsTaskRunner._TimerCheckTask.Interval);
            Assert.IsNotNull(csspWebToolsTaskRunner._User);
            Assert.AreEqual("Charles.LeBlanc@ec.gc.ca", csspWebToolsTaskRunner._User.Identity.Name);
        }
        [TestMethod]
        public void CSSPWebToolsTaskRunner_AppInit_Test()
        {
            // Arrange 
            SetupTest();

            // Already testing in constructor

        }
        [TestMethod]
        public void CSSPWebToolsTaskRunner_BWDoWork_Generate_Root_Excel_Test()
        {
            // Arrange 
            SetupTest();

            // Assert
            Assert.IsNotNull(csspWebToolsTaskRunner);

            // Arrange
            csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";

            // Act
            csspWebToolsTaskRunner.StopTimer();

            using (TransactionScope ts = new TransactionScope())
            {
                RemoveAllTask();

                // Act
                TVItemModel tvItemModelRoot = csspWebToolsTaskRunner._TaskRunnerBaseService._TVItemService.GetRootTVItemModelForLocationDB();

                FileGeneratorEnum fileGenerator = FileGeneratorEnum.Root;
                FileGeneratorTypeEnum fileGeneratorType = FileGeneratorTypeEnum.Excel;

                // Act
                List<FileGeneratorTypeAndText> fileGeneratorTypeAndTextList = csspWebToolsTaskRunner._TaskRunnerBaseService._TVFileService.GetFileGeneratorTypeAndText();
                string ext = fileGeneratorTypeAndTextList.Where(c => c.FileGeneratorType == fileGeneratorType).FirstOrDefault().FileGeneratorExtText;
                string retStr = csspWebToolsTaskRunner._TaskRunnerBaseService._TVFileService.FileGenerateDB(tvItemModelRoot.TVItemID, fileGenerator, fileGeneratorType);

                // Assert
                Assert.AreEqual("", retStr);

                string FileName = csspWebToolsTaskRunner._TaskRunnerBaseService._TVFileService.GenerateFileName(tvItemModelRoot, fileGenerator, fileGeneratorType);

                using (ShimsContext.Create())
                {
                    SetupShim();

                    shimTaskRunnerBaseService.GenerateDoc = () =>
                    {
                        return ""; // just so it does not start to Generate the documents
                    };
                    // Act
                    csspWebToolsTaskRunner.GetNextTask();

                    // Assert
                    //
                    Assert.IsNotNull(csspWebToolsTaskRunner._BWList[0].appTaskModel);
                    Assert.AreEqual(tvItemModelRoot.TVItemID, csspWebToolsTaskRunner._BWList[0].appTaskModel.TVItemID);
                    Assert.IsNotNull(csspWebToolsTaskRunner._BWList[0].bw);
                    Assert.IsFalse(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._BWList[0].appTaskCommand.ToString()));
                    Assert.IsTrue(csspWebToolsTaskRunner._BWList[0].Index > 0);
                    Assert.AreEqual(tvItemModelRoot.TVItemID, csspWebToolsTaskRunner._BWList[0].appTaskModel.TVItemID);
                    Assert.IsFalse(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._BWList[0].appTaskModel.BusyText));
                    Assert.IsTrue(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._BWList[0].appTaskModel.ErrorText));
                    Assert.IsTrue(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._BWList[0].appTaskModel.StatusText));
                    string FirstPart = "|||TVItemID," + tvItemModelRoot.TVItemID.ToString() +
                        "|||FileGenerator," + ((int)fileGenerator).ToString() +
                        "|||FileGeneratorType," + ((int)fileGeneratorType).ToString() +
                        "|||Generate,1" +
                        "|||Command,0" +
                        "|||FileName," + FileName;
                    Assert.AreEqual(FirstPart, csspWebToolsTaskRunner._BWList[0].appTaskModel.Parameters.Substring(0, FirstPart.Length));
                    string LastPart = "_" + csspWebToolsTaskRunner._BWList[0].appTaskModel.Language + ext.ToLower() + "|||";
                    int ParamLength = csspWebToolsTaskRunner._BWList[0].appTaskModel.Parameters.Length;
                    Assert.AreEqual(LastPart, csspWebToolsTaskRunner._BWList[0].appTaskModel.Parameters.Substring(ParamLength - LastPart.Length));
                    Assert.AreEqual(fileGenerator, csspWebToolsTaskRunner._BWList[0].fileGenerator);
                    Assert.AreEqual(fileGeneratorType, csspWebToolsTaskRunner._BWList[0].fileGeneratorType);
                    Assert.AreEqual(true, csspWebToolsTaskRunner._BWList[0].Generate);
                    Assert.AreEqual(false, csspWebToolsTaskRunner._BWList[0].Command);
                    Assert.AreEqual(FileName, csspWebToolsTaskRunner._BWList[0].FileName);
                    Assert.AreEqual(true, csspWebToolsTaskRunner._BWList[0].bw.IsBusy);
                    Assert.AreEqual(false, csspWebToolsTaskRunner._BWList[0 + 1].bw.IsBusy);
                }

                // Act
                AppTaskModel appTaskModel = csspWebToolsTaskRunner._TaskRunnerBaseService._AppTaskService.CheckAppTask();

                // Assert
                Assert.AreEqual("", appTaskModel.Error);

                // Act
                csspWebToolsTaskRunner.BWDoWork(csspWebToolsTaskRunner._BWList[0]);

                // Act
                List<TVItemModel> tvItemModelList = csspWebToolsTaskRunner._TaskRunnerBaseService._TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelRoot.TVItemID, TVTypeEnum.File);

                // Assert
                Assert.IsTrue(tvItemModelList.Count > 0);
                Assert.IsTrue(tvItemModelList.Where(c => c.TVText.Contains(FileName)).Any());

                // Act
                appTaskModel = csspWebToolsTaskRunner._TaskRunnerBaseService._AppTaskService.CheckAppTask();

                // Assert
                Assert.AreEqual(ServiceRes.NoNewTaskToRun, appTaskModel.Error);

                string ServerFilePath = csspWebToolsTaskRunner._TaskRunnerBaseService._TVFileService.GetServerFilePath(tvItemModelRoot.TVItemID);

                FileInfo fi = new FileInfo(ServerFilePath + FileName);

                if (fi.Exists)
                {
                    fi.Delete();
                }

            }
        }
        [TestMethod]
        public void CSSPWebToolsTaskRunner_BWDoWork_Command_Test()
        {
            // Arrange 
            SetupTest();

            // Assert
            Assert.IsNotNull(csspWebToolsTaskRunner);

            // Arrange
            csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";

            // Act
            csspWebToolsTaskRunner.StopTimer();

            using (TransactionScope ts = new TransactionScope())
            {
                RemoveAllTask();

                // Act
                TVItemModel tvItemModelRoot = csspWebToolsTaskRunner._TaskRunnerBaseService._TVItemService.GetRootTVItemModelForLocationDB();

                // Assert
                Assert.AreEqual("", tvItemModelRoot.Error);

                // Act
                TVItemModel tvItemModelBlacksHarbour = csspWebToolsTaskRunner._TaskRunnerBaseService._TVItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, "Blacks Harbour", TVTypeEnum.MikeScenario);

                // Assert
                Assert.AreEqual("", tvItemModelBlacksHarbour.Error);

                List<AppTaskParameter> appTaskParameterList = new List<AppTaskParameter>();
                appTaskParameterList.Add(new AppTaskParameter() { Name = "MikeScenarioTVItemID", Value = tvItemModelBlacksHarbour.TVItemID.ToString() });
                appTaskParameterList.Add(new AppTaskParameter() { Name = "Generate", Value = "0" });
                appTaskParameterList.Add(new AppTaskParameter() { Name = "Command", Value = "1" });

                StringBuilder sbParameters = new StringBuilder();
                int count = 0;
                foreach (AppTaskParameter atp in appTaskParameterList)
                {
                    if (count == 0)
                    {
                        sbParameters.Append("|||");
                    }
                    sbParameters.Append(atp.Name + "," + atp.Value + "|||");
                    count += 1;
                }

                AppTaskModel appTaskModelNew = new AppTaskModel()
                {
                    TVItemID = tvItemModelBlacksHarbour.TVItemID,
                    TVItemID2 = tvItemModelBlacksHarbour.TVItemID,
                    Command = AppTaskCommandEnum.MikeScenarioAskToRun,
                    BusyText = ServiceRes.AskingToRunMIKEScenario,
                    ErrorText = "",
                    StatusText = "",
                    Status = AppTaskStatusEnum.Created,
                    PercentCompleted = 1,
                    Parameters = sbParameters.ToString(),
                    Language = "en",
                    StartDateTime_UTC = DateTime.UtcNow,
                    EndDateTime_UTC = null,
                    EstimatedLength_second = null,
                    RemainingTime_second = null,
                };

                // Act
                AppTaskModel appTaskModelRet = csspWebToolsTaskRunner._TaskRunnerBaseService._AppTaskService.PostAddAppTask(appTaskModelNew);

                // Assert
                Assert.AreEqual("", appTaskModelRet.Error);

                using (ShimsContext.Create())
                {
                    SetupShim();

                    shimTaskRunnerBaseService.DoCommand = () =>
                    {
                        return ""; // just so it does not start to Generate the documents
                    };
                    // Act
                    csspWebToolsTaskRunner.GetNextTask();

                    // Assert
                    //
                    Assert.IsNotNull(csspWebToolsTaskRunner._BWList[0].appTaskModel);
                    Assert.AreEqual(tvItemModelBlacksHarbour.TVItemID, csspWebToolsTaskRunner._BWList[0].appTaskModel.TVItemID);
                    Assert.IsNotNull(csspWebToolsTaskRunner._BWList[0].bw);
                    Assert.IsFalse(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._BWList[0].appTaskCommand.ToString()));
                    Assert.IsTrue(csspWebToolsTaskRunner._BWList[0].Index > 0);
                    Assert.AreEqual(tvItemModelBlacksHarbour.TVItemID, csspWebToolsTaskRunner._BWList[0].appTaskModel.TVItemID);
                    Assert.IsFalse(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._BWList[0].appTaskModel.BusyText));
                    Assert.IsTrue(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._BWList[0].appTaskModel.ErrorText));
                    Assert.IsTrue(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._BWList[0].appTaskModel.StatusText));
                    string Parameters = "|||MikeScenarioTVItemID," + tvItemModelBlacksHarbour.TVItemID + "|||Generate,0|||Command,1|||";
                    Assert.AreEqual(Parameters, csspWebToolsTaskRunner._BWList[0].appTaskModel.Parameters);
                    Assert.AreEqual(FileGeneratorEnum.Error, csspWebToolsTaskRunner._BWList[0].fileGenerator);
                    Assert.AreEqual(FileGeneratorTypeEnum.Error, csspWebToolsTaskRunner._BWList[0].fileGeneratorType);
                    Assert.AreEqual(false, csspWebToolsTaskRunner._BWList[0].Generate);
                    Assert.AreEqual(true, csspWebToolsTaskRunner._BWList[0].Command);
                    Assert.AreEqual(true, csspWebToolsTaskRunner._BWList[0].bw.IsBusy);
                    Assert.AreEqual(false, csspWebToolsTaskRunner._BWList[0 + 1].bw.IsBusy);
                }

                // Act
                AppTaskModel appTaskModel = csspWebToolsTaskRunner._TaskRunnerBaseService._AppTaskService.CheckAppTask();

                // Assert
                Assert.AreEqual("", appTaskModel.Error);

                // Act
                csspWebToolsTaskRunner.BWDoWork(csspWebToolsTaskRunner._BWList[0]);

                // Act
                appTaskModel = csspWebToolsTaskRunner._TaskRunnerBaseService._AppTaskService.CheckAppTask();

                // Assert
                Assert.AreEqual(ServiceRes.NoNewTaskToRun, appTaskModel.Error);
            }
        }
        [TestMethod]
        public void CSSPWebToolsTaskRunner_BWProgressChanged_Test()
        {
            // Arrange 
            SetupTest();

            // Assert
            Assert.IsNotNull(csspWebToolsTaskRunner);

            // Arrange
            csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";

            // Act
            csspWebToolsTaskRunner.StopTimer();

            using (TransactionScope ts = new TransactionScope())
            {
                // Act
                UserStateObj userStateObj = new UserStateObj()
                {
                    ProgressText = "ProgressText",
                };
                int Percentage = 34;
                csspWebToolsTaskRunner.BWProgressChanged(Percentage, userStateObj);

                // Assert
                Assert.AreEqual("(" + Percentage + ") % --- " + userStateObj.ProgressText + "\n", csspWebToolsTaskRunner._RichTextBoxStatus.Text);
            }
        }
        [TestMethod]
        public void CSSPWebToolsTaskRunner_BWRunWorkerCompleted_Test()
        {
            // Arrange 
            SetupTest();

            // Assert
            Assert.IsNotNull(csspWebToolsTaskRunner);

            // Arrange
            csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";

            // Act
            csspWebToolsTaskRunner.StopTimer();

            using (TransactionScope ts = new TransactionScope())
            {
                // Act
                csspWebToolsTaskRunner.BWRunWorkerCompleted();

                // Assert
                Assert.AreEqual("\n", csspWebToolsTaskRunner._RichTextBoxStatus.Text);
            }
        }
        [TestMethod]
        public void CSSPWebToolsTaskRunner_GetNextTask_And_BWTryRunningTask_Test()
        {
            // Arrange 
            SetupTest();

            // Assert
            Assert.IsNotNull(csspWebToolsTaskRunner);

            // Arrange
            csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";

            // Act
            csspWebToolsTaskRunner.StopTimer();

            List<FileGeneratorTypeAndText> fileGeneratorTypeAndTextList = csspWebToolsTaskRunner._TaskRunnerBaseService._TVFileService.GetFileGeneratorTypeAndText();
            int count = 0;
            for (int i = 1, FileGeneratorEnumCount = Enum.GetNames(typeof(FileGeneratorEnum)).Count(); i < FileGeneratorEnumCount; i++)
            {
                for (int j = 1, FileGeneratorTypeEnumCount = Enum.GetNames(typeof(FileGeneratorTypeEnum)).Count(); j < FileGeneratorTypeEnumCount; j++)
                {
                    using (TransactionScope ts = new TransactionScope())
                    {
                        RemoveAllTask();

                        TVItemModel tvItemModel = new TVItemModel();

                        // Act
                        switch (((FileGeneratorEnum)i).ToString().Substring(0, 4))
                        {
                            case "Root":
                                {
                                    tvItemModel = csspWebToolsTaskRunner._TaskRunnerBaseService._TVItemService.GetTVItemModelListWithTVTypeDB(TVTypeEnum.Root).FirstOrDefault();
                                }
                                break;
                            case "Coun":
                                {
                                    tvItemModel = csspWebToolsTaskRunner._TaskRunnerBaseService._TVItemService.GetTVItemModelListWithTVTypeDB(TVTypeEnum.Country).FirstOrDefault();
                                }
                                break;
                            case "Prov":
                                {
                                    tvItemModel = csspWebToolsTaskRunner._TaskRunnerBaseService._TVItemService.GetTVItemModelListWithTVTypeDB(TVTypeEnum.Province).FirstOrDefault();
                                }
                                break;
                            case "Area":
                                {
                                    tvItemModel = csspWebToolsTaskRunner._TaskRunnerBaseService._TVItemService.GetTVItemModelListWithTVTypeDB(TVTypeEnum.Area).FirstOrDefault();
                                }
                                break;
                            case "Sect":
                                {
                                    tvItemModel = csspWebToolsTaskRunner._TaskRunnerBaseService._TVItemService.GetTVItemModelListWithTVTypeDB(TVTypeEnum.Sector).FirstOrDefault();
                                }
                                break;
                            case "Subs":
                                {
                                    tvItemModel = csspWebToolsTaskRunner._TaskRunnerBaseService._TVItemService.GetTVItemModelListWithTVTypeDB(TVTypeEnum.Subsector).FirstOrDefault();
                                }
                                break;
                            case "Muni":
                                {
                                    tvItemModel = csspWebToolsTaskRunner._TaskRunnerBaseService._TVItemService.GetTVItemModelListWithTVTypeDB(TVTypeEnum.Municipality).FirstOrDefault();
                                }
                                break;
                            case "Infr":
                                {
                                    tvItemModel = csspWebToolsTaskRunner._TaskRunnerBaseService._TVItemService.GetTVItemModelListWithTVTypeDB(TVTypeEnum.Infrastructure).FirstOrDefault();
                                }
                                break;
                            case "PolS":
                                {
                                    tvItemModel = csspWebToolsTaskRunner._TaskRunnerBaseService._TVItemService.GetTVItemModelListWithTVTypeDB(TVTypeEnum.PolSourceSite).FirstOrDefault();
                                }
                                break;
                            case "MWQM":
                                {
                                    tvItemModel = csspWebToolsTaskRunner._TaskRunnerBaseService._TVItemService.GetTVItemModelListWithTVTypeDB(TVTypeEnum.MWQMSite).FirstOrDefault();
                                }
                                break;
                            default:
                                {
                                    continue;
                                }
                        }

                        // Act
                        string ext = fileGeneratorTypeAndTextList.Where(c => c.FileGeneratorType == (FileGeneratorTypeEnum)j).FirstOrDefault().FileGeneratorExtText.ToLower();
                        string retStr = csspWebToolsTaskRunner._TaskRunnerBaseService._TVFileService.FileGenerateDB(tvItemModel.TVItemID, (FileGeneratorEnum)i, (FileGeneratorTypeEnum)j);

                        // Assert
                        Assert.AreEqual("", retStr);

                        PrivateObject privateObject = new PrivateObject(csspWebToolsTaskRunner._TaskRunnerBaseService._TVFileService);
                        string FileName = (string)privateObject.Invoke("GenerateFileName", tvItemModel, (FileGeneratorEnum)i, (FileGeneratorTypeEnum)j);

                        using (ShimsContext.Create())
                        {
                            SetupShim();

                            shimTaskRunnerBaseService.GenerateDoc = () =>
                            {
                                return ""; // just so it does not start to Generate the documents
                            };

                            // Act
                            csspWebToolsTaskRunner.GetNextTask();

                            // Assert
                            Assert.IsNotNull(csspWebToolsTaskRunner._BWList[count].appTaskModel);
                            Assert.AreEqual(tvItemModel.TVItemID, csspWebToolsTaskRunner._BWList[count].appTaskModel.TVItemID);
                            Assert.IsNotNull(csspWebToolsTaskRunner._BWList[count].bw);
                            Assert.IsFalse(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._BWList[count].appTaskCommand.ToString()));
                            Assert.IsTrue(csspWebToolsTaskRunner._BWList[count].Index > 0);
                            Assert.AreEqual(tvItemModel.TVItemID, csspWebToolsTaskRunner._BWList[count].appTaskModel.TVItemID);
                            Assert.IsFalse(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._BWList[count].appTaskModel.BusyText));
                            Assert.IsTrue(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._BWList[count].appTaskModel.ErrorText));
                            Assert.IsTrue(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._BWList[count].appTaskModel.StatusText));
                            string FirstPart = "|||TVItemID," + tvItemModel.TVItemID.ToString() +
                                "|||FileGenerator," + i.ToString() +
                                "|||FileGeneratorType," + j.ToString() +
                                "|||Generate,1" +
                                "|||Command,0" +
                                "|||FileName," + FileName;
                            Assert.AreEqual(FirstPart, csspWebToolsTaskRunner._BWList[count].appTaskModel.Parameters.Substring(0, FirstPart.Length));
                            string LastPart = "_" + csspWebToolsTaskRunner._BWList[count].appTaskModel.Language + ext.ToLower() + "|||";
                            int ParamLength = csspWebToolsTaskRunner._BWList[count].appTaskModel.Parameters.Length;
                            Assert.AreEqual(LastPart, csspWebToolsTaskRunner._BWList[count].appTaskModel.Parameters.Substring(ParamLength - LastPart.Length));
                            Assert.AreEqual((FileGeneratorEnum)i, csspWebToolsTaskRunner._BWList[count].fileGenerator);
                            Assert.AreEqual((FileGeneratorTypeEnum)j, csspWebToolsTaskRunner._BWList[count].fileGeneratorType);
                            Assert.AreEqual(true, csspWebToolsTaskRunner._BWList[count].Generate);
                            Assert.AreEqual(false, csspWebToolsTaskRunner._BWList[count].Command);
                            Assert.AreEqual(FileName, csspWebToolsTaskRunner._BWList[count].FileName);
                            Assert.AreEqual(true, csspWebToolsTaskRunner._BWList[count].bw.IsBusy);
                            Assert.AreEqual(false, csspWebToolsTaskRunner._BWList[count + 1].bw.IsBusy);
                        }

                        string ServerFilePath = csspWebToolsTaskRunner._TaskRunnerBaseService._TVFileService.GetServerFilePath(tvItemModel.TVItemID);

                        FileInfo fi = new FileInfo(ServerFilePath + FileName);

                        if (fi.Exists)
                        {
                            fi.Delete();
                        }
                    }
                    count += 1;
                }
            }
        }
        [TestMethod]
        public void CSSPWebToolsTaskRunner_DoTimerCheckTask_Test()
        {
            // Arrange 
            SetupTest();

            // Assert
            Assert.IsNotNull(csspWebToolsTaskRunner);

            // Assert
            Assert.IsNotNull(csspWebToolsTaskRunner._TimerCheckTask);
            Assert.IsTrue(csspWebToolsTaskRunner._TimerCheckTask.Enabled);
            Assert.AreEqual(1000, csspWebToolsTaskRunner._TimerCheckTask.Interval);

            using (ShimsContext.Create())
            {
                SetupShim();
                shimCSSPWebToolsTaskRunner.GetNextTask = () =>
                {
                    // Assert
                    Assert.IsFalse(csspWebToolsTaskRunner._TimerCheckTask.Enabled);
                };

                csspWebToolsTaskRunner.DoTimerCheckTask();

                // Assert
                Assert.IsTrue(csspWebToolsTaskRunner._TimerCheckTask.Enabled);
            }
        }
        [TestMethod]
        public void CSSPWebToolsTaskRunner_GetNextTask_Test()
        {
            // Arrange 
            SetupTest();

            // Assert
            Assert.IsNotNull(csspWebToolsTaskRunner);

            // Arrange
            csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";

            // Act
            csspWebToolsTaskRunner.StopTimer();

            using (TransactionScope ts = new TransactionScope())
            {
                RemoveAllTask();

                csspWebToolsTaskRunner.GetNextTask();

                // Assert
                Assert.IsTrue(csspWebToolsTaskRunner._LabelLastAppTaskCheckDate.Text.Contains(ServiceRes.NoNewTaskToRun));

                // Act
                TVItemModel tvItemModelRoot = csspWebToolsTaskRunner._TaskRunnerBaseService._TVItemService.GetRootTVItemModelForLocationDB();

                // Assert
                Assert.AreEqual("", tvItemModelRoot.Error);

                // Act
                TVItemModel tvItemModelBlacksHarbour = csspWebToolsTaskRunner._TaskRunnerBaseService._TVItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, "Blacks Harbour", TVTypeEnum.MikeScenario);

                // Assert
                Assert.AreEqual("", tvItemModelBlacksHarbour.Error);

                List<AppTaskParameter> appTaskParameterList = new List<AppTaskParameter>();
                appTaskParameterList.Add(new AppTaskParameter() { Name = "MikeScenarioTVItemID", Value = tvItemModelBlacksHarbour.TVItemID.ToString() });
                appTaskParameterList.Add(new AppTaskParameter() { Name = "Generate", Value = "0" });
                appTaskParameterList.Add(new AppTaskParameter() { Name = "Command", Value = "1" });

                StringBuilder sbParameters = new StringBuilder();
                int count = 0;
                foreach (AppTaskParameter atp in appTaskParameterList)
                {
                    if (count == 0)
                    {
                        sbParameters.Append("|||");
                    }
                    sbParameters.Append(atp.Name + "," + atp.Value + "|||");
                    count += 1;
                }

                AppTaskModel appTaskModelNew = new AppTaskModel()
                {
                    TVItemID = tvItemModelBlacksHarbour.TVItemID,
                    TVItemID2 = tvItemModelBlacksHarbour.TVItemID,
                    Command = AppTaskCommandEnum.MikeScenarioAskToRun,
                    BusyText = ServiceRes.AskingToRunMIKEScenario,
                    ErrorText = "",
                    StatusText = "",
                    Status = AppTaskStatusEnum.Created,
                    PercentCompleted = 1,
                    Parameters = sbParameters.ToString(),
                    Language = "en",
                    StartDateTime_UTC = DateTime.UtcNow,
                    EndDateTime_UTC = null,
                    EstimatedLength_second = null,
                    RemainingTime_second = null,
                };

                // Act
                AppTaskModel appTaskModelRet = csspWebToolsTaskRunner._TaskRunnerBaseService._AppTaskService.PostAddAppTask(appTaskModelNew);

                // Assert
                Assert.AreEqual("", appTaskModelRet.Error);

                using (ShimsContext.Create())
                {
                    SetupShim();

                    shimTaskRunnerBaseService.DoCommand = () =>
                    {
                        return ""; // just so it does not start to Generate the documents
                    };
                    // Act
                    csspWebToolsTaskRunner.GetNextTask();

                    // Assert
                    //
                    Assert.IsNotNull(csspWebToolsTaskRunner._BWList[0].appTaskModel);
                    Assert.AreEqual(tvItemModelBlacksHarbour.TVItemID, csspWebToolsTaskRunner._BWList[0].appTaskModel.TVItemID);
                    Assert.IsNotNull(csspWebToolsTaskRunner._BWList[0].bw);
                    Assert.IsFalse(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._BWList[0].appTaskCommand.ToString()));
                    Assert.IsTrue(csspWebToolsTaskRunner._BWList[0].Index > 0);
                    Assert.AreEqual(tvItemModelBlacksHarbour.TVItemID, csspWebToolsTaskRunner._BWList[0].appTaskModel.TVItemID);
                    Assert.IsFalse(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._BWList[0].appTaskModel.BusyText));
                    Assert.IsTrue(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._BWList[0].appTaskModel.ErrorText));
                    Assert.IsTrue(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._BWList[0].appTaskModel.StatusText));
                    string Parameters = "|||MikeScenarioTVItemID," + tvItemModelBlacksHarbour.TVItemID + "|||Generate,0|||Command,1|||";
                    Assert.AreEqual(Parameters, csspWebToolsTaskRunner._BWList[0].appTaskModel.Parameters);
                    Assert.AreEqual(FileGeneratorEnum.Error, csspWebToolsTaskRunner._BWList[0].fileGenerator);
                    Assert.AreEqual(FileGeneratorTypeEnum.Error, csspWebToolsTaskRunner._BWList[0].fileGeneratorType);
                    Assert.AreEqual(false, csspWebToolsTaskRunner._BWList[0].Generate);
                    Assert.AreEqual(true, csspWebToolsTaskRunner._BWList[0].Command);
                    Assert.AreEqual(true, csspWebToolsTaskRunner._BWList[0].bw.IsBusy);
                    Assert.AreEqual(false, csspWebToolsTaskRunner._BWList[0 + 1].bw.IsBusy);
                }
            }
        }
        [TestMethod]
        public void CSSPWebToolsTaskRunner_StartAndStop_Timer_Test()
        {
            // Arrange 
            SetupTest();

            // Assert
            Assert.IsNotNull(csspWebToolsTaskRunner);

            // Assert
            Assert.IsNotNull(csspWebToolsTaskRunner._TimerCheckTask);
            Assert.IsTrue(csspWebToolsTaskRunner._TimerCheckTask.Enabled);
            Assert.AreEqual(1000, csspWebToolsTaskRunner._TimerCheckTask.Interval);

            // Act
            csspWebToolsTaskRunner.StopTimer();

            // Assert
            Assert.IsNotNull(csspWebToolsTaskRunner._TimerCheckTask);
            Assert.IsFalse(csspWebToolsTaskRunner._TimerCheckTask.Enabled);
            Assert.AreEqual(1000, csspWebToolsTaskRunner._TimerCheckTask.Interval);

            // Act
            csspWebToolsTaskRunner.StartTimer();

            // Assert
            Assert.IsNotNull(csspWebToolsTaskRunner._TimerCheckTask);
            Assert.IsTrue(csspWebToolsTaskRunner._TimerCheckTask.Enabled);
            Assert.AreEqual(1000, csspWebToolsTaskRunner._TimerCheckTask.Interval);

            //
            csspWebToolsTaskRunner.StopTimer();

        }
        #endregion Testing Methods Public

        #region Testing Methods Private
        #endregion Testing Methods Private

        #region Functions private
        private void RemoveAllTask()
        {
            List<AppTaskModel> appTaskModelList = csspWebToolsTaskRunner._TaskRunnerBaseService._AppTaskService.GetAppTaskModelListDB();
            foreach (AppTaskModel appTaskModel in appTaskModelList)
            {
                // Act
                AppTaskModel appTaskModelRet = csspWebToolsTaskRunner._TaskRunnerBaseService._AppTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);

                // Assert
                Assert.AreEqual("", appTaskModelRet.Error);
            }
        }
        public void SetupTest()
        {
            csspWebToolsTaskRunner = new CSSPWebToolsTaskRunner();
        }
        private void SetupShim()
        {
            shimTaskRunnerBaseService = new ShimTaskRunnerBaseService(csspWebToolsTaskRunner._TaskRunnerBaseService);
            shimCSSPWebToolsTaskRunner = new ShimCSSPWebToolsTaskRunner(csspWebToolsTaskRunner);
        }
        #endregion Functions private
    }
}
