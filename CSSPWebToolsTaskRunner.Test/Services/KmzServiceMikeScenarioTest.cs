using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;
using CSSPWebToolsTaskRunner.Services;
using System.Linq;
using CSSPWebToolsDB.Models;
using System.ComponentModel;
using System.Transactions;
using CSSPWebToolsTaskRunner.Services.Fakes;
using Microsoft.QualityTools.Testing.Fakes;
using CSSPWebToolsDB.Services;
using System.IO;
using System.Web.Mvc;
using CSSPWebToolsDBDLL.Services;

namespace CSSPWebToolsTaskRunner.Test.Services
{
    /// <summary>
    /// Summary description for BaseMOdelServiceTest
    /// </summary>
    [TestClass]
    public class KmzServiceMikeScenarioTest
    {
        #region Variables
        private TestContext testContextInstance { get; set; }
        #endregion Variables

        #region Properties
        private CSSPWebToolsTaskRunner csspWebToolsTaskRunner { get; set; }
        private ShimTaskRunnerBaseService shimTaskRunnerBaseService { get; set; }
        private TVItemService tvItemService { get; set; }
        private AppTaskService appTaskService { get; set; }
        private TVFileService tvFileService { get; set; }
        private MapInfoService mapInfoService { get; set; }
        private KmzServiceMikeScenario kmzServiceMikeScenario { get; set; }
        private RandomService randomService { get; set; }
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
        public KmzServiceMikeScenarioTest()
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

        //#region Functions public
        //[TestMethod]
        //public void KmzServiceMikeScenario_Constructor_Test()
        //{
        //    foreach (string LanguageRequest in new List<string>() { "en", "fr" })
        //    {
        //        SetupTest(LanguageRequest);
        //        Assert.IsNotNull(csspWebToolsTaskRunner);

        //        csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";
        //        csspWebToolsTaskRunner.StopTimer();

        //        using (TransactionScope ts = new TransactionScope())
        //        {
        //            RemoveAllTask();

        //            TVItemModel tvItemModelRoot = tvItemService.GetRootTVItemModelForLocationDB();
        //            Assert.AreEqual("", tvItemModelRoot.Error);

        //            TVItemModel tvItemModelBouctouche = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, "Bouctouche", TVTypeEnum.Municipality);
        //            Assert.AreEqual("", tvItemModelBouctouche.Error);

        //            TVItemModel tvItemModelMikeScenario = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelBouctouche.TVItemID, TVTypeEnum.MikeScenario).FirstOrDefault();
        //            Assert.IsNotNull(tvItemModelMikeScenario);

        //            SetupBWObj(tvItemModelMikeScenario, FileGeneratorEnum.MIKEScenarioBoundaryConditions, FileGeneratorTypeEnum.KMZ, LanguageRequest, "", "", "", "", 0, "");

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            string PathName = tvFileService.GetServerFilePath(tvItemModelMikeScenario.TVItemID);
        //            string FileName = appTaskService.GetAppTaskParamStr(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "FileName");
        //            FileInfo fi = new FileInfo(PathName + FileName);
        //            Assert.IsNotNull(fi);

        //            kmzServiceMikeScenario = new KmzServiceMikeScenario(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(kmzServiceMikeScenario);
        //            Assert.IsNotNull(kmzServiceMikeScenario._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, kmzServiceMikeScenario._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
        //        }
        //    }
        //}
        //[TestMethod]
        //public void KmzServiceMikeScenario_GenerateMIKEScenarioBoundaryConditions_Test()
        //{
        //    foreach (string LanguageRequest in new List<string>() { "en", "fr" })
        //    {
        //        SetupTest(LanguageRequest);
        //        Assert.IsNotNull(csspWebToolsTaskRunner);

        //        csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";
        //        csspWebToolsTaskRunner.StopTimer();

        //        using (TransactionScope ts = new TransactionScope())
        //        {
        //            RemoveAllTask();

        //            TVItemModel tvItemModelRoot = tvItemService.GetRootTVItemModelForLocationDB();
        //            Assert.AreEqual("", tvItemModelRoot.Error);

        //            TVItemModel tvItemModelBouctouche = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, "Bouctouche", TVTypeEnum.Municipality);
        //            Assert.AreEqual("", tvItemModelBouctouche.Error);

        //            TVItemModel tvItemModelMikeScenario = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelBouctouche.TVItemID, TVTypeEnum.MikeScenario).FirstOrDefault();
        //            Assert.IsNotNull(tvItemModelMikeScenario);

        //            SetupBWObj(tvItemModelMikeScenario, FileGeneratorEnum.MIKEScenarioBoundaryConditions, FileGeneratorTypeEnum.KMZ, LanguageRequest, "", "", "", "", 0, "");

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            kmzServiceMikeScenario = new KmzServiceMikeScenario(csspWebToolsTaskRunner._TaskRunnerBaseService);

        //            string PathName = tvFileService.GetServerFilePath(tvItemModelMikeScenario.TVItemID);
        //            string FileName = appTaskService.GetAppTaskParamStr(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "FileName");
        //            FileInfo fi = new FileInfo(PathName + FileName);
        //            Assert.IsNotNull(fi);
        //            Assert.IsNotNull(kmzServiceMikeScenario);
        //            Assert.IsNotNull(kmzServiceMikeScenario._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, kmzServiceMikeScenario._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            kmzServiceMikeScenario.GenerateMikeScenarioBoundaryConditions(fi);
        //            Assert.IsTrue(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

        //            TVItemModel tvItemModelFile = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelMikeScenario.TVItemID, fi.Name, TVTypeEnum.File);
        //            Assert.AreEqual("", tvItemModelFile.Error);

        //            TVFileModel tvFileModel = tvFileService.GetTVFileModelWithTVFileTVItemIDDB(tvItemModelFile.TVItemID);
        //            Assert.AreEqual("", tvFileModel.Error);

        //            fi = new FileInfo(fi.FullName);
        //            Assert.IsTrue(fi.Exists);

        //            if (fi.Exists)
        //            {
        //                fi.Delete();
        //            }
        //        }
        //    }
        //}
        //[TestMethod]
        //public void KmzServiceMikeScenario_GenerateMikeScenarioConcentrationLimits_m21fm_Test()
        //{
        //    foreach (string LanguageRequest in new List<string>() { "en", "fr" })
        //    {
        //        SetupTest(LanguageRequest);
        //        Assert.IsNotNull(csspWebToolsTaskRunner);

        //        csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";
        //        csspWebToolsTaskRunner.StopTimer();

        //        using (TransactionScope ts = new TransactionScope())
        //        {
        //            RemoveAllTask();

        //            TVItemModel tvItemModelRoot = tvItemService.GetRootTVItemModelForLocationDB();
        //            Assert.AreEqual("", tvItemModelRoot.Error);

        //            TVItemModel tvItemModelBouctouche = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, "Bouctouche", TVTypeEnum.Municipality);
        //            Assert.AreEqual("", tvItemModelBouctouche.Error);

        //            TVItemModel tvItemModelMikeScenario = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelBouctouche.TVItemID, TVTypeEnum.MikeScenario).FirstOrDefault();
        //            Assert.IsNotNull(tvItemModelMikeScenario);

        //            SetupBWObj(tvItemModelMikeScenario, FileGeneratorEnum.MIKEScenarioConcentrationLimits, FileGeneratorTypeEnum.KMZ, LanguageRequest, "14,88", "", "", "", 0, "");

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            kmzServiceMikeScenario = new KmzServiceMikeScenario(csspWebToolsTaskRunner._TaskRunnerBaseService);

        //            string PathName = tvFileService.GetServerFilePath(tvItemModelMikeScenario.TVItemID);
        //            string FileName = appTaskService.GetAppTaskParamStr(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "FileName");
        //            FileInfo fi = new FileInfo(PathName + FileName);
        //            Assert.IsNotNull(fi);
        //            Assert.IsNotNull(kmzServiceMikeScenario);
        //            Assert.IsNotNull(kmzServiceMikeScenario._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, kmzServiceMikeScenario._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            kmzServiceMikeScenario.GenerateMikeScenarioConcentrationLimits(fi);
        //            Assert.IsTrue(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

        //            TVItemModel tvItemModelFile = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelMikeScenario.TVItemID, fi.Name, TVTypeEnum.File);
        //            Assert.AreEqual("", tvItemModelFile.Error);

        //            TVFileModel tvFileModel = tvFileService.GetTVFileModelWithTVFileTVItemIDDB(tvItemModelFile.TVItemID);
        //            Assert.AreEqual("", tvFileModel.Error);

        //            fi = new FileInfo(fi.FullName);
        //            Assert.IsTrue(fi.Exists);

        //            if (fi.Exists)
        //            {
        //                fi.Delete();
        //            }
        //        }
        //    }
        //}
        //#endregion Functions public

        #region Functions private
        //private void RemoveAllTask()
        //{
        //    List<AppTaskModel> appTaskModelList = appTaskService.GetAppTaskModelListDB();
        //    foreach (AppTaskModel appTaskModel in appTaskModelList)
        //    {
        //        // Act
        //        AppTaskModel appTaskModelRet = appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);

        //        // Assert
        //        Assert.AreEqual("", appTaskModelRet.Error);
        //    }
        //}
        //private void SetupBWObj(TVItemModel tvItemModel, FileGeneratorEnum fileGenerator, FileGeneratorTypeEnum fileGeneratorType,
        //    string LanguageRequest, string ContourValues, string SigmaLayerValues, string ZLayerValues, string DepthValues, 
        //    int VectorLength, string GoogleEarthPath)
        //{
        //    RemoveAllTask();

        //    // Act
        //    string ext = tvFileService.GetFileExtension(fileGeneratorType);

        //    // Assert
        //    Assert.AreEqual(".kmz", ext);

        //    // Act
        //    string ServerFilePath = tvFileService.GetServerFilePath(tvItemModel.TVItemID);

        //    // Assert
        //    Assert.IsTrue(!string.IsNullOrWhiteSpace(ServerFilePath));

        //    // Act
        //    string FileName = tvFileService.GenerateFileName(tvItemModel.TVItemID, fileGenerator, fileGeneratorType, ContourValues, LanguageRequest);

        //    // Assert
        //    Assert.IsTrue(!string.IsNullOrWhiteSpace(FileName));

        //    FileInfo fi = new FileInfo(tvFileService.ChoseEDriveOrCDrive(ServerFilePath) + FileName);
        //    if (tvFileService.GetFileExist(fi))
        //    {
        //        fi.Delete();
        //    }

        //    // Act
        //    fi = new FileInfo(tvFileService.ChoseEDriveOrCDrive(ServerFilePath) + FileName);

        //    // Assert
        //    Assert.IsFalse(tvFileService.GetFileExist(fi));

        //    // Act
        //    tvFileService.LanguageRequest = LanguageRequest;
        //    FormCollection fc = new FormCollection();
        //    fc.Add("TVItemID", tvItemModel.TVItemID.ToString());
        //    fc.Add("FileGenerator", ((int)fileGenerator).ToString());
        //    fc.Add("FileGeneratorType", ((int)fileGeneratorType).ToString());
        //    fc.Add("Generate", "1");
        //    fc.Add("Command", "0");
        //    fc.Add("ContourValues", ContourValues);
        //    fc.Add("SigmaLayerValues", SigmaLayerValues);
        //    fc.Add("ZLayerValues", ZLayerValues);
        //    fc.Add("DepthValues", DepthValues);
        //    fc.Add("VectorLength", VectorLength.ToString());
        //    fc.Add("GoogleEarthPath", GoogleEarthPath);
        //    fc.Add("FileName", fi.Name);
        //    string retStr = tvFileService.FileGenerateDB(fc);

        //    // Assert
        //    Assert.AreEqual("", retStr);

        //    using (ShimsContext.Create())
        //    {
        //        SetupShim();

        //        shimTaskRunnerBaseService.GenerateDoc = () =>
        //        {
        //            return; // just so it does not start to Generate the documents
        //        };
        //        // Act
        //        csspWebToolsTaskRunner.GetNextTask();

        //        // Assert
        //        //
        //        Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel);
        //        Assert.AreEqual(tvItemModel.TVItemID, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //        Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.bw);
        //        Assert.IsFalse(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskCommand.ToString()));
        //        Assert.IsTrue(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.Index > 0);
        //        Assert.IsTrue(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.ErrorText));
        //        Assert.IsTrue(!string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.StatusText));
        //        string Parameters = "|||TVItemID," + tvItemModel.TVItemID.ToString() +
        //            "|||FileGenerator," + ((int)fileGenerator).ToString() +
        //            "|||FileGeneratorType," + ((int)fileGeneratorType).ToString() +
        //            "|||Generate,1" +
        //            "|||Command,0" +
        //            "|||ContourValues," + ContourValues +
        //            "|||SigmaLayerValues," + SigmaLayerValues +
        //            "|||ZLayerValues," + ZLayerValues +
        //            "|||DepthValues," + DepthValues +
        //            "|||VectorLength," + VectorLength +
        //            "|||GoogleEarthPath," + GoogleEarthPath +
        //            "|||FileName," + FileName + "|||";
        //        Assert.AreEqual(Parameters, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters);
        //        Assert.AreEqual(tvItemModel.TVItemID, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //        Assert.AreEqual(fileGenerator, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.fileGenerator);
        //        Assert.AreEqual(fileGeneratorType, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.fileGeneratorType);
        //        Assert.AreEqual(true, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.Generate);
        //        Assert.AreEqual(false, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.Command);
        //        Assert.AreEqual(ContourValues, appTaskService.GetAppTaskParamStr(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "ContourValues"));
        //        Assert.AreEqual(SigmaLayerValues, appTaskService.GetAppTaskParamStr(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "SigmaLayerValues"));
        //        Assert.AreEqual(ZLayerValues, appTaskService.GetAppTaskParamStr(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "ZLayerValues"));
        //        Assert.AreEqual(DepthValues, appTaskService.GetAppTaskParamStr(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "DepthValues"));
        //        Assert.AreEqual(VectorLength.ToString(), appTaskService.GetAppTaskParamStr(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "VectorLength"));
        //        Assert.AreEqual(GoogleEarthPath, appTaskService.GetAppTaskParamStr(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "GoogleEarthPath"));
        //        Assert.AreEqual(FileName, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.FileName);
        //        Assert.AreEqual(false, csspWebToolsTaskRunner._BWList[0 + 1].bw.IsBusy);
        //    }

        //}
        //public void SetupTest(string LanguageRequest)
        //{
        //    csspWebToolsTaskRunner = new CSSPWebToolsTaskRunner();
        //    tvItemService = new TVItemService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        //    appTaskService = new AppTaskService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        //    tvFileService = new TVFileService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        //    mapInfoService = new MapInfoService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        //    randomService = new RandomService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        //}
        //private void SetupShim()
        //{
        //    shimTaskRunnerBaseService = new ShimTaskRunnerBaseService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //}
        #endregion Functions private

    }
}
