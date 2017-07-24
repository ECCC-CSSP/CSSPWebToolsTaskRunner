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
    public class HtmlServiceTest
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
        //private HtmlService htmlService { get; set; }
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
        public HtmlServiceTest()
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
        //public void HtmlService_Constructor_Test()
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

        //            SetupBWObj(tvItemModelRoot, FileGeneratorEnum.Root, FileGeneratorTypeEnum.HTML, LanguageRequest);

        //            // Act                   
        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            htmlService = new HtmlService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(htmlService);
        //            Assert.IsNotNull(htmlService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, htmlService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
        //        }
        //    }
        //}
        //[TestMethod]
        //public void HtmlService_GenerateRootHTML_TVItemID_1_Test()
        //{
        //    // Generated files will be located
        //    // \inetpub\wwwroot\csspwebtools\App_Data\1
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

        //            SetupBWObj(tvItemModelRoot, FileGeneratorEnum.Root, FileGeneratorTypeEnum.HTML, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            htmlService = new HtmlService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(htmlService);
        //            Assert.IsNotNull(htmlService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, htmlService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            htmlService.GenerateRootHTML();
        //            Assert.IsTrue(htmlService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

        //            string ServerFilePath = tvFileService.GetServerFilePath(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //            string ServerFileName = csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.FileName;

        //            TVItemModel tvItemModelFile = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, ServerFileName, TVTypeEnum.File);
        //            Assert.AreEqual("", tvItemModelFile.Error);

        //            TVFileModel tvFileModel = tvFileService.GetTVFileModelWithTVFileTVItemIDDB(tvItemModelFile.TVItemID);
        //            Assert.AreEqual("", tvFileModel.Error);

        //            FileInfo fi = new FileInfo(ServerFilePath + ServerFileName);
        //            Assert.IsTrue(fi.Exists);

        //            if (fi.Exists)
        //            {
        //                fi.Delete();
        //            }
        //        }
        //    }
        //}
        //[TestMethod]
        //public void HtmlService_GenerateCountryHTML_Canada_TVItemID_5_Test()
        //{
        //    // Generated files will be located
        //    // \inetpub\wwwroot\csspwebtools\App_Data\5           
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

        //            TVItemModel tvItemModelCanada = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, "Canada", TVTypeEnum.Country);

        //            SetupBWObj(tvItemModelCanada, FileGeneratorEnum.Country, FileGeneratorTypeEnum.HTML, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            htmlService = new HtmlService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(htmlService);
        //            Assert.IsNotNull(htmlService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, htmlService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            htmlService.GenerateCountryHTML();
        //            Assert.IsTrue(htmlService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

        //            string ServerFilePath = tvFileService.GetServerFilePath(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //            string ServerFileName = csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.FileName;

        //            TVItemModel tvItemModelFile = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelCanada.TVItemID, ServerFileName, TVTypeEnum.File);
        //            Assert.AreEqual("", tvItemModelFile.Error);

        //            TVFileModel tvFileModel = tvFileService.GetTVFileModelWithTVFileTVItemIDDB(tvItemModelFile.TVItemID);
        //            Assert.AreEqual("", tvFileModel.Error);

        //            FileInfo fi = new FileInfo(ServerFilePath + ServerFileName);
        //            Assert.IsTrue(fi.Exists);

        //            if (fi.Exists)
        //            {
        //                fi.Delete();
        //            }
        //        }
        //    }
        //}
        //[TestMethod]
        //public void HtmlService_GenerateProvinceHTML_New_Brunswick_TVItemID_7_Test()
        //{
        //    // Generated files will be located
        //    // \inetpub\wwwroot\csspwebtools\App_Data\7            
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

        //            TVItemModel tvItemModelNB = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, (LanguageRequest == "fr" ? "Nouveau-Brunswick" : "New Brunswick"), TVTypeEnum.Province);

        //            SetupBWObj(tvItemModelNB, FileGeneratorEnum.Province, FileGeneratorTypeEnum.HTML, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            htmlService = new HtmlService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(htmlService);
        //            Assert.IsNotNull(htmlService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, htmlService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            htmlService.GenerateProvinceHTML();
        //            Assert.IsTrue(htmlService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

        //            string ServerFilePath = tvFileService.GetServerFilePath(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //            string ServerFileName = csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.FileName;

        //            TVItemModel tvItemModelFile = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelNB.TVItemID, ServerFileName, TVTypeEnum.File);
        //            Assert.AreEqual("", tvItemModelFile.Error);

        //            TVFileModel tvFileModel = tvFileService.GetTVFileModelWithTVFileTVItemIDDB(tvItemModelFile.TVItemID);
        //            Assert.AreEqual("", tvFileModel.Error);

        //            FileInfo fi = new FileInfo(ServerFilePath + ServerFileName);
        //            Assert.IsTrue(fi.Exists);

        //            if (fi.Exists)
        //            {
        //                fi.Delete();
        //            }
        //        }
        //    }
        //}
        //[TestMethod]
        //public void HtmlService_GenerateAreaHTML_NB_06_TVItemID_629_Test()
        //{
        //    // Generated files will be located
        //    // \inetpub\wwwroot\csspwebtools\App_Data\629         
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

        //            TVItemModel tvItemModelArea = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, ("NB-06"), TVTypeEnum.Area);

        //            SetupBWObj(tvItemModelArea, FileGeneratorEnum.Area, FileGeneratorTypeEnum.HTML, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            htmlService = new HtmlService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(htmlService);
        //            Assert.IsNotNull(htmlService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, htmlService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            htmlService.GenerateAreaHTML();
        //            Assert.IsTrue(htmlService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

        //            string ServerFilePath = tvFileService.GetServerFilePath(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //            string ServerFileName = csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.FileName;

        //            TVItemModel tvItemModelFile = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelArea.TVItemID, ServerFileName, TVTypeEnum.File);
        //            Assert.AreEqual("", tvItemModelFile.Error);

        //            TVFileModel tvFileModel = tvFileService.GetTVFileModelWithTVFileTVItemIDDB(tvItemModelFile.TVItemID);
        //            Assert.AreEqual("", tvFileModel.Error);

        //            FileInfo fi = new FileInfo(ServerFilePath + ServerFileName);
        //            Assert.IsTrue(fi.Exists);

        //            if (fi.Exists)
        //            {
        //                fi.Delete();
        //            }
        //        }
        //    }
        //}
        //[TestMethod]
        //public void HtmlService_GenerateSectorHTML_NB_06_020_TVItemID_633_Test()
        //{
        //    // Generated files will be located
        //    // \inetpub\wwwroot\csspwebtools\App_Data\633        
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

        //            TVItemModel tvItemModelNB = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, ("NB-06-020"), TVTypeEnum.Sector);

        //            SetupBWObj(tvItemModelNB, FileGeneratorEnum.Sector, FileGeneratorTypeEnum.HTML, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            htmlService = new HtmlService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(htmlService);
        //            Assert.IsNotNull(htmlService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, htmlService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            htmlService.GenerateSectorHTML();
        //            Assert.IsTrue(htmlService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

        //            string ServerFilePath = tvFileService.GetServerFilePath(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //            string ServerFileName = csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.FileName;

        //            TVItemModel tvItemModelFile = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelNB.TVItemID, ServerFileName, TVTypeEnum.File);
        //            Assert.AreEqual("", tvItemModelFile.Error);

        //            TVFileModel tvFileModel = tvFileService.GetTVFileModelWithTVFileTVItemIDDB(tvItemModelFile.TVItemID);
        //            Assert.AreEqual("", tvFileModel.Error);

        //            FileInfo fi = new FileInfo(ServerFilePath + ServerFileName);
        //            Assert.IsTrue(fi.Exists);

        //            if (fi.Exists)
        //            {
        //                fi.Delete();
        //            }
        //        }
        //    }
        //}
        //[TestMethod]
        //public void HtmlService_GenerateSubsectorHTML_NB_06_020_002_TVItemID_635_Test()
        //{
        //    // Generated files will be located
        //    // \inetpub\wwwroot\csspwebtools\App_Data\635        
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

        //            TVItemModel tvItemModelNB = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, ("NB-06-020-002"), TVTypeEnum.Subsector);

        //            SetupBWObj(tvItemModelNB, FileGeneratorEnum.Subsector, FileGeneratorTypeEnum.HTML, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            htmlService = new HtmlService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(htmlService);
        //            Assert.IsNotNull(htmlService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, htmlService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            htmlService.GenerateSubsectorHTML();
        //            Assert.IsTrue(htmlService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

        //            string ServerFilePath = tvFileService.GetServerFilePath(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //            string ServerFileName = csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.FileName;

        //            TVItemModel tvItemModelFile = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelNB.TVItemID, ServerFileName, TVTypeEnum.File);
        //            Assert.AreEqual("", tvItemModelFile.Error);

        //            TVFileModel tvFileModel = tvFileService.GetTVFileModelWithTVFileTVItemIDDB(tvItemModelFile.TVItemID);
        //            Assert.AreEqual("", tvFileModel.Error);

        //            FileInfo fi = new FileInfo(ServerFilePath + ServerFileName);
        //            Assert.IsTrue(fi.Exists);

        //            if (fi.Exists)
        //            {
        //                fi.Delete();
        //            }
        //        }
        //    }
        //}
        //[TestMethod]
        //public void HtmlService_GenerateMunicipalityHTML_Bouctouche_TVItemID_27764_Test()
        //{
        //    // Generated files will be located
        //    // \inetpub\wwwroot\csspwebtools\App_Data\27764
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

        //            TVItemModel tvItemModelNB = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, ("Bouctouche"), TVTypeEnum.Municipality);

        //            SetupBWObj(tvItemModelNB, FileGeneratorEnum.Municipality, FileGeneratorTypeEnum.HTML, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            htmlService = new HtmlService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(htmlService);
        //            Assert.IsNotNull(htmlService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, htmlService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            htmlService.GenerateMunicipalityHTML();
        //            Assert.IsTrue(htmlService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

        //            string ServerFilePath = tvFileService.GetServerFilePath(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //            string ServerFileName = csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.FileName;

        //            TVItemModel tvItemModelFile = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelNB.TVItemID, ServerFileName, TVTypeEnum.File);
        //            Assert.AreEqual("", tvItemModelFile.Error);

        //            TVFileModel tvFileModel = tvFileService.GetTVFileModelWithTVFileTVItemIDDB(tvItemModelFile.TVItemID);
        //            Assert.AreEqual("", tvFileModel.Error);

        //            FileInfo fi = new FileInfo(ServerFilePath + ServerFileName);
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
        //        AppTaskModel appTaskModelRet = appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);

        //        Assert.AreEqual("", appTaskModelRet.Error);
        //    }
        //}
        //private void SetupBWObj(TVItemModel tvItemModel, FileGeneratorEnum fileGenerator, FileGeneratorTypeEnum fileGeneratorType, string LanguageRequest)
        //{
        //    RemoveAllTask();

        //    string ext = tvFileService.GetFileExtension(fileGeneratorType);
        //    Assert.AreEqual(".html", ext);

        //    string ServerFilePath = tvFileService.GetServerFilePath(tvItemModel.TVItemID);
        //    Assert.IsTrue(!string.IsNullOrWhiteSpace(ServerFilePath));

        //    string FileName = tvFileService.GenerateFileName(tvItemModel.TVItemID, fileGenerator, fileGeneratorType, "", LanguageRequest);
        //    Assert.IsTrue(!string.IsNullOrWhiteSpace(FileName));

        //    FileInfo fi = new FileInfo(ServerFilePath + "\\" + FileName);
        //    if (tvFileService.GetFileExist(fi))
        //    {
        //        fi.Delete();
        //    }

        //    fi = new FileInfo(ServerFilePath + "\\" + FileName);
        //    Assert.IsFalse(tvFileService.GetFileExist(fi));

        //    tvFileService.LanguageRequest = LanguageRequest;
        //    FormCollection fc = new FormCollection();
        //    fc.Add("TVItemID", tvItemModel.TVItemID.ToString());
        //    fc.Add("FileGenerator", ((int)fileGenerator).ToString());
        //    fc.Add("FileGeneratorType", ((int)fileGeneratorType).ToString());
        //    fc.Add("ContourValues", "");
        //    fc.Add("SigmaLayerValues", "");
        //    fc.Add("ZLayerValues", "");
        //    fc.Add("DepthValues", "");
        //    fc.Add("VectorLengthStr", "0");
        //    fc.Add("GoogleEarthPath", "");

        //    string retStr = tvFileService.FileGenerateDB(fc);
        //    Assert.AreEqual("", retStr);

        //    using (ShimsContext.Create())
        //    {
        //        SetupShim();

        //        shimTaskRunnerBaseService.GenerateDoc = () =>
        //        {
        //            return; // just so it does not start to Generate the documents
        //        };
        //        csspWebToolsTaskRunner.GetNextTask();

        //        //
        //        Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel);
        //        Assert.AreEqual(tvItemModel.TVItemID, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //        Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.bw);
        //        Assert.IsFalse(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskCommand.ToString()));
        //        Assert.IsTrue(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.Index > 0);
        //        Assert.AreEqual(tvItemModel.TVItemID, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //        Assert.IsTrue(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.ErrorText));
        //        Assert.IsTrue(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.StatusText));
        //        string FirstPart = "|||TVItemID," + tvItemModel.TVItemID.ToString() +
        //            "|||FileGenerator," + ((int)fileGenerator).ToString() +
        //            "|||FileGeneratorType," + ((int)fileGeneratorType).ToString() +
        //            "|||Generate,1" +
        //            "|||Command,0" +
        //            "|||FileName," + FileName;
        //        Assert.AreEqual(FirstPart, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters.Substring(0, FirstPart.Length));
        //        string LastPart = "_" + csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Language + ext.ToLower() + "|||";
        //        int ParamLength = csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters.Length;
        //        Assert.AreEqual(LastPart, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters.Substring(ParamLength - LastPart.Length));
        //        Assert.AreEqual(fileGenerator, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.fileGenerator);
        //        Assert.AreEqual(fileGeneratorType, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.fileGeneratorType);
        //        Assert.AreEqual(true, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.Generate);
        //        Assert.AreEqual(false, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.Command);
        //        Assert.AreEqual(FileName, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.FileName);
        //        Assert.AreEqual(true, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.bw.IsBusy);
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
        //}
        //private void SetupShim()
        //{
        //    shimTaskRunnerBaseService = new ShimTaskRunnerBaseService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //}
        #endregion Functions private

    }
}

