using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;
using CSSPWebToolsDBDLL.Models;
using CSSPWebToolsDBDLL.Services;
using CSSPWebToolsDBDLL.Services.Resources;
using CSSPWebToolsTaskRunner.Services;
using CSSPWebToolsTaskRunner.Services.Fakes;
using Microsoft.QualityTools.Testing.Fakes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Transactions;
using System.Web;
using System.Web.Mvc;

namespace CSSPWebToolsTaskRunner.Test.Services
{
    /// <summary>
    /// Summary description for BaseMOdelServiceTest
    /// </summary>
    [TestClass]
    public class XlsxServiceTest
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
        private XlsxService xlsxService { get; set; }
        private RandomService randomService { get; set; }
        private MWQMPlanService mwqmPlanService { get; set; }
        private MWQMPlanSubsectorService mwqmPlanSubsectorService { get; set; }
        private MWQMPlanSubsectorSiteService mwqmPlanSubsectorSiteService { get; set; }
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
        public XlsxServiceTest()
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
        //public void XlsxService_Constructor_Test()
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

        //            SetupBWObj(tvItemModelRoot, FileGeneratorEnum.Root, FileGeneratorTypeEnum.Excel, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            xlsxService = new XlsxService(csspWebToolsTaskRunner._TaskRunnerBaseService);

        //            Assert.IsNotNull(xlsxService);
        //            Assert.IsNotNull(xlsxService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, xlsxService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
        //        }
        //    }
        //}
        //[TestMethod]
        //public void XlsxService_GenerateRootXLSX_Root_TVItemID_1_Test()
        //{
        //    // Generated files will be located
        //    // \inetpub\wwwroot\csspwebtools\App_Data\1
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

        //            SetupBWObj(tvItemModelRoot, FileGeneratorEnum.Root, FileGeneratorTypeEnum.Excel, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            xlsxService = new XlsxService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(xlsxService);
        //            Assert.IsNotNull(xlsxService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, xlsxService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            xlsxService.GenerateRootXLSX();
        //            Assert.IsTrue(xlsxService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

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
        //public void XlsxService_GenerateRootXLSXAndPDF_Root_TVItemID_1_Test()
        //{
        //    // Generated files will be located
        //    // \inetpub\wwwroot\csspwebtools\App_Data\1
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

        //            SetupBWObj(tvItemModelRoot, FileGeneratorEnum.Root, FileGeneratorTypeEnum.ExcelAndPDF, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            xlsxService = new XlsxService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(xlsxService);
        //            Assert.IsNotNull(xlsxService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, xlsxService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            xlsxService.GenerateRootXLSXAndPDF();
        //            Assert.IsTrue(xlsxService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

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
        //public void XlsxService_GenerateCountryXLSX_Canada_TVItemID_5_Test()
        //{
        //    // Generated files will be located
        //    // \inetpub\wwwroot\csspwebtools\App_Data\5
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

        //            TVItemModel tvItemModelCountry = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, "Canada", TVTypeEnum.Country);
        //            Assert.AreEqual("", tvItemModelCountry.Error);

        //            SetupBWObj(tvItemModelCountry, FileGeneratorEnum.Country, FileGeneratorTypeEnum.Excel, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            xlsxService = new XlsxService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(xlsxService);
        //            Assert.IsNotNull(xlsxService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, xlsxService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            xlsxService.GenerateCountryXLSX();
        //            Assert.IsTrue(xlsxService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

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
        //public void XlsxService_GenerateCountryXLSXAndPDF_Canada_TVItemID_5_Test()
        //{
        //    // Generated files will be located
        //    // \inetpub\wwwroot\csspwebtools\App_Data\5
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

        //            TVItemModel tvItemModelCountry = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, "Canada", TVTypeEnum.Country);
        //            Assert.AreEqual("", tvItemModelCountry.Error);

        //            SetupBWObj(tvItemModelCountry, FileGeneratorEnum.Country, FileGeneratorTypeEnum.ExcelAndPDF, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            xlsxService = new XlsxService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(xlsxService);
        //            Assert.IsNotNull(xlsxService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, xlsxService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            xlsxService.GenerateCountryXLSXAndPDF();
        //            Assert.IsTrue(xlsxService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

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
        //public void XlsxService_GenerateProvinceXLSX_New_Brunswick_TVItemID_7_Test()
        //{
        //    // Generated files will be located
        //    // \inetpub\wwwroot\csspwebtools\App_Data\7
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

        //            TVItemModel tvItemModelProvince = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, (LanguageRequest == "fr" ? "Nouveau-Brunswick" : "New Brunswick"), TVTypeEnum.Province);
        //            Assert.AreEqual("", tvItemModelProvince.Error);

        //            SetupBWObj(tvItemModelProvince, FileGeneratorEnum.Province, FileGeneratorTypeEnum.Excel, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            xlsxService = new XlsxService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(xlsxService);
        //            Assert.IsNotNull(xlsxService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, xlsxService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            xlsxService.GenerateProvinceXLSX();
        //            Assert.IsTrue(xlsxService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

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
        //public void XlsxService_GenerateProvinceXLSXAndPDF_New_Brunswick_TVItemID_7_Test()
        //{
        //    // Generated files will be located
        //    // \inetpub\wwwroot\csspwebtools\App_Data\7
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

        //            TVItemModel tvItemModelProvince = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, (LanguageRequest == "fr" ? "Nouveau-Brunswick" : "New Brunswick"), TVTypeEnum.Province);
        //            Assert.AreEqual("", tvItemModelProvince.Error);

        //            SetupBWObj(tvItemModelProvince, FileGeneratorEnum.Province, FileGeneratorTypeEnum.ExcelAndPDF, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            xlsxService = new XlsxService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(xlsxService);
        //            Assert.IsNotNull(xlsxService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, xlsxService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            xlsxService.GenerateProvinceXLSXAndPDF();
        //            Assert.IsTrue(xlsxService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

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
        //public void XlsxService_GenerateAreaXLSX_NB_06_TVItemID_629_Test()
        //{
        //    // Generated files will be located
        //    // \inetpub\wwwroot\csspwebtools\App_Data\629
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

        //            TVItemModel tvItemModelArea = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, "NB-06", TVTypeEnum.Area);
        //            Assert.AreEqual("", tvItemModelArea.Error);

        //            SetupBWObj(tvItemModelArea, FileGeneratorEnum.Area, FileGeneratorTypeEnum.Excel, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            xlsxService = new XlsxService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(xlsxService);
        //            Assert.IsNotNull(xlsxService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, xlsxService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            xlsxService.GenerateAreaXLSX();
        //            Assert.IsTrue(xlsxService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

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
        //public void XlsxService_GenerateAreaXLSXAndPDF_NB_06_TVItemID_629_Test()
        //{
        //    // Generated files will be located
        //    // \inetpub\wwwroot\csspwebtools\App_Data\629
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

        //            TVItemModel tvItemModelArea = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, "NB-06", TVTypeEnum.Area);
        //            Assert.AreEqual("", tvItemModelArea.Error);

        //            SetupBWObj(tvItemModelArea, FileGeneratorEnum.Area, FileGeneratorTypeEnum.ExcelAndPDF, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            xlsxService = new XlsxService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(xlsxService);
        //            Assert.IsNotNull(xlsxService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, xlsxService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            xlsxService.GenerateAreaXLSXAndPDF();
        //            Assert.IsTrue(xlsxService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

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
        //public void XlsxService_GenerateSectorXLSX_NB_06_020_TVItemID_633_Test()
        //{
        //    // Generated files will be located
        //    // \inetpub\wwwroot\csspwebtools\App_Data\633
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

        //            TVItemModel tvItemModelSector = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, "NB-06-020", TVTypeEnum.Sector);
        //            Assert.AreEqual("", tvItemModelSector.Error);

        //            SetupBWObj(tvItemModelSector, FileGeneratorEnum.Sector, FileGeneratorTypeEnum.Excel, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            xlsxService = new XlsxService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(xlsxService);
        //            Assert.IsNotNull(xlsxService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, xlsxService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            xlsxService.GenerateSectorXLSX();
        //            Assert.IsTrue(xlsxService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

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
        //public void XlsxService_GenerateSectorXLSXAndPDF_NB_06_020_TVItemID_633_Test()
        //{
        //    // Generated files will be located
        //    // \inetpub\wwwroot\csspwebtools\App_Data\633
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

        //            TVItemModel tvItemModelSector = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, "NB-06-020", TVTypeEnum.Sector);
        //            Assert.AreEqual("", tvItemModelSector.Error);

        //            SetupBWObj(tvItemModelSector, FileGeneratorEnum.Sector, FileGeneratorTypeEnum.ExcelAndPDF, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            xlsxService = new XlsxService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(xlsxService);
        //            Assert.IsNotNull(xlsxService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, xlsxService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            xlsxService.GenerateSectorXLSXAndPDF();
        //            Assert.IsTrue(xlsxService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

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
        //public void XlsxService_GenerateSubsectorXLSX_NB_06_020_002_TVItemID_635_Test()
        //{
        //    // Generated files will be located
        //    // \inetpub\wwwroot\csspwebtools\App_Data\635
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

        //            TVItemModel tvItemModelSubsector = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, "NB-06-020-002", TVTypeEnum.Subsector);
        //            Assert.AreEqual("", tvItemModelSubsector.Error);

        //            SetupBWObj(tvItemModelSubsector, FileGeneratorEnum.Subsector, FileGeneratorTypeEnum.Excel, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            xlsxService = new XlsxService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(xlsxService);
        //            Assert.IsNotNull(xlsxService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, xlsxService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            xlsxService.GenerateSubsectorXLSX();
        //            Assert.IsTrue(xlsxService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

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
        //public void XlsxService_GenerateSubsectorXLSXAndPDF_NB_06_020_002_TVItemID_635_Test()
        //{
        //    // Generated files will be located
        //    // \inetpub\wwwroot\csspwebtools\App_Data\635
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

        //            TVItemModel tvItemModelSubsector = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, "NB-06-020-002", TVTypeEnum.Subsector);
        //            Assert.AreEqual("", tvItemModelSubsector.Error);

        //            SetupBWObj(tvItemModelSubsector, FileGeneratorEnum.Subsector, FileGeneratorTypeEnum.ExcelAndPDF, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            xlsxService = new XlsxService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(xlsxService);
        //            Assert.IsNotNull(xlsxService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, xlsxService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            xlsxService.GenerateSubsectorXLSXAndPDF();
        //            Assert.IsTrue(xlsxService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

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
        //public void XlsxService_GenerateMunicipalityXLSX_Bouctouche_TVItemID_27764_Test()
        //{
        //    // Generated files will be located
        //    // \inetpub\wwwroot\csspwebtools\App_Data\27764
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

        //            TVItemModel tvItemModelMunicipality = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, "Bouctouche", TVTypeEnum.Municipality);
        //            Assert.AreEqual("", tvItemModelMunicipality.Error);

        //            SetupBWObj(tvItemModelMunicipality, FileGeneratorEnum.Municipality, FileGeneratorTypeEnum.Excel, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            xlsxService = new XlsxService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(xlsxService);
        //            Assert.IsNotNull(xlsxService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, xlsxService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            xlsxService.GenerateMunicipalityXLSX();
        //            Assert.IsTrue(xlsxService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

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
        //public void XlsxService_GenerateMunicipalityXLSXAndPDF_Bouctouche_TVItemID_27764_Test()
        //{
        //    // Generated files will be located
        //    // \inetpub\wwwroot\csspwebtools\App_Data\27764
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

        //            TVItemModel tvItemModelMunicipality = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, "Bouctouche", TVTypeEnum.Municipality);
        //            Assert.AreEqual("", tvItemModelMunicipality.Error);

        //            SetupBWObj(tvItemModelMunicipality, FileGeneratorEnum.Municipality, FileGeneratorTypeEnum.ExcelAndPDF, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            xlsxService = new XlsxService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(xlsxService);
        //            Assert.IsNotNull(xlsxService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, xlsxService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            xlsxService.GenerateMunicipalityXLSXAndPDF();
        //            Assert.IsTrue(xlsxService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

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
        //public void XlsxService_GenerateMWQMPlanXLSX_Test()
        //{
        //    // Generated files will be located
        //    // \inetpub\wwwroot\csspwebtools\App_Data\7  New Brunswick
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

        //            TVItemModel tvItemModelProvince = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, (LanguageRequest == "fr" ? "Nouveau-Brunswick" : "New Brunswick"), TVTypeEnum.Province);
        //            Assert.AreEqual("", tvItemModelProvince.Error);

        //            MWQMPlanModel mwqmPlanModel = AddMWQMPlanModel(tvItemModelProvince);
        //            Assert.AreEqual("", mwqmPlanModel.Error);

        //            SetupBWObjForMWQMPlan(mwqmPlanModel, FileGeneratorEnum.MWQMPlan, FileGeneratorTypeEnum.Excel, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            xlsxService = new XlsxService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(xlsxService);
        //            Assert.IsNotNull(xlsxService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, xlsxService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            xlsxService.GenerateMWQMPlanXLSX();
        //            Assert.IsTrue(xlsxService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

        //            string ServerFilePath = tvFileService.GetServerFilePath(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //            string ServerFileName = csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.FileName;

        //            TVItemModel tvItemModelFile = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(mwqmPlanModel.ProvinceTVItemID, ServerFileName, TVTypeEnum.File);
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
        //public void XlsxService_GetCellColumnLetter_Test()
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

        //            SetupBWObj(tvItemModelRoot, FileGeneratorEnum.Root, FileGeneratorTypeEnum.Excel, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            XlsxServiceBase XlsxServiceBase = new XlsxServiceBase(csspWebToolsTaskRunner._TaskRunnerBaseService);

        //            for (UInt32 i = 1; i < 27; i++)
        //            {
        //                string retStr = XlsxServiceBase.GetCellReference(i, 1);

        //                Assert.AreEqual(((char)(64 + i)) + "1", retStr);
        //            }

        //            for (UInt32 i = 27; i < 52; i++)
        //            {
        //                string retStr = XlsxServiceBase.GetCellReference(i, 1);

        //                Assert.AreEqual("A" + ((char)(64 + i)) + "1", retStr);
        //            }

        //            for (UInt32 i = 53; i < 78; i++)
        //            {
        //                string retStr = XlsxServiceBase.GetCellReference(i, 1);

        //                Assert.AreEqual("B" + ((char)(64 + i)) + "1", retStr);
        //            }

        //            for (UInt32 i = 1; i < 27; i++)
        //            {
        //                string retStr = XlsxServiceBase.GetCellReference(1, i);

        //                Assert.AreEqual("A" + i.ToString(), retStr);
        //            }
        //        }
        //    }
        //}
        //[TestMethod]
        //public void XlsxService_GenerateSubsectorPollutionSourceFieldSheetXLSX_NB_06_020_002_TVItemID_635_Test()
        //{
        //    // Generated files will be located
        //    // \inetpub\wwwroot\csspwebtools\App_Data\635
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

        //            TVItemModel tvItemModelSubsector = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, "NB-06-020-002", TVTypeEnum.Subsector);
        //            Assert.AreEqual("", tvItemModelSubsector.Error);

        //            SetupBWObj(tvItemModelSubsector, FileGeneratorEnum.SubsectorPollutionSourceFieldSheet, FileGeneratorTypeEnum.Excel, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            xlsxService = new XlsxService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(xlsxService);
        //            Assert.IsNotNull(xlsxService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, xlsxService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            xlsxService.GenerateSubsectorPollutionSourceFieldSheetXLSX();
        //            Assert.IsTrue(xlsxService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

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
        //public void XlsxService_GenerateSubsectorPollutionSourceFieldSheetXLSX_NB_06_020_002_TVItemID_626_Test()
        //{
        //    // Generated files will be located
        //    // \inetpub\wwwroot\csspwebtools\App_Data\626  Aldouane
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

        //            TVItemModel tvItemModelSubsector = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, "NB-05-030-002", TVTypeEnum.Subsector);
        //            Assert.AreEqual("", tvItemModelSubsector.Error);

        //            SetupBWObj(tvItemModelSubsector, FileGeneratorEnum.SubsectorPollutionSourceFieldSheet, FileGeneratorTypeEnum.Excel, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            xlsxService = new XlsxService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(xlsxService);
        //            Assert.IsNotNull(xlsxService._TaskRunnerBaseService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, xlsxService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            xlsxService.GenerateSubsectorPollutionSourceFieldSheetXLSX();
        //            Assert.IsTrue(xlsxService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

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
        //#endregion Functions public

        #region Functions private
        //private MWQMPlanModel AddMWQMPlanModel(TVItemModel tvItemModelProvince)
        //{
        //    MWQMPlanModel mwqmPlanModelNew = new MWQMPlanModel()
        //    {
        //        ProvinceTVItemID = tvItemModelProvince.TVItemID,
        //        ConfigFileName = randomService.RandomString("", 20),
        //        ForGroupName = randomService.RandomString("", 20),
        //        Year = randomService.RandomInt(2015, 2020),
        //        SecretCode = randomService.RandomString("", 8),
        //        CreatorTVItemID = 2,
        //    };

        //    MWQMPlanModel mwqmPlanModelRet = mwqmPlanService.PostAddMWQMPlanDB(mwqmPlanModelNew);
        //    Assert.AreEqual("", mwqmPlanModelRet.Error);

        //    List<TVItemModel> tvItemModelSubsectorList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelProvince.TVItemID, TVTypeEnum.Subsector);

        //    int countSubsector = 0;
        //    foreach (TVItemModel tvItemModelSubsector in tvItemModelSubsectorList)
        //    {
        //        List<TVItemModel> tvItemModelSubsectorSiteList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.MWQMSite);
        //        if (tvItemModelSubsectorSiteList.Count > 0)
        //        {
        //            countSubsector += 1;
        //            if (countSubsector > 3)
        //            {
        //                break;
        //            }

        //            MWQMPlanSubsectorModel mwqmPlanSubsectorModelNew = new MWQMPlanSubsectorModel()
        //            {
        //                MWQMPlanID = mwqmPlanModelRet.MWQMPlanID,
        //                SubsectorTVItemID = tvItemModelSubsector.TVItemID,
        //            };

        //            MWQMPlanSubsectorModel mwqmPlanSubsectorModelRet = mwqmPlanSubsectorService.PostAddMWQMPlanSubsectorDB(mwqmPlanSubsectorModelNew);
        //            Assert.AreEqual("", mwqmPlanSubsectorModelRet.Error);

        //            int countSite = 0;
        //            foreach (TVItemModel tvItemModelSusectorSite in tvItemModelSubsectorSiteList)
        //            {
        //                countSite += 1;
        //                if (countSite > 3)
        //                {
        //                    break;
        //                }
        //                MWQMPlanSubsectorSiteModel mwqmPlanSubsectorSiteModelNew = new MWQMPlanSubsectorSiteModel()
        //                {
        //                    IsDuplicate = (countSite == 1 ? true : false),
        //                    MWQMPlanSubsectorID = mwqmPlanSubsectorModelRet.MWQMPlanSubsectorID,
        //                    MWQMSiteTVItemID = tvItemModelSusectorSite.TVItemID,
        //                };

        //                MWQMPlanSubsectorSiteModel mwqmPlanSubsectorSiteModelRet = mwqmPlanSubsectorSiteService.PostAddMWQMPlanSubsectorSiteDB(mwqmPlanSubsectorSiteModelNew);

        //                Assert.AreEqual("", mwqmPlanSubsectorSiteModelRet.Error);

        //            }
        //        }
        //    }

        //    return mwqmPlanModelRet;
        //}
        //private void RemoveAllTask()
        //{
        //    List<AppTaskModel> appTaskModelList = appTaskService.GetAppTaskModelListDB();
        //    foreach (AppTaskModel appTaskModel in appTaskModelList)
        //    {
        //        AppTaskModel appTaskModelRet = appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
        //        Assert.AreEqual("", appTaskModelRet.Error);
        //    }
        //}
        //private void SetupBWObj(TVItemModel tvItemModel, LanguageEnum LanguageRequest)
        //{
        //    RemoveAllTask();

        //    string ext = tvFileService.GetFileExtension(fileGeneratorType);
        //    Assert.AreEqual(".xlsx", ext);

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

        //    string retStr =  tvFileService.FileGenerateDB(fc);
        //    Assert.AreEqual("", retStr);

        //    using (ShimsContext.Create())
        //    {
        //        SetupShim();

        //        shimTaskRunnerBaseService.GenerateDoc = () =>
        //        {
        //            return; // just so it does not start to Generate the documents
        //        };

        //        csspWebToolsTaskRunner.GetNextTask();

        //        Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel);
        //        Assert.AreEqual(tvItemModel.TVItemID, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //        Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.bw);
        //        Assert.IsFalse(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskCommand.ToString()));
        //        Assert.IsTrue(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.Index > 0);
        //        Assert.IsTrue(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.ErrorText));
        //        Assert.IsTrue(!string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.StatusText));
        //        //string Parameters = "|||TVItemID," + tvItemModel.TVItemID.ToString() +
        //        //    "|||FileGenerator," + ((int)fileGenerator).ToString() +
        //        //    "|||FileGeneratorType," + ((int)fileGeneratorType).ToString() +
        //        //    "|||Generate,1" +
        //        //    "|||Command,0" +
        //        //    "|||ContourValues," + 
        //        //    "|||SigmaLayerValues," + 
        //        //    "|||ZLayerValues," +
        //        //    "|||DepthValues," +
        //        //    "|||VectorLength,0" +
        //        //    "|||GoogleEarthPath," +
        //        //    "|||FileName," + FileName + "|||";
        //        Assert.AreEqual("",
        //            appTaskService.GetAppTaskParamStr(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "ContourValues"));
        //        Assert.AreEqual("",
        //            appTaskService.GetAppTaskParamStr(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "SigmaLayerValues"));
        //        Assert.AreEqual("",
        //            appTaskService.GetAppTaskParamStr(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "ZLayerValues"));
        //        Assert.AreEqual("",
        //            appTaskService.GetAppTaskParamStr(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "DepthValues"));
        //        Assert.AreEqual("0",
        //            appTaskService.GetAppTaskParamStr(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "VectorLength"));
        //        Assert.AreEqual("",
        //            appTaskService.GetAppTaskParamStr(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "GoogleEarthPath"));
        //        //Assert.AreEqual(Parameters, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters);
        //        Assert.AreEqual(true, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.bw.IsBusy);
        //        Assert.AreEqual(false, csspWebToolsTaskRunner._BWList[0 + 1].bw.IsBusy);
        //    }

        //}
        //private void SetupBWObjForMWQMPlan(MWQMPlanModel mwqmPlanModel, LanguageEnum LanguageRequest)
        //{
        //    RemoveAllTask();

        //    string ServerFilePath = tvFileService.GetServerFilePath(mwqmPlanModel.ProvinceTVItemID);
        //    Assert.IsTrue(!string.IsNullOrWhiteSpace(ServerFilePath));

        //    TVItemModel tvItemModel = tvItemService.GetTVItemModelWithTVItemIDDB(mwqmPlanModel.ProvinceTVItemID);
        //    Assert.AreEqual("", tvItemModel.Error);

        //    string FileName = "config_" + mwqmPlanModel.ConfigFileName + tvFileService.GetFileExtension(fileGeneratorType);
        //    Assert.IsTrue(!string.IsNullOrWhiteSpace(FileName));

        //    FileInfo fi = new FileInfo(ServerFilePath + FileName);
        //    if (tvFileService.GetFileExist(fi))
        //    {
        //        fi.Delete();
        //    }

        //    fi = new FileInfo(ServerFilePath + FileName);
        //    Assert.IsFalse(tvFileService.GetFileExist(fi));

        //    mwqmPlanService.LanguageRequest = LanguageRequest;
        //    string retStr = mwqmPlanService.MWQMPlanFileGenerateDB(mwqmPlanModel.MWQMPlanID, fileGenerator, fileGeneratorType);
        //    Assert.AreEqual("", retStr);

        //    using (ShimsContext.Create())
        //    {
        //        SetupShim();

        //        shimTaskRunnerBaseService.GenerateDoc = () =>
        //        {
        //            return; // just so it does not start to Generate the documents
        //        };

        //        csspWebToolsTaskRunner.GetNextTask();

        //        Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel);
        //        Assert.AreEqual(mwqmPlanModel.ProvinceTVItemID, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //        Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.bw);
        //        Assert.IsFalse(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskCommand.ToString()));
        //        Assert.IsTrue(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.Index > 0);
        //        Assert.IsTrue(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.ErrorText));
        //        Assert.IsTrue(!string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.StatusText));
        //        Assert.AreEqual(mwqmPlanModel.ProvinceTVItemID, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
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
        //    randomService = new RandomService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        //    mwqmPlanService = new MWQMPlanService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        //    mwqmPlanSubsectorService = new MWQMPlanSubsectorService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        //    mwqmPlanSubsectorSiteService = new MWQMPlanSubsectorSiteService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        //}
        //private void SetupShim()
        //{
        //    shimTaskRunnerBaseService = new ShimTaskRunnerBaseService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //}
        #endregion Functions private

    }
}

