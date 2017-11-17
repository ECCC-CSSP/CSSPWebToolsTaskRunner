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
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;
using CSSPWebToolsTaskRunner.Services.Resources;

namespace CSSPWebToolsTaskRunner.Test.Services
{
    /// <summary>
    /// Summary description for BaseMOdelServiceTest
    /// </summary>
    [TestClass]
    public class ParameterServiceHTMLSubsectorTest
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
        public ParameterServiceHTMLSubsectorTest()
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
        public void Excel_Image_Test()
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = xlApp.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.Sheets.Add();

            worksheet.Shapes.AddChart().Select();
            xlApp.ActiveChart.ApplyLayout(9, Microsoft.Office.Interop.Excel.XlChartType.xl3DColumn);
            xlApp.ActiveChart.ChartTitle.Select();
            xlApp.Selection.Delete();
            xlApp.ActiveChart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue).AxisTitle.Select();
            xlApp.Selection.Delete();
            xlApp.ActiveChart.Legend.Select();
            xlApp.Selection.Delete();

            Microsoft.Office.Interop.Excel.SeriesCollection seriesCollection = (Microsoft.Office.Interop.Excel.SeriesCollection)xlApp.ActiveChart.SeriesCollection();
            Microsoft.Office.Interop.Excel.Series series = seriesCollection.NewSeries();

            xlApp.ActiveChart.Parent.Width = 600;
            xlApp.ActiveChart.Parent.Height = 100;

            series.Values = new double[] { 1d, 3d, 2d, 5d };
            series.XValues = new string[] { "A", "B", "C", "D" };

            xlApp.ActiveChart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = TaskRunnerServiceRes.YearsWithSamplesUsed;

            xlApp.ActiveChart.Export(@"C:\Users\leblancc\Desktop\test.png", "PNG", false);

            if (workbook != null)
            {
                workbook.Close(false);
            }
            if (xlApp != null)
            {
                xlApp.Quit();
            }
        }
        [TestMethod]
        public void GenerateHTMLSubsector_With_UniqueCode_FCSummaryStatDocx_Test()
        {
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                SetupTest(LanguageRequest);

                int SubsectorTVItemID = 635;
                int ReportTypeID = 23;
                int Year = 2016;

                FileInfo fi = new FileInfo(@"C:\Users\leblancc\Desktop\TestHTML\TestGenerateHTMLSubsector_FCSummaryStatDocx_" + LanguageRequest.ToString() + ".html");
                StringBuilder sbHTML = new StringBuilder();
                string Parameters = $"|||TVItemID,{ SubsectorTVItemID }|||ReportTypeID,{ ReportTypeID }|||Year,{ Year }|||";
                ReportTypeModel reportTypeModel = _ReportTypeService.GetReportTypeModelWithReportTypeIDDB(23);
                AppTaskModel appTaskModel = new AppTaskModel()
                {
                    AppTaskID = 10000,
                    TVItemID = SubsectorTVItemID,
                    TVItemID2 = SubsectorTVItemID,
                    AppTaskCommand = AppTaskCommandEnum.CreateDocumentFromParameters,
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
                ParametersService parameterService = new ParametersService(taskRunnerBaseService);
                bool retBool = parameterService.PublicGenerateHTMLSubsectorFCSummaryStatDocx(fi, sbHTML, Parameters, reportTypeModel);
                Assert.AreEqual(true, retBool);

                StreamWriter sw = fi.CreateText();
                sw.Write(sbHTML.ToString());
                sw.Flush();
                sw.Close();
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
        }
        #endregion Functions private

    }
}

