﻿using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;
using CSSPWebToolsTaskRunner.Services;
using System.Linq;
using System.ComponentModel;
using System.Transactions;
using Microsoft.QualityTools.Testing.Fakes;
using System.IO;
using System.Web.Mvc;
using CSSPWebToolsDBDLL.Services;
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;
using CSSPWebToolsTaskRunner.Services.Resources;
using System.Drawing;

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
        public void DrawPolSourceSitesPoints_test()
        {
            csspWebToolsTaskRunner = new CSSPWebToolsTaskRunner();
            MapInfoService mapInfoService = new MapInfoService(LanguageEnum.en, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
            TVItemService tvItemService = new TVItemService(LanguageEnum.en, csspWebToolsTaskRunner._TaskRunnerBaseService._User);

            Bitmap bitmap = new Bitmap(1280, 1200);
            Graphics g = Graphics.FromImage(bitmap);
            int GraphicWidth = bitmap.Width;
            int GraphicHeight = bitmap.Height;

            CoordMap coordMap = new CoordMap()
            {
                NorthEast = new Coord() { Lat = 46.5364151f, Lng = -64.55215f, Ordinal = 0 },
                SouthWest = new Coord() { Lat = 46.23907f, Lng = -64.99161f, Ordinal = 0 },
            };

            int SubsectorTVItemID = 635;

            List<MapInfoPointModel> mapInfoPointModelPolSourceSiteList = new List<MapInfoPointModel>();
            List<TVItemModel> tvItemModelPolSourceSiteList = new List<TVItemModel>();

            mapInfoPointModelPolSourceSiteList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithParentIDAndTVTypeAndMapInfoDrawTypeDB(SubsectorTVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);
            tvItemModelPolSourceSiteList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(SubsectorTVItemID, TVTypeEnum.PolSourceSite).Where(c => c.IsActive == true).ToList();

            GoogleMapToPNG googleMapToPNG = new GoogleMapToPNG(); // (new TaskRunnerBaseService(new List<BWObj>()), "", "", "", "");
            googleMapToPNG.DrawPolSourceSitesPoints(g, GraphicWidth, GraphicHeight, coordMap, mapInfoPointModelPolSourceSiteList, tvItemModelPolSourceSiteList);


        }
        [TestMethod]
        public void Excel_Image_Test()
        {
            Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = xlApp.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.Worksheets.get_Item(1);

            Microsoft.Office.Interop.Excel.ChartObjects xlCharts = (Microsoft.Office.Interop.Excel.ChartObjects)worksheet.ChartObjects();
            Microsoft.Office.Interop.Excel.ChartObject chart = xlCharts.Add(100, 100, 600, 100);
            Microsoft.Office.Interop.Excel.Chart chartPage = chart.Chart;

            chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;

            Microsoft.Office.Interop.Excel.SeriesCollection seriesCollection = chartPage.SeriesCollection();
            Microsoft.Office.Interop.Excel.Series series = seriesCollection.NewSeries();

            series.Values = new double[] { 1d, 3d, 2d, 5d, 6d, 7d };
            series.XValues = new double[] { 1d, 2d, 3d, 4d, 7d, 19d };

            chartPage.ApplyLayout(9, Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered);
            chartPage.ChartTitle.Select();
            xlApp.Selection.Delete();
            chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue).AxisTitle.Select();
            xlApp.Selection.Delete();
            chartPage.Legend.Select();
            xlApp.Selection.Delete();

            chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = TaskRunnerServiceRes.NumberOfRunsUsedByYear;

            chartPage.Export(@"C:\Users\leblancc\Desktop\TestHTML\test.png", "PNG", false);

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
        public void PublicGenerateHTMLSUBSECTOR_FULL_REPORT_COVER_PAGE_Test()
        {
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                SetupTest(LanguageRequest);

                int SubsectorTVItemID = 635;
                int ReportTypeID = 32;
                int Year = 2017;

                FileInfo fi = new FileInfo(@"C:\Users\leblancc\Desktop\TestHTML\GenerateHTMLSUBSECTOR_FULL_REPORT_COVER_PAGE_" + LanguageRequest.ToString() + ".html");
                StringBuilder sbHTML = new StringBuilder();
                string Parameters = $"|||TVItemID,{ SubsectorTVItemID }|||ReportTypeID,{ ReportTypeID }|||Year,{ Year }|||";
                ReportTypeModel reportTypeModel = _ReportTypeService.GetReportTypeModelWithReportTypeIDDB(ReportTypeID);
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
                parameterService.fi = fi;
                parameterService.sb = sbHTML;
                parameterService.Parameters = Parameters;
                parameterService.reportTypeModel = reportTypeModel;
                parameterService.TVItemID = SubsectorTVItemID;
                parameterService.Year = Year;
                StringBuilder sbTemp = new StringBuilder();
                bool retBool = parameterService.PublicGenerateHTMLSUBSECTOR_RE_EVALUATION_COVER_PAGE(sbTemp);
                Assert.AreEqual(true, retBool);

                StreamWriter sw = fi.CreateText();
                sw.Write(sbTemp.ToString());
                sw.Flush();
                sw.Close();
            }
        }
        [TestMethod]
        public void PublicGenerateHTMLSUBSECTOR_FC_SUMMARY_STAT_ALL_Test()
        {
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                SetupTest(LanguageRequest);

                int SubsectorTVItemID = 635;
                int ReportTypeID = 23;
                int Year = 2017;

                FileInfo fi = new FileInfo(@"C:\Users\leblancc\Desktop\TestHTML\PublicGenerateHTMLSUBSECTOR_FC_SUMMARY_STAT_ALL_" + LanguageRequest.ToString() + ".html");
                StringBuilder sbHTML = new StringBuilder();
                string Parameters = $"|||TVItemID,{ SubsectorTVItemID }|||ReportTypeID,{ ReportTypeID }|||Year,{ Year }|||";
                ReportTypeModel reportTypeModel = _ReportTypeService.GetReportTypeModelWithReportTypeIDDB(ReportTypeID);
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
                parameterService.fi = fi;
                parameterService.sb = sbHTML;
                parameterService.Parameters = Parameters;
                parameterService.reportTypeModel = reportTypeModel;
                parameterService.TVItemID = SubsectorTVItemID;
                parameterService.Year = Year;
                StringBuilder sbTemp = new StringBuilder();
                bool retBool = parameterService.PublicGenerateHTMLSUBSECTOR_FC_SUMMARY_STAT_ALL(sbTemp);
                Assert.AreEqual(true, retBool);

                StreamWriter sw = fi.CreateText();
                sw.Write(sbTemp.ToString());
                sw.Flush();
                sw.Close();
            }
        }
        [TestMethod]
        public void PublicGenerateHTMLSUBSECTOR_FC_SUMMARY_STAT_DRY_Test()
        {
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                SetupTest(LanguageRequest);

                int SubsectorTVItemID = 635;
                int ReportTypeID = 23;
                int Year = 2017;

                FileInfo fi = new FileInfo(@"C:\Users\leblancc\Desktop\TestHTML\PublicGenerateHTMLSUBSECTOR_FC_SUMMARY_STAT_DRY_" + LanguageRequest.ToString() + ".html");
                StringBuilder sbHTML = new StringBuilder();
                string Parameters = $"|||TVItemID,{ SubsectorTVItemID }|||ReportTypeID,{ ReportTypeID }|||Year,{ Year }|||";
                ReportTypeModel reportTypeModel = _ReportTypeService.GetReportTypeModelWithReportTypeIDDB(ReportTypeID);
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
                parameterService.fi = fi;
                parameterService.sb = sbHTML;
                parameterService.Parameters = Parameters;
                parameterService.reportTypeModel = reportTypeModel;
                parameterService.TVItemID = SubsectorTVItemID;
                parameterService.Year = Year;
                StringBuilder sbTemp = new StringBuilder();
                bool retBool = parameterService.PublicGenerateHTMLSUBSECTOR_FC_SUMMARY_STAT_DRY(sbTemp);
                Assert.AreEqual(true, retBool);

                StreamWriter sw = fi.CreateText();
                sw.Write(sbTemp.ToString());
                sw.Flush();
                sw.Close();
            }
        }
        [TestMethod]
        public void PublicGenerateHTMLSUBSECTOR_FC_SUMMARY_STAT_WET_Test()
        {
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                SetupTest(LanguageRequest);

                int SubsectorTVItemID = 635;
                int ReportTypeID = 23;
                int Year = 2017;

                FileInfo fi = new FileInfo(@"C:\Users\leblancc\Desktop\TestHTML\PublicGenerateHTMLSUBSECTOR_FC_SUMMARY_STAT_WET_" + LanguageRequest.ToString() + ".html");
                StringBuilder sbHTML = new StringBuilder();
                string Parameters = $"|||TVItemID,{ SubsectorTVItemID }|||ReportTypeID,{ ReportTypeID }|||Year,{ Year }|||";
                ReportTypeModel reportTypeModel = _ReportTypeService.GetReportTypeModelWithReportTypeIDDB(ReportTypeID);
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
                parameterService.fi = fi;
                parameterService.sb = sbHTML;
                parameterService.Parameters = Parameters;
                parameterService.reportTypeModel = reportTypeModel;
                parameterService.TVItemID = SubsectorTVItemID;
                parameterService.Year = Year;
                StringBuilder sbTemp = new StringBuilder();
                bool retBool = parameterService.PublicGenerateHTMLSUBSECTOR_FC_SUMMARY_STAT_WET(sbTemp);
                Assert.AreEqual(true, retBool);

                StreamWriter sw = fi.CreateText();
                sw.Write(sbTemp.ToString());
                sw.Flush();
                sw.Close();
            }
        }
        [TestMethod]
        public void PublicGenerateHTMLSubsectorTestObjectsOfFullReport_Test()
        {
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                SetupTest(LanguageRequest);

                int SubsectorTVItemID = 635;
                int ReportTypeID = 36;
                int Year = 1999;

                FileInfo fi = new FileInfo(@"C:\Users\leblancc\Desktop\TestHTML\PublicGenerateHTMLSubsectorTestObjectsOfFullReport_" + LanguageRequest.ToString() + ".html");
                StringBuilder sbHTML = new StringBuilder();
                string Parameters = $"|||TVItemID,{ SubsectorTVItemID }|||ReportTypeID,{ ReportTypeID }|||Year,{ Year }|||";
                ReportTypeModel reportTypeModel = _ReportTypeService.GetReportTypeModelWithReportTypeIDDB(ReportTypeID);
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
                parameterService.fi = fi;
                parameterService.sb = sbHTML;
                parameterService.Parameters = Parameters;
                parameterService.reportTypeModel = reportTypeModel;
                parameterService.TVItemID = SubsectorTVItemID;
                parameterService.Year = Year;
                bool retBool = parameterService.PublicGenerateHTMLDocx();
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

