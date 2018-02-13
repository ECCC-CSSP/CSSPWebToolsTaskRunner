using System;
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
using DHI.Generic.MikeZero.DFS.dfsu;

namespace CSSPWebToolsTaskRunner.Test.Services
{
    /// <summary>
    /// Summary description for BaseMOdelServiceTest
    /// </summary>
    [TestClass]
    public class ParameterServiceKMLMikeScenarioTest
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
        public ParameterServiceKMLMikeScenarioTest()
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
        public void FillElementLayerList_Test()
        {
            //AppTaskID TVItemID    TVItemID2 AppTaskCommand  AppTaskStatus PercentCompleted    Parameters Language    StartDateTime_UTC EndDateTime_UTC EstimatedLength_second RemainingTime_second    LastUpdateDate_UTC LastUpdateContactTVItemID
            //10371	336990	336990	19	2	1	|||TVItemID,336925|||ReportTypeID,30|||ContourValues,14 88|||	1	2018-02-12 14:25:06.863	NULL NULL    NULL	2018-02-12 14:25:07.663	2
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                SetupTest(LanguageRequest);

                int MikeScenarioTVItemID = 336990;
                int ReportTypeID = 30;
                string ContourValues = "14 88";
                int Year = 2017;

                FileInfo fi = new FileInfo(@"C:\Users\leblancc\Desktop\TestHTML\FillElementLayerList_" + LanguageRequest.ToString() + ".KML");
                StringBuilder sbKML = new StringBuilder();
                string Parameters = $"|||TVItemID,{ MikeScenarioTVItemID }|||ReportTypeID,{ ReportTypeID }|||CoutourValues,{ ContourValues }|||";
                ReportTypeModel reportTypeModel = _ReportTypeService.GetReportTypeModelWithReportTypeIDDB(ReportTypeID);

                AppTaskModel appTaskModel = new AppTaskModel()
                {
                    AppTaskID = 100000,
                    TVItemID = MikeScenarioTVItemID,
                    TVItemID2 = MikeScenarioTVItemID,
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
                parameterService.sb = sbKML;
                parameterService.Parameters = Parameters;
                parameterService.reportTypeModel = reportTypeModel;
                parameterService.TVItemID = MikeScenarioTVItemID;
                parameterService.Year = Year;

                DfsuFile dfsuFile = null;
                List<Element> elementList = new List<Element>();
                List<ElementLayer> elementLayerList = new List<ElementLayer>();
                List<NodeLayer> topNodeLayerList = new List<NodeLayer>();
                List<NodeLayer> bottomNodeLayerList = new List<NodeLayer>();
                List<Node> nodeList = new List<Node>();

                dfsuFile = parameterService.GetTransportDfsuFile();
                Assert.IsNotNull(dfsuFile);
                Assert.AreEqual(0, taskRunnerBaseService._BWObj.TextLanguageList.Count);

                StringBuilder sbTemp = new StringBuilder();
                bool retBool = parameterService.FillRequiredList(dfsuFile, elementList, elementLayerList, nodeList, topNodeLayerList, bottomNodeLayerList);
                Assert.AreEqual(true, retBool);
                Assert.AreEqual(15829, nodeList.Count);
                Assert.AreEqual(22340, elementLayerList.Count);
                Assert.AreEqual(22340, elementList.Count);

                try
                {
                    dfsuFile.Close();
                }
                catch (Exception)
                {
                    // nothing
                }

                break;
            }

        }
        [TestMethod]
        public void GenerateMikeScenarioPollutionLimitKMZ_Test()
        {
            //AppTaskID TVItemID    TVItemID2 AppTaskCommand  AppTaskStatus PercentCompleted    Parameters Language    StartDateTime_UTC EndDateTime_UTC EstimatedLength_second RemainingTime_second    LastUpdateDate_UTC LastUpdateContactTVItemID
            //10371	336990	336990	19	2	1	|||TVItemID,336925|||ReportTypeID,30|||ContourValues,14 88|||	1	2018-02-12 14:25:06.863	NULL NULL    NULL	2018-02-12 14:25:07.663	2
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                SetupTest(LanguageRequest);

                int MikeScenarioTVItemID = 336990;
                int ReportTypeID = 30;
                string ContourValues = "14 88";
                int Year = 2017;

                FileInfo fi = new FileInfo(@"C:\Users\leblancc\Desktop\TestHTML\TestPolSourceLimit" + LanguageRequest.ToString() + ".kml");
                StringBuilder sbHTML = new StringBuilder();
                string Parameters = $"|||TVItemID,{ MikeScenarioTVItemID }|||ReportTypeID,{ ReportTypeID }|||ContourValues,{ ContourValues }|||";
                ReportTypeModel reportTypeModel = _ReportTypeService.GetReportTypeModelWithReportTypeIDDB(ReportTypeID);

                AppTaskModel appTaskModel = new AppTaskModel()
                {
                    AppTaskID = 100000,
                    TVItemID = MikeScenarioTVItemID,
                    TVItemID2 = MikeScenarioTVItemID,
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
                parameterService.TVItemID = MikeScenarioTVItemID;
                parameterService.Year = Year;

                StringBuilder sbTemp = new StringBuilder();
                bool retBool = parameterService.GenerateMikeScenarioPollutionLimitKMZ();
                Assert.AreEqual(true, retBool);

                StreamWriter sw = fi.CreateText();
                sw.Write(parameterService.sb.ToString());
                sw.Close();

                break;
            }

        }
        [TestMethod]
        public void GenerateMikeScenarioPollutionAnimationKMZ_Test()
        {
            //AppTaskID TVItemID    TVItemID2 AppTaskCommand  AppTaskStatus PercentCompleted    Parameters Language    StartDateTime_UTC EndDateTime_UTC EstimatedLength_second RemainingTime_second    LastUpdateDate_UTC LastUpdateContactTVItemID
            //10371	336990	336990	19	2	1	|||TVItemID,336925|||ReportTypeID,31|||ContourValues,14 88|||	1	2018-02-12 14:25:06.863	NULL NULL    NULL	2018-02-12 14:25:07.663	2
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                SetupTest(LanguageRequest);

                int MikeScenarioTVItemID = 336990;
                int ReportTypeID = 31;
                string ContourValues = "14 88";
                int Year = 2017;

                FileInfo fi = new FileInfo(@"C:\Users\leblancc\Desktop\TestHTML\TestPolSourceAnimation" + LanguageRequest.ToString() + ".kml");
                StringBuilder sbHTML = new StringBuilder();
                string Parameters = $"|||TVItemID,{ MikeScenarioTVItemID }|||ReportTypeID,{ ReportTypeID }|||ContourValues,{ ContourValues }|||";
                ReportTypeModel reportTypeModel = _ReportTypeService.GetReportTypeModelWithReportTypeIDDB(ReportTypeID);

                AppTaskModel appTaskModel = new AppTaskModel()
                {
                    AppTaskID = 100000,
                    TVItemID = MikeScenarioTVItemID,
                    TVItemID2 = MikeScenarioTVItemID,
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
                parameterService.TVItemID = MikeScenarioTVItemID;
                parameterService.Year = Year;

                StringBuilder sbTemp = new StringBuilder();
                bool retBool = parameterService.GenerateMikeScenarioPollutionAnimationKMZ();
                Assert.AreEqual(true, retBool);

                StreamWriter sw = fi.CreateText();
                sw.Write(parameterService.sb.ToString());
                sw.Close();

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
        }
        #endregion Functions private

    }
}

