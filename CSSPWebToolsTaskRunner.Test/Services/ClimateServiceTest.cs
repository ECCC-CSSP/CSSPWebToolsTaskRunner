using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;
using CSSPWebToolsTaskRunner.Services;
using System.Linq;
using CSSPDBDLL.Models;
using System.ComponentModel;
using System.Transactions;
using Microsoft.QualityTools.Testing.Fakes;
using CSSPDBDLL.Services;
using System.IO;
using System.Windows.Forms;
using CSSPDBDLL;
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;

namespace CSSPWebToolsTaskRunner.Test.Services
{
    /// <summary>
    /// Summary description for BaseMOdelServiceTest
    /// </summary>
    [TestClass]
    public class ClimateServiceTest
    {
        #region Variables
        private TestContext testContextInstance { get; set; }
        #endregion Variables

        #region Properties
        private CSSPWebToolsTaskRunner csspWebToolsTaskRunner { get; set; }
        private BWObj bwObj { get; set; }
        private ClimateService climateService { get; set; }
        private AppTaskService appTaskService { get; set; }
        private ClimateSiteService climateSiteService { get; set; }
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
        public ClimateServiceTest()
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
        public void ClimateService_Constructor_Test()
        {
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                int CountryTVItemID = 5;

                AppTaskModel appTaskModel = new AppTaskModel()
                {
                    AppTaskID = 0,
                    TVItemID = CountryTVItemID,
                    TVItemID2 = CountryTVItemID,
                    AppTaskCommand = AppTaskCommandEnum.UpdateClimateSiteInformation,
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

                ClimateService _ClimateService = new ClimateService(taskRunnerBaseService);
                _ClimateService.UpdateClimateSitesInformationForCountryTVItemID();
                Assert.AreEqual(0, taskRunnerBaseService._BWObj.TextLanguageList.Count);

                break;
            }
        }
        [TestMethod]
        public void ClimateService_GetAllPrecipitationForYear_Test()
        {
            // AppTaskID	TVItemID	TVItemID2	AppTaskCommand	AppTaskStatus	PercentCompleted	Parameters	Language	StartDateTime_UTC	EndDateTime_UTC	EstimatedLength_second	RemainingTime_second	LastUpdateDate_UTC	LastUpdateContactTVItemID
            // 14922	7	2018	26	1	1	|||ProvinceTVItemID,7|||Year,2018|||	1	2018-10-23 12:13:31.377	NULL	NULL	NULL	2018-10-23 12:13:31.377	2            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                int ProvinceTVItemID = 7;
                int Year = 2018;

                AppTaskModel appTaskModel = new AppTaskModel()
                {
                    AppTaskID = 0,
                    TVItemID = ProvinceTVItemID,
                    TVItemID2 = ProvinceTVItemID,
                    AppTaskCommand = AppTaskCommandEnum.GetAllPrecipitationForYear,
                    AppTaskStatus = AppTaskStatusEnum.Created,
                    PercentCompleted = 1,
                    Parameters = $"|||ProvinceTVItemID,{ ProvinceTVItemID }|||Year,{ Year }|||",
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

                ClimateService _ClimateService = new ClimateService(taskRunnerBaseService);
                _ClimateService.GetAllPrecipitationForYear();
                Assert.AreEqual(0, taskRunnerBaseService._BWObj.TextLanguageList.Count);

                break;
            }
        }
        [TestMethod]
        public void ClimateService_GetClimateSitesDataForSubsectorRunsOfYear_Test()
        {
            // AppTaskID	TVItemID	TVItemID2	AppTaskCommand	AppTaskStatus	PercentCompleted	Parameters	Language	StartDateTime_UTC	EndDateTime_UTC	EstimatedLength_second	RemainingTime_second	LastUpdateDate_UTC	LastUpdateContactTVItemID
            // 17971	774	774	15	2	4	|||SubsectorTVItemID,774|||Year,2019|||	1	2019-10-02 14:07:54.543	NULL	NULL	NULL	2019-10-02 14:08:52.340	2
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                int SubsectorTVItemID = 774;
                int Year = 2019;

                AppTaskModel appTaskModel = new AppTaskModel()
                {
                    AppTaskID = 0,
                    TVItemID = SubsectorTVItemID,
                    TVItemID2 = SubsectorTVItemID,
                    AppTaskCommand = AppTaskCommandEnum.GetClimateSitesDataForRunsOfYear,
                    AppTaskStatus = AppTaskStatusEnum.Created,
                    PercentCompleted = 1,
                    Parameters = $"|||SubsectorTVItemID,{SubsectorTVItemID}|||Year,{Year}|||",
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

                ClimateService _ClimateService = new ClimateService(taskRunnerBaseService);
                _ClimateService.GetClimateSitesDataForSubsectorRunsOfYear(SubsectorTVItemID, Year);
                Assert.AreEqual(0, taskRunnerBaseService._BWObj.TextLanguageList.Count);

                break;
            }
        }
        [TestMethod]
        public void ClimateService_FillRunPrecipByClimateSitePriorityForYear_Test()
        {
            // AppTaskID	TVItemID	TVItemID2	AppTaskCommand	AppTaskStatus	PercentCompleted	Parameters	Language	StartDateTime_UTC	EndDateTime_UTC	EstimatedLength_second	RemainingTime_second	LastUpdateDate_UTC	LastUpdateContactTVItemID
            // 14922	7	2018	26	1	1	|||ProvinceTVItemID,7|||Year,2018|||	1	2018-10-23 12:13:31.377	NULL	NULL	NULL	2018-10-23 12:13:31.377	2            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                int ProvinceTVItemID = 7;
                int Year = 2019;

                AppTaskModel appTaskModel = new AppTaskModel()
                {
                    AppTaskID = 0,
                    TVItemID = ProvinceTVItemID,
                    TVItemID2 = ProvinceTVItemID,
                    AppTaskCommand = AppTaskCommandEnum.FillRunPrecipByClimateSitePriorityForYear,
                    AppTaskStatus = AppTaskStatusEnum.Created,
                    PercentCompleted = 1,
                    Parameters = $"|||ProvinceTVItemID,{ ProvinceTVItemID }|||Year,{ Year }|||",
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

                ClimateService _ClimateService = new ClimateService(taskRunnerBaseService);
                _ClimateService.FillRunPrecipByClimateSitePriorityForYear();
                Assert.AreEqual(0, taskRunnerBaseService._BWObj.TextLanguageList.Count);

                break;
            }
        }
        [TestMethod]
        public void ClimateService_LoadNewCoCoRaHSDataInDB_Test()
        {
            // AppTaskID	TVItemID	TVItemID2	AppTaskCommand	AppTaskStatus	PercentCompleted	Parameters	Language	StartDateTime_UTC	EndDateTime_UTC	EstimatedLength_second	RemainingTime_second	LastUpdateDate_UTC	LastUpdateContactTVItemID
            // 14922	7	2018	26	1	1	|||ProvinceTVItemID,7|||Year,2018|||	1	2018-10-23 12:13:31.377	NULL	NULL	NULL	2018-10-23 12:13:31.377	2            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                int ProvinceTVItemID = 7;
                int Year = 2019;

                AppTaskModel appTaskModel = new AppTaskModel()
                {
                    AppTaskID = 0,
                    TVItemID = ProvinceTVItemID,
                    TVItemID2 = ProvinceTVItemID,
                    AppTaskCommand = AppTaskCommandEnum.Error,
                    AppTaskStatus = AppTaskStatusEnum.Created,
                    PercentCompleted = 1,
                    Parameters = $"|||ProvinceTVItemID,{ ProvinceTVItemID }|||Year,{ Year }|||",
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

                ClimateService _ClimateService = new ClimateService(taskRunnerBaseService);
                _ClimateService.LoadNewCoCoRaHSDataInDB();
                Assert.AreEqual(0, taskRunnerBaseService._BWObj.TextLanguageList.Count);

                break;
            }
        }
        [TestMethod]
        public void ClimateService_ParseCoCoRaHSExportData_Test()
        {
            // AppTaskID	TVItemID	TVItemID2	AppTaskCommand	AppTaskStatus	PercentCompleted	Parameters	Language	StartDateTime_UTC	EndDateTime_UTC	EstimatedLength_second	RemainingTime_second	LastUpdateDate_UTC	LastUpdateContactTVItemID
            // 14922	7	2018	26	1	1	|||ProvinceTVItemID,7|||Year,2018|||	1	2018-10-23 12:13:31.377	NULL	NULL	NULL	2018-10-23 12:13:31.377	2            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                int ProvinceTVItemID = 7;
                int Year = 2019;

                AppTaskModel appTaskModel = new AppTaskModel()
                {
                    AppTaskID = 0,
                    TVItemID = ProvinceTVItemID,
                    TVItemID2 = ProvinceTVItemID,
                    AppTaskCommand = AppTaskCommandEnum.Error,
                    AppTaskStatus = AppTaskStatusEnum.Created,
                    PercentCompleted = 1,
                    Parameters = $"|||ProvinceTVItemID,{ ProvinceTVItemID }|||Year,{ Year }|||",
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

                ClimateService _ClimateService = new ClimateService(taskRunnerBaseService);

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("ObservationDate,ObservationTime,EntryDateTime,StationNumber,StationName,Latitude,Longitude,TotalPrecipAmt,NewSnowDepth,NewSnowSWE,TotalSnowDepth,TotalSnowSWE,DateTimeStamp");
                sb.AppendLine("2019-10-29, 09:00 AM, 2019-10-29 09:18 AM, ME-FR-22, Farmington 4.8 NNW, 44.733964, -70.180071, 0.01, NA, NA, NA, NA, 2019-10-29 06:18 AM");
                sb.AppendLine("2019-10-29, 08:50 AM, 2019-10-29 08:55 AM, ME-YK-46, Parsonsfield 3.9 NE, 43.765166, -70.868641, 0.03, NA, NA, NA, NA, 2019-10-29 05:55 AM");
                sb.AppendLine("2019-10-29, 08:35 AM, 2019-10-29 08:36 AM, ME-OX-14, Bethel 1.9 SSW, 44.377379, -70.801659, 0.06, NA, NA, NA, NA, 2019-10-29 05:36 AM");
                sb.AppendLine("2019-10-29, 08:25 AM, 2019-10-29 08:26 AM, ME-CM-124, Long Island 0.5 E, 43.68476, -70.162226, 0.05, NA, NA, NA, NA, 2019-10-29 05:26 AM");
                sb.AppendLine("2019-10-29, 08:17 AM, 2019-10-29 09:18 AM, ME-PN-45, Millinocket 2.7 ESE, 45.650028, -68.644722, 0.03, NA, NA, NA, NA, 2019-10-29 06:17 AM");
                sb.AppendLine("2019-10-29, 08:15 AM, 2019-10-29 08:26 AM, ME-YK-32, Wells 5.7 W, 43.331718, -70.718105, 0.05, NA, NA, NA, NA, 2019-10-29 05:26 AM");
                sb.AppendLine("2019-10-29, 08:04 AM, 2019-10-29 08:04 AM, ME-OX-2, Hartford 1.4 N, 44.39318334, -70.34755385, 0.01, NA, NA, NA, NA, 2019-10-29 05:04 AM");
                sb.AppendLine("2019-10-29, 08:00 AM, 2019-10-29 10:53 AM, ME-CM-129, Portland 5.0 WNW, 43.700591, -70.295207, 0.10, NA, NA, NA, NA, 2019-10-29 07:53 AM");
                sb.AppendLine("2019-10-29, 08:00 AM, 2019-10-29 08:46 AM, ME-CM-114, Yarmouth 1.8 E, 43.800727, -70.165461, 0.06, NA, NA, NA, NA, 2019-10-29 05:46 AM");
                sb.AppendLine("2019-10-29, 08:00 AM, 2019-10-29 08:26 AM, ME-LN-21, Westport Island 2.2 SSW, 43.88806, -69.70889, 0.05, NA, NA, NA, NA, 2019-10-29 05:26 AM");
                sb.AppendLine("2019-10-29, 08:00 AM, 2019-10-29 08:43 AM, ME-KB-44, Winthrop 5.8 NE, 44.382429, -69.900503, 0.02, NA, NA, NA, NA, 2019-10-29 05:43 AM");
                sb.AppendLine("2019-10-29, 08:00 AM, 2019-10-29 08:14 AM, ME-AR-35, Presque Isle 3.0 NNW, 46.728653, -67.999963, 0.02, NA, NA, NA, NA, 2019-10-29 05:14 AM");
                sb.AppendLine("2019-10-29, 08:00 AM, 2019-10-29 08:20 AM, ME-HN-4, Mariaville 1.4 ESE, 44.718103, -68.388139, 0.01, NA, NA, NA, NA, 2019-10-29 05:20 AM");
                sb.AppendLine("2019-10-29, 08:00 AM, 2019-10-29 08:10 AM, ME-HN-61, Frenchboro 0.2 SSW, 44.115072, -68.365293, 0.00, 0.0, NA, NA, NA, 2019-10-29 05:10 AM");
                sb.AppendLine("2019-10-29, 07:55 AM, 2019-10-29 07:56 AM, ME-LN-13, Waldoboro 1.5 NNE, 44.114819, -69.366445, 0.02, NA, NA, NA, NA, 2019-10-29 04:55 AM");
                sb.AppendLine("2019-10-29, 07:54 AM, 2019-10-29 07:55 AM, ME-YK-51, Saco 2.5 NNW, 43.5362, -70.453692, 0.09, NA, NA, NA, NA, 2019-10-29 04:54 AM");
                sb.AppendLine("2019-10-29, 07:45 AM, 2019-10-29 10:03 AM, ME-CM-86, Pownal 4.1 S, 43.8537, -70.1721, 0.05, NA, NA, NA, NA, 2019-10-29 07:03 AM");
                sb.AppendLine("2019-10-29, 07:30 AM, 2019-10-29 06:31 AM, ME-CM-57, Freeport 0.3 S, 43.852285, -70.100747, 0.14, NA, NA, NA, NA, 2019-10-29 03:31 AM");
                sb.AppendLine("2019-10-29, 07:30 AM, 2019-10-29 07:20 AM, ME-YK-63, York 4.7 NNW, 43.22303, -70.69262, 0.10, NA, NA, NA, NA, 2019-10-29 04:20 AM");
                sb.AppendLine("2019-10-29, 07:30 AM, 2019-10-29 07:33 AM, ME-CM-18, Portland 5.5 WNW, 43.702877, -70.305483, 0.06, 0.0, NA, 0.0, NA, 2019-10-29 04:33 AM");
                sb.AppendLine("2019-10-29, 07:30 AM, 2019-10-29 07:45 AM, ME-KX-6, Union 3.0 NW, 44.24435, -69.315389, 0.01, NA, NA, NA, NA, 2019-10-29 04:45 AM");
                sb.AppendLine("2019-10-29, 07:20 AM, 2019-10-29 07:28 AM, ME-YK-65, North Berwick 5.3 W, 43.298012, -70.839569, 0.07, 0.0, NA, 0.0, NA, 2019-10-29 04:28 AM");
                sb.AppendLine("2019-10-29, 07:20 AM, 2019-10-29 07:27 AM, ME-PN-37, Orono 1.0 E, 44.88482, -68.662713, 0.01, NA, NA, NA, NA, 2019-10-29 04:27 AM");
                sb.AppendLine("2019-10-29, 07:15 AM, 2019-10-29 09:07 AM, ME-WL-12, Searsmont 3.5 WNW, 44.3768, -69.2616, 0.02, NA, NA, NA, NA, 2019-10-29 06:07 AM");
                sb.AppendLine("2019-10-29, 07:15 AM, 2019-10-29 07:32 AM, ME-PN-51, Hermon 1.2 W, 44.8107147, -68.9385492, 0.02, NA, NA, NA, NA, 2019-10-29 04:32 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 10:23 AM, ME-YK-18, Wells 3.5 SW, 43.3014, -70.6447, 0.17, NA, NA, NA, NA, 2019-10-29 07:23 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 07:04 AM, ME-YK-70, Kennebunk 0.7 NW, 43.392222, -70.555, 0.09, NA, NA, NA, NA, 2019-10-29 04:04 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 08:58 AM, ME-YK-5, Kennebunk 1.8 WNW, 43.3932, -70.58057, 0.08, NA, NA, NA, NA, 2019-10-29 05:58 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 07:09 AM, ME-YK-52, York 5.2 N, 43.2358514517546, -70.6340458989143, 0.08, 0.0, NA, NA, NA, 2019-10-29 04:09 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 07:44 AM, ME-CM-128, South Portland 1.6 S, 43.609, -70.286, 0.07, NA, NA, NA, NA, 2019-10-29 04:44 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 07:02 AM, ME-YK-28, Biddeford 1.5 NNE, 43.48679, -70.43473, 0.07, NA, NA, NA, NA, 2019-10-29 04:02 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 07:11 AM, ME-AR-18, New Sweden 4.9 NNW, 47.0108, -68.1439, 0.05, 0.0, NA, 0.0, NA, 2019-10-29 04:11 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 08:41 AM, ME-CM-54, Windham 3.9 NW, 43.8578683, -70.4619924, 0.04, 0.0, 0.00, 0.0, NA, 2019-10-29 05:41 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 08:06 AM, ME-SG-16, Topsham 3.7 NE, 43.9654846191406, -69.8855590820313, 0.04, NA, NA, NA, NA, 2019-10-29 05:06 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 07:15 AM, ME-YK-3, Hollis Center 5.4 NW, 43.6475, -70.6751, 0.04, 0.0, NA, 0.0, NA, 2019-10-29 04:15 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 07:08 AM, ME-YK-68, Berwick 0.5 N, 43.27435, -70.86503, 0.04, NA, NA, NA, NA, 2019-10-29 04:08 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 07:01 AM, ME-CM-125, Cumberland Center 4.4 NW, 43.8464, -70.3134, 0.04, NA, NA, NA, NA, 2019-10-29 04:01 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 10:03 AM, ME-FR-4, New Sharon 2.0 NW, 44.665824, -70.027181, 0.03, NA, NA, NA, NA, 2019-10-29 07:03 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 08:09 AM, ME-YK-25, Parsonsfield 4.2 NE, 43.766873, -70.865721, 0.03, NA, NA, NA, NA, 2019-10-29 05:09 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 07:43 AM, ME-KB-10, Wayne 3.2 SSE, 44.3024727, -70.0554619, 0.03, NA, NA, NA, NA, 2019-10-29 04:43 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 07:40 AM, ME-LN-4, Dresden 1.5 NW, 44.105403, -69.740637, 0.03, NA, NA, NA, NA, 2019-10-29 04:40 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 07:33 AM, ME-LN-1, Newcastle 2.1 SW, 44.027672, -69.568607, 0.03, NA, NA, NA, NA, 2019-10-29 04:33 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 07:00 AM, ME-AN-43, Lisbon 0.6 S, 44.0232, -70.1055, 0.03, NA, NA, NA, NA, 2019-10-29 04:00 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 06:46 AM, ME-KB-47, Windsor 1.1 W, 44.312812, -69.603041, 0.03, NA, NA, NA, NA, 2019-10-29 03:46 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 09:54 AM, ME-AR-33, Westmanland 2.9 N, 47.01713, -68.2242, 0.02, NA, NA, NA, NA, 2019-10-29 06:54 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 07:35 AM, ME-KB-25, Sidney 2.6 NNW, 44.449852, -69.738428, 0.02, NA, NA, NA, NA, 2019-10-29 04:35 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 07:26 AM, ME-CM-3, New Gloucester 3.0 SE, 43.92372, -70.25789, 0.02, NA, NA, NA, NA, 2019-10-29 04:26 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 07:20 AM, ME-KX-10, Union 2.1 NNE, 44.2369256168604, -69.2717143893242, 0.02, NA, NA, NA, NA, 2019-10-29 04:20 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 07:03 AM, ME-AR-30, Houlton 0.7 N, 46.12775, -67.83418, 0.02, NA, NA, NA, NA, 2019-10-29 04:03 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 06:43 AM, ME-FR-23, Farmington 4.2 NW, 44.7169, -70.1967, 0.02, NA, NA, NA, NA, 2019-10-29 03:43 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 10:11 AM, ME-WS-11, Whiting 2.3 WSW, 44.775633, -67.217733, 0.01, NA, NA, NA, NA, 2019-10-29 07:11 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 07:45 AM, ME-SM-3, Palmyra 3.5 NW, 44.871801, -69.419524, 0.01, 0.0, 0.00, 0.0, 0.00, 2019-10-29 04:45 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 07:08 AM, ME-SM-43, Detroit 0.8 SSE, 44.78201, -69.2879, 0.01, NA, NA, NA, NA, 2019-10-29 04:08 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 06:13 AM, ME-SM-10, North New Portland 0.3 WSW, 44.923537, -70.027881, 0.01, NA, NA, NA, NA, 2019-10-29 03:13 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 08:15 AM, ME-HN-20, Ellsworth 4.5 N, 44.650278, -68.499236, 0.00, 0.0, NA, NA, NA, 2019-10-29 05:15 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 08:09 AM, ME-WS-31, Eastport 1.4 ESE, 44.91044, -66.985721, 0.00, 0.0, 0.00, 0.0, 0.00, 2019-10-29 05:08 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 07:26 AM, ME-HN-7, Ellsworth 7.4 NW, 44.637419, -68.491978, 0.00, 0.0, NA, NA, NA, 2019-10-29 04:26 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 07:09 AM, ME-HN-17, Ellsworth 8.2 NNW, 44.64628333, -68.50273333, 0.00, NA, NA, NA, NA, 2019-10-29 04:09 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 07:01 AM, ME-WL-27, Prospect 2.6 W, 44.55719, -68.91732, 0.00, 0.0, NA, NA, NA, 2019-10-29 04:01 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 06:58 AM, ME-HN-57, Ellsworth 2.9 E, 44.58545, -68.437577, 0.00, 0.0, NA, NA, NA, 2019-10-29 03:58 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 08:16 AM, ME-PN-48, Lowell 0.1 SW, 45.211497, -68.474963, T, NA, NA, NA, NA, 2019-10-29 06:15 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 08:27 AM, ME-WS-10, Pembroke 5.4 SSE, 44.909473, -67.126765, T, NA, NA, NA, NA, 2019-10-29 05:27 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 08:04 AM, ME-OX-28, Oxford 2.1 W, 44.13489, -70.54943, T, NA, NA, NA, NA, 2019-10-29 05:04 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 07:26 AM, ME-WL-8, Winterport 2.9 N, 44.693691, -68.847972, T, 0.0, NA, NA, NA, 2019-10-29 04:26 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 07:15 AM, ME-LN-27, Pemaquid Point 0.3 E, 43.840527, -69.502774, T, NA, NA, NA, NA, 2019-10-29 04:15 AM");
                sb.AppendLine("2019-10-29, 07:00 AM, 2019-10-29 07:10 AM, ME-PN-54, Orono 0.6 N, 44.8925, -68.6833, T, NA, NA, NA, NA, 2019-10-29 04:10 AM");
                sb.AppendLine("2019-10-29, 06:36 AM, 2019-10-29 06:37 AM, ME-KX-12, Union 3.0 W, 44.213204, -69.341576, 0.02, NA, NA, NA, NA, 2019-10-29 03:36 AM");
                sb.AppendLine("2019-10-29, 06:30 AM, 2019-10-29 06:25 AM, ME-PS-9, Abbot 4.6 WNW, 45.218051, -69.547912, 0.04, NA, NA, NA, NA, 2019-10-29 03:25 AM");
                sb.AppendLine("2019-10-29, 06:30 AM, 2019-10-29 10:54 AM, ME-SG-1, Bath 1.1 WSW, 43.9306274, -69.8573092, 0.01, 0.0, NA, 0.0, NA, 2019-10-29 07:54 AM");
                sb.AppendLine("2019-10-29, 06:15 AM, 2019-10-29 06:22 AM, ME-AN-32, Winthrop 9.4 W, 44.304768, -70.158354, 0.05, 0.0, 0.00, 0.0, 0.00, 2019-10-29 03:22 AM");
                sb.AppendLine("2019-10-29, 06:15 AM, 2019-10-29 06:19 AM, ME-HN-2, East Surry, 44.493673, -68.45775, T, 0.0, NA, 0.0, NA, 2019-10-29 03:19 AM");
                sb.AppendLine("2019-10-29, 06:00 AM, 2019-10-29 07:41 AM, ME-FR-2, Temple 1.8 W, 44.68329, -70.264179, 0.10, NA, NA, NA, NA, 2019-10-29 04:41 AM");
                sb.AppendLine("2019-10-29, 06:00 AM, 2019-10-29 05:46 AM, ME-AR-23, Madawaska 2.5 E, 47.349337, -68.281411, 0.07, NA, NA, NA, NA, 2019-10-29 02:45 AM");
                sb.AppendLine("2019-10-29, 06:00 AM, 2019-10-29 06:33 AM, ME-YK-57, North Waterboro 1.2 NE, 43.630646, -70.714211, 0.04, NA, NA, NA, NA, 2019-10-29 03:33 AM");
                sb.AppendLine("2019-10-29, 06:00 AM, 2019-10-29 05:22 AM, ME-YK-67, Cornish 5.6 ESE, 43.783085, -70.695395, 0.02, 0.0, 0.00, 0.0, 0.00, 2019-10-29 02:21 AM");
                sb.AppendLine("2019-10-29, 06:00 AM, 2019-10-29 07:33 AM, ME-WL-3, Belmont 2.7 SSE, 44.3835, -69.1425, T, NA, NA, NA, NA, 2019-10-29 04:33 AM");
                sb.AppendLine("2019-10-29, 05:35 AM, 2019-10-29 05:39 AM, ME-PN-26, Etna 3.2 WSW, 44.8000165075064, -69.1683608293533, 0.02, 0.0, 0.00, NA, NA, 2019-10-29 02:39 AM");
                sb.AppendLine("2019-10-29, 05:30 AM, 2019-10-29 08:09 AM, ME-YK-48, Acton 2.7 NW, 43.5658107697964, -70.9429733455181, 0.04, NA, NA, NA, NA, 2019-10-29 05:09 AM");
                sb.AppendLine("2019-10-28, 05:00 PM, 2019-10-28 07:44 PM, ME-AN-40, Mechanic Falls 2.7 S, 44.073426, -70.401514, T, NA, NA, NA, NA, 2019-10-28 04:44 PM");
                sb.AppendLine("2019-10-28, 01:00 PM, 2019-10-28 07:57 PM, ME-HN-26, Brooklin 2.8 SE, 44.23945, -68.52774, 0.92, 0.0, NA, 0.0, NA, 2019-10-28 04:57 PM");
                sb.AppendLine("2019-10-28, 10:30 AM, 2019-10-29 07:19 AM, ME-KX-1, Rockport 0.1 SW, 44.183903, -69.078577, 1.29, NA, NA, NA, NA, 2019-10-29 04:19 AM");
                sb.AppendLine("2019-10-28, 09:00 AM, 2019-10-28 09:00 AM, ME-YK-46, Parsonsfield 3.9 NE, 43.765166, -70.868641, 1.44, NA, NA, NA, NA, 2019-10-28 06:00 AM");
                sb.AppendLine("2019-10-28, 09:00 AM, 2019-10-28 09:23 AM, ME-FR-22, Farmington 4.8 NNW, 44.733964, -70.180071, 0.67, NA, NA, NA, NA, 2019-10-28 06:23 AM");
                sb.AppendLine("2019-10-28, 08:25 AM, 2019-10-28 08:26 AM, ME-OX-2, Hartford 1.4 N, 44.39318334, -70.34755385, 1.26, NA, NA, NA, NA, 2019-10-28 05:26 AM");
                sb.AppendLine("2019-10-28, 08:15 AM, 2019-10-28 08:24 AM, ME-YK-32, Wells 5.7 W, 43.331718, -70.718105, 1.42, NA, NA, NA, NA, 2019-10-28 05:24 AM");
                sb.AppendLine("2019-10-28, 08:00 AM, 2019-10-28 09:43 AM, ME-KX-5, Rockport 2.8 SW, 44.15192, -69.11159, 1.50, NA, NA, NA, NA, 2019-10-28 06:43 AM");
                sb.AppendLine("2019-10-28, 08:00 AM, 2019-10-28 04:18 PM, ME-CM-8, Portland 5.1 NW, 43.71213, -70.288869, 1.40, 0.0, NA, NA, NA, 2019-10-28 01:18 PM");
                sb.AppendLine("2019-10-28, 08:00 AM, 2019-10-28 08:09 AM, ME-YK-52, York 5.2 N, 43.2358514517546, -70.6340458989143, 1.34, 0.0, NA, NA, NA, 2019-10-28 05:09 AM");
                sb.AppendLine("2019-10-28, 08:00 AM, 2019-10-28 01:01 PM, ME-CM-114, Yarmouth 1.8 E, 43.800727, -70.165461, 1.27, NA, NA, NA, NA, 2019-10-28 10:01 AM");
                sb.AppendLine("2019-10-28, 08:00 AM, 2019-10-28 08:32 AM, ME-LN-21, Westport Island 2.2 SSW, 43.88806, -69.70889, 1.25, NA, NA, NA, NA, 2019-10-28 05:32 AM");
                sb.AppendLine("2019-10-28, 08:00 AM, 2019-10-28 08:07 AM, ME-OX-14, Bethel 1.9 SSW, 44.377379, -70.801659, 1.14, NA, NA, NA, NA, 2019-10-28 05:07 AM");
                sb.AppendLine("2019-10-28, 08:00 AM, 2019-10-28 09:22 AM, ME-KB-44, Winthrop 5.8 NE, 44.382429, -69.900503, 1.11, NA, NA, NA, NA, 2019-10-28 06:22 AM");
                sb.AppendLine("2019-10-28, 08:00 AM, 2019-10-28 09:46 AM, ME-HN-21, Ellsworth 4.6 NNE, 44.64703333, -68.45741667, 0.91, NA, NA, NA, NA, 2019-10-28 06:46 AM");
                sb.AppendLine("2019-10-28, 08:00 AM, 2019-10-28 08:25 AM, ME-HN-2, East Surry, 44.493673, -68.45775, 0.85, 0.0, NA, 0.0, NA, 2019-10-28 05:25 AM");
                sb.AppendLine("2019-10-28, 08:00 AM, 2019-10-28 09:28 AM, ME-HN-4, Mariaville 1.4 ESE, 44.718103, -68.388139, 0.82, NA, NA, NA, NA, 2019-10-28 06:28 AM");
                sb.AppendLine("2019-10-28, 08:00 AM, 2019-10-28 07:56 AM, ME-HN-7, Ellsworth 7.4 NW, 44.637419, -68.491978, 0.75, NA, NA, NA, NA, 2019-10-28 04:56 AM");
                sb.AppendLine("2019-10-28, 08:00 AM, 2019-10-28 12:23 PM, ME-WS-25, Pembroke 4.8 SSE, 44.886773, -67.136843, 0.71, NA, NA, NA, NA, 2019-10-28 09:22 AM");
                sb.AppendLine("2019-10-28, 08:00 AM, 2019-10-28 08:16 AM, ME-WS-8, Lubec 4.1 W, 44.8569, -67.0683, 0.71, NA, NA, NA, NA, 2019-10-28 05:16 AM");
                sb.AppendLine("2019-10-28, 08:00 AM, 2019-10-28 07:52 AM, ME-AR-35, Presque Isle 3.0 NNW, 46.728653, -67.999963, 0.24, NA, NA, NA, NA, 2019-10-28 04:52 AM");
                sb.AppendLine("2019-10-28, 07:45 AM, 2019-10-28 07:47 AM, ME-PN-51, Hermon 1.2 W, 44.8107147, -68.9385492, 0.94, NA, NA, NA, NA, 2019-10-28 04:46 AM");
                sb.AppendLine("2019-10-28, 07:42 AM, 2019-10-28 07:43 AM, ME-HN-58, Sullivan 2.4 SSE, 44.49426, -68.13693, 0.74, NA, NA, NA, NA, 2019-10-28 04:43 AM");
                sb.AppendLine("2019-10-28, 07:40 AM, 2019-10-28 07:54 AM, ME-WL-12, Searsmont 3.5 WNW, 44.3768, -69.2616, 1.17, NA, NA, NA, NA, 2019-10-28 04:54 AM");
                sb.AppendLine("2019-10-28, 07:38 AM, 2019-10-28 07:39 AM, ME-PN-45, Millinocket 2.7 ESE, 45.650028, -68.644722, 0.56, NA, NA, NA, NA, 2019-10-28 04:38 AM");
                sb.AppendLine("2019-10-28, 07:30 AM, 2019-10-28 07:39 AM, ME-CM-98, Sebago 2.4 ESE, 43.88, -70.63, 1.48, NA, NA, NA, NA, 2019-10-28 04:38 AM");
                sb.AppendLine("2019-10-28, 07:30 AM, 2019-10-28 07:34 AM, ME-YK-65, North Berwick 5.3 W, 43.298012, -70.839569, 1.36, 0.0, NA, 0.0, NA, 2019-10-28 04:34 AM");
                sb.AppendLine("2019-10-28, 07:30 AM, 2019-10-28 06:33 AM, ME-CM-57, Freeport 0.3 S, 43.852285, -70.100747, 1.22, NA, NA, NA, NA, 2019-10-28 03:33 AM");
                sb.AppendLine("2019-10-28, 07:30 AM, 2019-10-28 07:48 AM, ME-PN-30, Old Town 4.1 ESE, 44.93543, -68.65845, 0.75, 0.0, 0.00, 0.0, 0.00, 2019-10-28 04:48 AM");
                sb.AppendLine("2019-10-28, 07:30 AM, 2019-10-28 07:44 AM, ME-HN-59, Bar Harbor 7.5 N, 44.489375, -68.233608, 0.75, NA, NA, NA, NA, 2019-10-28 04:44 AM");
                sb.AppendLine("2019-10-28, 07:20 AM, 2019-10-28 07:25 AM, ME-PN-37, Orono 1.0 E, 44.88482, -68.662713, 0.94, NA, NA, NA, NA, 2019-10-28 04:25 AM");
                sb.AppendLine("2019-10-28, 07:20 AM, 2019-10-28 09:51 AM, ME-HN-56, Surry 2.5 SSE, 44.4614, -68.4869, 0.87, NA, NA, NA, NA, 2019-10-28 06:51 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 07:29 AM, ME-YK-3, Hollis Center 5.4 NW, 43.6475, -70.6751, 1.53, 0.0, NA, 0.0, NA, 2019-10-28 04:29 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 07:11 AM, ME-YK-70, Kennebunk 0.7 NW, 43.392222, -70.555, 1.52, NA, NA, NA, NA, 2019-10-28 04:10 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 07:44 AM, ME-YK-5, Kennebunk 1.8 WNW, 43.3932, -70.58057, 1.51, NA, NA, NA, NA, 2019-10-28 04:44 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 07:18 AM, ME-YK-57, North Waterboro 1.2 NE, 43.630646, -70.714211, 1.47, NA, NA, NA, NA, 2019-10-28 04:18 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 01:26 PM, ME-YK-25, Parsonsfield 4.2 NE, 43.766873, -70.865721, 1.46, NA, NA, NA, NA, 2019-10-28 10:26 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 09:53 AM, ME-YK-18, Wells 3.5 SW, 43.3014, -70.6447, 1.44, NA, NA, NA, NA, 2019-10-28 06:53 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 10:01 AM, ME-YK-63, York 4.7 NNW, 43.22303, -70.69262, 1.42, NA, NA, NA, NA, 2019-10-28 07:01 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 07:05 AM, ME-YK-68, Berwick 0.5 N, 43.27435, -70.86503, 1.42, NA, NA, NA, NA, 2019-10-28 04:04 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 07:48 AM, ME-LN-1, Newcastle 2.1 SW, 44.027672, -69.568607, 1.41, NA, NA, NA, NA, 2019-10-28 04:48 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 08:01 AM, ME-CM-128, South Portland 1.6 S, 43.609, -70.286, 1.37, NA, NA, NA, NA, 2019-10-28 05:01 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 12:59 PM, ME-YK-8, South Berwick 4.2 NE, 43.260797, -70.734544, 1.34, NA, NA, NA, NA, 2019-10-28 09:58 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 07:19 AM, ME-YK-28, Biddeford 1.5 NNE, 43.48679, -70.43473, 1.33, NA, NA, NA, NA, 2019-10-28 04:18 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 08:32 AM, ME-CM-133, Standish 1.3 SSW, 43.71756, -70.559235, 1.32, NA, NA, NA, NA, 2019-10-28 05:32 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 07:53 AM, ME-CM-18, Portland 5.5 WNW, 43.702877, -70.305483, 1.32, 0.0, NA, 0.0, NA, 2019-10-28 04:53 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 07:06 AM, ME-CM-125, Cumberland Center 4.4 NW, 43.8464, -70.3134, 1.32, NA, NA, NA, NA, 2019-10-28 04:06 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-29 08:38 AM, ME-CM-54, Windham 3.9 NW, 43.8578683, -70.4619924, 1.27, 0.0, 0.00, 0.0, NA, 2019-10-29 05:38 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 07:45 AM, ME-LN-4, Dresden 1.5 NW, 44.105403, -69.740637, 1.23, NA, NA, NA, NA, 2019-10-28 04:45 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 07:00 AM, ME-AN-43, Lisbon 0.6 S, 44.0232, -70.1055, 1.23, NA, NA, NA, NA, 2019-10-28 04:00 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 09:08 AM, ME-CM-3, New Gloucester 3.0 SE, 43.92372, -70.25789, 1.23, NA, NA, NA, NA, 2019-10-28 06:08 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 06:09 AM, ME-CM-130, Gray 1.9 NNE, 43.911689, -70.322615, 1.23, 0.0, NA, 0.0, NA, 2019-10-28 03:09 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 11:28 AM, ME-LN-14, Whitefield 2.4 NNE, 44.203667, -69.612768, 1.22, NA, NA, NA, NA, 2019-10-28 08:28 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 11:09 AM, ME-WL-24, Swanville 2.7 S, 44.483743, -69.006467, 1.22, NA, NA, NA, NA, 2019-10-28 08:08 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 07:35 AM, ME-LN-26, Trevett 0.3 NE, 43.887802, -69.669032, 1.18, NA, NA, NA, NA, 2019-10-28 04:35 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 11:33 AM, ME-CM-86, Pownal 4.1 S, 43.8537, -70.1721, 1.16, NA, NA, NA, NA, 2019-10-28 08:33 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 09:14 AM, ME-OX-28, Oxford 2.1 W, 44.13489, -70.54943, 1.16, NA, NA, NA, NA, 2019-10-28 06:14 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 07:15 AM, ME-SG-1, Bath 1.1 WSW, 43.9306274, -69.8573092, 1.15, 0.0, NA, 0.0, NA, 2019-10-28 04:15 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 06:56 AM, ME-SG-16, Topsham 3.7 NE, 43.9654846191406, -69.8855590820313, 1.12, NA, NA, NA, NA, 2019-10-28 03:56 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 08:02 AM, ME-KB-10, Wayne 3.2 SSE, 44.3024727, -70.0554619, 1.10, NA, NA, NA, NA, 2019-10-28 05:02 AM");
                sb.AppendLine("2019-10-28, 07:00 AM, 2019-10-28 06:41 AM, ME-FR-23, Farmington 4.2 NW, 44.7169, -70.1967, 1.10, NA, NA, NA, NA, 2019-10-28 03:40 AM");

                //_ClimateService.ParseCoCoRaHSExportData(sb.ToString());
                Assert.AreEqual(0, taskRunnerBaseService._BWObj.TextLanguageList.Count);

                break;
            }
        }
        //    [TestMethod]
        //    public void ClimateService_UpdateClimateSiteInformation_Test()
        //    {
        //        foreach (string LanguageRequest in new List<string>() { "en" /*, "fr" */ })
        //        {
        //            // Arrange 
        //            SetupTest(LanguageRequest);
        //            Assert.IsNotNull(csspWebToolsTaskRunner);

        //            csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";

        //            csspWebToolsTaskRunner.StopTimer();

        //            //using (TransactionScope ts = new TransactionScope())
        //            //{
        //            RemoveAllTask();

        //            TVItemModel tvItemModelRoot = tvItemService.GetRootTVItemModelForLocationDB();
        //            Assert.AreEqual("", tvItemModelRoot.Error);

        //            TVItemModel tvItemModelProvince = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID,
        //"Prince Edward Island", TVTypeEnum.Province);
        //            Assert.AreEqual("", tvItemModelProvince.Error);

        //            climateSiteService.LanguageRequest = LanguageRequest;
        //            string retStr = climateSiteService.UpdateClimateSitesInformationForProvinceTVItemIDDB(tvItemModelProvince.TVItemID);
        //            Assert.AreEqual("", retStr);

        //            SetupBWObj(tvItemModelProvince.TVItemID, LanguageRequest, AppTaskCommandEnum.UpdateClimateSiteInformation, DateTime.Now, DateTime.Now);

        //            climateService = new ClimateService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(climateService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, climateService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            climateService.UpdateClimateSitesInformationForProvinceTVItemID();
        //            Assert.IsTrue(climateService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

        //            RemoveAllTask();

        //            //}
        //        }
        //    }
        //    [TestMethod]
        //    public void ClimateService_UpdateClimateSiteDailyAndHourlyFromStartDateToEndDate_Test()
        //    {
        //        foreach (string LanguageRequest in new List<string>() { "en" /*, "fr" */ })
        //        {
        //            // Arrange 
        //            SetupTest(LanguageRequest);
        //            Assert.IsNotNull(csspWebToolsTaskRunner);

        //            csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";

        //            csspWebToolsTaskRunner.StopTimer();

        //            //using (TransactionScope ts = new TransactionScope())
        //            //{
        //            RemoveAllTask();

        //            TVItemModel tvItemModelRoot = tvItemService.GetRootTVItemModelForLocationDB();
        //            Assert.AreEqual("", tvItemModelRoot.Error);

        //            TVItemModel tvItemModelProvince = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID,
        //(LanguageRequest == "fr" ? "Île-du-Prince-Édouard" : "Prince Edward Island"), TVTypeEnum.Province);
        //            Assert.AreEqual("", tvItemModelProvince.Error);

        //            TVItemModel tvItemModelClimateSite = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelProvince.TVItemID,
        //(LanguageRequest == "fr" ? "NORTH CAPE" : "NORTH CAPE"), TVTypeEnum.ClimateSite);
        //            Assert.AreEqual("", tvItemModelClimateSite.Error);

        //            DateTime StartDate = new DateTime(2015, 1, 1);
        //            DateTime EndDate = new DateTime(2015, 2, 1);

        //            climateSiteService.LanguageRequest = LanguageRequest;
        //            string retStr = climateSiteService.UpdateClimateSiteDailyAndHourlyFromStartDateToEndDateDB(tvItemModelClimateSite.TVItemID, StartDate, EndDate);
        //            Assert.AreEqual("", retStr);

        //            SetupBWObj(tvItemModelClimateSite.TVItemID, LanguageRequest, AppTaskCommandEnum.UpdateClimateSiteDailyAndHourlyFromStartDateToEndDate, StartDate, EndDate);

        //            climateService = new ClimateService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //            Assert.IsNotNull(climateService);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, climateService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //            climateService.UpdateClimateSiteDailyAndHourlyFromStartDateToEndDate();
        //            Assert.IsTrue(climateService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);
        //            //}
        //        }
        //    }
        //    [TestMethod]
        //    public void ClimateService_UpdateClimateSiteDailyAndHourlyForSubsectorFromStartDateToEndDate_Test()
        //    {
        //        foreach (string LanguageRequest in new List<string>() { "en" /*, "fr" */ })
        //        {
        //            // Arrange 
        //            SetupTest(LanguageRequest);
        //            Assert.IsNotNull(csspWebToolsTaskRunner);

        //            csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";

        //            csspWebToolsTaskRunner.StopTimer();

        //            //using (TransactionScope ts = new TransactionScope())
        //            //{
        //            RemoveAllTask();

        //            TVItemModel tvItemModelRoot = tvItemService.GetRootTVItemModelForLocationDB();
        //            Assert.AreEqual("", tvItemModelRoot.Error);

        //            List<string> provList = new List<string>() { "Prince Edward Island", "New Brunswick", "Nova Scotia", "Newfoundland and Labrador", "Québec", "British Columbia" };

        //            foreach (string Prov in provList)
        //            {
        //                TVItemModel tvItemModelProvince = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, Prov, TVTypeEnum.Province);

        //                Assert.AreEqual("", tvItemModelProvince.Error);

        //                List<TVItemModel> tvItemModelSubsectorList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelProvince.TVItemID, TVTypeEnum.Subsector);

        //                Assert.IsTrue(tvItemModelSubsectorList.Count > 0);

        //                //bool proceed = false;
        //                foreach (TVItemModel tvItemModelSubsector in tvItemModelSubsectorList)
        //                {
        //                    //if (tvItemModelSubsector.TVItemID == 1483)
        //                    //    proceed = true;

        //                    //if (!proceed)
        //                    //    continue;

        //                    List<YearMinMaxDate> yearMinMaxDateList = new List<YearMinMaxDate>();

        //                    Application.DoEvents();

        //                    List<int> mwqmSampleModelYearList = (from c in mwqmSampleService.db.TVItems
        //                                                         from si in mwqmSampleService.db.MWQMSites
        //                                                         from sa in mwqmSampleService.db.MWQMSamples
        //                                                         where c.ParentID == tvItemModelSubsector.TVItemID
        //                                                         && c.TVItemID == si.MWQMSiteTVItemID
        //                                                         && si.MWQMSiteTVItemID == sa.MWQMSiteTVItemID
        //                                                         orderby sa.SampleDateTime_Local
        //                                                         select sa.SampleDateTime_Local.Year).Distinct().ToList();

        //                    foreach (int year in mwqmSampleModelYearList)
        //                    {
        //                        DateTime Earliest = (from c in mwqmSampleService.db.TVItems
        //                                             from si in mwqmSampleService.db.MWQMSites
        //                                             from sa in mwqmSampleService.db.MWQMSamples
        //                                             where c.ParentID == tvItemModelSubsector.TVItemID
        //                                             && c.TVItemID == si.MWQMSiteTVItemID
        //                                             && si.MWQMSiteTVItemID == sa.MWQMSiteTVItemID
        //                                             && sa.SampleDateTime_Local.Year == year
        //                                             orderby sa.SampleDateTime_Local
        //                                             select sa.SampleDateTime_Local).FirstOrDefault();

        //                        DateTime Latest = (from c in mwqmSampleService.db.TVItems
        //                                           from si in mwqmSampleService.db.MWQMSites
        //                                           from sa in mwqmSampleService.db.MWQMSamples
        //                                           where c.ParentID == tvItemModelSubsector.TVItemID
        //                                           && c.TVItemID == si.MWQMSiteTVItemID
        //                                           && si.MWQMSiteTVItemID == sa.MWQMSiteTVItemID
        //                                           && sa.SampleDateTime_Local.Year == year
        //                                           orderby sa.SampleDateTime_Local descending
        //                                           select sa.SampleDateTime_Local).FirstOrDefault();

        //                        YearMinMaxDate yearMinMaxDate = (from c in yearMinMaxDateList
        //                                                         where c.Year == year
        //                                                         select c).FirstOrDefault();

        //                        if (yearMinMaxDate == null)
        //                        {
        //                            YearMinMaxDate yearMinMaxDateNew = new YearMinMaxDate()
        //                            {
        //                                Year = year,
        //                                MinDateTime = Earliest,
        //                                MaxDateTime = Latest,
        //                            };

        //                            yearMinMaxDateList.Add(yearMinMaxDateNew);
        //                        }
        //                        else
        //                        {
        //                            if (yearMinMaxDate.MinDateTime > Earliest)
        //                            {
        //                                yearMinMaxDate.MinDateTime = Earliest;
        //                            }
        //                            if (yearMinMaxDate.MaxDateTime < Latest)
        //                            {
        //                                yearMinMaxDate.MaxDateTime = Latest;
        //                            }
        //                        }
        //                    }
        //                    yearMinMaxDateList = (from c in yearMinMaxDateList orderby c.Year select c).ToList();

        //                    foreach (YearMinMaxDate yearMinMaxDate in yearMinMaxDateList)
        //                    {
        //                        RemoveAllTask();

        //                        climateSiteService.LanguageRequest = LanguageRequest;
        //                        string retStr = climateSiteService.UpdateClimateSiteDailyAndHourlyForSubsectorFromStartDateToEndDateDB(tvItemModelSubsector.TVItemID, yearMinMaxDate.MinDateTime, yearMinMaxDate.MaxDateTime);
        //                        Assert.AreEqual("", retStr);

        //                        SetupBWObj(tvItemModelSubsector.TVItemID, LanguageRequest, AppTaskCommandEnum.UpdateClimateSiteDailyAndHourlyForSubsectorFromStartDateToEndDate, yearMinMaxDate.MinDateTime, yearMinMaxDate.MaxDateTime);

        //                        climateService = new ClimateService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //                        Assert.IsNotNull(climateService);
        //                        Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, climateService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

        //                        climateService.UpdateClimateSiteDailyAndHourlyForSubsectorFromStartDateToEndDate();
        //                        Assert.IsTrue(climateService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

        //                        RemoveAllTask();

        //                    }
        //                }
        //            }
        //            //}
        //        }
        //    }
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
        //        if (appTaskCommand == AppTaskCommandEnum.UpdateClimateSiteInformation)
        //        {
        //            string FirstPart = "|||ProvinceTVItemID," + TVItemID.ToString() +
        //              "|||Generate,0" +
        //              "|||Command,1|||";
        //            Assert.AreEqual(FirstPart, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters);
        //            Assert.AreEqual(AppTaskCommandEnum.UpdateClimateSiteInformation, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskCommand);
        //        }
        //        else if (appTaskCommand == AppTaskCommandEnum.UpdateClimateSiteDailyAndHourlyFromStartDateToEndDate)
        //        {
        //            string FirstPart = "|||ClimateSiteTVItemID," + TVItemID.ToString() +
        //              "|||StartYear," + StartDate.Year.ToString() +
        //              "|||StartMonth," + StartDate.Month.ToString() +
        //              "|||StartDay," + StartDate.Day.ToString() +
        //              "|||EndYear," + EndDate.Year.ToString() +
        //              "|||EndMonth," + EndDate.Month.ToString() +
        //              "|||EndDay," + EndDate.Day.ToString() +
        //              "|||Generate,0" +
        //              "|||Command,1|||";
        //            Assert.AreEqual(FirstPart, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.Parameters);
        //            Assert.AreEqual(AppTaskCommandEnum.UpdateClimateSiteDailyAndHourlyFromStartDateToEndDate, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskCommand);
        //        }
        //        else if (appTaskCommand == AppTaskCommandEnum.UpdateClimateSiteDailyAndHourlyForSubsectorFromStartDateToEndDate)
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
        //            Assert.AreEqual(AppTaskCommandEnum.UpdateClimateSiteDailyAndHourlyForSubsectorFromStartDateToEndDate, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskCommand);
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
        //    climateSiteService = new ClimateSiteService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        //    mwqmSiteService = new MWQMSiteService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        //    mwqmSampleService = new MWQMSampleService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        //}
        //private void SetupShim()
        //{
        //    shimTaskRunnerBaseService = new ShimTaskRunnerBaseService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        //}
        #endregion Functions private

    }

    public class YearMinMaxDate
    {
        public int Year { get; set; }
        public DateTime MinDateTime { get; set; }
        public DateTime MaxDateTime { get; set; }
    }
}

