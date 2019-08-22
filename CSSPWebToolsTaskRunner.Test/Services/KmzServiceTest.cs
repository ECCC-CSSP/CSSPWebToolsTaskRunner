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
//using CSSPWebToolsDBDLL.Services;
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;
using CSSPDBDLL.Services;

namespace CSSPWebToolsTaskRunner.Test.Services
{
    /// <summary>
    /// Summary description for BaseMOdelServiceTest
    /// </summary>
    [TestClass]
    public class KmzServiceTest
    {
        #region Variables
        private TestContext testContextInstance { get; set; }
        #endregion Variables

        #region Properties
        private CSSPWebToolsTaskRunner csspWebToolsTaskRunner { get; set; }
   
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
        public KmzServiceTest()
        {
            csspWebToolsTaskRunner = new CSSPWebToolsTaskRunner();
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
        public void KmzService_GenerateMikeScenarioEstimatedDroguePathsAnimationKMZ_Test()
        {
            AppTaskModel appTaskModel = new AppTaskModel();
            appTaskModel.AppTaskID = 17384;
            appTaskModel.TVItemID = 28475;
            appTaskModel.TVItemID2 = 28475;
            appTaskModel.AppTaskCommand = AppTaskCommandEnum.CreateDocumentFromParameters;
            appTaskModel.AppTaskStatus = AppTaskStatusEnum.Created;
            appTaskModel.PercentCompleted = 1;
            appTaskModel.Parameters = @"|||TVItemID,28475|||ReportTypeID,43|||GoogleEarthPath,!!!!!?xml version=""1.0"" encoding=""UTF-8""?@@@@@" +
@"!!!!!kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom""@@@@@" +
@"!!!!!Document@@@@@" +
@" !!!!!name@@@@@KmlFile!!!!!/name@@@@@" +
@" !!!!!Style id=""s_ylw-pushpin_hl""@@@@@" +
@"  !!!!!IconStyle@@@@@" +
@"   !!!!!scale@@@@@1.3!!!!!/scale@@@@@" +
@"   !!!!!Icon@@@@@" +
@"    !!!!!href@@@@@http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png!!!!!/href@@@@@" +
@"   !!!!!/Icon@@@@@" +
@"   !!!!!hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/@@@@@" +
@"  !!!!!/IconStyle@@@@@" +
@" !!!!!/Style@@@@@" +
@" !!!!!Style id=""s_ylw-pushpin""@@@@@" +
@"  !!!!!IconStyle@@@@@" +
@"   !!!!!scale@@@@@1.1!!!!!/scale@@@@@" +
@"   !!!!!Icon@@@@@" +
@"    !!!!!href@@@@@http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png!!!!!/href@@@@@" +
@"   !!!!!/Icon@@@@@" +
@"   !!!!!hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/@@@@@" +
@"  !!!!!/IconStyle@@@@@" +
@" !!!!!/Style@@@@@" +
@" !!!!!StyleMap id=""m_ylw-pushpin""@@@@@" +
@"  !!!!!Pair@@@@@" +
@"   !!!!!key@@@@@normal!!!!!/key@@@@@" +
@"   !!!!!styleUrl@@@@@#s_ylw-pushpin!!!!!/styleUrl@@@@@" +
@"  !!!!!/Pair@@@@@" +
@"  !!!!!Pair@@@@@" +
@"   !!!!!key@@@@@highlight!!!!!/key@@@@@" +
@"   !!!!!styleUrl@@@@@#s_ylw-pushpin_hl!!!!!/styleUrl@@@@@" +
@"  !!!!!/Pair@@@@@" +
@" !!!!!/StyleMap@@@@@" +
@" !!!!!Placemark@@@@@" +
@"  !!!!!name@@@@@Untitled Path!!!!!/name@@@@@" +
@"  !!!!!styleUrl@@@@@#m_ylw-pushpin!!!!!/styleUrl@@@@@" +
@"  !!!!!LineString@@@@@" +
@"   !!!!!tessellate@@@@@1!!!!!/tessellate@@@@@" +
@"   !!!!!coordinates@@@@@" +
@"    -64.70513588369528%%%%%46.4717867228814%%%%%0 -64.69737256625811%%%%%46.47342948729938%%%%%0 " +
@"   !!!!!/coordinates@@@@@" +
@"  !!!!!/LineString@@@@@" +
@" !!!!!/Placemark@@@@@" +
@"!!!!!/Document@@@@@" +
@"!!!!!/kml@@@@@" +
@" |||";

            appTaskModel.Language = LanguageEnum.en;
            appTaskModel.StartDateTime_UTC = new DateTime(2019, 8, 21, 17, 2, 15);
            appTaskModel.EndDateTime_UTC = null;
            appTaskModel.EstimatedLength_second = null;
            appTaskModel.RemainingTime_second = null;
            appTaskModel.LastUpdateDate_UTC = new DateTime(2019, 8, 21, 17, 4, 17);
            appTaskModel.LastUpdateContactTVItemID = 2;


            // Generated files will be located
            // \inetpub\wwwroot\csspwebtools\App_Data\28475
            foreach (string LanguageRequest in new List<string>() { "en"/*, "fr"*/ })
            {
                csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj = new BWObj()
                {
                    appTaskCommand = AppTaskCommandEnum.CreateDocumentFromParameters,
                    appTaskModel = appTaskModel,
                    bw = null,
                    Index = 0,
                    TextLanguageList = new List<TextLanguage>(),
                };

                ParametersService parametersService = new ParametersService(csspWebToolsTaskRunner._TaskRunnerBaseService);
                parametersService.Generate();
                if (csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                {
                    int ThereWasAnError = 34;
                }
                else
                {
                    int EverythingOK = 34;
                }
            }
        }
        #endregion Functions public

        #region Functions private
        #endregion Functions private

    }
}

