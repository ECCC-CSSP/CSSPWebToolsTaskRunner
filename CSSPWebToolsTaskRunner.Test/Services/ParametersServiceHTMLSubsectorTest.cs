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
        public void GenerateHTMLSubsector_With_UniqueCode_FCSummaryStatDocx_Test()
        {
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                SetupTest(LanguageRequest);

                FileInfo fi = new FileInfo(@"C:\Users\leblancc\Desktop\TestHTML\TestGenerateHTMLSubsector_FCSummaryStatDocx_" + LanguageRequest.ToString() + ".html");
                StringBuilder sbHTML = new StringBuilder();
                string Parameters = "|||TVItemID,635|||ReportTypeID,23|||Year,2017|||";
                ReportTypeModel reportTypeModel = new ReportTypeModel(); // don't really need it for testing
                TaskRunnerBaseService taskRunnerBaseService = new TaskRunnerBaseService(new List<BWObj>()); // don't really need it for testing
                ParametersService parameterService = new ParametersService(taskRunnerBaseService);
                bool retBool = parameterService.PublicGenerateHTMLSubsector_SubsectorTestDocx(fi, sbHTML, Parameters, reportTypeModel);
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
            tvItemService = new TVItemService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
            appTaskService = new AppTaskService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
            tvFileService = new TVFileService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
            mapInfoService = new MapInfoService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        }
        #endregion Functions private

    }
}

