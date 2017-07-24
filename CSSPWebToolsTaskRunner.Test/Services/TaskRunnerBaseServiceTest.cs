using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using CSSPWebToolsDBDLL.Models;
using System.Security.Principal;
using CSSPWebToolsDBDLL.Services;
using CSSPWebToolsDBDLL.Services.Resources;
using System.Transactions;
using CSSPWebToolsDBDLL.Services.Fakes;
using Microsoft.QualityTools.Testing.Fakes;
using System.Threading;
using System.Globalization;
using CSSPWebToolsTaskRunner.Services;
using CSSPWebToolsTaskRunner;

namespace CSSPWebToolsTaskRunner.Test.Services
{
    /// <summary>
    /// Summary description for AddressServiceTest
    /// </summary>
    [TestClass]
    public class TestRunnerBaseServiceTest
    {
        #region Variables
        private TestContext testContextInstance { get; set; }
        #endregion Variables

        #region Properties
        private CSSPWebToolsTaskRunner csspWebToolsTaskRunner { get; set; }
        private ShimAppTaskLanguageService shimAppTaskLanguageService { get; set; }
        private ShimAppTaskService shimAppTaskService { get; set; }
        private ShimMapInfoService shimMapInfoService { get; set; }
        private ShimMikeBoundaryConditionService shimMikeBoundaryConditionService { get; set; }
        private ShimMikeScenarioService shimMikeScenarioService { get; set; }
        private ShimMikeSourceService shimMikeSourceService { get; set; }
        private ShimMikeSourceStartEndService shimMikeSourceStartEndService { get; set; }
        private ShimTVFileService shimTVFileService { get; set; }
        private ShimTVItemLanguageService shimTVItemLanguageService { get; set; }
        private ShimTVItemService shimTVItemService { get; set; }
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
        public TestRunnerBaseServiceTest()
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
        public void TaskRunnerBaseService_Constructor_Test()
        {
            // Arrange 
            SetupTest();

            // in Arrange
            Assert.IsNotNull(csspWebToolsTaskRunner);
            Assert.IsNotNull(csspWebToolsTaskRunner._User);
            Assert.AreEqual("Charles.LeBlanc2@Canada.ca", csspWebToolsTaskRunner._User.Identity.Name);
        }
        [TestMethod]
        public void TaskRunnerBaseService_CleanText_Empty_Test()
        {
            // Arrange 
            SetupTest();

            string TextToClean = "";
            string retStr = csspWebToolsTaskRunner._TaskRunnerBaseService.CleanText(TextToClean);
            Assert.IsNotNull("", retStr);
        }
        [TestMethod]
        public void TaskRunnerBaseService_CleanText_HasSlash_Test()
        {
            // Arrange 
            SetupTest();

            string TextToClean = "/allo/";
            string retStr = csspWebToolsTaskRunner._TaskRunnerBaseService.CleanText(TextToClean);
            Assert.IsNotNull("-allo-", retStr);
        }
        [TestMethod]
        public void TaskRunnerBaseService_CleanText_HasTabs_Test()
        {
            // Arrange 
            SetupTest();

            string TextToClean = "/allo\t/";
            string retStr = csspWebToolsTaskRunner._TaskRunnerBaseService.CleanText(TextToClean);
            Assert.IsNotNull("-allo -", retStr);
        }
        [TestMethod]
        public void TaskRunnerBaseService_CleanText_HasReturn_Test()
        {
            // Arrange 
            SetupTest();

            string TextToClean = "/allo\t\r";
            string retStr = csspWebToolsTaskRunner._TaskRunnerBaseService.CleanText(TextToClean);
            Assert.IsNotNull("-allo", retStr);
        }
        [TestMethod]
        public void TaskRunnerBaseService_CleanText_HasNewLine_Test()
        {
            // Arrange 
            SetupTest();

            string TextToClean = "/allo\t\r\n";
            string retStr = csspWebToolsTaskRunner._TaskRunnerBaseService.CleanText(TextToClean);
            Assert.IsNotNull("-allo", retStr);
        }
        [TestMethod]
        public void TaskRunnerBaseService_CleanText_HasAnd_Test()
        {
            // Arrange 
            SetupTest();

            string TextToClean = "/allo&a";
            string retStr = csspWebToolsTaskRunner._TaskRunnerBaseService.CleanText(TextToClean);
            Assert.IsNotNull("-alloAnda", retStr);
        }
        [TestMethod]
        public void TaskRunnerBaseService_CleanText_HasPound_Test()
        {
            // Arrange 
            SetupTest();

            string TextToClean = "/allo#a";
            string retStr = csspWebToolsTaskRunner._TaskRunnerBaseService.CleanText(TextToClean);
            Assert.IsNotNull("-allo a", retStr);
        }
        [TestMethod]
        public void TaskRunnerBaseService_CleanText_HasManySpaces_Test()
        {
            // Arrange 
            SetupTest();

            string TextToClean = "/allo            a";
            string retStr = csspWebToolsTaskRunner._TaskRunnerBaseService.CleanText(TextToClean);
            Assert.IsNotNull("-allo a", retStr);
        }
        [TestMethod]
        public void TaskRunnerBaseService_GenerateDoc_Test()
        {
            // Arrange 
            SetupTest();

            string TextToClean = "/allo            a";
            string retStr = csspWebToolsTaskRunner._TaskRunnerBaseService.CleanText(TextToClean);
            Assert.IsNotNull("-allo a", retStr);
        }
        #endregion Testing Methods Public

        #region Testing Methods Private
        #endregion Testing Methods Private

        #region Functions private
        public void SetupTest()
        {
            csspWebToolsTaskRunner = new CSSPWebToolsTaskRunner();
        }
        private void SetupShim()
        {
        }
        #endregion Functions private
    }
}

