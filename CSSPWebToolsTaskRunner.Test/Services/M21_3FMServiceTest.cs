using CSSPWebToolsTaskRunner.Services;
using CSSPWebToolsTaskRunner.Services.Fakes;
using DHI.Generic.MikeZero;
using DHI.Generic.MikeZero.DFS;
using DHI.Generic.MikeZero.DFS.dfsu;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;

namespace CSSPWebToolsTaskRunner.Test.Services
{
    /// <summary>
    /// Summary description for BaseMOdelServiceTest
    /// </summary>
    [TestClass]
    public class M21_3FMServiceTest
    {
        #region Variables
        private TestContext testContextInstance { get; set; }
        string directoryName = @"C:\CSSPWebToolsTaskRunner\CSSPWebToolsTaskRunner\";
        string FileName_m21fm = "Cap-Pele.m21fm";
        string FileName_m3fm = "Sechelt.m3fm";
        string FileName_log = "Cap-Pele.log";
        #endregion Variables

        #region Properties
        private CSSPWebToolsTaskRunner csspWebToolsTaskRunner { get; set; }
        private ShimTaskRunnerBaseService shimTaskRunnerBaseService { get; set; }
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
        public M21_3FMServiceTest()
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
        public void M21_3FMService_Constructor_Test()
        {
            foreach (string LanguageRequest in new List<string>() { "en", "fr" })
            {
                SetupTest(LanguageRequest);
                Assert.IsNotNull(csspWebToolsTaskRunner);

                csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";
                csspWebToolsTaskRunner.StopTimer();

                FileInfo fi = new FileInfo(directoryName + FileName_m21fm);
                Assert.IsTrue(fi.Exists);

                M21_3FMService m21_3FMService = new M21_3FMService(csspWebToolsTaskRunner._TaskRunnerBaseService);
                Assert.IsNotNull(m21_3FMService);
                Assert.IsNotNull(m21_3FMService._TaskRunnerBaseService);
                Assert.IsNotNull(m21_3FMService._TaskRunnerBaseService._BWObj);
                Assert.IsNotNull(m21_3FMService.topfileinfo);
                Assert.IsNotNull(m21_3FMService.femEngineHD);
                Assert.IsNotNull(m21_3FMService.system);
                Assert.IsNotNull(m21_3FMService.sbFileTxt);

                fi = new FileInfo(directoryName + FileName_m3fm);
                Assert.IsTrue(fi.Exists);

                m21_3FMService = new M21_3FMService(csspWebToolsTaskRunner._TaskRunnerBaseService);
                Assert.IsNotNull(m21_3FMService);
                Assert.IsNotNull(m21_3FMService._TaskRunnerBaseService);
                Assert.IsNotNull(m21_3FMService._TaskRunnerBaseService._BWObj);
                Assert.IsNotNull(m21_3FMService.topfileinfo);
                Assert.IsNotNull(m21_3FMService.femEngineHD);
                Assert.IsNotNull(m21_3FMService.system);
                Assert.IsNotNull(m21_3FMService.sbFileTxt);
            }
        }
        [TestMethod]
        public void M21_3FMService_Read_M21_3FM_File_Test()
        {
            foreach (string LanguageRequest in new List<string>() { "en", "fr" })
            {
                SetupTest(LanguageRequest);
                Assert.IsNotNull(csspWebToolsTaskRunner);

                csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";
                csspWebToolsTaskRunner.StopTimer();

                FileInfo fi = new FileInfo(directoryName + FileName_m21fm);
                Assert.IsTrue(fi.Exists);

                M21_3FMService m21_3FMService = new M21_3FMService(csspWebToolsTaskRunner._TaskRunnerBaseService);
                Assert.IsNotNull(m21_3FMService);
                Assert.IsNotNull(m21_3FMService._TaskRunnerBaseService);
                Assert.IsNotNull(m21_3FMService._TaskRunnerBaseService._BWObj);
                Assert.IsNotNull(m21_3FMService.topfileinfo);
                Assert.IsNotNull(m21_3FMService.femEngineHD);
                Assert.IsNotNull(m21_3FMService.system);
                Assert.IsNotNull(m21_3FMService.sbFileTxt);

                m21_3FMService.Read_M21_3FM_File(fi);
                Assert.IsTrue(m21_3FMService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

                fi = new FileInfo(directoryName + FileName_m3fm);
                Assert.IsTrue(fi.Exists);

                m21_3FMService = new M21_3FMService(csspWebToolsTaskRunner._TaskRunnerBaseService);
                Assert.IsNotNull(m21_3FMService);
                Assert.IsNotNull(m21_3FMService._TaskRunnerBaseService);
                Assert.IsNotNull(m21_3FMService._TaskRunnerBaseService._BWObj);
                Assert.IsNotNull(m21_3FMService.topfileinfo);
                Assert.IsNotNull(m21_3FMService.femEngineHD);
                Assert.IsNotNull(m21_3FMService.system);
                Assert.IsNotNull(m21_3FMService.sbFileTxt);

                m21_3FMService.Read_M21_3FM_File(fi);
                Assert.IsTrue(m21_3FMService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);
            }
        }
        [TestMethod]
        public void M21_3FMService_Write_M21_3FM_File_Test()
        {
            foreach (string LanguageRequest in new List<string>() { "en", "fr" })
            {
                SetupTest(LanguageRequest);
                Assert.IsNotNull(csspWebToolsTaskRunner);

                csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";
                csspWebToolsTaskRunner.StopTimer();

                FileInfo fi = new FileInfo(directoryName + FileName_m21fm);
                Assert.IsTrue(fi.Exists);

                M21_3FMService m21_3FMService = new M21_3FMService(csspWebToolsTaskRunner._TaskRunnerBaseService);
                Assert.IsNotNull(m21_3FMService);
                Assert.IsNotNull(m21_3FMService._TaskRunnerBaseService);
                Assert.IsNotNull(m21_3FMService._TaskRunnerBaseService._BWObj);
                Assert.IsNotNull(m21_3FMService.topfileinfo);
                Assert.IsNotNull(m21_3FMService.femEngineHD);
                Assert.IsNotNull(m21_3FMService.system);
                Assert.IsNotNull(m21_3FMService.sbFileTxt);

                m21_3FMService.Read_M21_3FM_File(fi);
                Assert.IsTrue(m21_3FMService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

                FileInfo fiWrite = new FileInfo(directoryName + "New" + FileName_m21fm);

                m21_3FMService.Write_M21_3FM_File(fiWrite, m21_3FMService);
                Assert.IsTrue(m21_3FMService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);
                Assert.IsTrue(fiWrite.Exists);

                if (fiWrite.Exists)
                {
                    //fiWrite.Delete();
                }

                fi = new FileInfo(directoryName + FileName_m3fm);
                Assert.IsTrue(fi.Exists);

                m21_3FMService = new M21_3FMService(csspWebToolsTaskRunner._TaskRunnerBaseService);
                Assert.IsNotNull(m21_3FMService);
                Assert.IsNotNull(m21_3FMService._TaskRunnerBaseService);
                Assert.IsNotNull(m21_3FMService._TaskRunnerBaseService._BWObj);
                Assert.IsNotNull(m21_3FMService.topfileinfo);
                Assert.IsNotNull(m21_3FMService.femEngineHD);
                Assert.IsNotNull(m21_3FMService.system);
                Assert.IsNotNull(m21_3FMService.sbFileTxt);

                m21_3FMService.Read_M21_3FM_File(fi);
                Assert.IsTrue(m21_3FMService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

                fiWrite = new FileInfo(directoryName + "New" + FileName_m3fm);

                m21_3FMService.Write_M21_3FM_File(fiWrite, m21_3FMService);
                Assert.IsTrue(m21_3FMService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);

                Assert.IsTrue(fiWrite.Exists);

                if (fiWrite.Exists)
                {
                    //fiWrite.Delete();
                }

            }
        }
        [TestMethod]
        public void M21_3FMService_Read_M21_3FM_Log_Test()
        {
            foreach (string LanguageRequest in new List<string>() { "en", "fr" })
            {
                SetupTest(LanguageRequest);
                Assert.IsNotNull(csspWebToolsTaskRunner);

                csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";
                csspWebToolsTaskRunner.StopTimer();

                FileInfo fi = new FileInfo(directoryName + FileName_log);
                Assert.IsTrue(fi.Exists);

                M21_3FMLogService m21_3FMLogService = new M21_3FMLogService(csspWebToolsTaskRunner._TaskRunnerBaseService);
                Assert.IsNotNull(m21_3FMLogService);
                Assert.IsNotNull(m21_3FMLogService._TaskRunnerBaseService);
                Assert.IsNotNull(m21_3FMLogService._TaskRunnerBaseService._BWObj);
                Assert.IsNotNull(m21_3FMLogService.StartExecutionDate);
                Assert.IsNotNull(m21_3FMLogService.TotalElapseTimeInSeconds);
                Assert.IsNotNull(m21_3FMLogService.CompletionTxt);

                m21_3FMLogService.Read_M21_3FM_Log(fi);
                Assert.IsTrue(m21_3FMLogService._TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0);
                Assert.AreEqual(new DateTime(2015, 7, 31, 10, 49, 31), m21_3FMLogService.StartExecutionDate);
                Assert.AreEqual(137, m21_3FMLogService.TotalElapseTimeInSeconds);
                Assert.AreEqual("Normal run completion", m21_3FMLogService.CompletionTxt);
            }
        }
        #endregion Functions public

        #region Functions private
        public void SetupTest(string LanguageRequest)
        {
            csspWebToolsTaskRunner = new CSSPWebToolsTaskRunner();

            csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj = new BWObj() { TextLanguageList = new List<TextLanguage>() };

            FileInfo fi = new FileInfo(directoryName + FileName_m21fm);
            Assert.IsTrue(fi.Exists);

            fi = new FileInfo(directoryName + FileName_m3fm);
            Assert.IsTrue(fi.Exists);
        }
        private void SetupShim()
        {
            shimTaskRunnerBaseService = new ShimTaskRunnerBaseService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        }
        #endregion Functions private
    }
}
