using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;
using CSSPWebToolsTaskRunner.Services;
using System.Linq;
using CSSPWebToolsDB.Models;
using System.ComponentModel;
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;

namespace CSSPWebToolsTaskRunner.Test.Services
{
    /// <summary>
    /// Summary description for BaseMOdelServiceTest
    /// </summary>
    [TestClass]
    public class BaseModelServiceTest
    {
        public BaseModelServiceTest()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        private TestContext testContextInstance;

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

        #region Additional test attributes
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
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion

        [TestMethod]
        public void BWObj_Test()
        {
            PropertyInfo[] propertyInfoList = typeof(BWObj).GetProperties();

            Assert.AreEqual(11, propertyInfoList.Count());
            Assert.AreEqual("bw", propertyInfoList[0].Name);
            Assert.AreEqual(typeof(BackgroundWorker), propertyInfoList[0].PropertyType);
            Assert.AreEqual("Index", propertyInfoList[1].Name);
            Assert.AreEqual(typeof(System.Int32), propertyInfoList[1].PropertyType);
            Assert.AreEqual("appTaskCommand", propertyInfoList[2].Name);
            Assert.AreEqual(typeof(AppTaskCommandEnum), propertyInfoList[2].PropertyType);
            Assert.AreEqual("appTaskModel", propertyInfoList[3].Name);
            Assert.AreEqual(typeof(AppTaskModel), propertyInfoList[3].PropertyType);
            Assert.AreEqual("MWQMPlanID", propertyInfoList[4].Name);
            Assert.AreEqual(typeof(System.Int32), propertyInfoList[4].PropertyType);
            Assert.AreEqual("Generate", propertyInfoList[7].Name);
            Assert.AreEqual(typeof(System.Boolean), propertyInfoList[7].PropertyType);
            Assert.AreEqual("Command", propertyInfoList[8].Name);
            Assert.AreEqual(typeof(System.Boolean), propertyInfoList[8].PropertyType);
            Assert.AreEqual("FileName", propertyInfoList[9].Name);
            Assert.AreEqual(typeof(System.String), propertyInfoList[9].PropertyType);
        }
        [TestMethod]
        public void CurrentResult_Test()
        {
            PropertyInfo[] propertyInfoList = typeof(CurrentResult).GetProperties();

            Assert.AreEqual(3, propertyInfoList.Count());
            Assert.AreEqual("Date", propertyInfoList[0].Name);
            Assert.AreEqual(typeof(System.DateTime), propertyInfoList[0].PropertyType);
            Assert.AreEqual("x_velocity", propertyInfoList[1].Name);
            Assert.AreEqual(typeof(System.Double), propertyInfoList[1].PropertyType);
            Assert.AreEqual("y_velocity", propertyInfoList[2].Name);
            Assert.AreEqual(typeof(System.Double), propertyInfoList[2].PropertyType);
        }
        [TestMethod]
        public void OtherFileInfo_Test()
        {
            PropertyInfo[] propertyInfoList = typeof(OtherFileInfo).GetProperties();

            Assert.AreEqual(4, propertyInfoList.Count());
            Assert.AreEqual("TVFileID", propertyInfoList[0].Name);
            Assert.AreEqual(typeof(System.Int32), propertyInfoList[0].PropertyType);
            Assert.AreEqual("ClientFullFileName", propertyInfoList[1].Name);
            Assert.AreEqual(typeof(System.String), propertyInfoList[1].PropertyType);
            Assert.AreEqual("ServerFullFileName", propertyInfoList[2].Name);
            Assert.AreEqual(typeof(System.String), propertyInfoList[2].PropertyType);
            Assert.AreEqual("IsOutput", propertyInfoList[3].Name);
            Assert.AreEqual(typeof(System.Boolean), propertyInfoList[3].PropertyType);
        }
        [TestMethod]
        public void PeakDifference_Test()
        {
            PropertyInfo[] propertyInfoList = typeof(PeakDifference).GetProperties();

            Assert.AreEqual(3, propertyInfoList.Count());
            Assert.AreEqual("StartDate", propertyInfoList[0].Name);
            Assert.AreEqual(typeof(System.DateTime), propertyInfoList[0].PropertyType);
            Assert.AreEqual("EndDate", propertyInfoList[1].Name);
            Assert.AreEqual(typeof(System.DateTime), propertyInfoList[1].PropertyType);
            Assert.AreEqual("Value", propertyInfoList[2].Name);
            Assert.AreEqual(typeof(System.Single), propertyInfoList[2].PropertyType);
        }
        [TestMethod]
        public void UserStateObj_Test()
        {
            PropertyInfo[] propertyInfoList = typeof(UserStateObj).GetProperties();

            Assert.AreEqual(1, propertyInfoList.Count());
            Assert.AreEqual("ProgressText", propertyInfoList[0].Name);
            Assert.AreEqual(typeof(System.String), propertyInfoList[0].PropertyType);
        }
        [TestMethod]
        public void WaterLevelResult_Test()
        {
            PropertyInfo[] propertyInfoList = typeof(WaterLevelResult).GetProperties();

            Assert.AreEqual(2, propertyInfoList.Count());
            Assert.AreEqual("Date", propertyInfoList[0].Name);
            Assert.AreEqual(typeof(System.DateTime), propertyInfoList[0].PropertyType);
            Assert.AreEqual("WaterLevel", propertyInfoList[1].Name);
            Assert.AreEqual(typeof(System.Double), propertyInfoList[1].PropertyType);
        }
    }
}
