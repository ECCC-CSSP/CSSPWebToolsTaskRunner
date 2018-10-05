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
        private KmzService kmzService { get; set; }
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
        #endregion Functions public

        #region Functions private
        #endregion Functions private

    }
}

