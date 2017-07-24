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
    public class DfsuServiceTest
    {
        #region Variables
        private TestContext testContextInstance { get; set; }
        string directoryName = @"C:\CSSPWebToolsTaskRunner\CSSPWebToolsTaskRunner\";
        string FileName_Hydro_21 = "hydro_21.dfsu";
        string FileName_Trans_21 = "trans_21.dfsu";
        string FileName_Hydro_3 = "hydro_3.dfsu";
        string FileName_Trans_3 = "trans_3.dfsu";

        #region Variables Hydro_21 res
        #endregion Variables Hydro_21 res
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
        public DfsuServiceTest()
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
        public void DfsuService_Constructor_Test()
        {
            foreach (string LanguageRequest in new List<string>() { "en", "fr" })
            {
                SetupTest(LanguageRequest);
                Assert.IsNotNull(csspWebToolsTaskRunner);

                csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";
                csspWebToolsTaskRunner.StopTimer();

                FileInfo fi = new FileInfo(directoryName + FileName_Hydro_21);
                Assert.IsTrue(fi.Exists);

                DfsuService dfsuService = new DfsuService(csspWebToolsTaskRunner._TaskRunnerBaseService, fi.FullName);
                Assert.IsNotNull(dfsuService);
                Assert.IsFalse(string.IsNullOrWhiteSpace(dfsuService._DfsuFileName));
                Assert.IsNotNull(dfsuService.InterpolatedContourNodeList);
                Assert.IsTrue(dfsuService.InterpolatedContourNodeList.Count == 0);
                Assert.IsNotNull(dfsuService.ForwardVector);
                Assert.IsTrue(dfsuService.ForwardVector.Count == 0);
                Assert.IsNotNull(dfsuService.BackwardVector);
                Assert.IsTrue(dfsuService.BackwardVector.Count == 0);
                Assert.IsNotNull(dfsuService.ElementList);
                Assert.IsTrue(dfsuService.ElementList.Count == 0);
            }
        }
        [TestMethod]
        public void DfsuService_Hydro_21_open_Test()
        {
            foreach (string LanguageRequest in new List<string>() { "en", "fr" })
            {
                SetupTest(LanguageRequest);
                Assert.IsNotNull(csspWebToolsTaskRunner);

                csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";
                csspWebToolsTaskRunner.StopTimer();

                FileInfo fiDfsu = new FileInfo(directoryName + FileName_Hydro_21);
                Assert.IsTrue(fiDfsu.Exists);

                DfsuService dfsuService = new DfsuService(csspWebToolsTaskRunner._TaskRunnerBaseService, fiDfsu.FullName);
                dfsuService._DfsuFile = DfsuFile.Open(fiDfsu.FullName);
                Assert.IsNotNull(dfsuService._DfsuFile);
                Assert.AreEqual("", dfsuService._DfsuFile.ApplicationTitle);
                Assert.AreEqual(0, dfsuService._DfsuFile.ApplicationVersion);
                Assert.AreEqual(1766, dfsuService._DfsuFile.Code.Length);
                Assert.AreEqual(1, dfsuService._DfsuFile.Code[0]);
                Assert.AreEqual(DfsuFileType.Dfsu2D, dfsuService._DfsuFile.DfsuFileType);
                Assert.AreEqual(3104, dfsuService._DfsuFile.ElementIds.Length);
                Assert.AreEqual(1, dfsuService._DfsuFile.ElementIds[0]);
                Assert.AreEqual(3104, dfsuService._DfsuFile.ElementTable.Length);
                Assert.AreEqual(3, dfsuService._DfsuFile.ElementTable[0].Length);
                Assert.AreEqual(1, dfsuService._DfsuFile.ElementTable[0][0]);
                Assert.AreEqual(2, dfsuService._DfsuFile.ElementTable[0][1]);
                Assert.AreEqual(3, dfsuService._DfsuFile.ElementTable[0][2]);
                Assert.AreEqual(3104, dfsuService._DfsuFile.ElementType.Length);
                Assert.AreEqual(21, dfsuService._DfsuFile.ElementType[0]);
                Assert.AreEqual(@"C:\CSSPWebToolsTaskRunner\CSSPWebToolsTaskRunner\hydro_21.dfsu", dfsuService._DfsuFile.FileName);
                Assert.AreEqual("Hydrodynamic", dfsuService._DfsuFile.FileTitle);
                Assert.AreEqual(23, dfsuService._DfsuFile.ItemInfo.Count);

                // ItemInfo 0
                int i = 0;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(1, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Surface elevation", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumISurfaceElevation, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Surface Elevation", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100078, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUmeter, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("meter", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(1000, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 1
                i = 1;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(2, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Still water depth", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIStillWaterDepth, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Still Water Depth", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(110017, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUmeter, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("meter", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(1000, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 2
                i = 2;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(3, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Total water depth", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIWaterDepth, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Water Depth", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100280, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUmeter, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("meter", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(1000, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 3
                i = 3;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(4, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("U velocity", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIuVelocity, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("u-velocity component", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100269, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUmeterPerSec, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2000, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 4
                i = 4;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(5, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("V velocity", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIvVelocity, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("v-velocity component", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100270, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUmeterPerSec, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2000, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 5
                i = 5;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(6, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("P flux", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIFlowFlux, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Flow Flux", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100080, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUm3PerSecPerM, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m^3/s/m", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("m^3/s/m", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(4700, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 6
                i = 6;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(7, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Q flux", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIFlowFlux, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Flow Flux", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100080, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUm3PerSecPerM, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m^3/s/m", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("m^3/s/m", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(4700, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 7
                i = 7;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(8, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Current speed", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumICurrentSpeed, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Current Speed", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100242, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUmeterPerSec, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2000, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 8
                i = 8;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(9, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Current direction", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumICurrentDirection, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Current Direction", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100243, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUdegree, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("deg", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("degree", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2401, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 9
                i = 9;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(10, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Density", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIDensity, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Density", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100017, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUkiloGramPerM3, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("kg/m^3", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("kg/m^3", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2200, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 10
                i = 10;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(11, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Temperature", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumITemperature, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Temperature", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100006, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUdegreeCelsius, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("deg C", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("degree Celsius", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2800, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 11
                i = 11;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(12, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Salinity", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumISalinity, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Salinity", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100024, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUPSU, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("PSU", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("PSU", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(6200, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 12
                i = 12;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(13, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Wind U velocity", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIWindVelocity, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Wind Velocity", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100002, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUmeterPerSec, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2000, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 13
                i = 13;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(14, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Wind V velocity", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIWindVelocity, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Wind Velocity", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100002, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUmeterPerSec, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2000, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 14
                i = 14;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(15, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Air pressure", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIPressureSI, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Pressure", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100196, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUPascal, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("Pa", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("Pascal", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(6100, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 15
                i = 15;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(16, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Precipitation rate", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIRainfallRate, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Rainfall rate", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100088, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUmeterPerSec, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2000, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 16
                i = 16;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(17, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Evaporation rate", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIItemUndefined, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Undefined", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(999, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("-", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("undefined", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(0, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 17
                i = 17;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(18, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Drag coefficient", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIItemUndefined, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Undefined", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(999, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("-", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("undefined", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(0, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 18
                i = 18;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(19, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Eddy viscosity", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIEddyViscosity, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Eddy viscosity", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100123, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUm2PerSec, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m^2/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("m^2/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(4702, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 19
                i = 19;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(20, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("CFL number (HD)", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIItemUndefined, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Undefined", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(999, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("-", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("undefined", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(0, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 20
                i = 20;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(21, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("CFL number (AD)", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIItemUndefined, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Undefined", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(999, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("-", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("undefined", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(0, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 21
                i = 21;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(22, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Conv. angle", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIAngle, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Angle", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100184, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUdegree, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("deg", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("degree", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2401, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 22
                i = 22;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(23, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Element area", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIElementArea, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Element area", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(110030, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUm2, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m^2", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("m^2", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(3200, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);
                Assert.AreEqual(1766, dfsuService._DfsuFile.NodeIds.Length);
                Assert.AreEqual(1, dfsuService._DfsuFile.NodeIds[0]);
                Assert.AreEqual(3104, dfsuService._DfsuFile.NumberOfElements);
                Assert.AreEqual(0, dfsuService._DfsuFile.NumberOfLayers);
                Assert.AreEqual(1766, dfsuService._DfsuFile.NumberOfNodes);
                Assert.AreEqual(0, dfsuService._DfsuFile.NumberOfSigmaLayers);
                Assert.AreEqual(25, dfsuService._DfsuFile.NumberOfTimeSteps);
                Assert.AreEqual(0.0D, dfsuService._DfsuFile.Projection.Latitude, 0.0001);
                Assert.AreEqual(0.0D, dfsuService._DfsuFile.Projection.Longitude, 0.0001);
                Assert.AreEqual(0.0D, dfsuService._DfsuFile.Projection.Orientation, 0.0001);
                Assert.AreEqual(ProjectionType.Projection, dfsuService._DfsuFile.Projection.Type);
                Assert.AreEqual("LONG/LAT", dfsuService._DfsuFile.Projection.WKTString);
                Assert.AreEqual(new DateTime(2012, 1, 1, 0, 0, 0), dfsuService._DfsuFile.StartDateTime);
                Assert.AreEqual(3600.0D, dfsuService._DfsuFile.TimeStepInSeconds, 0.0001);
                Assert.AreEqual(1766, dfsuService._DfsuFile.X.Length);
                Assert.AreEqual(-64.2767258, dfsuService._DfsuFile.X[0], 0.000001);
                Assert.AreEqual(1766, dfsuService._DfsuFile.Y.Length);
                Assert.AreEqual(46.2240829, dfsuService._DfsuFile.Y[0], 0.000001);
                Assert.AreEqual(1766, dfsuService._DfsuFile.Z.Length);
                Assert.AreEqual(-0.89819, dfsuService._DfsuFile.Z[0], 0.000001);
                dfsuService._DfsuFile.Close();

            }
        }
        [TestMethod]
        public void DfsuService_Trans_21_open_Test()
        {
            foreach (string LanguageRequest in new List<string>() { "en", "fr" })
            {
                SetupTest(LanguageRequest);
                Assert.IsNotNull(csspWebToolsTaskRunner);

                csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";
                csspWebToolsTaskRunner.StopTimer();

                FileInfo fiDfsu = new FileInfo(directoryName + FileName_Trans_21);
                Assert.IsTrue(fiDfsu.Exists);

                DfsuService dfsuService = new DfsuService(csspWebToolsTaskRunner._TaskRunnerBaseService, fiDfsu.FullName);
                dfsuService._DfsuFile = DfsuFile.Open(fiDfsu.FullName);
                Assert.IsNotNull(dfsuService._DfsuFile);
                Assert.AreEqual("", dfsuService._DfsuFile.ApplicationTitle);
                Assert.AreEqual(0, dfsuService._DfsuFile.ApplicationVersion);
                Assert.AreEqual(1766, dfsuService._DfsuFile.Code.Length);
                Assert.AreEqual(1, dfsuService._DfsuFile.Code[0]);
                Assert.AreEqual(DfsuFileType.Dfsu2D, dfsuService._DfsuFile.DfsuFileType);
                Assert.AreEqual(3104, dfsuService._DfsuFile.ElementIds.Length);
                Assert.AreEqual(1, dfsuService._DfsuFile.ElementIds[0]);
                Assert.AreEqual(3104, dfsuService._DfsuFile.ElementTable.Length);
                Assert.AreEqual(3, dfsuService._DfsuFile.ElementTable[0].Length);
                Assert.AreEqual(1, dfsuService._DfsuFile.ElementTable[0][0]);
                Assert.AreEqual(2, dfsuService._DfsuFile.ElementTable[0][1]);
                Assert.AreEqual(3, dfsuService._DfsuFile.ElementTable[0][2]);
                Assert.AreEqual(3104, dfsuService._DfsuFile.ElementType.Length);
                Assert.AreEqual(21, dfsuService._DfsuFile.ElementType[0]);
                Assert.AreEqual(@"C:\CSSPWebToolsTaskRunner\CSSPWebToolsTaskRunner\trans_21.dfsu", dfsuService._DfsuFile.FileName);
                Assert.AreEqual("Transport", dfsuService._DfsuFile.FileTitle);
                Assert.AreEqual(4, dfsuService._DfsuFile.ItemInfo.Count);

                // ItemInfo 0
                int i = 0;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(1, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Concentration - component 1", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIConcentration, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Concentration", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100007, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUkiloGramPerM3, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("kg/m^3", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("kg/m^3", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2200, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 1
                i = 1;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(2, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("U-velocity", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIuVelocity, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("u-velocity component", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100269, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUmeterPerSec, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2000, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 2
                i = 2;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(3, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("V-velocity", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIvVelocity, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("v-velocity component", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100270, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUmeterPerSec, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2000, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 3
                i = 3;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(4, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("CFL number (AD)", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIItemUndefined, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Undefined", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(999, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("-", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("undefined", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(0, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);             

                Assert.AreEqual(3104, dfsuService._DfsuFile.NumberOfElements);
                Assert.AreEqual(0, dfsuService._DfsuFile.NumberOfLayers);
                Assert.AreEqual(1766, dfsuService._DfsuFile.NumberOfNodes);
                Assert.AreEqual(0, dfsuService._DfsuFile.NumberOfSigmaLayers);
                Assert.AreEqual(25, dfsuService._DfsuFile.NumberOfTimeSteps);
                Assert.AreEqual(0.0D, dfsuService._DfsuFile.Projection.Latitude, 0.0001);
                Assert.AreEqual(0.0D, dfsuService._DfsuFile.Projection.Longitude, 0.0001);
                Assert.AreEqual(0.0D, dfsuService._DfsuFile.Projection.Orientation, 0.0001);
                Assert.AreEqual(ProjectionType.Projection, dfsuService._DfsuFile.Projection.Type);
                Assert.AreEqual("LONG/LAT", dfsuService._DfsuFile.Projection.WKTString);
                Assert.AreEqual(new DateTime(2012, 1, 1, 0, 0, 0), dfsuService._DfsuFile.StartDateTime);
                Assert.AreEqual(3600.0D, dfsuService._DfsuFile.TimeStepInSeconds, 0.0001);
                Assert.AreEqual(1766, dfsuService._DfsuFile.X.Length);
                Assert.AreEqual(-64.2767258, dfsuService._DfsuFile.X[0], 0.000001);
                Assert.AreEqual(1766, dfsuService._DfsuFile.Y.Length);
                Assert.AreEqual(46.2240829, dfsuService._DfsuFile.Y[0], 0.000001);
                Assert.AreEqual(1766, dfsuService._DfsuFile.Z.Length);
                Assert.AreEqual(-0.89819, dfsuService._DfsuFile.Z[0], 0.000001);
                dfsuService._DfsuFile.Close();

            }
        }
        [TestMethod]
        public void DfsuService_Hydro_3_open_Test()
        {
            foreach (string LanguageRequest in new List<string>() { "en", "fr" })
            {
                SetupTest(LanguageRequest);
                Assert.IsNotNull(csspWebToolsTaskRunner);

                csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";
                csspWebToolsTaskRunner.StopTimer();

                FileInfo fiDfsu = new FileInfo(directoryName + FileName_Hydro_3);
                Assert.IsTrue(fiDfsu.Exists);

                DfsuService dfsuService = new DfsuService(csspWebToolsTaskRunner._TaskRunnerBaseService, fiDfsu.FullName);
                dfsuService._DfsuFile = DfsuFile.Open(fiDfsu.FullName);
                Assert.IsNotNull(dfsuService._DfsuFile);
                Assert.AreEqual("", dfsuService._DfsuFile.ApplicationTitle);
                Assert.AreEqual(0, dfsuService._DfsuFile.ApplicationVersion);
                Assert.AreEqual(6439, dfsuService._DfsuFile.Code.Length);
                Assert.AreEqual(1, dfsuService._DfsuFile.Code[0]);
                Assert.AreEqual(DfsuFileType.Dfsu3DSigmaZ, dfsuService._DfsuFile.DfsuFileType);
                Assert.AreEqual(9558, dfsuService._DfsuFile.ElementIds.Length);
                Assert.AreEqual(1, dfsuService._DfsuFile.ElementIds[0]);
                Assert.AreEqual(9558, dfsuService._DfsuFile.ElementTable.Length);
                Assert.AreEqual(6, dfsuService._DfsuFile.ElementTable[0].Length);
                Assert.AreEqual(1, dfsuService._DfsuFile.ElementTable[0][0]);
                Assert.AreEqual(6, dfsuService._DfsuFile.ElementTable[0][1]);
                Assert.AreEqual(11, dfsuService._DfsuFile.ElementTable[0][2]);
                Assert.AreEqual(2, dfsuService._DfsuFile.ElementTable[0][3]);
                Assert.AreEqual(7, dfsuService._DfsuFile.ElementTable[0][4]);
                Assert.AreEqual(12, dfsuService._DfsuFile.ElementTable[0][5]);
                Assert.AreEqual(9558, dfsuService._DfsuFile.ElementType.Length);
                Assert.AreEqual(32, dfsuService._DfsuFile.ElementType[0]);
                Assert.AreEqual(@"C:\CSSPWebToolsTaskRunner\CSSPWebToolsTaskRunner\hydro_3.dfsu", dfsuService._DfsuFile.FileName);
                Assert.AreEqual("Hydrodynamic", dfsuService._DfsuFile.FileTitle);
                Assert.AreEqual(18, dfsuService._DfsuFile.ItemInfo.Count);

                // ItemInfo 0
                int i = 0;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(1, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Z coordinate", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIItemGeometry3D, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Item geometry 3-dimensional", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100188, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUmeter, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("meter", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(1000, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 1
                i = 1;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(2, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("U velocity", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIuVelocity, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("u-velocity component", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100269, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUmeterPerSec, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2000, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 2
                i = 2;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(3, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("V velocity", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIvVelocity, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("v-velocity component", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100270, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUmeterPerSec, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2000, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 3
                i = 3;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(4, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("W velocity", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIwVelocity, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("w-velocity component", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100271, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUmeterPerSec, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2000, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 4
                i = 4;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(5, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("WS velocity", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIwVelocity, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("w-velocity component", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100271, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUmeterPerSec, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2000, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 5
                i = 5;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(6, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Current speed", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumICurrentSpeed, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Current Speed", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100242, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUmeterPerSec, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2000, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 6
                i = 6;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(7, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Current direction (Horizontal)", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumICurrentDirection, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Current Direction", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100243, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUdegree, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("deg", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("degree", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2401, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 7
                i = 7;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(8, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Current direction (Vertical)", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumICurrentDirection, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Current Direction", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100243, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUdegree, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("deg", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("degree", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2401, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 8
                i = 8;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(9, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Density", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIDensity, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Density", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100017, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUkiloGramPerM3, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("kg/m^3", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("kg/m^3", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2200, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 9
                i = 9;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(10, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Temperature", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumITemperature, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Temperature", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100006, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUdegreeCelsius, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("deg C", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("degree Celsius", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2800, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 10
                i = 10;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(11, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Salinity", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumISalinity, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Salinity", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100024, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUPSU, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("PSU", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("PSU", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(6200, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 11
                i = 11;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(12, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Turbulent kinetic energy", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumITurbulentKineticEnergy, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Turbulent kinetic energy", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100197, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUm2PerSec2, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m^2/s^2", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("m^2/s^2", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(6400, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 12
                i = 12;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(13, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Dissipation of TKE", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIDissipationTKE, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Dissipation of TKE", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100198, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUm2PerSec3, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m^2/s^3", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("m^2/s^3", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(6401, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 13
                i = 13;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(14, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Horizontal eddy viscosity", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIEddyViscosity, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Eddy viscosity", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100123, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUm2PerSec, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m^2/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("m^2/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(4702, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 14
                i = 14;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(15, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Vertical eddy viscosity", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIEddyViscosity, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Eddy viscosity", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100123, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUm2PerSec, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m^2/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("m^2/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(4702, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 15
                i = 15;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(16, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("CFL number (HD)", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIItemUndefined, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Undefined", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(999, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("-", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("undefined", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(0, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 16
                i = 16;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(17, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("CFL number (AD)", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIItemUndefined, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Undefined", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(999, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("-", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("undefined", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(0, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 17
                i = 17;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(18, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Element volume", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIElementVolume, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Element Volume", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(110231, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUm3, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m^3", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("m^3", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(1600, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                Assert.AreEqual(6439, dfsuService._DfsuFile.NodeIds.Length);
                Assert.AreEqual(1, dfsuService._DfsuFile.NodeIds[0]);
                Assert.AreEqual(9558, dfsuService._DfsuFile.NumberOfElements);
                Assert.AreEqual(8, dfsuService._DfsuFile.NumberOfLayers);
                Assert.AreEqual(6439, dfsuService._DfsuFile.NumberOfNodes);
                Assert.AreEqual(4, dfsuService._DfsuFile.NumberOfSigmaLayers);
                Assert.AreEqual(19, dfsuService._DfsuFile.NumberOfTimeSteps);
                Assert.AreEqual(0.0D, dfsuService._DfsuFile.Projection.Latitude, 0.0001);
                Assert.AreEqual(0.0D, dfsuService._DfsuFile.Projection.Longitude, 0.0001);
                Assert.AreEqual(0.0D, dfsuService._DfsuFile.Projection.Orientation, 0.0001);
                Assert.AreEqual(ProjectionType.Projection, dfsuService._DfsuFile.Projection.Type);
                Assert.AreEqual("LONG/LAT", dfsuService._DfsuFile.Projection.WKTString);
                Assert.AreEqual(new DateTime(2013, 3, 8, 15, 0, 0), dfsuService._DfsuFile.StartDateTime);
                Assert.AreEqual(3600.0D, dfsuService._DfsuFile.TimeStepInSeconds, 0.0001);
                Assert.AreEqual(6439, dfsuService._DfsuFile.X.Length);
                Assert.AreEqual(-123.821945D, dfsuService._DfsuFile.X[0], 0.000001);
                Assert.AreEqual(6439, dfsuService._DfsuFile.Y.Length);
                Assert.AreEqual(49.4700851D, dfsuService._DfsuFile.Y[0], 0.000001);
                Assert.AreEqual(6439, dfsuService._DfsuFile.Z.Length);
                Assert.AreEqual(-0.0D, dfsuService._DfsuFile.Z[0], 0.000001);
                dfsuService._DfsuFile.Close();

            }
        }
        [TestMethod]
        public void DfsuService_Trans_3_open_Test()
        {
            foreach (string LanguageRequest in new List<string>() { "en", "fr" })
            {
                SetupTest(LanguageRequest);
                Assert.IsNotNull(csspWebToolsTaskRunner);

                csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";
                csspWebToolsTaskRunner.StopTimer();

                FileInfo fiDfsu = new FileInfo(directoryName + FileName_Trans_3);
                Assert.IsTrue(fiDfsu.Exists);

                DfsuService dfsuService = new DfsuService(csspWebToolsTaskRunner._TaskRunnerBaseService, fiDfsu.FullName);
                dfsuService._DfsuFile = DfsuFile.Open(fiDfsu.FullName);
                Assert.IsNotNull(dfsuService._DfsuFile);
                Assert.AreEqual("", dfsuService._DfsuFile.ApplicationTitle);
                Assert.AreEqual(0, dfsuService._DfsuFile.ApplicationVersion);
                Assert.AreEqual(6439, dfsuService._DfsuFile.Code.Length);
                Assert.AreEqual(1, dfsuService._DfsuFile.Code[0]);
                Assert.AreEqual(DfsuFileType.Dfsu3DSigmaZ, dfsuService._DfsuFile.DfsuFileType);
                Assert.AreEqual(9558, dfsuService._DfsuFile.ElementIds.Length);
                Assert.AreEqual(1, dfsuService._DfsuFile.ElementIds[0]);
                Assert.AreEqual(9558, dfsuService._DfsuFile.ElementTable.Length);
                Assert.AreEqual(6, dfsuService._DfsuFile.ElementTable[0].Length);
                Assert.AreEqual(1, dfsuService._DfsuFile.ElementTable[0][0]);
                Assert.AreEqual(6, dfsuService._DfsuFile.ElementTable[0][1]);
                Assert.AreEqual(11, dfsuService._DfsuFile.ElementTable[0][2]);
                Assert.AreEqual(2, dfsuService._DfsuFile.ElementTable[0][3]);
                Assert.AreEqual(7, dfsuService._DfsuFile.ElementTable[0][4]);
                Assert.AreEqual(12, dfsuService._DfsuFile.ElementTable[0][5]);
                Assert.AreEqual(9558, dfsuService._DfsuFile.ElementType.Length);
                Assert.AreEqual(32, dfsuService._DfsuFile.ElementType[0]);
                Assert.AreEqual(@"C:\CSSPWebToolsTaskRunner\CSSPWebToolsTaskRunner\trans_3.dfsu", dfsuService._DfsuFile.FileName);
                Assert.AreEqual("Transport", dfsuService._DfsuFile.FileTitle);
                Assert.AreEqual(6, dfsuService._DfsuFile.ItemInfo.Count);

                // ItemInfo 0
                int i = 0;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(1, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Z coordinate", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIItemGeometry3D, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Item geometry 3-dimensional", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100188, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUmeter, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("meter", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(1000, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 1
                i = 1;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(2, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("Concentration - component 1", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIConcentration, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Concentration", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100007, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUkiloGramPerM3, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("kg/m^3", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("kg/m^3", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2200, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 2
                i = 2;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(3, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("U-velocity", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIuVelocity, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("u-velocity component", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100269, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUmeterPerSec, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2000, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 3
                i = 3;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(4, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("V-velocity", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIvVelocity, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("v-velocity component", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100270, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUmeterPerSec, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2000, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 4
                i = 4;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(5, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("W-velocity", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIwVelocity, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("w-velocity component", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(100271, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUmeterPerSec, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("m/s", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(2000, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);

                // ItemInfo 5
                i = 5;
                Assert.AreEqual(UnitConversionType.NoConversion, dfsuService._DfsuFile.ItemInfo[i].ConversionType);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].ConversionUnit);
                Assert.AreEqual(DfsSimpleType.Float, dfsuService._DfsuFile.ItemInfo[i].DataType);
                Assert.AreEqual(6, dfsuService._DfsuFile.ItemInfo[i].ItemNumber);
                Assert.AreEqual("CFL number (AD)", dfsuService._DfsuFile.ItemInfo[i].Name);
                Assert.AreEqual(DataValueType.Instantaneous, dfsuService._DfsuFile.ItemInfo[i].ValueType);
                Assert.AreEqual(eumItem.eumIItemUndefined, dfsuService._DfsuFile.ItemInfo[i].Quantity.Item);
                Assert.AreEqual("Undefined", dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemDescription);
                Assert.AreEqual(999, dfsuService._DfsuFile.ItemInfo[i].Quantity.ItemInt);
                Assert.AreEqual(eumUnit.eumUUnitUndefined, dfsuService._DfsuFile.ItemInfo[i].Quantity.Unit);
                Assert.AreEqual("-", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitAbbreviation);
                Assert.AreEqual("undefined", dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitDescription);
                Assert.AreEqual(0, dfsuService._DfsuFile.ItemInfo[i].Quantity.UnitInt);
              

                Assert.AreEqual(6439, dfsuService._DfsuFile.NodeIds.Length);
                Assert.AreEqual(1, dfsuService._DfsuFile.NodeIds[0]);
                Assert.AreEqual(9558, dfsuService._DfsuFile.NumberOfElements);
                Assert.AreEqual(8, dfsuService._DfsuFile.NumberOfLayers);
                Assert.AreEqual(6439, dfsuService._DfsuFile.NumberOfNodes);
                Assert.AreEqual(4, dfsuService._DfsuFile.NumberOfSigmaLayers);
                Assert.AreEqual(19, dfsuService._DfsuFile.NumberOfTimeSteps);
                Assert.AreEqual(0.0D, dfsuService._DfsuFile.Projection.Latitude, 0.0001);
                Assert.AreEqual(0.0D, dfsuService._DfsuFile.Projection.Longitude, 0.0001);
                Assert.AreEqual(0.0D, dfsuService._DfsuFile.Projection.Orientation, 0.0001);
                Assert.AreEqual(ProjectionType.Projection, dfsuService._DfsuFile.Projection.Type);
                Assert.AreEqual("LONG/LAT", dfsuService._DfsuFile.Projection.WKTString);
                Assert.AreEqual(new DateTime(2013, 3, 8, 15, 0, 0), dfsuService._DfsuFile.StartDateTime);
                Assert.AreEqual(3600.0D, dfsuService._DfsuFile.TimeStepInSeconds, 0.0001);
                Assert.AreEqual(6439, dfsuService._DfsuFile.X.Length);
                Assert.AreEqual(-123.821945D, dfsuService._DfsuFile.X[0], 0.000001);
                Assert.AreEqual(6439, dfsuService._DfsuFile.Y.Length);
                Assert.AreEqual(49.4700851D, dfsuService._DfsuFile.Y[0], 0.000001);
                Assert.AreEqual(6439, dfsuService._DfsuFile.Z.Length);
                Assert.AreEqual(-0.0D, dfsuService._DfsuFile.Z[0], 0.000001);
                dfsuService._DfsuFile.Close();

            }
        }
        #endregion Functions public

        #region Functions private
        public void SetupTest(string LanguageRequest)
        {
            csspWebToolsTaskRunner = new CSSPWebToolsTaskRunner();

            FileInfo fi = new FileInfo(directoryName + FileName_Hydro_21);
            Assert.IsTrue(fi.Exists);

            fi = new FileInfo(directoryName + FileName_Trans_21);
            Assert.IsTrue(fi.Exists);

            fi = new FileInfo(directoryName + FileName_Hydro_3);
            Assert.IsTrue(fi.Exists);

            fi = new FileInfo(directoryName + FileName_Trans_3);
            Assert.IsTrue(fi.Exists);

        }
        private void SetupShim()
        {
            shimTaskRunnerBaseService = new ShimTaskRunnerBaseService(csspWebToolsTaskRunner._TaskRunnerBaseService);
        }
        #endregion Functions private
    }
}
