using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;
using CSSPWebToolsTaskRunner.Services;
using System.Linq;
using CSSPModelsDLL.Models;
using System.ComponentModel;
using System.Transactions;
using CSSPModelsDLL.Services;
using System.IO;
using CSSPWebToolsDBDLL.Services;
using CSSPEnumsDLL.Enums;
using DHI.PFS;

namespace CSSPWebToolsTaskRunner.Test.Services
{
    /// <summary>
    /// Summary description for BaseMOdelServiceTest
    /// </summary>
    [TestClass]
    public class MikeScenarioFileServiceTest
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
        private MikeScenarioService _MikeScenarioService { get; set; }
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
        public MikeScenarioFileServiceTest()
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

        //#region Functions public
        ////[TestMethod]
        ////public void Testing_Remove_Duplicate_Infrastructure()
        ////{
        ////    foreach (string LanguageRequest in new List<string>() { "en" })
        ////    {
        ////        SetupTest(LanguageRequest);

        ////        InfrastructureService infrastructureService = new InfrastructureService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        ////        BoxModelService boxModelService = new BoxModelService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        ////        VPScenarioService VPScenarioService = new VPScenarioService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);

        ////        TVItemModel tvItemModelRoot = tvItemService.GetRootTVItemModelForLocationDB();
        ////        Assert.AreEqual("", tvItemModelRoot.Error);

        ////        List<TVItemModel> tvItemModelMunicipalityList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelRoot.TVItemID, TVTypeEnum.Municipality);

        ////        int total = tvItemModelMunicipalityList.Count;
        ////        int count = 0;
        ////        foreach (TVItemModel tvItemModelMuni in tvItemModelMunicipalityList)
        ////        {
        ////            count += 1;

        ////            List<TVItemModel> tvItemModelInfrastructureList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelMuni.TVItemID, TVTypeEnum.Infrastructure).OrderBy(c => c.TVText).ThenBy(c => c.TVItemID).ToList();

        ////            TVItemModel OldTVItemInf = new TVItemModel();
        ////            foreach (TVItemModel tvItemModelInf in tvItemModelInfrastructureList)
        ////            {
        ////                int TVItemID = tvItemModelInf.TVItemID;
        ////                if (OldTVItemInf.TVText == tvItemModelInf.TVText)
        ////                {
        ////                    InfrastructureModel infrastructureModel = infrastructureService.GetInfrastructureModelWithInfrastructureTVItemIDDB(tvItemModelInf.TVItemID);
        ////                    if (string.IsNullOrWhiteSpace(infrastructureModel.Error))
        ////                    {
        ////                        List<BoxModel> boxModelList = (from c in tvFileService.db.BoxModels
        ////                                                       where c.InfrastructureTVItemID == TVItemID
        ////                                                       select c).ToList();

        ////                        foreach (BoxModel boxModel in boxModelList)
        ////                        {
        ////                            tvFileService.db.BoxModels.Remove(boxModel);
        ////                        }

        ////                        try
        ////                        {
        ////                            tvFileService.db.SaveChanges();
        ////                        }
        ////                        catch (Exception ex)
        ////                        {
        ////                            int a = 34;
        ////                            a += 1;
        ////                        }

        ////                        List<VPScenario> vpScenarioList = (from c in tvFileService.db.VPScenarios
        ////                                                           where c.InfrastructureTVItemID == TVItemID
        ////                                                           select c).ToList();

        ////                        foreach (VPScenario vpScenario in vpScenarioList)
        ////                        {
        ////                            tvFileService.db.VPScenarios.Remove(vpScenario);
        ////                        }

        ////                        try
        ////                        {
        ////                            tvFileService.db.SaveChanges();
        ////                        }
        ////                        catch (Exception ex)
        ////                        {
        ////                            int a = 34;
        ////                            a += 1;
        ////                        }

        ////                        InfrastructureModel infrastructureModelRet = infrastructureService.PostDeleteInfrastructureWithInfrastructureTVItemIDDB(tvItemModelInf.TVItemID);
        ////                        if (!string.IsNullOrWhiteSpace(infrastructureModelRet.Error))
        ////                        {
        ////                            int a = 34;
        ////                            a += 1;
        ////                        }
        ////                    }
        ////                    else
        ////                    {
        ////                        Infrastructure infrastructureNew = new Infrastructure()
        ////                        {
        ////                            InfrastructureTVItemID = TVItemID,
        ////                            LastUpdateDate_UTC = DateTime.Now,
        ////                            LastUpdateContactTVItemID = 2,
        ////                        };

        ////                        try
        ////                        {
        ////                            tvItemService.db.Infrastructures.Add(infrastructureNew);
        ////                            tvItemService.db.SaveChanges();
        ////                        }
        ////                        catch (Exception ex)
        ////                        {
        ////                            int a = 34;
        ////                            a += 1;
        ////                        }

        ////                        InfrastructureLanguage infrastructureLanguageNewEN = new InfrastructureLanguage()
        ////                        {
        ////                            Comment = "empty",
        ////                            InputDataComment = "empty",
        ////                            Language = "en",
        ////                            InfrastructureID = infrastructureNew.InfrastructureID,
        ////                            TranslationStatus = (int)StatusTranslationEnum.Translated,
        ////                            LastUpdateDate_UTC = DateTime.Now,
        ////                            LastUpdateContactTVItemID = 2,
        ////                        };

        ////                        try
        ////                        {
        ////                            tvItemService.db.InfrastructureLanguages.Add(infrastructureLanguageNewEN);
        ////                            tvItemService.db.SaveChanges();
        ////                        }
        ////                        catch (Exception ex)
        ////                        {
        ////                            int a = 34;
        ////                            a += 1;
        ////                        }

        ////                        InfrastructureLanguage infrastructureLanguageNewFR = new InfrastructureLanguage()
        ////                        {
        ////                            Comment = "empty",
        ////                            InputDataComment = "empty",
        ////                            Language = "fr",
        ////                            InfrastructureID = infrastructureNew.InfrastructureID,
        ////                            TranslationStatus = (int)StatusTranslationEnum.NotTranslated,
        ////                            LastUpdateDate_UTC = DateTime.Now,
        ////                            LastUpdateContactTVItemID = 2,
        ////                        };

        ////                        try
        ////                        {
        ////                            tvItemService.db.InfrastructureLanguages.Add(infrastructureLanguageNewFR);
        ////                            tvItemService.db.SaveChanges();
        ////                        }
        ////                        catch (Exception ex)
        ////                        {
        ////                            int a = 34;
        ////                            a += 1;
        ////                        }
        ////                    }
        ////                }
        ////                else
        ////                {
        ////                    InfrastructureModel infrastructureModel = infrastructureService.GetInfrastructureModelWithInfrastructureTVItemIDDB(tvItemModelInf.TVItemID);
        ////                    if (!string.IsNullOrWhiteSpace(infrastructureModel.Error))
        ////                    {
        ////                        Infrastructure infrastructureNew = new Infrastructure()
        ////                        {
        ////                            InfrastructureTVItemID = TVItemID,
        ////                            LastUpdateDate_UTC = DateTime.Now,
        ////                            LastUpdateContactTVItemID = 2,
        ////                        };

        ////                        try
        ////                        {
        ////                            tvItemService.db.Infrastructures.Add(infrastructureNew);
        ////                            tvItemService.db.SaveChanges();
        ////                        }
        ////                        catch (Exception ex)
        ////                        {
        ////                            int a = 34;
        ////                            a += 1;
        ////                        }

        ////                        InfrastructureLanguage infrastructureLanguageNewEN = new InfrastructureLanguage()
        ////                        {
        ////                            Comment = "empty",
        ////                            InputDataComment = "empty",
        ////                            Language = "en",
        ////                            InfrastructureID = infrastructureNew.InfrastructureID,
        ////                            TranslationStatus = (int)StatusTranslationEnum.Translated,
        ////                            LastUpdateDate_UTC = DateTime.Now,
        ////                            LastUpdateContactTVItemID = 2,
        ////                        };

        ////                        try
        ////                        {
        ////                            tvItemService.db.InfrastructureLanguages.Add(infrastructureLanguageNewEN);
        ////                            tvItemService.db.SaveChanges();
        ////                        }
        ////                        catch (Exception ex)
        ////                        {
        ////                            int a = 34;
        ////                            a += 1;
        ////                        }

        ////                        InfrastructureLanguage infrastructureLanguageNewFR = new InfrastructureLanguage()
        ////                        {
        ////                            Comment = "empty",
        ////                            InputDataComment = "empty",
        ////                            Language = "fr",
        ////                            InfrastructureID = infrastructureNew.InfrastructureID,
        ////                            TranslationStatus = (int)StatusTranslationEnum.NotTranslated,
        ////                            LastUpdateDate_UTC = DateTime.Now,
        ////                            LastUpdateContactTVItemID = 2,
        ////                        };

        ////                        try
        ////                        {
        ////                            tvItemService.db.InfrastructureLanguages.Add(infrastructureLanguageNewFR);
        ////                            tvItemService.db.SaveChanges();
        ////                        }
        ////                        catch (Exception ex)
        ////                        {
        ////                            int a = 34;
        ////                            a += 1;
        ////                        }
        ////                    }
        ////                    OldTVItemInf = tvItemModelInf;
        ////                }
        ////            }
        ////        }
        ////    }
        ////}
        //[TestMethod]
        //public void Testing_CleanAllTextFieldInDB_Remove_undesirable_characters_important_for_KML_files()
        //{
        //    foreach (string LanguageRequest in new List<string>() { "en" })
        //    {
        //        SetupTest(LanguageRequest);

        //        List<InfrastructureLanguage> InfrastructureLanguageList = (from c in tvFileService.db.InfrastructureLanguages
        //                                                                   select c).ToList();

        //        foreach (InfrastructureLanguage a in InfrastructureLanguageList)
        //        {
        //            a.Comment = Clean(a.Comment);
        //            a.InputDataComment = Clean(a.InputDataComment);
        //        }

        //        try
        //        {
        //            tvFileService.db.SaveChanges();
        //        }
        //        catch (Exception)
        //        {
        //            int a = 34;
        //            a = a + 1;
        //        }

        //        InfrastructureLanguageList = new List<InfrastructureLanguage>();

        //        List<MWQMRunLanguage> MWQMRunLanguageList = (from c in tvFileService.db.MWQMRunLanguages
        //                                                     select c).ToList();

        //        foreach (MWQMRunLanguage a in MWQMRunLanguageList)
        //        {
        //            a.MWQMRunComment = Clean(a.MWQMRunComment);
        //        }

        //        try
        //        {
        //            tvFileService.db.SaveChanges();
        //        }
        //        catch (Exception)
        //        {
        //            int a = 34;
        //            a = a + 1;
        //        }

        //        MWQMRunLanguageList = new List<MWQMRunLanguage>();

        //        List<MWQMSampleLanguage> MWQMSampleLanguageList = (from c in tvFileService.db.MWQMSampleLanguages
        //                                                           select c).ToList();

        //        foreach (MWQMSampleLanguage a in MWQMSampleLanguageList)
        //        {
        //            a.MWQMSampleNote = Clean(a.MWQMSampleNote);
        //        }

        //        try
        //        {
        //            tvFileService.db.SaveChanges();
        //        }
        //        catch (Exception)
        //        {
        //            int a = 34;
        //            a = a + 1;
        //        }

        //        MWQMSampleLanguageList = new List<MWQMSampleLanguage>();

        //        List<MWQMSubsectorLanguage> MWQMSubsectorLanguageList = (from c in tvFileService.db.MWQMSubsectorLanguages
        //                                                                 select c).ToList();

        //        foreach (MWQMSubsectorLanguage a in MWQMSubsectorLanguageList)
        //        {
        //            a.SubsectorDesc = Clean(a.SubsectorDesc);
        //        }

        //        try
        //        {
        //            tvFileService.db.SaveChanges();
        //        }
        //        catch (Exception)
        //        {
        //            int a = 34;
        //            a = a + 1;
        //        }

        //        MWQMSubsectorLanguageList = new List<MWQMSubsectorLanguage>();

        //        List<PolSourceObservation> PolSourceObservationList = (from c in tvFileService.db.PolSourceObservations
        //                                                               select c).ToList();

        //        foreach (PolSourceObservation pso in PolSourceObservationList)
        //        {
        //            pso.Observation_ToBeDeleted = Clean(pso.Observation_ToBeDeleted);
        //        }

        //        try
        //        {
        //            tvFileService.db.SaveChanges();
        //        }
        //        catch (Exception)
        //        {
        //            int a = 34;
        //            a = a + 1;
        //        }

        //        PolSourceObservationList = new List<PolSourceObservation>();
        //    }
        //}

        //private string Clean(string psoObs)
        //{
        //    StringBuilder sb = new StringBuilder();
        //    foreach (char c in psoObs)
        //    {
        //        if (c < 32 && c != 10 && c != 13)
        //        {
        //        }
        //        else
        //        {
        //            sb.Append(c);
        //        }
        //    }

        //    return sb.ToString();
        //}

        //[TestMethod]
        //public void Testing_Changing_M21FM_And_MDF_file_name()
        //{
        //    List<string> FileTypeList = new List<string>() { ".m21fm", ".m3fm", ".mdf" };

        //    int MikeScenarioTVItemID = 0;
        //    int count = 0;
        //    foreach (string LanguageRequest in new List<string>() { "en" })
        //    {
        //        SetupTest(LanguageRequest);

        //        TVItemModel tvItemModelRoot = tvItemService.GetRootTVItemModelForLocationDB();
        //        Assert.AreEqual("", tvItemModelRoot.Error);

        //        List<TVItemModel> tvItemModelMikeScenarioList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelRoot.TVItemID, TVTypeEnum.MikeScenario);
        //        Assert.IsTrue(tvItemModelMikeScenarioList.Count > 0);

        //        int totalCount = tvItemModelMikeScenarioList.Count;
        //        foreach (TVItemModel tvItemModel in tvItemModelMikeScenarioList)
        //        {
        //            count += 1;
        //            MikeScenarioTVItemID = tvItemModel.TVItemID;

        //            List<TVFileModel> tvFileModelList = tvFileService.GetTVFileModelListWithParentTVItemIDDB(MikeScenarioTVItemID);
        //            Assert.IsTrue(tvFileModelList.Count > 0);

        //            tvFileModelList = (from c in tvFileModelList
        //                               from fc in FileTypeList
        //                               where c.ServerFileName.Contains(fc)
        //                               select c).ToList();

        //            foreach (TVFileModel tvFileModel in tvFileModelList)
        //            {
        //                FileInfo fi = new FileInfo(tvFileService.ChoseEDriveOrCDrive(tvFileModel.ServerFilePath) + tvFileModel.ServerFileName);

        //                if (fi.Extension == ".m21fm" || fi.Extension == ".m3fm")
        //                {
        //                    if (!fi.Exists)
        //                    {
        //                        Assert.IsTrue(false);
        //                    }

        //                    FileInfo fiSave = new FileInfo(tvFileService.ChoseEDriveOrCDrive(tvFileModel.ServerFilePath) + tvFileModel.ServerFileName + "2");

        //                    //FileInfo fiSave = new FileInfo(@"C:\Users\leblancc\Desktop\" + fi.Name);

        //                    StreamReader sr = new StreamReader(fi.FullName, Encoding.GetEncoding("iso-8859-1"));
        //                    StreamWriter sw = new StreamWriter(fiSave.FullName, false, Encoding.GetEncoding("iso-8859-1"));
        //                    while (!sr.EndOfStream)
        //                    {
        //                        string LineText = sr.ReadLine();
        //                        if (LineText.Contains("file_name = |") && !LineText.Contains("||"))
        //                        {
        //                            string FirstPart = LineText.Substring(0, LineText.IndexOf("|") + 1);
        //                            string LastPart = LineText.Substring(LineText.LastIndexOf("|"));
        //                            string MidPart = LineText.Substring(LineText.IndexOf("|") + 1, LineText.LastIndexOf("|") - LineText.IndexOf("|") - 1);
        //                            MidPart = MidPart.Substring(MidPart.LastIndexOf("\\") + 1);

        //                            LineText = FirstPart + @".\" + MidPart + LastPart;
        //                        }
        //                        else if (LineText.Contains("ResultRootFolder = |"))
        //                        {
        //                            LineText = "   ResultRootFolder = ||";
        //                        }
        //                        else if (LineText.Contains("UseCustomResultFolder ="))
        //                        {
        //                            LineText = "   UseCustomResultFolder = true";
        //                        }
        //                        else if (LineText.Contains("CustomResultFolder = |"))
        //                        {
        //                            LineText = "   CustomResultFolder = |.\\|";
        //                        }
        //                        sw.WriteLine(LineText);
        //                    }
        //                    sw.Close();
        //                    sr.Close();

        //                    fiSave = new FileInfo(fiSave.FullName);

        //                    if (fiSave.Exists)
        //                    {
        //                        if (fi.Exists)
        //                        {
        //                            fi.Delete();
        //                        }

        //                        File.Copy(fiSave.FullName, fi.FullName);
        //                    }

        //                    fiSave.Delete();
        //                    fiSave = new FileInfo(fiSave.FullName);
        //                    Assert.IsFalse(fiSave.Exists);


        //                }

        //                if (fi.Extension == ".mdf")
        //                {
        //                    if (!fi.Exists)
        //                    {
        //                        Assert.IsTrue(false);
        //                    }

        //                    FileInfo fiSave = new FileInfo(tvFileService.ChoseEDriveOrCDrive(tvFileModel.ServerFilePath) + tvFileModel.ServerFileName + "2");

        //                    //FileInfo fiSave = new FileInfo(@"C:\Users\leblancc\Desktop\" + fi.Name);

        //                    StreamReader sr = new StreamReader(fi.FullName, Encoding.GetEncoding("iso-8859-1"));
        //                    StreamWriter sw = new StreamWriter(fiSave.FullName, false, Encoding.GetEncoding("iso-8859-1"));
        //                    while (!sr.EndOfStream)
        //                    {
        //                        string LineText = sr.ReadLine();
        //                        if (LineText.Contains("SCATTER_DATA_I = |"))
        //                        {
        //                            string FirstPart = LineText.Substring(0, LineText.IndexOf("|") + 1);
        //                            string LastPart = LineText.Substring(LineText.LastIndexOf("|"));
        //                            string MidPart = LineText.Substring(LineText.IndexOf("|") + 1, LineText.LastIndexOf("|") - LineText.IndexOf("|") - 1);
        //                            MidPart = MidPart.Substring(MidPart.LastIndexOf("\\") + 1);

        //                            LineText = FirstPart + @".\" + MidPart + LastPart;
        //                        }
        //                        sw.WriteLine(LineText);
        //                    }

        //                    sw.Close();
        //                    sr.Close();

        //                    fiSave = new FileInfo(fiSave.FullName);

        //                    if (fiSave.Exists)
        //                    {
        //                        if (fi.Exists)
        //                        {
        //                            fi.Delete();
        //                        }

        //                        File.Copy(fiSave.FullName, fi.FullName);
        //                    }

        //                    fiSave.Delete();
        //                    fiSave = new FileInfo(fiSave.FullName);
        //                    Assert.IsFalse(fiSave.Exists);

        //                }
        //            }
        //        }
        //    }
        //}
        //[TestMethod]
        //public void MikeScenarioFileService_Constructor_Test()
        //{
        //    foreach (string LanguageRequest in new List<string>() { "en", "fr" })
        //    {
        //        SetupTest(LanguageRequest);
        //        Assert.IsNotNull(csspWebToolsTaskRunner);

        //        csspWebToolsTaskRunner._RichTextBoxStatus.Text = "";
        //        csspWebToolsTaskRunner.StopTimer();

        //        TVItemModel tvItemModelRoot = tvItemService.GetRootTVItemModelForLocationDB();
        //        Assert.AreEqual("", tvItemModelRoot.Error);

        //        TVItemModel tvItemModelAcadianVillage = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelRoot.TVItemID, (LanguageRequest == "en" ? "Acadian Village" : "Village Acadien"), TVTypeEnum.Municipality);
        //        Assert.AreEqual("", tvItemModelAcadianVillage.Error);

        //        List<TVItemModel> tvItemModelMikeScenarioList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelAcadianVillage.TVItemID, TVTypeEnum.MikeScenario);
        //        Assert.IsTrue(tvItemModelMikeScenarioList.Count > 0);

        //        TVItemModel tvItemModelMikeScenario = tvItemModelMikeScenarioList[0];
        //        Assert.IsTrue(tvItemModelMikeScenario.TVItemID > 0);

        //        using (TransactionScope ts = new TransactionScope())
        //        {
        //            RemoveAllTask();

        //            SetupBWObj(tvItemModelMikeScenario, AppTaskCommandEnum.MikeScenarioAskToRun, LanguageRequest);

        //            int countAppTask = appTaskService.GetAppTaskModelCountDB();
        //            Assert.AreEqual(1, countAppTask);

        //            mikeScenarioFileService = new MikeScenarioFileService(csspWebToolsTaskRunner._TaskRunnerBaseService);

        //            Assert.IsNotNull(mikeScenarioFileService);
        //            Assert.IsNotNull(mikeScenarioFileService._TaskRunnerBaseService);
        //            Assert.IsNotNull(mikeScenarioFileService.tidesAndCurrentsService);
        //            Assert.IsNotNull(mikeScenarioFileService.kmzServiceMikeScenario);
        //            Assert.AreEqual(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, mikeScenarioFileService._TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
        //        }
        //    }
        //}
        [TestMethod]
        public void MikeScenarioFileService_MikeScenarioAskToRun_Test()
        {
            // AppTaskID	TVItemID	TVItemID2	AppTaskCommand	AppTaskStatus	PercentCompleted	Parameters	Language	StartDateTime_UTC	EndDateTime_UTC	EstimatedLength_second	RemainingTime_second	LastUpdateDate_UTC	LastUpdateContactTVItemID
            // 12643   344316  344316  2   1   1 ||| MikeScenarioTVItemID,344316 ||| 1   2018 - 03 - 20 14:49:58.440 NULL NULL    NULL    2018 - 03 - 20 14:49:58.443 2
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en, LanguageEnum.fr })
            {
                SetupTest(LanguageRequest);

                int MikeScenarioTVItemID = 344316;

                FileInfo fi = new FileInfo(@"C:\Users\leblancc\Desktop\TestHTML\TestingMikeScenario344316.m21fm");
                StringBuilder sbKML = new StringBuilder();
                string Parameters = $"|||MikeScenarioTVItemID,{ MikeScenarioTVItemID }|||";

                AppTaskModel appTaskModel = new AppTaskModel()
                {
                    AppTaskID = 100000,
                    TVItemID = MikeScenarioTVItemID,
                    TVItemID2 = MikeScenarioTVItemID,
                    AppTaskCommand = AppTaskCommandEnum.MikeScenarioAskToRun,
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

                MikeScenarioFileService _MikeScenarioFileService = new MikeScenarioFileService(taskRunnerBaseService);
                MikeScenarioModel mikeScenarioModel = _MikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(MikeScenarioTVItemID);


                TVFileModel tvFileModelM21_3FM = _TVFileService.GetTVFileModelWithTVItemIDAndTVFileTypeM21FMOrM3FMDB(MikeScenarioTVItemID);
                Assert.AreEqual("", tvFileModelM21_3FM.Error);

                string ServerPath = _TVFileService.GetServerFilePath(MikeScenarioTVItemID);
                Assert.IsFalse(string.IsNullOrWhiteSpace(ServerPath));

                FileInfo fiM21_M3 = new FileInfo(ServerPath + tvFileModelM21_3FM.ServerFileName);
                Assert.IsTrue(fiM21_M3.Exists);

                PFSFile pfsFile = new PFSFile(fiM21_M3.FullName);
                Assert.IsNotNull(pfsFile);

                _MikeScenarioFileService.DoSources(pfsFile, fiM21_M3, mikeScenarioModel, MikeScenarioTVItemID, tvFileModelM21_3FM);
                Assert.AreEqual(0, taskRunnerBaseService._BWObj.TextLanguageList.Count);

                try
                {
                    pfsFile.Write(ServerPath + mikeScenarioModel.MikeScenarioTVText + fiM21_M3.Extension);
                }
                catch (Exception ex)
                {
                    int seilfj = 345;
                    // nothing
                }

                pfsFile.Close();
                _MikeScenarioFileService.FixPFSFileSystemPart(ServerPath + mikeScenarioModel.MikeScenarioTVText + fiM21_M3.Extension);

                break;
            }

        }
        //#endregion Functions public

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
        //private void SetupBWObj(TVItemModel tvItemModelMikeScenario, AppTaskCommandEnum appTaskCommand, string LanguageRequest)
        //{
        //    RemoveAllTask();

        //    if (appTaskCommand == AppTaskCommandEnum.MikeScenarioAskToRun)
        //    {
        //        AppTaskModel appTaskModel = mikeScenarioService.PostMikeScenarioAskToRunDB(tvItemModelMikeScenario.TVItemID);

        //        Assert.AreEqual("", appTaskModel.Error);
        //    }
        //    else
        //    {
        //        Assert.IsTrue(false);
        //    }

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
        //        Assert.AreEqual(tvItemModelMikeScenario.TVItemID, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //        Assert.IsNotNull(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.bw);
        //        Assert.IsFalse(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskCommand.ToString()));
        //        Assert.IsTrue(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.Index > 0);
        //        Assert.AreEqual(tvItemModelMikeScenario.TVItemID, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //        Assert.IsTrue(string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.ErrorText));
        //        Assert.IsTrue(!string.IsNullOrWhiteSpace(csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskModel.StatusText));
        //        string FirstPart = "|||MikeScenarioTVItemID," + tvItemModelMikeScenario.TVItemID.ToString() +
        //            "|||Generate,1" +
        //            "|||Command,0";
        //        Assert.AreEqual(false, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.Generate);
        //        Assert.AreEqual(true, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.Command);
        //        Assert.AreEqual(appTaskCommand, csspWebToolsTaskRunner._TaskRunnerBaseService._BWObj.appTaskCommand);
        //    }

        //}
        public void SetupTest(LanguageEnum LanguageRequest)
        {
            csspWebToolsTaskRunner = new CSSPWebToolsTaskRunner();
            _TVItemService = new TVItemService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
            _AppTaskService = new AppTaskService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
            _TVFileService = new TVFileService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
            _MapInfoService = new MapInfoService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
            _ReportTypeService = new ReportTypeService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
            _MikeScenarioService = new MikeScenarioService(LanguageRequest, csspWebToolsTaskRunner._TaskRunnerBaseService._User);
        }
        #endregion Functions private
    }
}

