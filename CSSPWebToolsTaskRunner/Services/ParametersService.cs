﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using CSSPWebToolsTaskRunner.Services;
using CSSPWebToolsTaskRunner.Services.Resources;
using CSSPWebToolsTaskRunner;
using System.Transactions;
using System.Text;
using CSSPDBDLL.Models;
using CSSPDBDLL.Services;
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;
using CSSPEnumsDLL.Services;
using Microsoft.Office.Interop.Word;
using System.Threading;
using System.Globalization;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        #region Variables
        private List<Node> InterpolatedContourNodeList = new List<Node>();
        private Dictionary<String, Vector> ForwardVector = new Dictionary<String, Vector>();
        private Dictionary<String, Vector> BackwardVector = new Dictionary<String, Vector>();
        #endregion Variables

        #region Properties
        private TaskRunnerBaseService _TaskRunnerBaseService { get; set; }
        private BaseEnumService _BaseEnumService { get; set; }

        private AddressService _AddressService { get; set; }
        private ContactService _ContactService { get; set; }
        private TVFileService _TVFileService { get; set; }
        private TVItemService _TVItemService { get; set; }
        private TVItemLinkService _TVItemLinkService { get; set; }
        private ReportTypeService _ReportTypeService { get; set; }
        private ReportSectionService _ReportSectionService { get; set; }
        private MWQMSubsectorService _MWQMSubsectorService { get; set; }
        private MWQMAnalysisReportParameterService _MWQMAnalysisReportParameterService { get; set; }
        private MapInfoService _MapInfoService { get; set; }
        private MikeScenarioService _MikeScenarioService { get; set; }
        private MikeBoundaryConditionService _MikeBoundaryConditionService { get; set; }
        private MikeSourceService _MikeSourceService { get; set; }
        private MikeSourceStartEndService _MikeSourceStartEndService { get; set; }
        private MWQMRunService _MWQMRunService { get; set; }
        private PolSourceSiteService _PolSourceSiteService { get; set; }
        private PolSourceObservationService _PolSourceObservationService { get; set; }
        private PolSourceObservationIssueService _PolSourceObservationIssueService { get; set; }
        private MWQMSiteService _MWQMSiteService { get; set; }
        private MWQMSampleService _MWQMSampleService { get; set; }
        private InfrastructureService _InfrastructureService { get; set; }
        private UseOfSiteService _UseOfSiteService { get; set; }
        private ClimateSiteService _ClimateSiteService { get; set; }
        private ClimateDataValueService _ClimateDataValueService { get; set; }
        private TidesAndCurrentsService _TidesAndCurrentsService { get; set; }
        private TideSiteService _TideSiteService { get; set; }
        private TelService _TelService { get; set; }
        private EmailService _EmailService { get; set; }

        public FileInfo fi { get; set; }
        public ReportTypeModel reportTypeModel { get; set; }
        public int TVItemID { get; set; }
        public int Year { get; set; }
        public int StatStartYear { get; set; }
        public int StatEndYear { get; set; }
        public string Parameters { get; set; }
        public StringBuilder sb { get; set; }
        public string FileNameExtra { get; set; }
        public Microsoft.Office.Interop.Excel._Application xlApp { get; set; }
        public Microsoft.Office.Interop.Excel.Workbook workbook { get; set; }
        public Microsoft.Office.Interop.Excel.Worksheet worksheet { get; set; }
        public Microsoft.Office.Interop.Excel.ChartObjects xlCharts { get; set; }
        public List<RunSiteInfo> RunSiteInfoList { get; set; }
        #endregion Properties

        #region Constructors
        public ParametersService(TaskRunnerBaseService taskRunnerBaseService)
        {
            xlApp = null;
            workbook = null;
            worksheet = null;
            xlCharts = null;

            reportTypeModel = new ReportTypeModel();
            TVItemID = 0;
            Year = 0;
            Parameters = "";
            fi = new FileInfo(@"C:\DoesNotExist.txt");
            sb = new StringBuilder();
            FileNameExtra = "";

            Random random = new Random();
            for (int i = 0; i < 10; i++)
            {
                FileNameExtra = FileNameExtra + (char)random.Next(97, 122);
            }

            _TaskRunnerBaseService = taskRunnerBaseService;
            _BaseEnumService = new BaseEnumService(_TaskRunnerBaseService._BWObj.appTaskModel.Language);

            _AddressService = new AddressService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _ContactService = new ContactService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _TVFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _TVItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _TVItemLinkService = new TVItemLinkService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _ReportTypeService = new ReportTypeService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _ReportSectionService = new ReportSectionService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _MWQMSubsectorService = new MWQMSubsectorService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _MWQMAnalysisReportParameterService = new MWQMAnalysisReportParameterService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _MapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _MikeScenarioService = new MikeScenarioService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _MikeBoundaryConditionService = new MikeBoundaryConditionService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _MikeSourceService = new MikeSourceService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _MikeSourceStartEndService = new MikeSourceStartEndService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _MWQMRunService = new MWQMRunService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _PolSourceSiteService = new PolSourceSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _PolSourceObservationService = new PolSourceObservationService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _PolSourceObservationIssueService = new PolSourceObservationIssueService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _MWQMSiteService = new MWQMSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _MWQMSampleService = new MWQMSampleService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _InfrastructureService = new InfrastructureService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _UseOfSiteService = new UseOfSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _ClimateSiteService = new ClimateSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _ClimateDataValueService = new ClimateDataValueService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _TidesAndCurrentsService = new TidesAndCurrentsService(_TaskRunnerBaseService);
            _TideSiteService = new TideSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _TelService = new TelService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _EmailService = new EmailService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            RunSiteInfoList = new List<RunSiteInfo>();
        }
        #endregion Constructors

        #region Functions public
        public void Generate()
        {
            string NotUsed = "";
            string ReportTypeIDText = "";
            int ReportTypeID = 0;
            string TVItemIDText = "";
            string YearText = "";
            int TempInt = 0;

            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            string ServerFilePath = _TVFileService.GetServerFilePath(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);

            Parameters = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters;
            List<string> ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            if (_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.TVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.TVItemID);
            }

            // doing ReportTypeID
            ReportTypeIDText = GetParameters("ReportTypeID", ParamValueList);
            if (string.IsNullOrWhiteSpace(ReportTypeIDText))
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.ReportTypeID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.ReportTypeID);
                return;
            }

            if (!int.TryParse(ReportTypeIDText, out ReportTypeID))
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.ReportTypeID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.ReportTypeID);
                return;
            }

            reportTypeModel = _ReportTypeService.GetReportTypeModelWithReportTypeIDDB(ReportTypeID);
            if (!string.IsNullOrWhiteSpace(reportTypeModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.ReportTypeID, TaskRunnerServiceRes.ReportTypeID, ReportTypeID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.ReportTypeID, TaskRunnerServiceRes.ReportTypeID, ReportTypeID.ToString());
                return;
            }

            if (reportTypeModel.Language == LanguageEnum.fr)
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-CA");
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("fr-CA");
                _TaskRunnerBaseService._BWObj.appTaskModel.Language = LanguageEnum.fr;
            }
            else
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-CA");
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-CA");
                _TaskRunnerBaseService._BWObj.appTaskModel.Language = LanguageEnum.en;
            }

            // doing TVItemID
            TVItemIDText = GetParameters("TVItemID", ParamValueList);

            if (string.IsNullOrWhiteSpace(TVItemIDText))
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.TVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.TVItemID);
                return;
            }

            if (!int.TryParse(TVItemIDText, out TempInt))
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.TVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.TVItemID);
                return;
            }

            TVItemID = TempInt;

            if (reportTypeModel.FileType != FileTypeEnum.KMZ)
            {
                // doing Year
                YearText = GetParameters("Year", ParamValueList);

                if (string.IsNullOrWhiteSpace(YearText))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.Year);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.Year);
                    return;
                }

                if (!int.TryParse(YearText, out TempInt))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.Year);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.Year);
                    return;
                }

                Year = TempInt;
            }
            else
            {
                Year = DateTime.Now.Year;
            }

            DateTime CD = DateTime.Now;
            //string Language = "_" + _TaskRunnerBaseService._BWObj.appTaskModel.Language;

            string DateText = CD.Year.ToString() +
                "_" + (CD.Month > 9 ? CD.Month.ToString() : "0" + CD.Month.ToString()) +
                "_" + (CD.Day > 9 ? CD.Day.ToString() : "0" + CD.Day.ToString()) +
                "_" + (CD.Hour > 9 ? CD.Hour.ToString() : "0" + CD.Hour.ToString()) +
                "_" + (CD.Minute > 9 ? CD.Minute.ToString() : "0" + CD.Minute.ToString());

            if (!RenameStartOfFileName(reportTypeModel, TVItemID, TVItemIDText, ParamValueList))
            {
                return;
            }

            fi = new FileInfo(ServerFilePath + reportTypeModel.StartOfFileName.Replace("{datecreated}", DateText) + ".html");

            if (fi.Exists)
            {
                try
                {
                    fi.Delete();
                    fi = new FileInfo(ServerFilePath + reportTypeModel.StartOfFileName.Replace("{datecreated}", DateText) + ".html");
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDeleteFile_Error_, fi.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDeleteFile_Error_", fi.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return;
                }
            }

            if (reportTypeModel.FileType == FileTypeEnum.DOCX || reportTypeModel.FileType == FileTypeEnum.XLSX)
            {
                if (!GenerateHTML())
                {
                    if (workbook != null)
                    {
                        workbook.Close(false);
                    }
                    if (xlApp != null)
                    {
                        xlApp.Quit();
                    }

                    return;
                }

                if (workbook != null)
                {
                    workbook.Close(false);
                }
                if (xlApp != null)
                {
                    xlApp.Quit();
                }

                DirectoryInfo di = new DirectoryInfo(fi.DirectoryName);

                if (!di.Exists)
                {
                    try
                    {
                        di.Create();
                    }
                    catch (Exception ex)
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateDirectory__, di.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateDirectory__", di.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                        return;
                    }
                }

                StreamWriter sw = fi.CreateText();
                sw.Write(sb.ToString());
                sw.Close();

                if (reportTypeModel.FileType == FileTypeEnum.XLSX)
                {
                    CreateXlsxWithHTML();
                }
                else
                {
                    CreateDocxWithHTML();
                }
            }
            else if (reportTypeModel.FileType == FileTypeEnum.KMZ)
            {
                if (!GenerateKMZ())
                {
                    return;
                }
            }
            else
            {
                NotUsed = string.Format(TaskRunnerServiceRes._NotImplemented, reportTypeModel.FileType.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_NotImplemented", reportTypeModel.FileType.ToString());
                return;
            }
        }

        #endregion Functions public

        #region Functions private
        private bool CreateDocxWithHTML()
        {
            string NotUsed = "";

            Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
            appWord.Visible = false;
            Microsoft.Office.Interop.Word.Document _Document = appWord.Documents.Open(fi.FullName);

            if (_Document.ActiveWindow.View.SplitSpecial == Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneNone)
            {
                _Document.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView;
            }
            else
            {
                _Document.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView;
            }

            CreateDocxWithHTMLDoPROVINCE_NAMETag(appWord, _Document);
            appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
            appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);

            CreateDocxWithHTMLDoPROVINCE_INITIALTag(appWord, _Document);
            appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
            appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);

            CreateDocxWithHTMLDoMUNICIPALITY_NAMETag(appWord, _Document);
            appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
            appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);

            CreateDocxWithHTMLDoSUBSECTOR_NAME_SHORTTag(appWord, _Document);
            appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
            appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);

            CreateDocxWithHTMLDoSUBSECTOR_NAME_LONGTag(appWord, _Document);
            appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
            appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);

            CreateDocxWithHTMLDoSUBSECTOR_NAME_TEXTTag(appWord, _Document);
            appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
            appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);

            CreateDocxWithHTMLDoREPORT_YEARTag(appWord, _Document);
            appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
            appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);

            CreateDocxWithHTMLDoNL_AUTHORSTag(appWord, _Document);
            appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
            appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);

            CreateDocxWithHTMLDoNB_AUTHORSTag(appWord, _Document);
            appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
            appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);

            CreateDocxWithHTMLDoNS_AUTHORSTag(appWord, _Document);
            appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
            appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);

            CreateDocxWithHTMLDoPE_AUTHORSTag(appWord, _Document);
            appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
            appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);

            CreateDocxWithHTMLDoSTATISTICS_LAST_YEARTag(appWord, _Document);
            appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
            appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);

            CreateDocxWithHTMLDoSTATISTICS_PERIODTag(appWord, _Document);
            appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
            appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);

            CreateDocxWithHTMLDoPAGE_BREAKTag(appWord, _Document);
            appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
            appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);

            CreateDocxWithHTMLDoLANDSCAPETag(appWord, _Document);
            appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
            appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);

            CreateDocxWithHTMLDoPORTRAITTag(appWord, _Document);
            appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
            appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);

            CreateDocxWithHTMLDoPAGE_NUMBERING_ROMANTag(appWord, _Document);
            appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
            appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);

            CreateDocxWithHTMLDoPAGE_NUMBERING_NORMALTag(appWord, _Document);
            appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
            appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);

            CreateDocxWithHTMLDoImageTag(appWord, _Document);
            appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
            appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);

            if (Thread.CurrentThread.CurrentCulture.TwoLetterISOLanguageName == "fr")
            {
                CreateDocxWithHTMLDoTableRenamingInFRTag(appWord, _Document);
                appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
                appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
                appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            }

            // renumbering the Tables and Figures
            appWord.Selection.WholeStory();
            appWord.Selection.Fields.Update();

            CreateDocxWithHTMLDoLIST_OF_TABLESTag(appWord, _Document);
            appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
            appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);

            CreateDocxWithHTMLDoLIST_OF_FIGURESTag(appWord, _Document);
            appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
            appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);

            CreateDocxWithHTMLDoTABLE_OF_CONTENTSTag(appWord, _Document);
            appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
            appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);

            string NewDocxFileName = fi.FullName.Replace(".html", ".docx");
            _Document.SaveAs2(NewDocxFileName, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument);
            _Document.Close();
            appWord.Quit();

            fi = new FileInfo(fi.FullName);

            if (fi.Exists)
            {
                try
                {
                    fi.Delete();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDeleteFile_Error_, fi.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDeleteFile_Error_", fi.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return false;
                }
            }

            fi = new FileInfo(NewDocxFileName);

            TVItemModel tvItemModel = _TVItemService.PostAddChildTVItemDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi.Name.Replace(fi.Extension, ""), TVTypeEnum.File);
            if (!string.IsNullOrWhiteSpace(tvItemModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVItem, tvItemModel.Error.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVItem, tvItemModel.Error.ToString());
                return false;
            }

            TVFileModel tvFileModelNew = new TVFileModel();
            tvFileModelNew.TVFileTVItemID = tvItemModel.TVItemID;
            tvFileModelNew.TemplateTVType = 0;
            tvFileModelNew.ReportTypeID = reportTypeModel.ReportTypeID;
            tvFileModelNew.Parameters = Parameters;
            tvFileModelNew.ServerFileName = fi.Name;
            tvFileModelNew.FilePurpose = FilePurposeEnum.ReportGenerated;
            tvFileModelNew.Language = _TaskRunnerBaseService._BWObj.appTaskModel.Language;
            tvFileModelNew.Year = DateTime.Now.Year;
            tvFileModelNew.FileDescription = reportTypeModel.Description;
            tvFileModelNew.FileType = _TVFileService.GetFileType(fi.Extension);
            tvFileModelNew.FileSize_kb = (((int)fi.Length / 1024) == 0 ? 1 : (int)fi.Length / 1024);
            tvFileModelNew.FileInfo = TaskRunnerServiceRes.FileName + "[" + fi.Name + "]\r\n" + TaskRunnerServiceRes.FileType + "[" + fi.Extension + "]\r\n";
            tvFileModelNew.FileCreatedDate_UTC = fi.LastWriteTimeUtc;
            tvFileModelNew.ServerFilePath = (fi.DirectoryName + @"\").Replace(@"C:\", @"E:\");
            tvFileModelNew.LastUpdateDate_UTC = DateTime.UtcNow;
            tvFileModelNew.LastUpdateContactTVItemID = _TaskRunnerBaseService._BWObj.appTaskModel.LastUpdateContactTVItemID;

            TVFileModel tvFileModelRet = _TVFileService.PostAddTVFileDB(tvFileModelNew);
            if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                return false;
            }

            if (!string.IsNullOrWhiteSpace(FileNameExtra))
            {
                DirectoryInfo di = new DirectoryInfo(fi.Directory + @"\");
                List<FileInfo> fiList = di.GetFiles().Where(c => c.Name.Contains(FileNameExtra)).ToList();

                foreach (FileInfo fiToDelete in fiList)
                {
                    try
                    {
                        fiToDelete.Delete();
                    }
                    catch (Exception ex)
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDeleteFile_Error_, fiToDelete.FullName, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDeleteFile_Error_", fiToDelete.FullName, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                        return false;
                    }
                }
            }

            return true;
        }

        private void CreateDocxWithHTMLDoLIST_OF_FIGURESTag(Application appWord, Document document)
        {
            // turn all |||PageBreak||| into word page break
            string SearchMarker = "|||LIST_OF_FIGURES|||";
            bool Found = true;
            while (Found)
            {
                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                if (appWord.Selection.Find.Execute(SearchMarker))
                {
                    appWord.Selection.Font.Name = "Arial";
                    appWord.Selection.Font.Size = 14;
                    appWord.Selection.Text = TaskRunnerServiceRes.ListOfFigures;
                    appWord.Selection.Start = appWord.Selection.End;
                    appWord.Selection.InsertParagraph();

                    document.TablesOfFigures.Add(appWord.Selection.Range, "Figure");

                    appWord.Selection.Find.MatchWildcards = false;
                    appWord.Selection.Find.ClearFormatting();
                    if (appWord.Selection.Find.Execute("^p"))
                    {
                        appWord.Selection.Delete();
                    }
                }
                else
                {
                    Found = false;
                }

                appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
                appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
                appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            }
        }
        private void CreateDocxWithHTMLDoLIST_OF_TABLESTag(Application appWord, Document document)
        {
            // turn all |||PageBreak||| into word page break
            string SearchMarker = "|||LIST_OF_TABLES|||";
            bool Found = true;
            while (Found)
            {
                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                if (appWord.Selection.Find.Execute(SearchMarker))
                {
                    appWord.Selection.Font.Name = "Arial";
                    appWord.Selection.Font.Size = 14;
                    appWord.Selection.Text = TaskRunnerServiceRes.ListOfTables;
                    appWord.Selection.Start = appWord.Selection.End;
                    appWord.Selection.InsertParagraph();

                    document.TablesOfFigures.Add(appWord.Selection.Range, "Table");

                    appWord.Selection.Find.MatchWildcards = false;
                    appWord.Selection.Find.ClearFormatting();
                    if (appWord.Selection.Find.Execute("^p"))
                    {
                        appWord.Selection.Delete();
                    }
                }
                else
                {
                    Found = false;
                }

                appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
                appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
                appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            }
        }
        private void CreateDocxWithHTMLDoImageTag(Application appWord, Document document)
        {
            // importing images/graphics where we find |||Image|||
            bool Found = true;
            while (Found)
            {
                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                appWord.Selection.Find.MatchWildcards = true;
                appWord.Selection.Find.Text = "|||*|||";
                if (appWord.Selection.Find.Execute())
                {
                    if (appWord.Selection.Find.Found)
                    {
                        string textFound = appWord.Selection.Text;

                        // --------------------------------------------------
                        //           |||Image|
                        // --------------------------------------------------
                        if (textFound.StartsWith("|||Image|"))
                        {

                            appWord.Selection.Text = "";

                            string FileName = "";
                            float width = 0;
                            float height = 0;
                            textFound = textFound.Substring("|||Image|".Length).Replace("|||", "");

                            List<string> ImageParamList = textFound.Split("|".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

                            foreach (string s in ImageParamList)
                            {
                                List<string> ParamValue = s.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

                                if (ParamValue.Count() == 2)
                                {
                                    if (ParamValue[0] == "FileName")
                                    {
                                        FileName = ParamValue[1];
                                    }
                                    else if (ParamValue[0] == "width")
                                    {
                                        width = float.Parse(ParamValue[1]);
                                    }
                                    else if (ParamValue[0] == "height")
                                    {
                                        height = float.Parse(ParamValue[1]);
                                    }
                                }
                            }
                            Microsoft.Office.Interop.Word.InlineShape shape = appWord.Selection.InlineShapes.AddPicture(FileName, false, true);
                            shape.Width = width;
                            shape.Height = height;
                        }
                        else if (textFound.StartsWith("|||TableCaption|"))
                        {
                            appWord.Selection.Text = "";

                            textFound = textFound.Substring("|||TableCaption|".Length).Replace("|||", "");

                            if (Thread.CurrentThread.CurrentCulture.TwoLetterISOLanguageName == "fr")
                            {
                                appWord.Selection.Range.InsertCaption("Table", textFound, "", WdCaptionPosition.wdCaptionPositionBelow, 0);
                            }
                            else
                            {
                                appWord.Selection.Range.InsertCaption("Table", textFound, "", WdCaptionPosition.wdCaptionPositionBelow, 0);
                            }
                            appWord.Selection.HomeKey(WdUnits.wdLine);
                            appWord.Selection.EndKey(WdUnits.wdLine, WdMovementType.wdExtend);
                            appWord.Selection.Font.Italic = 0;
                            appWord.Selection.Font.Name = "Arial";
                            appWord.Selection.Font.Size = 11;
                            appWord.Selection.Font.Color = 0;
                            appWord.Selection.MoveRight();
                        }
                        else if (textFound.StartsWith("|||FigureCaption|"))
                        {
                            appWord.Selection.Text = "";

                            textFound = textFound.Substring("|||FigureCaption|".Length).Replace("|||", "");

                            appWord.Selection.Range.InsertCaption("Figure", textFound, "", WdCaptionPosition.wdCaptionPositionBelow, 0);
                            appWord.Selection.HomeKey(WdUnits.wdLine);
                            appWord.Selection.EndKey(WdUnits.wdLine, WdMovementType.wdExtend);
                            appWord.Selection.Font.Italic = 0;
                            appWord.Selection.Font.Name = "Arial";
                            appWord.Selection.Font.Size = 11;
                            appWord.Selection.Font.Color = 0;
                            appWord.Selection.MoveRight();
                        }
                    }
                }
                else
                {
                    Found = false;
                }
            }
        }
        private void CreateDocxWithHTMLDoTableRenamingInFRTag(Application appWord, Document document)
        {
            // Renaming Table to Tableau in french
            string SearchMarker = "Table ";
            bool Found = true;
            while (Found)
            {
                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                if (appWord.Selection.Find.Execute(SearchMarker))
                {
                    appWord.Selection.Text = "Tableau ";
                }
                else
                {
                    Found = false;
                }

                appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
                appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
                appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            }
        }
        private void CreateDocxWithHTMLDoMUNICIPALITY_NAMETag(Application appWord, Document document)
        {
            string MUNICIPALITY_NAME = "ERROR: MUNICIPALITY_NAME";
            TVItemModel tvItemModel = _TVItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModel.Error))
            {
                MUNICIPALITY_NAME = tvItemModel.Error;
            }

            List<TVItemModel> tvItemModelList = _TVItemService.GetParentsTVItemModelList(tvItemModel.TVPath);
            foreach (TVItemModel tvItemModelMuni in tvItemModelList)
            {
                if (tvItemModelMuni.TVType == TVTypeEnum.Municipality)
                {
                    MUNICIPALITY_NAME = tvItemModel.TVText.Substring(0, tvItemModel.TVText.IndexOf(" "));
                    break;
                }
            }

            // turn all |||MUNICIPALITY_NAME||| into word page break
            string SearchMarker = "|||MUNICIPALITY_NAME|||";
            bool Found = true;
            while (Found)
            {
                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                if (appWord.Selection.Find.Execute(SearchMarker))
                {
                    appWord.Selection.Text = MUNICIPALITY_NAME;
                }
                else
                {
                    Found = false;
                }

                appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
                appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
                appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            }
        }
        private void CreateDocxWithHTMLDoPAGE_BREAKTag(Application appWord, Document document)
        {
            // turn all |||PageBreak||| into word page break
            string SearchMarker = "|||PAGE_BREAK|||";
            bool Found = true;
            while (Found)
            {
                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                if (appWord.Selection.Find.Execute(SearchMarker))
                {
                    appWord.Selection.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);

                    appWord.Selection.Find.MatchWildcards = false;
                    appWord.Selection.Find.ClearFormatting();
                    if (appWord.Selection.Find.Execute("^p"))
                    {
                        appWord.Selection.Delete();
                    }
                }
                else
                {
                    Found = false;
                }

                appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
                appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
                appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            }
        }
        private void CreateDocxWithHTMLDoLANDSCAPETag(Application appWord, Document document)
        {
            string SearchMarker = "|||LANDSCAPE|||";
            bool Found = true;
            while (Found)
            {
                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                appWord.Selection.Find.MatchWildcards = true;
                if (appWord.Selection.Find.Execute(SearchMarker))
                {
                    appWord.Selection.Delete();

                    appWord.Selection.Find.MatchWildcards = false;
                    appWord.Selection.Find.ClearFormatting();
                    if (appWord.Selection.Find.Execute("^p"))
                    {
                        appWord.Selection.Delete();
                    }

                    appWord.Selection.InsertBreak(WdBreakType.wdSectionBreakNextPage);
                    appWord.Selection.PageSetup.SectionStart = WdSectionStart.wdSectionNewPage;
                    appWord.Selection.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
                }
                else
                {
                    Found = false;
                }

                appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
                appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
                appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            }
        }
        private void CreateDocxWithHTMLDoPORTRAITTag(Application appWord, Document document)
        {
            string SearchMarker = "|||PORTRAIT|||";
            bool Found = true;
            while (Found)
            {
                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                appWord.Selection.Find.MatchWildcards = true;
                if (appWord.Selection.Find.Execute(SearchMarker))
                {
                    appWord.Selection.Delete();

                    appWord.Selection.Find.MatchWildcards = false;
                    appWord.Selection.Find.ClearFormatting();
                    if (appWord.Selection.Find.Execute("^p"))
                    {
                        appWord.Selection.Delete();
                    }

                    appWord.Selection.InsertBreak(WdBreakType.wdSectionBreakNextPage);
                    appWord.Selection.PageSetup.SectionStart = WdSectionStart.wdSectionNewPage;
                    appWord.Selection.PageSetup.Orientation = WdOrientation.wdOrientPortrait;
                }
                else
                {
                    Found = false;
                }

                appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
                appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
                appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            }
        }
        private void CreateDocxWithHTMLDoPAGE_NUMBERING_ROMANTag(Application appWord, Document document)
        {
            string SearchMarker = "|||PAGE_NUMBERING_ROMAN|||";
            bool Found = true;
            while (Found)
            {
                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                appWord.Selection.Find.MatchWildcards = true;
                if (appWord.Selection.Find.Execute(SearchMarker))
                {
                    appWord.Selection.Delete();

                    appWord.Selection.Find.MatchWildcards = false;
                    appWord.Selection.Find.ClearFormatting();
                    if (appWord.Selection.Find.Execute("^p"))
                    {
                        appWord.Selection.Delete();
                    }

                    appWord.Selection.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakContinuous);
                    for (int i = 1; i <= 2; i++)
                    {
                        HeaderFooter headerFooter = appWord.ActiveDocument.Sections[i].Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];

                        headerFooter.PageNumbers.NumberStyle = Microsoft.Office.Interop.Word.WdPageNumberStyle.wdPageNumberStyleLowercaseRoman;
                        headerFooter.PageNumbers.HeadingLevelForChapter = 0;
                        headerFooter.PageNumbers.IncludeChapterNumber = false;
                        headerFooter.PageNumbers.ChapterPageSeparator = Microsoft.Office.Interop.Word.WdSeparatorType.wdSeparatorHyphen;
                        if (i == 1)
                        {
                            headerFooter.PageNumbers.RestartNumberingAtSection = false;
                            //headerFooter.PageNumbers.StartingNumber = 1;

                            Microsoft.Office.Interop.Word.Range footerRange = appWord.ActiveDocument.Sections[i].Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                            for (int j = 1; j <= footerRange.Fields.Count; j++)
                            {
                                footerRange.Text = "";
                                footerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            }
                        }
                        else
                        {
                            headerFooter.PageNumbers.RestartNumberingAtSection = true;
                            headerFooter.PageNumbers.StartingNumber = 1;

                            Microsoft.Office.Interop.Word.Range footerRange = appWord.ActiveDocument.Sections[i].Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                            footerRange.Fields.Add(footerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                            footerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        }
                    }
                }
                else
                {
                    Found = false;
                }

                appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
                appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
                appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            }
        }
        private void CreateDocxWithHTMLDoPAGE_NUMBERING_NORMALTag(Application appWord, Document document)
        {
            string SearchMarker = "|||PAGE_NUMBERING_NORMAL|||";
            bool Found = true;
            while (Found)
            {
                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                appWord.Selection.Find.MatchWildcards = true;
                if (appWord.Selection.Find.Execute(SearchMarker))
                {
                    appWord.Selection.Delete();

                    appWord.Selection.Find.MatchWildcards = false;
                    appWord.Selection.Find.ClearFormatting();
                    if (appWord.Selection.Find.Execute("^p"))
                    {
                        appWord.Selection.Delete();
                    }

                    appWord.Selection.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakContinuous);

                    for (int i = 3; i <= appWord.ActiveDocument.Sections.Count; i++)
                    {
                        HeaderFooter headerFooter = appWord.ActiveDocument.Sections[i].Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];

                        headerFooter.PageNumbers.NumberStyle = Microsoft.Office.Interop.Word.WdPageNumberStyle.wdPageNumberStyleArabic;
                        headerFooter.PageNumbers.HeadingLevelForChapter = 0;
                        headerFooter.PageNumbers.IncludeChapterNumber = false;
                        headerFooter.PageNumbers.ChapterPageSeparator = Microsoft.Office.Interop.Word.WdSeparatorType.wdSeparatorHyphen;
                        if (i == 3)
                        {
                            headerFooter.PageNumbers.RestartNumberingAtSection = true;
                            headerFooter.PageNumbers.StartingNumber = 1;
                        }
                        else
                        {
                            headerFooter.PageNumbers.RestartNumberingAtSection = false;
                            //headerFooter.PageNumbers.StartingNumber = 1;
                        }

                        Microsoft.Office.Interop.Word.Range footerRange = appWord.ActiveDocument.Sections[i].Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        footerRange.Fields.Add(footerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                        footerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                }
                else
                {
                    Found = false;
                }

                appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
                appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
                appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            }
        }
        private void CreateDocxWithHTMLDoPROVINCE_NAMETag(Application appWord, Document document)
        {
            string PROVINCE_NAME = "ERROR: PROVINCE_NAME";
            TVItemModel tvItemModel = _TVItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModel.Error))
            {
                PROVINCE_NAME = tvItemModel.Error;
            }

            List<TVItemModel> tvItemModelList = _TVItemService.GetParentsTVItemModelList(tvItemModel.TVPath);
            foreach (TVItemModel tvItemModelProv in tvItemModelList)
            {
                if (tvItemModelProv.TVType == TVTypeEnum.Province)
                {
                    PROVINCE_NAME = tvItemModelProv.TVText;
                    break;
                }
            }

            // turn all |||PROVINCE_NAME||| into word page break
            string SearchMarker = "|||PROVINCE_NAME|||";
            bool Found = true;
            while (Found)
            {
                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                if (appWord.Selection.Find.Execute(SearchMarker))
                {
                    appWord.Selection.Text = PROVINCE_NAME;
                }
                else
                {
                    Found = false;
                }

                appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
                appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
                appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            }
        }
        private void CreateDocxWithHTMLDoPROVINCE_INITIALTag(Application appWord, Document document)
        {
            string PROVINCE_INITIAL = "ERROR: PROVINCE_INITIAL";
            TVItemModel tvItemModel = _TVItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModel.Error))
            {
                PROVINCE_INITIAL = tvItemModel.Error;
            }

            List<TVItemModel> tvItemModelList = _TVItemService.GetParentsTVItemModelList(tvItemModel.TVPath);
            foreach (TVItemModel tvItemModelProv in tvItemModelList)
            {
                if (tvItemModelProv.TVType == TVTypeEnum.Province)
                {
                    switch (tvItemModelProv.TVText)
                    {
                        case "British Columbia":
                            {
                                PROVINCE_INITIAL = "BC";
                            }
                            break;
                        case "Colombie-Britannique":
                            {
                                PROVINCE_INITIAL = "CB";
                            }
                            break;
                        case "New Brunswick":
                            {
                                PROVINCE_INITIAL = "NB";
                            }
                            break;
                        case "Nouveau-Brunswick":
                            {
                                PROVINCE_INITIAL = "NB";
                            }
                            break;
                        case "Nova Scotia":
                            {
                                PROVINCE_INITIAL = "NS";
                            }
                            break;
                        case "Nouvelle-Écosse":
                            {
                                PROVINCE_INITIAL = "NE";
                            }
                            break;
                        case "Newfoundland and Labrador":
                            {
                                PROVINCE_INITIAL = "NL";
                            }
                            break;
                        case "Terre-Neuve-et-Labrador":
                            {
                                PROVINCE_INITIAL = "TL";
                            }
                            break;
                        case "Québec":
                            {
                                PROVINCE_INITIAL = "QC";
                            }
                            break;
                        case "Quebec":
                            {
                                PROVINCE_INITIAL = "QC";
                            }
                            break;
                        case "Prince Edward Island":
                            {
                                PROVINCE_INITIAL = "PEI";
                            }
                            break;
                        case "Île-du-Prince-Édouard":
                            {
                                PROVINCE_INITIAL = "IPE";
                            }
                            break;
                        default:
                            break;
                    }
                    break;
                }
            }

            // turn all |||PROVINCE_INITIAL||| into word page break
            string SearchMarker = "|||PROVINCE_INITIAL|||";
            bool Found = true;
            while (Found)
            {
                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                if (appWord.Selection.Find.Execute(SearchMarker))
                {
                    appWord.Selection.Text = PROVINCE_INITIAL;
                }
                else
                {
                    Found = false;
                }

                appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
                appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
                appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            }
        }
        private void CreateDocxWithHTMLDoSUBSECTOR_NAME_SHORTTag(Application appWord, Document document)
        {
            string SUBSECTOR_NAME_SHORT = "ERROR: SUBSECTOR_NAME_SHORT";
            TVItemModel tvItemModel = _TVItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModel.Error))
            {
                SUBSECTOR_NAME_SHORT = tvItemModel.Error;
            }

            List<TVItemModel> tvItemModelList = _TVItemService.GetParentsTVItemModelList(tvItemModel.TVPath);
            foreach (TVItemModel tvItemModelSS in tvItemModelList)
            {
                if (tvItemModelSS.TVType == TVTypeEnum.Subsector)
                {
                    SUBSECTOR_NAME_SHORT = tvItemModelSS.TVText.Substring(0, tvItemModelSS.TVText.IndexOf(" "));
                    break;
                }
            }

            // turn all |||SUBSECTOR_NAME_SHORT||| into word page break
            string SearchMarker = "|||SUBSECTOR_NAME_SHORT|||";
            bool Found = true;
            while (Found)
            {
                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                if (appWord.Selection.Find.Execute(SearchMarker))
                {
                    appWord.Selection.Text = SUBSECTOR_NAME_SHORT;
                }
                else
                {
                    Found = false;
                }

                appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
                appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
                appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            }
        }
        private void CreateDocxWithHTMLDoSUBSECTOR_NAME_LONGTag(Application appWord, Document document)
        {
            string SUBSECTOR_NAME_LONG = "ERROR: SUBSECTOR_NAME_LONG";
            TVItemModel tvItemModel = _TVItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModel.Error))
            {
                SUBSECTOR_NAME_LONG = tvItemModel.Error;
            }

            List<TVItemModel> tvItemModelList = _TVItemService.GetParentsTVItemModelList(tvItemModel.TVPath);
            foreach (TVItemModel tvItemModelSS in tvItemModelList)
            {
                if (tvItemModelSS.TVType == TVTypeEnum.Subsector)
                {
                    SUBSECTOR_NAME_LONG = tvItemModelSS.TVText;
                    break;
                }
            }

            // turn all |||SUBSECTOR_NAME_LONG||| into word page break
            string SearchMarker = "|||SUBSECTOR_NAME_LONG|||";
            bool Found = true;
            while (Found)
            {
                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                if (appWord.Selection.Find.Execute(SearchMarker))
                {
                    appWord.Selection.Text = SUBSECTOR_NAME_LONG;
                }
                else
                {
                    Found = false;
                }

                appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
                appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
                appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            }
        }
        private void CreateDocxWithHTMLDoSUBSECTOR_NAME_TEXTTag(Application appWord, Document document)
        {
            string SUBSECTOR_NAME_TEXT = "ERROR: SUBSECTOR_NAME_TEXT";
            TVItemModel tvItemModel = _TVItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModel.Error))
            {
                SUBSECTOR_NAME_TEXT = tvItemModel.Error;
            }

            List<TVItemModel> tvItemModelList = _TVItemService.GetParentsTVItemModelList(tvItemModel.TVPath);
            foreach (TVItemModel tvItemModelSS in tvItemModelList)
            {
                if (tvItemModelSS.TVType == TVTypeEnum.Subsector)
                {
                    SUBSECTOR_NAME_TEXT = tvItemModelSS.TVText.Substring(tvItemModelSS.TVText.IndexOf(" ") + 1).Trim();
                    if (SUBSECTOR_NAME_TEXT.Substring(0, 1) == "(")
                    {
                        SUBSECTOR_NAME_TEXT = SUBSECTOR_NAME_TEXT.Substring(1);
                    }
                    if (SUBSECTOR_NAME_TEXT.Substring(SUBSECTOR_NAME_TEXT.Length - 1) == ")")
                    {
                        SUBSECTOR_NAME_TEXT = SUBSECTOR_NAME_TEXT.Substring(0, SUBSECTOR_NAME_TEXT.Length - 1);
                    }
                    break;
                }
            }

            string SearchMarker = "|||SUBSECTOR_NAME_TEXT|||";
            bool Found = true;
            while (Found)
            {
                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                if (appWord.Selection.Find.Execute(SearchMarker))
                {
                    appWord.Selection.Text = SUBSECTOR_NAME_TEXT;
                }
                else
                {
                    Found = false;
                }

                appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
                appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
                appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            }
        }
        private void CreateDocxWithHTMLDoREPORT_YEARTag(Application appWord, Document document)
        {
            string REPORT_YEAR = Year.ToString();

            string SearchMarker = "|||REPORT_YEAR|||";
            bool Found = true;
            while (Found)
            {
                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                if (appWord.Selection.Find.Execute(SearchMarker))
                {
                    appWord.Selection.Text = REPORT_YEAR;
                }
                else
                {
                    Found = false;
                }

                appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
                appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
                appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            }
        }
        private void CreateDocxWithHTMLDoNL_AUTHORSTag(Application appWord, Document document)
        {
            string SearchMarker = "|||NL_AUTHORS|||";
            bool Found = true;
            while (Found)
            {
                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                if (appWord.Selection.Find.Execute(SearchMarker))
                {
                    appWord.Selection.Text = $"D. Curtis, G. Perchard, M. Glavine";
                }
                else
                {
                    Found = false;
                }

                appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
                appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
                appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            }
        }
        private void CreateDocxWithHTMLDoNB_AUTHORSTag(Application appWord, Document document)
        {
            string SearchMarker = "|||NB_AUTHORS|||";
            bool Found = true;
            while (Found)
            {
                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                if (appWord.Selection.Find.Execute(SearchMarker))
                {
                    appWord.Selection.Text = $"B. Richard, P. Godin, K. Martell, J. Pomeroy, C. LeBlanc, J.A. Richard";
                }
                else
                {
                    Found = false;
                }

                appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
                appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
                appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            }
        }
        private void CreateDocxWithHTMLDoNS_AUTHORSTag(Application appWord, Document document)
        {
            string SearchMarker = "|||NS_AUTHORS|||";
            bool Found = true;
            while (Found)
            {
                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                if (appWord.Selection.Find.Execute(SearchMarker))
                {
                    appWord.Selection.Text = $"D. MacArthur, L. Pothier, R. Alexanders, P. Densmore";
                }
                else
                {
                    Found = false;
                }

                appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
                appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
                appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            }
        }
        private void CreateDocxWithHTMLDoPE_AUTHORSTag(Application appWord, Document document)
        {
            string SearchMarker = "|||PE_AUTHORS|||";
            bool Found = true;
            while (Found)
            {
                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                if (appWord.Selection.Find.Execute(SearchMarker))
                {
                    appWord.Selection.Text = $"D. MacArthur, L. Pothier, R. Alexanders";
                }
                else
                {
                    Found = false;
                }

                appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
                appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
                appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            }
        }
        private void CreateDocxWithHTMLDoSTATISTICS_PERIODTag(Application appWord, Document document)
        {
            string SearchMarker = "|||STATISTICS_PERIOD|||";
            bool Found = true;
            while (Found)
            {
                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                if (appWord.Selection.Find.Execute(SearchMarker))
                {
                    appWord.Selection.Text = $"{StatEndYear} - {StatStartYear}";
                }
                else
                {
                    Found = false;
                }

                appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
                appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
                appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            }
        }
        private void CreateDocxWithHTMLDoSTATISTICS_LAST_YEARTag(Application appWord, Document document)
        {
            string SearchMarker = "|||STATISTICS_LAST_YEAR|||";
            bool Found = true;
            while (Found)
            {
                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                if (appWord.Selection.Find.Execute(SearchMarker))
                {
                    appWord.Selection.Text = $"{StatEndYear}";
                }
                else
                {
                    Found = false;
                }

                appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
                appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
                appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            }
        }
        private void CreateDocxWithHTMLDoTABLE_OF_CONTENTSTag(Application appWord, Document document)
        {
            // turn all |||PageBreak||| into word page break
            string SearchMarker = "|||TABLE_OF_CONTENTS|||";
            bool Found = true;
            while (Found)
            {
                appWord.Selection.Find.ClearFormatting();
                appWord.Selection.Find.Replacement.ClearFormatting();
                if (appWord.Selection.Find.Execute(SearchMarker))
                {
                    appWord.Selection.Font.Size = 14;
                    appWord.Selection.Text = TaskRunnerServiceRes.TableOfContents;
                    appWord.Selection.Start = appWord.Selection.End;
                    appWord.Selection.InsertParagraph();

                    document.TablesOfContents.Add(appWord.Selection.Range, true, 1, 6, true, "", true, true, null, true, true, true);

                    appWord.Selection.Find.MatchWildcards = false;
                    appWord.Selection.Find.ClearFormatting();
                    if (appWord.Selection.Find.Execute("^p"))
                    {
                        appWord.Selection.Delete();
                    }
                }
                else
                {
                    Found = false;
                }

                appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
                appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
                appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            }
        }

        private bool CreateXlsxWithHTML()
        {
            string NotUsed = "";

            Microsoft.Office.Interop.Excel._Application appExcel = new Microsoft.Office.Interop.Excel.Application();
            appExcel.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = appExcel.Workbooks.Open(fi.FullName);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            string NewXlsxFileName = fi.FullName.Replace(".html", ".xlsx");
            appExcel.ActiveWorkbook.SaveAs(NewXlsxFileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
            xlWorkbook.Close();
            appExcel.Quit();

            fi = new FileInfo(fi.FullName);

            if (fi.Exists)
            {
                try
                {
                    fi.Delete();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDeleteFile_Error_, fi.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDeleteFile_Error_", fi.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return false;
                }
            }

            fi = new FileInfo(NewXlsxFileName);

            TVItemModel tvItemModel = _TVItemService.PostAddChildTVItemDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi.Name.Replace(fi.Extension, ""), TVTypeEnum.File);
            if (!string.IsNullOrWhiteSpace(tvItemModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVItem, tvItemModel.Error.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVItem, tvItemModel.Error.ToString());
                return false;
            }

            TVFileModel tvFileModelNew = new TVFileModel();
            tvFileModelNew.TVFileTVItemID = tvItemModel.TVItemID;
            tvFileModelNew.TemplateTVType = 0;
            tvFileModelNew.ReportTypeID = reportTypeModel.ReportTypeID;
            tvFileModelNew.Parameters = Parameters;
            tvFileModelNew.ServerFileName = fi.Name;
            tvFileModelNew.FilePurpose = FilePurposeEnum.ReportGenerated;
            tvFileModelNew.Language = _TaskRunnerBaseService._BWObj.appTaskModel.Language;
            tvFileModelNew.Year = DateTime.Now.Year;
            tvFileModelNew.FileDescription = reportTypeModel.Description;
            tvFileModelNew.FileType = _TVFileService.GetFileType(fi.Extension);
            tvFileModelNew.FileSize_kb = (((int)fi.Length / 1024) == 0 ? 1 : (int)fi.Length / 1024);
            tvFileModelNew.FileInfo = TaskRunnerServiceRes.FileName + "[" + fi.Name + "]\r\n" + TaskRunnerServiceRes.FileType + "[" + fi.Extension + "]\r\n";
            tvFileModelNew.FileCreatedDate_UTC = fi.LastWriteTimeUtc;
            tvFileModelNew.ServerFilePath = (fi.DirectoryName + @"\").Replace(@"C:\", @"E:\");
            tvFileModelNew.LastUpdateDate_UTC = DateTime.UtcNow;
            tvFileModelNew.LastUpdateContactTVItemID = _TaskRunnerBaseService._BWObj.appTaskModel.LastUpdateContactTVItemID;

            TVFileModel tvFileModelRet = _TVFileService.PostAddTVFileDB(tvFileModelNew);
            if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                return false;
            }

            return true;
        }
        private string GetParameters(string Parameter, List<string> ParamValueList)
        {
            foreach (string pv in ParamValueList)
            {
                List<string> ParamValue = pv.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();
                if (Parameter == ParamValue[0])
                {
                    return ParamValue[1];
                }
            }

            return "";
        }
        private List<string> GetParametersArray(string Parameter, List<string> ParamValueList)
        {
            foreach (string pv in ParamValueList)
            {
                string Param = pv.Substring(0, pv.IndexOf(","));
                if (Param == Parameter)
                {
                    string Value = pv.Substring(pv.IndexOf(",") + 1);

                    List<string> valueList = Value.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

                    return valueList;
                }
            }

            return new List<string>();
        }
        private bool RenameStartOfFileName(ReportTypeModel reportTypeModel, int TVItemID, string TVItemIDText, List<string> ParamValueList)
        {
            string NotUsed = "";
            string subsector = "";
            string mikescenarioname = "";
            string year = "";
            string contourvalues = "";
            string lat = "";
            string lng = "";
            string municipality = "";

            TVItemModel tvItemModel = _TVItemService.GetTVItemModelWithTVItemIDDB(TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                return false;
            }

            if (reportTypeModel.TVType == TVTypeEnum.Subsector)
            {
                subsector = tvItemModel.TVText;
                int pos = subsector.IndexOf(" ");
                if (pos > 0)
                {
                    subsector = subsector.Substring(0, subsector.IndexOf(" "));
                }

                if (!string.IsNullOrWhiteSpace(subsector))
                {
                    reportTypeModel.StartOfFileName = reportTypeModel.StartOfFileName.Replace("{subsector}", subsector);
                }
                else
                {
                    reportTypeModel.StartOfFileName = reportTypeModel.StartOfFileName.Replace("{subsector}", "ERROR subsector");
                }

            }
            else if (reportTypeModel.TVType == TVTypeEnum.MikeScenario)
            {
                mikescenarioname = tvItemModel.TVText;

                if (!string.IsNullOrWhiteSpace(mikescenarioname))
                {
                    reportTypeModel.StartOfFileName = reportTypeModel.StartOfFileName.Replace("{mikescenarioname}", mikescenarioname);
                }
                else
                {
                    reportTypeModel.StartOfFileName = reportTypeModel.StartOfFileName.Replace("{mikescenarioname}", "ERROR mikescenarioname");
                }
            }
            else if (reportTypeModel.TVType == TVTypeEnum.Municipality)
            {
                municipality = tvItemModel.TVText;

                if (!string.IsNullOrWhiteSpace(municipality))
                {
                    reportTypeModel.StartOfFileName = reportTypeModel.StartOfFileName.Replace("{municipality}", municipality);
                }
                else
                {
                    reportTypeModel.StartOfFileName = reportTypeModel.StartOfFileName.Replace("{municipality}", "ERROR municipality");
                }
            }
            else
            {
                reportTypeModel.StartOfFileName = "Not_Implemented_" + reportTypeModel.StartOfFileName;
            }

            year = GetParameters("Year", ParamValueList);
            if (!string.IsNullOrWhiteSpace(year))
            {
                reportTypeModel.StartOfFileName = reportTypeModel.StartOfFileName.Replace("{year}", year);
            }
            else
            {
                reportTypeModel.StartOfFileName = reportTypeModel.StartOfFileName.Replace("{year}", "ERROR year");
            }

            contourvalues = GetParameters("ContourValues", ParamValueList);
            contourvalues = contourvalues.Trim().Replace(" ", "_");

            if (!string.IsNullOrWhiteSpace(contourvalues))
            {
                reportTypeModel.StartOfFileName = reportTypeModel.StartOfFileName.Replace("{contourvalues}", contourvalues);
            }
            else
            {
                reportTypeModel.StartOfFileName = reportTypeModel.StartOfFileName.Replace("{contourvalues}", "ERROR contourvalues");
            }

            lat = GetParameters("Lat", ParamValueList);
            lat = lat.Trim().Replace(",", "_");
            lat = lat.Trim().Replace(".", "_");

            if (!string.IsNullOrWhiteSpace(lat))
            {
                reportTypeModel.StartOfFileName = reportTypeModel.StartOfFileName.Replace("{lat}", lat);
            }
            else
            {
                reportTypeModel.StartOfFileName = reportTypeModel.StartOfFileName.Replace("{lat}", "ERROR lat");
            }

            lng = GetParameters("Lng", ParamValueList);
            lng = lng.Trim().Replace(",", "_");
            lng = lng.Trim().Replace(".", "_");

            if (!string.IsNullOrWhiteSpace(lng))
            {
                reportTypeModel.StartOfFileName = reportTypeModel.StartOfFileName.Replace("{lng}", lng);
            }
            else
            {
                reportTypeModel.StartOfFileName = reportTypeModel.StartOfFileName.Replace("{lng}", "ERROR lng");
            }

            return true;
        }
        private string GetTideInitial(TideTextEnum? tideText)
        {
            if (tideText == null)
            {
                return "--";
            }

            switch (tideText)
            {
                case TideTextEnum.LowTide:
                    return "LT";
                case TideTextEnum.LowTideFalling:
                    return "LF";
                case TideTextEnum.LowTideRising:
                    return "LR";
                case TideTextEnum.MidTide:
                    return "MT";
                case TideTextEnum.MidTideFalling:
                    return "MF";
                case TideTextEnum.MidTideRising:
                    return "MR";
                case TideTextEnum.HighTide:
                    return "HT";
                case TideTextEnum.HighTideFalling:
                    return "HF";
                case TideTextEnum.HighTideRising:
                    return "HR";
                default:
                    return "--";
            }
        }
        #endregion Functions private
    }

    public class RunSiteInfo
    {
        public RunSiteInfo()
        {

        }

        public int RunTVItemID { get; set; }
        public int SiteTVItemID { get; set; }
    }
}
