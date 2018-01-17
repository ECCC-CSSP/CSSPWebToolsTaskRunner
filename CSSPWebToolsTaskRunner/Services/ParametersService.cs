using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using CSSPWebToolsTaskRunner.Services;
using CSSPWebToolsTaskRunner.Services.Resources;
using CSSPWebToolsTaskRunner;
using System.Transactions;
using System.Text;
using CSSPWebToolsDBDLL.Models;
using CSSPWebToolsDBDLL.Services;
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;
using CSSPEnumsDLL.Services;
using Microsoft.Office.Interop.Word;

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
        private TVFileService _TVFileService { get; set; }
        private TVItemService _TVItemService { get; set; }
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
        private BaseEnumService _BaseEnumService { get; set; }
        public FileInfo fi { get; set; }
        public ReportTypeModel reportTypeModel { get; set; }
        public int TVItemID { get; set; }
        public int Year { get; set; }
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
            _TVFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _TVItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
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
            _BaseEnumService = new BaseEnumService(_TaskRunnerBaseService._BWObj.appTaskModel.Language);

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
            string Language = "_" + _TaskRunnerBaseService._BWObj.appTaskModel.Language;

            string DateText = "_" + CD.Year.ToString() +
                "_" + (CD.Month > 9 ? CD.Month.ToString() : "0" + CD.Month.ToString()) +
                "_" + (CD.Day > 9 ? CD.Day.ToString() : "0" + CD.Day.ToString()) +
                "_" + (CD.Hour > 9 ? CD.Hour.ToString() : "0" + CD.Hour.ToString()) +
                "_" + (CD.Minute > 9 ? CD.Minute.ToString() : "0" + CD.Minute.ToString());

            if (!RenameStartOfFileName(reportTypeModel, TVItemID, TVItemIDText, ParamValueList))
            {
                return;
            }

            fi = new FileInfo(ServerFilePath + reportTypeModel.StartOfFileName + DateText + Language + ".html");

            if (fi.Exists)
            {
                try
                {
                    fi.Delete();
                    fi = new FileInfo(ServerFilePath + reportTypeModel.StartOfFileName + DateText + Language + ".html");
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

            CreateDocxWithHTMLDoPAGE_BREAKTag(appWord, _Document);
            appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
            appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);

            CreateDocxWithHTMLDoImageTag(appWord, _Document);
            appWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);
            appWord.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);
            appWord.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1);

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

                            appWord.Selection.Range.InsertCaption("Table", textFound);
                        }
                        else if (textFound.StartsWith("|||FigureCaption|"))
                        {

                            appWord.Selection.Text = "";

                            textFound = textFound.Substring("|||FigureCaption|".Length).Replace("|||", "");

                            appWord.Selection.Range.InsertCaption("Figure", textFound);
                        }
                    }
                }
                else
                {
                    Found = false;
                }
            }
        }

        private bool CreateXlsxWithHTML()
        {
            string NotUsed = "";

            Microsoft.Office.Interop.Excel._Application appExcel = new Microsoft.Office.Interop.Excel.Application();
            appExcel.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = appExcel.Workbooks.Open(fi.FullName);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            string NewDocxFileName = fi.FullName.Replace(".html", ".xlsx");
            appExcel.ActiveWorkbook.SaveAs(NewDocxFileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
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

            return true;
        }
        private string GetParameters(string Parameter, List<string> ParamValueList)
        {
            foreach (string pv in ParamValueList)
            {
                List<string> ParamValue = pv.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();
                if (ParamValue.Count != 2)
                {
                    return "";
                }
                if (Parameter == ParamValue[0])
                {
                    return ParamValue[1];
                }
            }

            return "";
        }
        private bool RenameStartOfFileName(ReportTypeModel reportTypeModel, int TVItemID, string TVItemIDText, List<string> ParamValueList)
        {
            string NotUsed = "";

            switch (reportTypeModel.TVType)
            {
                case TVTypeEnum.Subsector:
                    {
                        TVItemModel tvItemModelSS = _TVItemService.GetTVItemModelWithTVItemIDDB(TVItemID);
                        if (!string.IsNullOrWhiteSpace(tvItemModelSS.Error))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                            return false;
                        }
                        string Subsector = tvItemModelSS.TVText;
                        int pos = Subsector.IndexOf(" ");
                        if (pos > 0)
                        {
                            Subsector = Subsector.Substring(0, Subsector.IndexOf(" "));
                        }
                        reportTypeModel.StartOfFileName = reportTypeModel.StartOfFileName.Replace("{subsector}", Subsector);

                        string YearText = GetParameters("Year", ParamValueList);
                        if (!string.IsNullOrWhiteSpace(YearText))
                        {
                            int Year = 0;

                            if (string.IsNullOrWhiteSpace(TVItemIDText))
                            {
                                reportTypeModel.StartOfFileName = reportTypeModel.StartOfFileName.Replace("{year}", "");
                            }
                            else
                            {
                                if (!int.TryParse(YearText, out Year))
                                {
                                    NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.Year);
                                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.Year);
                                    return false;
                                }
                                reportTypeModel.StartOfFileName = reportTypeModel.StartOfFileName.Replace("{year}", Year.ToString());
                            }
                        }
                    }
                    break;
                case TVTypeEnum.MikeScenario:
                    {
                        string TVItemIDStr = "";
                        string ContourValues = "";

                        TVItemIDStr = GetParameters("TVItemID", ParamValueList);
                        if (string.IsNullOrWhiteSpace(TVItemIDStr))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.TVItemID);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.TVItemID);
                        }

                        int.TryParse(TVItemIDStr, out TVItemID);
                        if (TVItemID == 0)
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.TVItemID);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.TVItemID);
                        }

                        ContourValues = GetParameters("ContourValues", ParamValueList);
                        ContourValues = ContourValues.Trim().Replace(" ", "_");

                        TVItemModel tvItemModelMikeScenario = _TVItemService.GetTVItemModelWithTVItemIDDB(TVItemID);
                        if (!string.IsNullOrWhiteSpace(tvItemModelMikeScenario.Error))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                            return false;
                        }
                        string MikeScenarioName = tvItemModelMikeScenario.TVText;
                        int pos = MikeScenarioName.IndexOf(" ");
                        if (pos > 0)
                        {
                            MikeScenarioName = MikeScenarioName.Trim();
                        }
                        reportTypeModel.StartOfFileName = reportTypeModel.StartOfFileName.Replace("{MikeScenarioName}", MikeScenarioName);


                        // it is possible that ContourValues parameter does not exist or is empty
                        ContourValues = GetParameters("ContourValues", ParamValueList);
                        ContourValues = ContourValues.Trim().Replace(" ", "_");

                        if (!string.IsNullOrWhiteSpace(ContourValues))
                        {

                            reportTypeModel.StartOfFileName = reportTypeModel.StartOfFileName.Replace("{ContourValues}", ContourValues);
                        }
                    }
                    break;
                default:
                    break;
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
