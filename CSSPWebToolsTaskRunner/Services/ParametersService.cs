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

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        #region Variables
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        #endregion Variables

        #region Constructors
        public ParametersService(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
        }
        #endregion Constructors

        #region Functions public
        public void Generate()
        {
            string NotUsed = "";
            string ReportTypeIDText = "";
            int ReportTypeID = 0;
            string TVItemIDText = "";
            int TVItemID = 0;
            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            ReportTypeService reportTypeService = new ReportTypeService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);

            string Parameters = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters;
            List<string> ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            if (_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.TVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.TVItemID);
            }

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

            ReportTypeModel reportTypeModel = reportTypeService.GetReportTypeModelWithReportTypeIDDB(ReportTypeID);
            if (!string.IsNullOrWhiteSpace(reportTypeModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.ReportTypeID, TaskRunnerServiceRes.ReportTypeID, ReportTypeID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.ReportTypeID, TaskRunnerServiceRes.ReportTypeID, ReportTypeID.ToString());
                return;
            }

            TVItemIDText = GetParameters("TVItemID", ParamValueList);

            if (string.IsNullOrWhiteSpace(TVItemIDText))
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.TVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.TVItemID);
                return;
            }

            if (!int.TryParse(TVItemIDText, out TVItemID))
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.TVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.TVItemID);
                return;
            }

            switch (reportTypeModel.TVType)
            {
                case TVTypeEnum.Subsector:
                    {
                        TVItemModel tvItemModelSS = tvItemService.GetTVItemModelWithTVItemIDDB(TVItemID);
                        if (!string.IsNullOrWhiteSpace(tvItemModelSS.Error))
                        {
                            NotUsed = tvItemModelSS.Error;
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(tvItemModelSS.Error);
                            return;
                        }
                        string Subsector = tvItemModelSS.TVText;
                        int pos = Subsector.IndexOf(" ");
                        if (pos > 0)
                        {
                            Subsector = Subsector.Substring(0, Subsector.IndexOf(" "));
                        }
                        reportTypeModel.StartOfFileName = reportTypeModel.StartOfFileName.Replace("{subsector}", Subsector);


                        string YearText = GetParameters("Year", ParamValueList);
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
                                return;
                            }
                            reportTypeModel.StartOfFileName = reportTypeModel.StartOfFileName.Replace("{year}", Year.ToString());
                        }
                    }
                    break;
                default:
                    break;
            }

            if (reportTypeModel.FileType == FileTypeEnum.DOCX || reportTypeModel.FileType == FileTypeEnum.XLSX)
            {
                StringBuilder sbHTML = new StringBuilder();

                DateTime CD = DateTime.Now;
                string Language = "_" + _TaskRunnerBaseService._BWObj.appTaskModel.Language;

                string DateText = "_" + CD.Year.ToString() +
                    "_" + (CD.Month > 9 ? CD.Month.ToString() : "0" + CD.Month.ToString()) +
                    "_" + (CD.Day > 9 ? CD.Day.ToString() : "0" + CD.Day.ToString()) +
                    "_" + (CD.Hour > 9 ? CD.Hour.ToString() : "0" + CD.Hour.ToString()) +
                    "_" + (CD.Minute > 9 ? CD.Minute.ToString() : "0" + CD.Minute.ToString());

                FileInfo fi = new FileInfo(ServerFilePath + reportTypeModel.StartOfFileName + DateText + Language + ".html");

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

                if (!GenerateHTML(fi, sbHTML, Parameters, reportTypeModel))
                {
                    return;
                }

                StreamWriter sw = fi.CreateText();
                sw.Write(sbHTML.ToString());
                sw.Close();

                if (reportTypeModel.FileType == FileTypeEnum.XLSX)
                {
                    CreateXlsxWithHTML(fi, Parameters, reportTypeModel);
                }
                else
                {
                    CreateDocxWithHTML(fi, Parameters, reportTypeModel);
                }
            }
            else if (reportTypeModel.FileType == FileTypeEnum.KMZ)
            {
                if (!GenerateKMZ(ServerFilePath, Parameters, reportTypeModel))
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
        private bool CreateDocxWithHTML(FileInfo fi, string Parameters, ReportTypeModel reportTypeModel)
        {
            string NotUsed = "";

            Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
            appWord.Visible = false;
            Microsoft.Office.Interop.Word.Document _Document = appWord.Documents.Open(fi.FullName);
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

            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemModel tvItemModel = tvItemService.PostAddChildTVItemDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi.Name.Replace(fi.Extension, ""), TVTypeEnum.File);
            if (!string.IsNullOrWhiteSpace(tvItemModel.Error))
            {
                NotUsed = tvItemModel.Error;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(tvItemModel.Error);
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
            tvFileModelNew.FileDescription = reportTypeModel.Description;
            tvFileModelNew.FileType = tvFileService.GetFileType(fi.Extension);
            tvFileModelNew.FileSize_kb = (((int)fi.Length / 1024) == 0 ? 1 : (int)fi.Length / 1024);
            tvFileModelNew.FileInfo = TaskRunnerServiceRes.FileName + "[" + fi.Name + "]\r\n" + TaskRunnerServiceRes.FileType + "[" + fi.Extension + "]\r\n";
            tvFileModelNew.FileCreatedDate_UTC = fi.LastWriteTimeUtc;
            tvFileModelNew.ServerFilePath = (fi.DirectoryName + @"\").Replace(@"C:\", @"E:\");
            tvFileModelNew.LastUpdateDate_UTC = DateTime.UtcNow;
            tvFileModelNew.LastUpdateContactTVItemID = _TaskRunnerBaseService._BWObj.appTaskModel.LastUpdateContactTVItemID;

            TVFileModel tvFileModelRet = tvFileService.PostAddTVFileDB(tvFileModelNew);
            if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                return false;
            }

            return true;
        }
        private bool CreateXlsxWithHTML(FileInfo fi, string Parameters, ReportTypeModel reportTypeModel)
        {
            string NotUsed = "";

            Microsoft.Office.Interop.Excel.Application appExcel = new Microsoft.Office.Interop.Excel.Application();
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

            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemModel tvItemModel = tvItemService.PostAddChildTVItemDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi.Name.Replace(fi.Extension, ""), TVTypeEnum.File);
            if (!string.IsNullOrWhiteSpace(tvItemModel.Error))
            {
                NotUsed = tvItemModel.Error;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(tvItemModel.Error);
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
            tvFileModelNew.FileDescription = reportTypeModel.Description;
            tvFileModelNew.FileType = tvFileService.GetFileType(fi.Extension);
            tvFileModelNew.FileSize_kb = (((int)fi.Length / 1024) == 0 ? 1 : (int)fi.Length / 1024);
            tvFileModelNew.FileInfo = TaskRunnerServiceRes.FileName + "[" + fi.Name + "]\r\n" + TaskRunnerServiceRes.FileType + "[" + fi.Extension + "]\r\n";
            tvFileModelNew.FileCreatedDate_UTC = fi.LastWriteTimeUtc;
            tvFileModelNew.ServerFilePath = (fi.DirectoryName + @"\").Replace(@"C:\", @"E:\");
            tvFileModelNew.LastUpdateDate_UTC = DateTime.UtcNow;
            tvFileModelNew.LastUpdateContactTVItemID = _TaskRunnerBaseService._BWObj.appTaskModel.LastUpdateContactTVItemID;

            TVFileModel tvFileModelRet = tvFileService.PostAddTVFileDB(tvFileModelNew);
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
        #endregion Functions private
    }
}
