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
using CSSPDBDLL.Models;
using CSSPDBDLL.Services;
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;
using CSSPEnumsDLL.Services;
using CSSPDBDLL;

namespace CSSPWebToolsTaskRunner.Services
{
    public class XlsxService
    {
        #region Variables
        string BasePathForExportToArcGIS = @"\\int.ec.gc.ca\sys\InGEO\GW\EC1210WQAEH_QESEA\CSSP_NAT\Data\CSSP_Web_Tools_Imports\";
        #endregion Variables

        #region Properties
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        public SamplingPlanService _SamplingPlanService { get; private set; }
        public LabSheetService _LabSheetService { get; private set; }
        public LabSheetDetailService _LabSheetDetailService { get; private set; }
        public LabSheetTubeMPNDetailService _LabSheetTubeMPNDetailService { get; private set; }
        public TVFileService _TVFileService { get; private set; }
        public AppTaskService _AppTaskService { get; private set; }
        public TVItemService _TVItemService { get; private set; }
        public ContactService _ContactService { get; private set; }
        public EmailDistributionListService _EmailDistributionListService { get; private set; }
        public EmailDistributionListContactService _EmailDistributionListContactService { get; private set; }
        public MWQMAnalysisReportParameterService _MWQMAnalysisReportParameterService { get; private set; }
        public MWQMSubsectorService _MWQMSubsectorService { get; private set; }
        public BaseEnumService _BaseEnumService { get; private set; }
        #endregion Properties

        #region Constructors
        public XlsxService(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
            _SamplingPlanService = new SamplingPlanService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _LabSheetService = new LabSheetService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _LabSheetDetailService = new LabSheetDetailService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _LabSheetTubeMPNDetailService = new LabSheetTubeMPNDetailService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _TVFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _AppTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _TVItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _ContactService = new ContactService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _EmailDistributionListService = new EmailDistributionListService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _EmailDistributionListContactService = new EmailDistributionListContactService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _MWQMAnalysisReportParameterService = new MWQMAnalysisReportParameterService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _MWQMAnalysisReportParameterService = new MWQMAnalysisReportParameterService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _MWQMSubsectorService = new MWQMSubsectorService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _BaseEnumService = new BaseEnumService(_TaskRunnerBaseService._BWObj.appTaskModel.Language);
        }
        #endregion Constructors

        #region Functions public
        public void CreateDocumentFromTemplateXlsx(DocTemplateModel docTemplateModel)
        {
            string NotUsed = "";

            TVFileModel tvFileModelTemplate = _TVFileService.GetTVFileModelWithTVFileTVItemIDDB(docTemplateModel.TVFileTVItemID);
            if (!string.IsNullOrWhiteSpace(tvFileModelTemplate.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.TVFileTVItemID, docTemplateModel.TVFileTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.TVFileTVItemID, docTemplateModel.TVFileTVItemID.ToString());
                return;
            }

            string ServerNewFilePath = _TVFileService.ChoseEDriveOrCDrive(_TVFileService.GetServerFilePath(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID));

            DateTime CD = DateTime.Now;
            string Language = "_" + _TaskRunnerBaseService._BWObj.appTaskModel.Language;

            string DateText = "_" + CD.Year.ToString() +
                "_" + (CD.Month > 9 ? CD.Month.ToString() : "0" + CD.Month.ToString()) +
                "_" + (CD.Day > 9 ? CD.Day.ToString() : "0" + CD.Day.ToString()) +
                "_" + (CD.Hour > 9 ? CD.Hour.ToString() : "0" + CD.Hour.ToString()) +
                "_" + (CD.Minute > 9 ? CD.Minute.ToString() : "0" + CD.Minute.ToString());

            FileInfo fi = new FileInfo(docTemplateModel.FileName);
            string NewFileName = docTemplateModel.FileName.Replace(fi.Extension.ToLower(), DateText + Language + fi.Extension.ToLower());

            fi = new FileInfo(ServerNewFilePath + NewFileName);

            if (fi.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.FileAlreadyExist_, fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("FileAlreadyExist_", fi.FullName);
                return;
            }

            try
            {
                File.Copy(_TVFileService.ChoseEDriveOrCDrive(tvFileModelTemplate.ServerFilePath) + tvFileModelTemplate.ServerFileName, fi.FullName);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateDocumentFromTemplateError_, ex.Message + " - " + (ex.InnerException != null ? ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotCreateDocumentFromTemplateError_", ex.Message + " - " + (ex.InnerException != null ? ex.InnerException.Message : ""));
                return;
            }

            TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            string retStr = ParseXlsx(fi);
            if (!string.IsNullOrWhiteSpace(retStr))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateDocumentFromTemplateError_, retStr);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotCreateDocumentFromTemplateError_", retStr);
                return;
            }

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.FileGeneratedFromTemplate, FilePurposeEnum.TemplateGenerated);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;
        }
        public void CreateExcelFileForAnalysisReportParameter()
        {
            string NotUsed = "";

            string Parameters = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters;
            string[] ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            int SubsectorTVItemID = 0;
            int MWQMAnalysisReportParameterID = 0;
            foreach (string s in ParamValueList)
            {
                string[] ParamValue = s.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                if (ParamValue.Length != 2)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, TaskRunnerServiceRes.Parameters);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", TaskRunnerServiceRes.Parameters);
                    return;
                }

                if (ParamValue[0] == "SubsectorTVItemID")
                {
                    SubsectorTVItemID = int.Parse(ParamValue[1]);
                }
                else if (ParamValue[0] == "MWQMAnalysisReportParameterID")
                {
                    MWQMAnalysisReportParameterID = int.Parse(ParamValue[1]);
                }
                else
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind__, TaskRunnerServiceRes.Parameter, ParamValue[0]);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind__", TaskRunnerServiceRes.Parameter, ParamValue[0].ToString());
                    return;
                }
            }

            if (SubsectorTVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.SubsectorTVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.SubsectorTVItemID.ToString());
                return;
            }

            if (MWQMAnalysisReportParameterID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.MWQMAnalysisReportParameterID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.MWQMAnalysisReportParameterID.ToString());
                return;
            }

            TVItemModel tvItemModelSubsector = _TVItemService.GetTVItemModelWithTVItemIDDB(SubsectorTVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelSubsector.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.Subsector, SubsectorTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.Subsector, SubsectorTVItemID.ToString());
                return;
            }

            MWQMAnalysisReportParameterModel mwqmAnalysisReportParameterModel = _MWQMAnalysisReportParameterService.GetMWQMAnalysisReportParameterModelWithMWQMAnalysisReportParameterIDDB(MWQMAnalysisReportParameterID);
            if (!string.IsNullOrWhiteSpace(mwqmAnalysisReportParameterModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MWQMAnalysisReportParameter, TaskRunnerServiceRes.MWQMAnalysisReportParameterID, MWQMAnalysisReportParameterID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MWQMAnalysisReportParameter, TaskRunnerServiceRes.MWQMAnalysisReportParameterID, MWQMAnalysisReportParameterID.ToString());
                return;
            }

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                NotUsed = TaskRunnerServiceRes.ExcelCouldNotBeStartedCheckOfficeInstallation;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("ExcelCouldNotBeStartedCheckOfficeInstallation");
                return;
            }
            xlApp.Visible = false;

            Microsoft.Office.Interop.Excel.Workbook wb = xlApp.Workbooks.Add();
            //Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];
            if (wb == null)
            {
                NotUsed = TaskRunnerServiceRes.ExcelCouldNotBeStartedCheckOfficeInstallation;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("ExcelCouldNotBeStartedCheckOfficeInstallation");
                return;
            }

            SetupParametersAndBasicTextOnSheet2(xlApp, wb, mwqmAnalysisReportParameterModel);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
            {
                return;
            }

            SetupParametersAndBasicTextOnSheet1(xlApp, wb, mwqmAnalysisReportParameterModel);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
            {
                return;
            }

            SetupStatOnSheet1(xlApp, wb, mwqmAnalysisReportParameterModel);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
            {
                return;
            }

            int Percent = 50;
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);

            string FilePath = _TVFileService.GetServerFilePath(SubsectorTVItemID);

            DirectoryInfo di = new DirectoryInfo(FilePath);
            if (!di.Exists)
            {
                try
                {
                    di.Create();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateDirectory__, FilePath, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateDirectory__", FilePath, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return;
                }
            }

            string dateText = DateTime.Now.ToString("_yyyy_MM_dd_HH_mm_") + (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr ? "fr" : "en");
            FileInfo fi = new FileInfo(FilePath + mwqmAnalysisReportParameterModel.AnalysisName + dateText + ".xlsx");

            TVItemModel tvItemModelTVFileExist = _TVItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(SubsectorTVItemID, FilePath + mwqmAnalysisReportParameterModel.AnalysisName + dateText, TVTypeEnum.File);
            if (string.IsNullOrWhiteSpace(tvItemModelTVFileExist.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes._AlreadyExist, FilePath + mwqmAnalysisReportParameterModel.AnalysisName + dateText);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_AlreadyExist", FilePath + mwqmAnalysisReportParameterModel.AnalysisName + dateText);
                return;
            }

            TVItemModel tvItemModelTVFile = _TVItemService.PostAddChildTVItemDB(SubsectorTVItemID, FilePath + mwqmAnalysisReportParameterModel.AnalysisName + dateText, TVTypeEnum.File);
            if (!string.IsNullOrWhiteSpace(tvItemModelTVFile.Error))
            {
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(tvItemModelTVFile.Error);
                return;
            }

            Percent = 60;
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);

            try
            {
                wb.SaveAs(fi.FullName);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotSaveExcelFile_Error_, fi.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotSaveExcelFile_Error_", fi.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                return;
            }

            try
            {
                wb.Close();
                xlApp.Quit();
            }
            catch (Exception)
            {
                // nothing
            }

            Percent = 70;
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);

            fi = new FileInfo(fi.FullName);
            if (!fi.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.File_DoesNotExist, fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("File_DoesNotExist", fi.FullName);
                return;
            }

            TVFileModel tvFileModelNew = new TVFileModel();
            tvFileModelNew.TVFileTVItemID = tvItemModelTVFile.TVItemID;
            tvFileModelNew.FileCreatedDate_UTC = fi.CreationTimeUtc;
            tvFileModelNew.FileDescription = "File created from analysis Excel export";
            tvFileModelNew.FilePurpose = FilePurposeEnum.Analysis;
            tvFileModelNew.FileSize_kb = (int)(fi.Length / 1024);
            tvFileModelNew.FileType = FileTypeEnum.XLSX;
            tvFileModelNew.FileInfo = "File created from analysis Excel export";
            tvFileModelNew.Language = _TaskRunnerBaseService._BWObj.appTaskModel.Language;
            tvFileModelNew.Year = DateTime.Now.Year;
            tvFileModelNew.ServerFileName = fi.Name;
            tvFileModelNew.ServerFilePath = fi.Directory + @"\";

            TVFileModel tvFileModelRet = _TVFileService.PostAddTVFileDB(tvFileModelNew);
            if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
            {
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(tvFileModelRet.Error);
                return;
            }

            mwqmAnalysisReportParameterModel.ExcelTVFileTVItemID = tvItemModelTVFile.TVItemID;

            MWQMAnalysisReportParameterModel mwqmAnalysisReportParameterModelRet = _MWQMAnalysisReportParameterService.PostUpdateMWQMAnalysisReportParameterDB(mwqmAnalysisReportParameterModel);
            if (!string.IsNullOrWhiteSpace(mwqmAnalysisReportParameterModelRet.Error))
            {
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(mwqmAnalysisReportParameterModelRet.Error);
                return;
            }

            Percent = 80;
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);


        }
        public void CreateXlsxOfMWQMSitesAndSamples()
        {
            string NotUsed = "";
            string ProvInit = "";
            List<string> ProvInitList = new List<string>()
            {
                "BC", "ME", "NB", "NL", "NS", "PE", "QC",
            };
            List<string> ProvList = new List<string>()
            {
                "British Columbia", "Maine", "New Brunswick", "Newfoundland and Labrador", "Nova Scotia", "Prince Edward Island", "Québec",
            };
            int TVItemID = _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID;

            if (TVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.Parameter_NotFound, "TVItemID");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Parameter_NotFound", "TVItemID");
                return;
            }

            TVItemModel tvItemModel = _TVItemService.GetTVItemModelWithTVItemIDDB(TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotDeleteFile_Error_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                return;
            }

            for (int i = 0, countProv = ProvList.Count; i < countProv; i++)
            {
                if (ProvList[i] == tvItemModel.TVText)
                {
                    ProvInit = ProvInitList[i];
                    break;
                }
            }

            string ServerFilePath = _TVFileService.GetServerFilePath(TVItemID);

            FileInfo fi = new FileInfo(_TVFileService.ChoseEDriveOrCDrive(ServerFilePath) + $"Sites_Samples_Echantillons_{ProvInit}.xlsx");

            TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            if (fi.Exists)
            {
                try
                {
                    fi.Delete();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDeleteFile_Error_, fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDeleteFile_Error_", fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    return;
                }
            }

            //StringBuilder sb = new StringBuilder();

            Microsoft.Office.Interop.Excel._Application appExcel = new Microsoft.Office.Interop.Excel.Application();
            appExcel.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = appExcel.Workbooks.Add();
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet2 = xlWorkbook.Sheets[2];

            xlWorksheet.Name = "Sites";

            xlWorksheet.Cells[1, 1].Value = "Site_ID";
            xlWorksheet.Cells[1, 2].Value = "Lat";
            xlWorksheet.Cells[1, 3].Value = "Long";

            // Doing Sheet #1
            int rowCount = 1;
            using (CSSPDBEntities db = new CSSPDBEntities())
            {
                var tvItemProv = (from c in db.TVItems
                                  from cl in db.TVItemLanguages
                                  where c.TVItemID == cl.TVItemID
                                  && c.TVItemID == TVItemID
                                  && cl.Language == (int)LanguageEnum.en
                                  && c.TVType == (int)TVTypeEnum.Province
                                  select new { c, cl }).FirstOrDefault();

                if (tvItemProv == null)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                    return;
                }

                for (int i = 0, countProv = ProvList.Count; i < countProv; i++)
                {
                    if (ProvList[i] == tvItemProv.cl.TVText)
                    {
                        ProvInit = ProvInitList[i];
                        break;
                    }
                }

                NotUsed = string.Format(TaskRunnerServiceRes.Creating_, fi.Name);
                List<TextLanguage> TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", fi.Name);

                _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);

                var tvItemSSList = (from t in db.TVItems
                                    from tl in db.TVItemLanguages
                                    where t.TVItemID == tl.TVItemID
                                    && tl.Language == (int)LanguageEnum.en
                                    && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                    && t.TVType == (int)TVTypeEnum.Subsector
                                    orderby tl.TVText
                                    select new { t, tl }).ToList();

                var MonitoringSiteList = (from t in db.TVItems
                                          from mi in db.MapInfos
                                          from mip in db.MapInfoPoints
                                          let hasSample = (from c in db.MWQMSamples
                                                           where c.MWQMSiteTVItemID == t.TVItemID
                                                           && c.UseForOpenData == true
                                                           select c).Any()
                                          where mi.TVItemID == t.TVItemID
                                          && mip.MapInfoID == mi.MapInfoID
                                          && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                          && t.TVType == (int)TVTypeEnum.MWQMSite
                                          && mi.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Point
                                          && mi.TVType == (int)TVTypeEnum.MWQMSite
                                          && hasSample == true
                                          select new { t, mip, hasSample }).ToList();


                int TotalCount2 = tvItemSSList.Count;
                int Count2 = 0;
                foreach (var tvItemSS in tvItemSSList)
                {
                    if (Count2 % 20 == 0)
                    {
                        _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)(50.0f * ((float)Count2 / (float)TotalCount2)));

                        NotUsed = string.Format(TaskRunnerServiceRes.Creating_, fi.Name + " --- doing " + tvItemSS.tl.TVText + "");
                        TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", fi.Name + " --- doing " + tvItemSS.tl.TVText + "");

                        _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);
                    }

                    Count2 += 1;

                    foreach (var mwqmSite in MonitoringSiteList.Where(c => c.t.ParentID == tvItemSS.t.TVItemID))
                    {
                        string MS = mwqmSite.t.TVItemID.ToString();
                        string Lat = (mwqmSite.mip != null ? mwqmSite.mip.Lat.ToString("F6").Replace(",", ".") : "");
                        string Lng = (mwqmSite.mip != null ? mwqmSite.mip.Lng.ToString("F6").Replace(",", ".") : "");

                        rowCount += 1;

                        xlWorksheet.Cells[rowCount, 1].Value = MS;
                        xlWorksheet.Cells[rowCount, 2].Value = Lat;
                        xlWorksheet.Cells[rowCount, 3].Value = Lng;
                    }
                }
            }


            // Doing Sheet #2
            xlWorksheet2.Name = "Samples_Echantillons";

            xlWorksheet2.Cells[1, 1].Value = "Site_ID";
            xlWorksheet2.Cells[1, 2].Value = "Date";
            xlWorksheet2.Cells[1, 3].Value = "FC_MPN_CF_NPP";
            xlWorksheet2.Cells[1, 4].Value = "Temp_°C";
            xlWorksheet2.Cells[1, 5].Value = "Sal_PPT_PPM";
            //xlWorksheet2.Cells[1, 6].Value = "pH";
            //xlWorksheet2.Cells[1, 7].Value = "Depth_Profondeur_m";

            rowCount = 1;

            using (CSSPDBEntities db = new CSSPDBEntities())
            {
                var tvItemProv = (from c in db.TVItems
                                  from cl in db.TVItemLanguages
                                  where c.TVItemID == cl.TVItemID
                                  && c.TVItemID == TVItemID
                                  && cl.Language == (int)LanguageEnum.en
                                  && c.TVType == (int)TVTypeEnum.Province
                                  select new { c, cl }).FirstOrDefault();

                if (tvItemProv == null)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                    return;
                }

                for (int i = 0, countProv = ProvList.Count; i < countProv; i++)
                {
                    if (ProvList[i] == tvItemProv.cl.TVText)
                    {
                        ProvInit = ProvInitList[i];
                        break;
                    }
                }

                NotUsed = string.Format(TaskRunnerServiceRes.Creating_, fi.Name);
                List<TextLanguage> TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", fi.Name);

                _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);

                var tvItemSSList = (from t in db.TVItems
                                    from tl in db.TVItemLanguages
                                    where t.TVItemID == tl.TVItemID
                                    && tl.Language == (int)LanguageEnum.en
                                    && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                    && t.TVType == (int)TVTypeEnum.Subsector
                                    orderby tl.TVText
                                    select new { t, tl }).ToList();

                var MonitoringSiteList = (from t in db.TVItems
                                          from mi in db.MapInfos
                                          from mip in db.MapInfoPoints
                                          let hasSample = (from c in db.MWQMSamples
                                                           where c.MWQMSiteTVItemID == t.TVItemID
                                                           && c.UseForOpenData == true
                                                           select c).Any()
                                          where mi.TVItemID == t.TVItemID
                                          && mip.MapInfoID == mi.MapInfoID
                                          && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                          && t.TVType == (int)TVTypeEnum.MWQMSite
                                          && mi.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Point
                                          && mi.TVType == (int)TVTypeEnum.MWQMSite
                                          && hasSample == true
                                          select new { t, mip, hasSample }).ToList();

                int TotalCount2 = tvItemSSList.Count;
                int Count2 = 0;
                foreach (var tvItemSS in tvItemSSList)
                {
                    if (Count2 % 2 == 0)
                    {
                        _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 50 + (int)(50.0f * ((float)Count2 / (float)TotalCount2)));

                        NotUsed = string.Format(TaskRunnerServiceRes.Creating_, fi.Name + " --- doing " + tvItemSS.tl.TVText + "");
                        TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", fi.Name + " --- doing " + tvItemSS.tl.TVText + "");

                        _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);
                    }

                    Count2 += 1;

                    foreach (var mwqmSite in MonitoringSiteList.Where(c => c.t.ParentID == tvItemSS.t.TVItemID))
                    {
                        using (CSSPDBEntities db2 = new CSSPDBEntities())
                        {
                            List<MWQMSample> sampleList = (from c in db2.MWQMSamples
                                                           where c.MWQMSiteTVItemID == mwqmSite.t.TVItemID
                                                           && c.UseForOpenData == true
                                                           orderby c.SampleDateTime_Local ascending
                                                           select c).ToList();

                            foreach (MWQMSample mwqmSample in sampleList)
                            {
                                string MS = mwqmSite.t.TVItemID.ToString();
                                string D = mwqmSample.SampleDateTime_Local.ToString("yyy-MM-dd");
                                string FC = (mwqmSample.FecCol_MPN_100ml < 2 ? "< 2" : (mwqmSample.FecCol_MPN_100ml > 1600 ? "> 1600" : mwqmSample.FecCol_MPN_100ml.ToString().Replace(",", ".")));
                                string Temp = (mwqmSample.WaterTemp_C != null ? ((double)mwqmSample.WaterTemp_C).ToString("F1").Replace(",", ".") : "");
                                string Sal = (mwqmSample.Salinity_PPT != null ? ((double)mwqmSample.Salinity_PPT).ToString("F1").Replace(",", ".") : "");
                                //string pH = (mwqmSample.PH != null ? ((double)mwqmSample.PH).ToString("F1").Replace(",", ".") : "");
                                //string Depth = (mwqmSample.Depth_m != null ? ((double)mwqmSample.Depth_m).ToString("F1").Replace(",", ".") : "");

                                rowCount += 1;
                                xlWorksheet2.Cells[rowCount, 1].Value = MS;
                                xlWorksheet2.Cells[rowCount, 2].Value = D;
                                xlWorksheet2.Cells[rowCount, 3].Value = FC;
                                xlWorksheet2.Cells[rowCount, 4].Value = Temp;
                                xlWorksheet2.Cells[rowCount, 5].Value = Sal;
                                //xlWorksheet2.Cells[rowCount, 6].Value = pH;
                                //xlWorksheet2.Cells[rowCount, 7].Value = Depth;
                            }
                        }
                    }
                }
            }

            appExcel.ActiveWorkbook.SaveAs(fi.FullName, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
            xlWorkbook.Close();
            appExcel.Quit();

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.XlsxOfMWQMSamples, FilePurposeEnum.OpenData);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;
        }
        //public void CreateXlsxOfMWQMSites()
        //{
        //    string NotUsed = "";
        //    string ProvInit = "";
        //    List<string> ProvInitList = new List<string>()
        //    {
        //        "BC", "ME", "NB", "NL", "NS", "PE", "QC",
        //    };
        //    List<string> ProvList = new List<string>()
        //    {
        //        "British Columbia", "Maine", "New Brunswick", "Newfoundland and Labrador", "Nova Scotia", "Prince Edward Island", "Québec",
        //    };
        //    int TVItemID = _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID;

        //    if (TVItemID == 0)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.Parameter_NotFound, "TVItemID");
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Parameter_NotFound", "TVItemID");
        //        return;
        //    }

        //    TVItemModel tvItemModel = _TVItemService.GetTVItemModelWithTVItemIDDB(TVItemID);
        //    if (!string.IsNullOrWhiteSpace(tvItemModel.Error))
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotDeleteFile_Error_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
        //        return;
        //    }

        //    for (int i = 0, countProv = ProvList.Count; i < countProv; i++)
        //    {
        //        if (ProvList[i] == tvItemModel.TVText)
        //        {
        //            ProvInit = ProvInitList[i];
        //            break;
        //        }
        //    }

        //    string ServerFilePath = _TVFileService.GetServerFilePath(TVItemID);

        //    FileInfo fi = new FileInfo(_TVFileService.ChoseEDriveOrCDrive(ServerFilePath) + $"Sites_{ProvInit}.xlsx");

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    if (fi.Exists)
        //    {
        //        try
        //        {
        //            fi.Delete();
        //        }
        //        catch (Exception ex)
        //        {
        //            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDeleteFile_Error_, fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
        //            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDeleteFile_Error_", fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
        //            return;
        //        }
        //    }

        //    StringBuilder sb = new StringBuilder();

        //    Microsoft.Office.Interop.Excel._Application appExcel = new Microsoft.Office.Interop.Excel.Application();
        //    appExcel.Visible = false;
        //    Microsoft.Office.Interop.Excel.Workbook xlWorkbook = appExcel.Workbooks.Add();
        //    Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];

        //    // loop through all the MWQMSites etc...

        //    xlWorksheet.Cells[1, 1].Value = "ID";
        //    xlWorksheet.Cells[1, 2].Value = "Lat";
        //    xlWorksheet.Cells[1, 3].Value = "Lng";

        //    int rowCount = 1;
        //    using (CSSPDBEntities db = new CSSPDBEntities())
        //    {
        //        var tvItemProv = (from c in db.TVItems
        //                          from cl in db.TVItemLanguages
        //                          where c.TVItemID == cl.TVItemID
        //                          && c.TVItemID == TVItemID
        //                          && cl.Language == (int)LanguageEnum.en
        //                          && c.TVType == (int)TVTypeEnum.Province
        //                          select new { c, cl }).FirstOrDefault();

        //        if (tvItemProv == null)
        //        {
        //            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
        //            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
        //            return;
        //        }

        //        for (int i = 0, countProv = ProvList.Count; i < countProv; i++)
        //        {
        //            if (ProvList[i] == tvItemProv.cl.TVText)
        //            {
        //                ProvInit = ProvInitList[i];
        //                break;
        //            }
        //        }

        //        NotUsed = string.Format(TaskRunnerServiceRes.Creating_, fi.Name);
        //        List<TextLanguage> TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", fi.Name);

        //        _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);

        //        var tvItemSSList = (from t in db.TVItems
        //                            from tl in db.TVItemLanguages
        //                            where t.TVItemID == tl.TVItemID
        //                            && tl.Language == (int)LanguageEnum.en
        //                            && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
        //                            && t.TVType == (int)TVTypeEnum.Subsector
        //                            orderby tl.TVText
        //                            select new { t, tl }).ToList();

        //        var MonitoringSiteList = (from t in db.TVItems
        //                                  from mi in db.MapInfos
        //                                  from mip in db.MapInfoPoints
        //                                  let hasSample = (from c in db.MWQMSamples
        //                                                   where c.MWQMSiteTVItemID == t.TVItemID
        //                                                   && c.UseForOpenData == true
        //                                                   select c).Any()
        //                                  where mi.TVItemID == t.TVItemID
        //                                  && mip.MapInfoID == mi.MapInfoID
        //                                  && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
        //                                  && t.TVType == (int)TVTypeEnum.MWQMSite
        //                                  && mi.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Point
        //                                  && mi.TVType == (int)TVTypeEnum.MWQMSite
        //                                  && hasSample == true
        //                                  select new { t, mip, hasSample }).ToList();


        //        int TotalCount2 = tvItemSSList.Count;
        //        int Count2 = 0;
        //        foreach (var tvItemSS in tvItemSSList)
        //        {
        //            if (Count2 % 20 == 0)
        //            {
        //                _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)(100.0f * ((float)Count2 / (float)TotalCount2)));

        //                NotUsed = string.Format(TaskRunnerServiceRes.Creating_, fi.Name + " --- doing " + tvItemSS.tl.TVText + "");
        //                TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", fi.Name + " --- doing " + tvItemSS.tl.TVText + "");

        //                _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);
        //            }

        //            Count2 += 1;

        //            foreach (var mwqmSite in MonitoringSiteList.Where(c => c.t.ParentID == tvItemSS.t.TVItemID))
        //            {
        //                string MS = mwqmSite.t.TVItemID.ToString();
        //                string Lat = (mwqmSite.mip != null ? mwqmSite.mip.Lat.ToString("F6").Replace(",", ".") : "");
        //                string Lng = (mwqmSite.mip != null ? mwqmSite.mip.Lng.ToString("F6").Replace(",",".") : "");

        //                rowCount += 1;

        //                xlWorksheet.Cells[rowCount, 1].Value = MS;
        //                xlWorksheet.Cells[rowCount, 2].Value = Lat;
        //                xlWorksheet.Cells[rowCount, 3].Value = Lng;
        //            }
        //        }
        //    }


        //    appExcel.ActiveWorkbook.SaveAs(fi.FullName, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
        //    xlWorkbook.Close();
        //    appExcel.Quit();

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.XlsxOfMWQMSites, FilePurposeEnum.OpenData);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;
        //}

        public void CreateSamplingPlanConfigExcelFile()
        {
            string NotUsed = "";

            bool t = true;
            if (t)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.NotImplementedYet);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("NotImplementedYet");
                return;
            }
        }
        public void ExportToArcGIS()
        {
            string NotUsed = "";

            string Parameters = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters;
            string[] ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            string ProvinceTVItemIDText = "";
            List<int> ProvinceTVItemIDList = new List<int>();
            bool Active = false;
            bool Inactive = false;
            string DocType = "";

            // example: |||ProvinceTVItemIDText,_7_10|||Active,True|||Inactive,True|||DocType,monitoringsites|||
            foreach (string s in ParamValueList)
            {
                string[] ParamValue = s.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                if (ParamValue.Length != 2)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, TaskRunnerServiceRes.Parameters);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", TaskRunnerServiceRes.Parameters);
                    return;
                }

                if (ParamValue[0] == "ProvinceTVItemIDText")
                {
                    ProvinceTVItemIDText = ParamValue[1];
                }
                else if (ParamValue[0] == "Active")
                {
                    Active = bool.Parse(ParamValue[1]);
                }
                else if (ParamValue[0] == "Inactive")
                {
                    Inactive = bool.Parse(ParamValue[1]);
                }
                else if (ParamValue[0] == "DocType")
                {
                    DocType = ParamValue[1];
                }
                else
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind__, TaskRunnerServiceRes.Parameter, ParamValue[0]);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind__", TaskRunnerServiceRes.Parameter, ParamValue[0].ToString());
                    return;
                }
            }

            ProvinceTVItemIDList = ProvinceTVItemIDText.Split("_".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Select(c => int.Parse(c)).ToList();

            int count = 0;
            foreach (int ProvinceTVItemID in ProvinceTVItemIDList)
            {
                _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)((100.0f * ((float)count + 0.5f)) / (float)ProvinceTVItemIDList.Count));

                count += 1;

                if (Active)
                {
                    switch (DocType)
                    {
                        case "monitoringsites":
                            {
                                CreateMonitoringSitesDocumentForArcGIS(ProvinceTVItemID, count, ProvinceTVItemIDList.Count, true);
                            }
                            break;
                        case "pollutionsourcesites":
                            {
                                CreatePollutionSourceSitesDocumentForArcGIS(ProvinceTVItemID, count, ProvinceTVItemIDList.Count, true);
                            }
                            break;
                        case "municipalities":
                            {
                                CreateMunicipalitiesDocumentForArcGIS(ProvinceTVItemID, count, ProvinceTVItemIDList.Count, true);
                            }
                            break;
                        case "wwtp":
                            {
                                CreateWWTPsDocumentForArcGIS(ProvinceTVItemID, count, ProvinceTVItemIDList.Count, true);
                            }
                            break;
                        case "liftstations":
                            {
                                CreateLiftStationsDocumentForArcGIS(ProvinceTVItemID, count, ProvinceTVItemIDList.Count, true);
                            }
                            break;
                        default:
                            break;
                    }
                }

                if (Inactive)
                {
                    switch (DocType)
                    {
                        case "monitoringsites":
                            {
                                CreateMonitoringSitesDocumentForArcGIS(ProvinceTVItemID, count, ProvinceTVItemIDList.Count, false);
                            }
                            break;
                        case "pollutionsourcesites":
                            {
                                CreatePollutionSourceSitesDocumentForArcGIS(ProvinceTVItemID, count, ProvinceTVItemIDList.Count, false);
                            }
                            break;
                        case "municipalities":
                            {
                                CreateMunicipalitiesDocumentForArcGIS(ProvinceTVItemID, count, ProvinceTVItemIDList.Count, false);
                            }
                            break;
                        case "wwtp":
                            {
                                CreateWWTPsDocumentForArcGIS(ProvinceTVItemID, count, ProvinceTVItemIDList.Count, false);
                            }
                            break;
                        case "liftstations":
                            {
                                CreateLiftStationsDocumentForArcGIS(ProvinceTVItemID, count, ProvinceTVItemIDList.Count, false);
                            }
                            break;
                        default:
                            break;
                    }
                }
            }
        }

        private void CreateLiftStationsDocumentForArcGIS(int ProvinceTVItemID, int count, int TotalCount, bool Active)
        {
            string BasePath = "";
            string ProvInit = "";
            List<string> ProvInitList = new List<string>()
            {
                "BC", "ME", "NB", "NL", "NS", "PE", "QC",
            };
            List<string> ProvList = new List<string>()
            {
                "British Columbia", "Maine", "New Brunswick", "Newfoundland and Labrador", "Nova Scotia", "Prince Edward Island", "Québec",
            };
            string NotUsed = "";
            StringBuilder sb = new StringBuilder();

            sb.AppendLine("Muni,LSName,Lat,Lng,PercFlow,OutLat,OutLng,AlarmSystemType,CanOverFlow,Address,CSSPUrl,Comment");

            using (CSSPDBEntities db = new CSSPDBEntities())
            {
                var tvItemProv = (from c in db.TVItems
                                  from cl in db.TVItemLanguages
                                  where c.TVItemID == cl.TVItemID
                                  && c.TVItemID == ProvinceTVItemID
                                  && cl.Language == (int)LanguageEnum.en
                                  && c.TVType == (int)TVTypeEnum.Province
                                  select new { c, cl }).FirstOrDefault();

                if (tvItemProv == null)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, ProvinceTVItemID.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, ProvinceTVItemID.ToString());
                    return;
                }

                for (int i = 0, countProv = ProvList.Count; i < countProv; i++)
                {
                    if (ProvList[i] == tvItemProv.cl.TVText)
                    {
                        ProvInit = ProvInitList[i];
                        break;
                    }
                }

                switch (ProvInit)
                {
                    case "NB":
                    case "NL":
                    case "NS":
                    case "PE":
                        {
                            BasePath = BasePathForExportToArcGIS.Replace("CSSP_NAT", "CSSP_ATL");
                        }
                        break;
                    case "BC":
                        {
                            BasePath = BasePathForExportToArcGIS.Replace("CSSP_NAT", "CSSP_NAT"); // will have to change to "CSSP_PYR"
                        }
                        break;
                    case "QC":
                        {
                            BasePath = BasePathForExportToArcGIS.Replace("CSSP_NAT", "CSSP_QC");
                        }
                        break;
                    default:
                        break;
                }

                NotUsed = string.Format(TaskRunnerServiceRes.Creating_, BasePath + @"LiftStations_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv");
                List<TextLanguage> TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", BasePath + @"LiftStations_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv");

                _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);

                var tvItemMuniList = (from t in db.TVItems
                                      from tl in db.TVItemLanguages
                                      where t.TVItemID == tl.TVItemID
                                      && tl.Language == (int)LanguageEnum.en
                                      && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                      && t.TVType == (int)TVTypeEnum.Municipality
                                      && t.IsActive == Active
                                      orderby tl.TVText
                                      select new { t, tl }).ToList();

                var LiftStationList = (from c in db.TVItems
                                       from cl in db.TVItemLanguages
                                       let inf = (from a in db.Infrastructures
                                                  where a.InfrastructureTVItemID == c.TVItemID
                                                  && a.InfrastructureType == (int)InfrastructureTypeEnum.LiftStation
                                                  select a).FirstOrDefault()
                                       let infLang = (from b in db.InfrastructureLanguages
                                                      where b.InfrastructureID == inf.InfrastructureID
                                                      && b.Language == (int)LanguageEnum.en
                                                      select b).FirstOrDefault()
                                       let latLng = (from mi in db.MapInfos
                                                     from mip in db.MapInfoPoints
                                                     where mi.TVItemID == c.TVItemID
                                                     && mi.MapInfoID == mip.MapInfoID
                                                     && mi.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Point
                                                     && mi.TVType == (int)TVTypeEnum.LiftStation
                                                     select mip).FirstOrDefault()
                                       let latLngOut = (from mi in db.MapInfos
                                                        from mip in db.MapInfoPoints
                                                        where mi.TVItemID == c.TVItemID
                                                        && mi.MapInfoID == mip.MapInfoID
                                                        && mi.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Point
                                                        && mi.TVType == (int)TVTypeEnum.Outfall
                                                        select mip).FirstOrDefault()
                                       let address = (from a in db.Addresses
                                                      let muni = (from cl in db.TVItemLanguages where cl.TVItemID == a.MunicipalityTVItemID && cl.Language == (int)LanguageEnum.en select cl.TVText).FirstOrDefault<string>()
                                                      let add = a.StreetNumber + " " + a.StreetName
                                                      where a.AddressTVItemID == inf.CivicAddressTVItemID
                                                      select new { add }).FirstOrDefault()
                                       where c.TVItemID == cl.TVItemID
                                       && c.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                       && c.TVType == (int)TVTypeEnum.Infrastructure
                                       && c.IsActive == Active
                                       && cl.Language == (int)LanguageEnum.en
                                       && inf.InfrastructureType == (int)InfrastructureTypeEnum.LiftStation
                                       orderby cl.TVText
                                       select new { c, cl, inf, infLang, latLng, latLngOut, address.add }).ToList();


                int TotalCount2 = tvItemMuniList.Count;
                int Count2 = 0;
                foreach (var tvItemMuni in tvItemMuniList)
                {
                    if (Count2 % 20 == 0)
                    {
                        _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)((100.0f * ((float)count - 0.5f + ((float)Count2 / (float)TotalCount2) - 0.5f)) / (float)TotalCount));

                        NotUsed = string.Format(TaskRunnerServiceRes.Creating_, BasePath + @"LiftStations_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv --- doing " + tvItemMuni.tl.TVText + "");
                        TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", BasePath + @"LiftStations_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv --- doing " + tvItemMuni.tl.TVText + "");

                        _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);
                    }

                    Count2 += 1;

                    foreach (var liftStation in LiftStationList.Where(c => c.c.ParentID == tvItemMuni.t.TVItemID))
                    {
                        string MN = tvItemMuni.tl.TVText.Replace(",", "_");
                        string LN = liftStation.cl.TVText.Replace(",", "_");
                        string Lat = (liftStation.latLng != null ? liftStation.latLng.Lat.ToString("F5") : "");
                        string Lng = (liftStation.latLng != null ? liftStation.latLng.Lng.ToString("F5") : "");
                        string PF = (liftStation.inf != null && liftStation.inf.PercFlowOfTotal != null ? ((double)liftStation.inf.PercFlowOfTotal).ToString("F2") : "");
                        string LatOut = (liftStation.latLngOut != null ? liftStation.latLngOut.Lat.ToString("F5") : "");
                        string LngOut = (liftStation.latLngOut != null ? liftStation.latLngOut.Lng.ToString("F5") : "");
                        string AS = (liftStation.inf.AlarmSystemType != null && liftStation.inf.AlarmSystemType > 0 ? (_BaseEnumService.GetEnumText_AlarmSystemTypeEnum((AlarmSystemTypeEnum)liftStation.inf.AlarmSystemType)) : "");
                        string CO = (liftStation.inf.CanOverflow != null ? ((bool)liftStation.inf.CanOverflow).ToString() : "");
                        string AD = (!string.IsNullOrWhiteSpace(liftStation.add) ? liftStation.add.Replace(",", "_") + " --- " + MN : "");
                        string URL = @"http://wmon01dtchlebl2/csspwebtools/en-CA/#!View/a|||" + tvItemMuni.t.TVItemID.ToString() + @"|||30010100004000000000000000000000";
                        string C = (liftStation.infLang != null && !string.IsNullOrWhiteSpace(liftStation.infLang.Comment) ? liftStation.infLang.Comment.Replace(",", "_") : "");
                        sb.AppendLine($"{MN},{LN},{Lat},{Lng},{PF},{LatOut},{LngOut},{AS},{CO},{AD},{URL},{C}".Replace("\r", "   ").Replace("\n", "").Replace("empty", "").Replace("Empty", ""));
                    }
                }
            }

            FileInfo fi = new FileInfo(BasePath + @"LiftStations_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv");

            if (fi.Exists)
            {
                try
                {
                    fi.Delete();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDeleteFile_Error_, fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDeleteFile_Error_", fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    return;
                }
            }

            fi = new FileInfo(BasePath + @"LiftStations_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv");

            try
            {
                File.WriteAllText(fi.FullName, sb.ToString(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateFile_Error_, fi.FullName, ex.Message);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateFile_Error_", fi.FullName, ex.Message);
                return;
            }
        }
        private void CreateMonitoringSitesDocumentForArcGIS(int ProvinceTVItemID, int count, int TotalCount, bool Active)
        {
            string BasePath = "";
            string ProvInit = "";
            List<string> ProvInitList = new List<string>()
            {
                "BC", "ME", "NB", "NL", "NS", "PE", "QC",
            };
            List<string> ProvList = new List<string>()
            {
                "British Columbia", "Maine", "New Brunswick", "Newfoundland and Labrador", "Nova Scotia", "Prince Edward Island", "Québec",
            };
            string NotUsed = "";
            StringBuilder sb = new StringBuilder();

            sb.AppendLine("Subsector,SiteName,SiteDesc,Lat,Lng,CurrentClass,Letter,Color,AsOf,StatPeriod,GMean,Median,P90,PercOver43,CSSPUrl");

            using (CSSPDBEntities db = new CSSPDBEntities())
            {
                var tvItemProv = (from c in db.TVItems
                                  from cl in db.TVItemLanguages
                                  where c.TVItemID == cl.TVItemID
                                  && c.TVItemID == ProvinceTVItemID
                                  && cl.Language == (int)LanguageEnum.en
                                  && c.TVType == (int)TVTypeEnum.Province
                                  select new { c, cl }).FirstOrDefault();

                if (tvItemProv == null)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, ProvinceTVItemID.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, ProvinceTVItemID.ToString());
                    return;
                }

                for (int i = 0, countProv = ProvList.Count; i < countProv; i++)
                {
                    if (ProvList[i] == tvItemProv.cl.TVText)
                    {
                        ProvInit = ProvInitList[i];
                        break;
                    }
                }

                switch (ProvInit)
                {
                    case "NB":
                    case "NL":
                    case "NS":
                    case "PE":
                        {
                            BasePath = BasePathForExportToArcGIS.Replace("CSSP_NAT", "CSSP_ATL");
                        }
                        break;
                    case "BC":
                        {
                            BasePath = BasePathForExportToArcGIS.Replace("CSSP_NAT", "CSSP_NAT"); // will have to change to "CSSP_PYR"
                        }
                        break;
                    case "QC":
                        {
                            BasePath = BasePathForExportToArcGIS.Replace("CSSP_NAT", "CSSP_QC");
                        }
                        break;
                    default:
                        break;
                }

                NotUsed = string.Format(TaskRunnerServiceRes.Creating_, BasePath + @"MonitoringSites_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv");
                List<TextLanguage> TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", BasePath + @"MonitoringSites_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv");

                _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);

                var tvItemSSList = (from t in db.TVItems
                                    from tl in db.TVItemLanguages
                                    where t.TVItemID == tl.TVItemID
                                    && tl.Language == (int)LanguageEnum.en
                                    && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                    && t.TVType == (int)TVTypeEnum.Subsector
                                    && t.IsActive == Active
                                    orderby tl.TVText
                                    select new { t, tl }).ToList();

                var MonitoringSiteList = (from t in db.TVItems
                                          from tl in db.TVItemLanguages
                                          from mi in db.MapInfos
                                          from mip in db.MapInfoPoints
                                          from s in db.MWQMSites
                                          where t.TVItemID == tl.TVItemID
                                          && mi.TVItemID == t.TVItemID
                                          && mip.MapInfoID == mi.MapInfoID
                                          && s.MWQMSiteTVItemID == t.TVItemID
                                          && tl.Language == (int)LanguageEnum.en
                                          && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                          && t.TVType == (int)TVTypeEnum.MWQMSite
                                          && t.IsActive == Active
                                          && mi.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Point
                                          && mi.TVType == (int)TVTypeEnum.MWQMSite
                                          orderby tl.TVText
                                          select new { t, tl, mip, s }).ToList();


                int TotalCount2 = tvItemSSList.Count;
                int Count2 = 0;
                foreach (var tvItemSS in tvItemSSList)
                {
                    if (Count2 % 2 == 0)
                    {
                        _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)((100.0f * ((float)count - 0.5f + ((float)Count2 / (float)TotalCount2) - 0.5f)) / (float)TotalCount));

                        NotUsed = string.Format(TaskRunnerServiceRes.Creating_, BasePath + @"MonitoringSites_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv --- doing " + tvItemSS.tl.TVText + "");
                        TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", BasePath + @"MonitoringSites_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv --- doing " + tvItemSS.tl.TVText + "");

                        _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);
                    }

                    Count2 += 1;

                    foreach (var mwqmSite in MonitoringSiteList.Where(c => c.t.ParentID == tvItemSS.t.TVItemID))
                    {
                        TVItemMoreInfoMWQMSiteModel tvItemMoreInfoMWQMSiteModel = _TVItemService.GetTVItemMoreInfoMWQMSiteTVItemIDDB(mwqmSite.t.TVItemID, 30);

                        string SS = tvItemSS.tl.TVText.Replace(",", "_");
                        string MS = "S" + mwqmSite.tl.TVText.Replace(",", "_");
                        string DESC = mwqmSite.s.MWQMSiteDescription.Replace(",", "_").Replace("\r", "   ").Replace("\n", "  ");
                        string Lat = (mwqmSite.mip != null ? mwqmSite.mip.Lat.ToString("F5") : "");
                        string Lng = (mwqmSite.mip != null ? mwqmSite.mip.Lng.ToString("F5") : "");
                        string CC = (mwqmSite.s.MWQMSiteLatestClassification != 0 ? _BaseEnumService.GetEnumText_MWQMSiteLatestClassificationEnum((MWQMSiteLatestClassificationEnum)mwqmSite.s.MWQMSiteLatestClassification) : "");
                        string L = (tvItemMoreInfoMWQMSiteModel != null ? tvItemMoreInfoMWQMSiteModel.Letter : "");
                        string COL = (tvItemMoreInfoMWQMSiteModel != null ? tvItemMoreInfoMWQMSiteModel.Coloring : "");
                        string AsOf = (tvItemMoreInfoMWQMSiteModel != null && tvItemMoreInfoMWQMSiteModel.LastSampleDate != null ? ((DateTime)tvItemMoreInfoMWQMSiteModel.LastSampleDate).ToString("yyyy MMM dd") : "");
                        string StatPeriod = (tvItemMoreInfoMWQMSiteModel != null && tvItemMoreInfoMWQMSiteModel.LastSampleDate != null ? (tvItemMoreInfoMWQMSiteModel.StatMinYear > 0 ? (tvItemMoreInfoMWQMSiteModel.StatMinYear.ToString() + " - " + tvItemMoreInfoMWQMSiteModel.StatMaxYear.ToString()) : "") : "");
                        string GM = (tvItemMoreInfoMWQMSiteModel != null && tvItemMoreInfoMWQMSiteModel.GeoMean != null ? ((double)tvItemMoreInfoMWQMSiteModel.GeoMean).ToString("F1") : "");
                        string MED = (tvItemMoreInfoMWQMSiteModel != null && tvItemMoreInfoMWQMSiteModel.Median != null ? ((double)tvItemMoreInfoMWQMSiteModel.Median).ToString("F1") : "");
                        string P90 = (tvItemMoreInfoMWQMSiteModel != null && tvItemMoreInfoMWQMSiteModel.P90 != null ? ((double)tvItemMoreInfoMWQMSiteModel.P90).ToString("F1") : "");
                        string PO43 = (tvItemMoreInfoMWQMSiteModel != null && tvItemMoreInfoMWQMSiteModel.PercOver43 != null ? ((double)tvItemMoreInfoMWQMSiteModel.PercOver43).ToString("F1") : "");
                        string URL = @"http://wmon01dtchlebl2/csspwebtools/en-CA/#!View/a|||" + mwqmSite.t.TVItemID.ToString() + @"|||30010100004000000000000000000000";
                        sb.AppendLine($"{SS},{MS},{DESC},{Lat},{Lng},{CC},{L},{COL},{AsOf},{StatPeriod},{GM},{MED},{P90},{PO43},{URL}".Replace("\r", "   ").Replace("\n", " ").Replace("empty", "").Replace("Empty", ""));
                    }
                }
            }

            FileInfo fi = new FileInfo(BasePath + @"MonitoringSites_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv");

            if (fi.Exists)
            {
                try
                {
                    fi.Delete();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDeleteFile_Error_, fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDeleteFile_Error_", fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    return;
                }
            }

            fi = new FileInfo(BasePath + @"MonitoringSites_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv");

            try
            {
                File.WriteAllText(fi.FullName, sb.ToString(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateFile_Error_, fi.FullName, ex.Message);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateFile_Error_", fi.FullName, ex.Message);
                return;
            }

            //Microsoft.Office.Interop.Excel._Application appExcel = new Microsoft.Office.Interop.Excel.Application();
            //appExcel.Visible = false;
            //appExcel.DisplayAlerts = false;
            //Microsoft.Office.Interop.Excel.Workbook xlWorkbook = appExcel.Workbooks.Open(fi.FullName);
            //Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            //string NewDocxFileName = fi.FullName.Replace("_.csv", ".csv");
            //appExcel.ActiveWorkbook.SaveAs(NewDocxFileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSVWindows);
            //xlWorkbook.Close();
            //appExcel.Quit();

        }
        private void CreateMunicipalitiesDocumentForArcGIS(int ProvinceTVItemID, int count, int TotalCount, bool Active)
        {
            string BasePath = "";
            string ProvInit = "";
            List<string> ProvInitList = new List<string>()
            {
                "BC", "ME", "NB", "NL", "NS", "PE", "QC",
            };
            List<string> ProvList = new List<string>()
            {
                "British Columbia", "Maine", "New Brunswick", "Newfoundland and Labrador", "Nova Scotia", "Prince Edward Island", "Québec",
            };
            string NotUsed = "";
            StringBuilder sb = new StringBuilder();

            sb.AppendLine("Municipality,Lat,Lng,CSSPUrl");

            using (CSSPDBEntities db = new CSSPDBEntities())
            {
                var tvItemProv = (from c in db.TVItems
                                  from cl in db.TVItemLanguages
                                  where c.TVItemID == cl.TVItemID
                                  && c.TVItemID == ProvinceTVItemID
                                  && cl.Language == (int)LanguageEnum.en
                                  && c.TVType == (int)TVTypeEnum.Province
                                  select new { c, cl }).FirstOrDefault();

                if (tvItemProv == null)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, ProvinceTVItemID.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, ProvinceTVItemID.ToString());
                    return;
                }

                for (int i = 0, countProv = ProvList.Count; i < countProv; i++)
                {
                    if (ProvList[i] == tvItemProv.cl.TVText)
                    {
                        ProvInit = ProvInitList[i];
                        break;
                    }
                }

                switch (ProvInit)
                {
                    case "NB":
                    case "NL":
                    case "NS":
                    case "PE":
                        {
                            BasePath = BasePathForExportToArcGIS.Replace("CSSP_NAT", "CSSP_ATL");
                        }
                        break;
                    case "BC":
                        {
                            BasePath = BasePathForExportToArcGIS.Replace("CSSP_NAT", "CSSP_NAT"); // will have to change to "CSSP_PYR"
                        }
                        break;
                    case "QC":
                        {
                            BasePath = BasePathForExportToArcGIS.Replace("CSSP_NAT", "CSSP_QC");
                        }
                        break;
                    default:
                        break;
                }

                NotUsed = string.Format(TaskRunnerServiceRes.Creating_, BasePath + @"Municipalities_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv");
                List<TextLanguage> TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", BasePath + @"Municipalities_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv");

                _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);

                var tvItemMuniList = (from t in db.TVItems
                                      from tl in db.TVItemLanguages
                                      from mi in db.MapInfos
                                      from mip in db.MapInfoPoints
                                      where t.TVItemID == tl.TVItemID
                                      && mi.TVItemID == t.TVItemID
                                      && mip.MapInfoID == mi.MapInfoID
                                      && tl.Language == (int)LanguageEnum.en
                                      && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                      && t.TVType == (int)TVTypeEnum.Municipality
                                      && t.IsActive == Active
                                      && mi.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Point
                                      && mi.TVType == (int)TVTypeEnum.Municipality
                                      orderby tl.TVText
                                      select new { t, tl, mip }).ToList();

                int TotalCount2 = tvItemMuniList.Count;
                int Count2 = 0;
                foreach (var tvItemMuni in tvItemMuniList)
                {
                    if (Count2 % 20 == 0)
                    {
                        _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)((100.0f * ((float)count - 0.5f + ((float)Count2 / (float)TotalCount2) - 0.5f)) / (float)TotalCount));

                        NotUsed = string.Format(TaskRunnerServiceRes.Creating_, BasePath + @"Municipalities_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv --- doing " + tvItemMuni.tl.TVText + "");
                        TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", BasePath + @"Municipalities_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv --- doing " + tvItemMuni.tl.TVText + "");

                        _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);
                    }

                    Count2 += 1;

                    string MN = tvItemMuni.tl.TVText.Replace(",", "_");
                    string Lat = (tvItemMuni.mip != null ? tvItemMuni.mip.Lat.ToString("F5") : "");
                    string Lng = (tvItemMuni.mip != null ? tvItemMuni.mip.Lng.ToString("F5") : "");
                    string URL = @"http://wmon01dtchlebl2/csspwebtools/en-CA/#!View/a|||" + tvItemMuni.t.TVItemID.ToString() + @"|||30010100004000000000000000000000";
                    sb.AppendLine($"{MN},{Lat},{Lng},{URL}".Replace("\r", "   ").Replace("\n", "").Replace("empty", "").Replace("Empty", ""));
                }
            }

            FileInfo fi = new FileInfo(BasePath + @"Municipalities_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv");

            if (fi.Exists)
            {
                try
                {
                    fi.Delete();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDeleteFile_Error_, fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDeleteFile_Error_", fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    return;
                }
            }

            fi = new FileInfo(BasePath + @"Municipalities_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv");

            try
            {
                File.WriteAllText(fi.FullName, sb.ToString(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateFile_Error_, fi.FullName, ex.Message);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateFile_Error_", fi.FullName, ex.Message);
                return;
            }
        }
        private void CreatePollutionSourceSitesDocumentForArcGIS(int ProvinceTVItemID, int count, int TotalCount, bool Active)
        {
            string BasePath = "";
            List<string> startWithList = new List<string>() { "101", "143", "910" };
            string ProvInit = "";
            List<string> ProvInitList = new List<string>()
            {
                "BC", "ME", "NB", "NL", "NS", "PE", "QC",
            };
            List<string> ProvList = new List<string>()
            {
                "British Columbia", "Maine", "New Brunswick", "Newfoundland and Labrador", "Nova Scotia", "Prince Edward Island", "Québec",
            };
            string NotUsed = "";
            StringBuilder sb = new StringBuilder();

            //sb.AppendLine("Subsector,Site,Type,SubType,Risk,Lat,Lng,OBS,Address,CSSPUrl,C0,C250,C500,C750,C1000,C1250,C1500,C1750,C2000,C2250,C2500,C2750,C3000,C3250,C3500,C3750");
            //sb.AppendLine("Subsector,Site,Type,SubType,Risk,Lat,Lng,OBS,Address,CSSPUrl");
            sb.AppendLine("Subsector,Site,Type,SubType,Risk,Lat,Lng,ObsDate,CSSPUrl");

            using (CSSPDBEntities db = new CSSPDBEntities())
            {
                var tvItemProv = (from c in db.TVItems
                                  from cl in db.TVItemLanguages
                                  where c.TVItemID == cl.TVItemID
                                  && c.TVItemID == ProvinceTVItemID
                                  && cl.Language == (int)LanguageEnum.en
                                  && c.TVType == (int)TVTypeEnum.Province
                                  select new { c, cl }).FirstOrDefault();

                if (tvItemProv == null)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, ProvinceTVItemID.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, ProvinceTVItemID.ToString());
                    return;
                }

                for (int i = 0, countProv = ProvList.Count; i < countProv; i++)
                {
                    if (ProvList[i] == tvItemProv.cl.TVText)
                    {
                        ProvInit = ProvInitList[i];
                        break;
                    }
                }

                switch (ProvInit)
                {
                    case "NB":
                    case "NL":
                    case "NS":
                    case "PE":
                        {
                            BasePath = BasePathForExportToArcGIS.Replace("CSSP_NAT", "CSSP_ATL");
                        }
                        break;
                    case "BC":
                        {
                            BasePath = BasePathForExportToArcGIS.Replace("CSSP_NAT", "CSSP_NAT"); // will have to change to "CSSP_PYR"
                        }
                        break;
                    case "QC":
                        {
                            BasePath = BasePathForExportToArcGIS.Replace("CSSP_NAT", "CSSP_QC");
                        }
                        break;
                    default:
                        break;
                }

                NotUsed = string.Format(TaskRunnerServiceRes.Creating_, BasePath + @"PollutionSourceSites_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv");
                List<TextLanguage> TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", BasePath + @"PollutionSourceSites_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv");

                _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);

                var tvItemSSList = (from t in db.TVItems
                                    from tl in db.TVItemLanguages
                                    where t.TVItemID == tl.TVItemID
                                    && tl.Language == (int)LanguageEnum.en
                                    && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                    && t.TVType == (int)TVTypeEnum.Subsector
                                    && t.IsActive == Active
                                    orderby tl.TVText
                                    select new { t, tl }).ToList();

                var PollutionSourceSiteList = (from t in db.TVItems
                                               from tl in db.TVItemLanguages
                                               from mi in db.MapInfos
                                               from mip in db.MapInfoPoints
                                               from pss in db.PolSourceSites
                                               let address = (from a in db.Addresses
                                                              let muni = (from cl in db.TVItemLanguages where cl.TVItemID == a.MunicipalityTVItemID && cl.Language == (int)LanguageEnum.en select cl.TVText).FirstOrDefault<string>()
                                                              let add = a.StreetNumber + " " + a.StreetName + " --- " + muni
                                                              where a.AddressTVItemID == t.TVItemID
                                                              select new { add }).FirstOrDefault()
                                               let pso = (from pso in db.PolSourceObservations
                                                          where pso.PolSourceSiteID == pss.PolSourceSiteID
                                                          orderby pso.ObservationDate_Local descending
                                                          select new { pso }).FirstOrDefault()
                                               let psi = (from psi in db.PolSourceObservationIssues
                                                          where psi.PolSourceObservationID == pso.pso.PolSourceObservationID
                                                          orderby psi.Ordinal ascending
                                                          select new { psi }).ToList()
                                               where t.TVItemID == tl.TVItemID
                                               && mi.TVItemID == t.TVItemID
                                               && mip.MapInfoID == mi.MapInfoID
                                               && t.TVItemID == pss.PolSourceSiteTVItemID
                                               && tl.Language == (int)LanguageEnum.en
                                               && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                               && t.TVType == (int)TVTypeEnum.PolSourceSite
                                               && t.IsActive == Active
                                               && mi.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Point
                                               && mi.TVType == (int)TVTypeEnum.PolSourceSite
                                               orderby tl.TVText
                                               select new { t, tl, mip, address.add, pss, pso, psi }).ToList();

                int TotalCount2 = tvItemSSList.Count;
                int Count2 = 0;
                foreach (var tvItemSS in tvItemSSList)
                {
                    if (Count2 % 20 == 0)
                    {
                        _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)((100.0f * ((float)count - 0.5f + ((float)Count2 / (float)TotalCount2) - 0.5f)) / (float)TotalCount));

                        NotUsed = string.Format(TaskRunnerServiceRes.Creating_, BasePath + @"PollutionSourceSites_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv --- doing " + tvItemSS.tl.TVText + "");
                        TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", BasePath + @"PollutionSourceSites_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv --- doing " + tvItemSS.tl.TVText + "");

                        _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);
                    }

                    Count2 += 1;

                    foreach (var polSourceSite in PollutionSourceSiteList.Where(c => c.t.ParentID == tvItemSS.t.TVItemID))
                    {
                        string SS = tvItemSS.tl.TVText.Replace(",", "_");
                        string PSS = "P" + (polSourceSite.pss != null && polSourceSite.pss.Site != null ? polSourceSite.pss.Site.ToString().Replace(",", "_") : "");
                        string OBSDate = (polSourceSite.pso != null && polSourceSite.pso.pso.ObservationDate_Local != null ? polSourceSite.pso.pso.ObservationDate_Local.ToString("yyyy-MM-dd") : "");
                        string PSTVT = polSourceSite.tl.TVText;
                        string[] PSArr = PSTVT.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToArray();
                        string PSType = "";
                        string PSSubtype = "";
                        string PSRisk = "";
                        if (PSArr.Length > 0)
                        {
                            PSType = PSArr[0];
                            if (PSType.Contains(" - "))
                            {
                                PSType = PSType.Substring(PSType.IndexOf(" - ") + 3);
                            }
                        }
                        if (PSArr.Length > 1)
                        {
                            PSSubtype = PSArr[1];
                        }
                        if (PSArr.Length > 2)
                        {
                            PSRisk = PSArr[2];
                        }
                        string Lat = (polSourceSite.mip != null ? polSourceSite.mip.Lat.ToString("F5") : "");
                        string Lng = (polSourceSite.mip != null ? polSourceSite.mip.Lng.ToString("F5") : "");
                        //string AD = (!string.IsNullOrWhiteSpace(polSourceSite.add) ? polSourceSite.add.Replace(",", "_") : "");
                        string URL = @"http://wmon01dtchlebl2/csspwebtools/en-CA/#!View/a|||" + polSourceSite.t.TVItemID.ToString() + @"|||30010100004000000000000000000000";

                        string TVText = "";

                        foreach (var psi in polSourceSite.psi)
                        {
                            if (psi != null && psi.psi != null)
                            {
                                List<string> ObservationInfoList = (string.IsNullOrWhiteSpace(psi.psi.ObservationInfo) ? new List<string>() : psi.psi.ObservationInfo.Trim().Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList());

                                for (int i = 0, countObs = ObservationInfoList.Count; i < countObs; i++)
                                {
                                    string Temp = _BaseEnumService.GetEnumText_PolSourceObsInfoReportEnum((PolSourceObsInfoEnum)int.Parse(ObservationInfoList[i]));
                                    switch (ObservationInfoList[i].Substring(0, 3))
                                    {
                                        case "101":
                                            {
                                                Temp = Temp.Replace("Source", "     Source");
                                            }
                                            break;
                                        //case "153":
                                        //    {
                                        //        Temp = Temp.Replace("Dilution Analyses", "     Dilution Analyses");
                                        //    }
                                        //    break;
                                        case "250":
                                            {
                                                Temp = Temp.Replace("Pathway", "     Pathway");
                                            }
                                            break;
                                        case "900":
                                            {
                                                Temp = Temp.Replace("Status", "     Status");
                                            }
                                            break;
                                        case "910":
                                            {
                                                Temp = Temp.Replace("Risk", "     Risk");
                                            }
                                            break;
                                        case "110":
                                        case "120":
                                        case "122":
                                        case "151":
                                        case "152":
                                        case "153":
                                        case "155":
                                        case "156":
                                        case "157":
                                        case "163":
                                        case "166":
                                        case "167":
                                        case "170":
                                        case "171":
                                        case "172":
                                        case "173":
                                        case "176":
                                        case "178":
                                        case "181":
                                        case "182":
                                        case "183":
                                        case "185":
                                        case "186":
                                        case "187":
                                        case "190":
                                        case "191":
                                        case "192":
                                        case "193":
                                        case "194":
                                        case "196":
                                        case "198":
                                        case "199":
                                        case "220":
                                        case "930":
                                            {
                                                //Temp = @"<span class=""hidden"">" + Temp + "</span>";
                                            }
                                            break;
                                        default:
                                            break;
                                    }
                                    TVText = TVText + Temp;
                                }
                            }

                            string TVT = (polSourceSite.pso != null && polSourceSite.pso.pso.Observation_ToBeDeleted != null ? polSourceSite.pso.pso.Observation_ToBeDeleted : "");

                            string TempISS = (!string.IsNullOrWhiteSpace(TVT) ? TVT.Replace(",", "_") + " ----- " : "") + TVText;

                            string ISS = TempISS.Replace("\r", "   ").Replace("\n", "").Replace("empty", "").Replace("Empty", "").Replace("\r", "   ").Replace("\n", "");

                            if (ISS.Length > 4000)
                            {
                                int lsief = 354;
                            }

                            //string C0 = " ";
                            //string C250 = " ";
                            //string C500 = " ";
                            //string C750 = " ";
                            //string C1000 = " ";
                            //string C1250 = " ";
                            //string C1500 = " ";
                            //string C1750 = " ";
                            //string C2000 = " ";
                            //string C2250 = " ";
                            //string C2500 = " ";
                            //string C2750 = " ";
                            //string C3000 = " ";
                            //string C3250 = " ";
                            //string C3500 = " ";
                            //string C3750 = " ";
                            //if (ISS.Length > 0)
                            //{
                            //    if (ISS.Length < 250)
                            //    {
                            //        C0 = ISS.Substring(0);
                            //    }
                            //    else
                            //    {
                            //        C0 = ISS.Substring(0, 250);
                            //    }
                            //}
                            //if (ISS.Length > 250)
                            //{
                            //    if (ISS.Length < 500)
                            //    {
                            //        C250 = ISS.Substring(250);
                            //    }
                            //    else
                            //    {
                            //        C250 = ISS.Substring(250, 250);
                            //    }
                            //}
                            //if (ISS.Length > 500)
                            //{
                            //    if (ISS.Length < 750)
                            //    {
                            //        C500 = ISS.Substring(500);
                            //    }
                            //    else
                            //    {
                            //        C500 = ISS.Substring(500, 250);
                            //    }
                            //}
                            //if (ISS.Length > 750)
                            //{
                            //    if (ISS.Length < 1000)
                            //    {
                            //        C750 = ISS.Substring(750);
                            //    }
                            //    else
                            //    {
                            //        C750 = ISS.Substring(750, 250);
                            //    }
                            //}
                            //if (ISS.Length > 1000)
                            //{
                            //    if (ISS.Length < 1250)
                            //    {
                            //        C1000 = ISS.Substring(1000);
                            //    }
                            //    else
                            //    {
                            //        C1000 = ISS.Substring(1000, 250);
                            //    }
                            //}
                            //if (ISS.Length > 1250)
                            //{
                            //    if (ISS.Length < 1500)
                            //    {
                            //        C1250 = ISS.Substring(1250);
                            //    }
                            //    else
                            //    {
                            //        C1250 = ISS.Substring(1250, 250);
                            //    }
                            //}
                            //if (ISS.Length > 1500)
                            //{
                            //    if (ISS.Length < 1750)
                            //    {
                            //        C1500 = ISS.Substring(1500);
                            //    }
                            //    else
                            //    {
                            //        C1500 = ISS.Substring(1500, 250);
                            //    }
                            //}
                            //if (ISS.Length > 1750)
                            //{
                            //    if (ISS.Length < 2000)
                            //    {
                            //        C1750 = ISS.Substring(1750);
                            //    }
                            //    else
                            //    {
                            //        C1750 = ISS.Substring(1750, 250);
                            //    }
                            //}
                            //if (ISS.Length > 2000)
                            //{
                            //    if (ISS.Length < 2250)
                            //    {
                            //        C2000 = ISS.Substring(2000);
                            //    }
                            //    else
                            //    {
                            //        C2000 = ISS.Substring(2000, 250);
                            //    }
                            //}
                            //if (ISS.Length > 2250)
                            //{
                            //    if (ISS.Length < 2500)
                            //    {
                            //        C2250 = ISS.Substring(2250);
                            //    }
                            //    else
                            //    {
                            //        C2250 = ISS.Substring(2250, 250);
                            //    }
                            //}
                            //if (ISS.Length > 2500)
                            //{
                            //    if (ISS.Length < 2750)
                            //    {
                            //        C2500 = ISS.Substring(2500);
                            //    }
                            //    else
                            //    {
                            //        C2500 = ISS.Substring(2500, 250);
                            //    }
                            //}
                            //if (ISS.Length > 2750)
                            //{
                            //    if (ISS.Length < 3000)
                            //    {
                            //        C2750 = ISS.Substring(2750);
                            //    }
                            //    else
                            //    {
                            //        C2750 = ISS.Substring(2750, 250);
                            //    }
                            //}
                            //if (ISS.Length > 3000)
                            //{
                            //    if (ISS.Length < 3250)
                            //    {
                            //        C3000 = ISS.Substring(3000);
                            //    }
                            //    else
                            //    {
                            //        C3000 = ISS.Substring(3000, 350);
                            //    }
                            //}
                            //if (ISS.Length > 3250)
                            //{
                            //    if (ISS.Length < 3500)
                            //    {
                            //        C3250 = ISS.Substring(3250);
                            //    }
                            //    else
                            //    {
                            //        C3250 = ISS.Substring(3250, 350);
                            //    }
                            //}
                            //if (ISS.Length > 3500)
                            //{
                            //    if (ISS.Length < 3750)
                            //    {
                            //        C3500 = ISS.Substring(3500);
                            //    }
                            //    else
                            //    {
                            //        C3500 = ISS.Substring(3500, 350);
                            //    }
                            //}
                            //if (ISS.Length > 3750)
                            //{
                            //    if (ISS.Length < 3000)
                            //    {
                            //        C3750 = ISS.Substring(3750);
                            //    }
                            //    else
                            //    {
                            //        C3750 = ISS.Substring(3750, 350);
                            //    }
                            //}
                            //sb.AppendLine($"{SS},{PSS},{PSType},{PSSubtype},{PSRisk},{Lat},{Lng},{OBS},{AD},{URL},{C0},{C250},{C500},{C750},{C1000},{C1250},{C1500},{C1750},{C2000},{C2250},{C2500},{C2750},{C3000},{C3250},{C3500},{C3750}");
                            if (SS.Length == 0)
                            {
                                SS = " ";
                            }
                            if (PSS.Length == 0)
                            {
                                PSS = " ";
                            }
                            if (PSType.Length == 0)
                            {
                                PSType = " ";
                            }
                            if (PSSubtype.Length == 0)
                            {
                                PSSubtype = " ";
                            }
                            if (Lat.Length == 0)
                            {
                                Lat = " ";
                            }
                            if (Lng.Length == 0)
                            {
                                Lng = " ";
                            }
                            if (OBSDate.Length == 0)
                            {
                                OBSDate = " ";
                            }
                            //if (AD.Length == 0)
                            //{
                            //    AD = " ";
                            //}
                            if (URL.Length == 0)
                            {
                                URL = " ";
                            }
                            //sb.AppendLine($"{SS},{PSS},{PSType},{PSSubtype},{PSRisk},{Lat},{Lng},{OBS},{AD},{URL}");
                            sb.AppendLine($"{SS},{PSS},{PSType},{PSSubtype},{PSRisk},{Lat},{Lng},{OBSDate},{URL}");
                        }
                    }
                }
            }

            FileInfo fi = new FileInfo(BasePath + @"PollutionSourceSites_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv");

            if (fi.Exists)
            {
                try
                {
                    fi.Delete();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDeleteFile_Error_, fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDeleteFile_Error_", fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    return;
                }
            }

            fi = new FileInfo(BasePath + @"PollutionSourceSites_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv");

            try
            {
                File.WriteAllText(fi.FullName, sb.ToString(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateFile_Error_, fi.FullName, ex.Message);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateFile_Error_", fi.FullName, ex.Message);
                return;
            }
        }
        private void CreateWWTPsDocumentForArcGIS(int ProvinceTVItemID, int count, int TotalCount, bool Active)
        {
            string BasePath = "";
            string ProvInit = "";
            List<string> ProvInitList = new List<string>()
            {
                "BC", "ME", "NB", "NL", "NS", "PE", "QC",
            };
            List<string> ProvList = new List<string>()
            {
                "British Columbia", "Maine", "New Brunswick", "Newfoundland and Labrador", "Nova Scotia", "Prince Edward Island", "Québec",
            };
            string NotUsed = "";
            StringBuilder sb = new StringBuilder();

            sb.AppendLine("Municipality,WWTPName,Lat,Lng,OutLat,OutLng,CanOverflow,AlarmSystemType,FacilityType,AerationType,DisinfectionType,CollectionSystemType,AverageFlowInM3PerD,PeakFlowInM3PerD,Address,CSSPUrl,Comment");

            using (CSSPDBEntities db = new CSSPDBEntities())
            {
                var tvItemProv = (from c in db.TVItems
                                  from cl in db.TVItemLanguages
                                  where c.TVItemID == cl.TVItemID
                                  && c.TVItemID == ProvinceTVItemID
                                  && cl.Language == (int)LanguageEnum.en
                                  && c.TVType == (int)TVTypeEnum.Province
                                  select new { c, cl }).FirstOrDefault();

                if (tvItemProv == null)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, ProvinceTVItemID.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, ProvinceTVItemID.ToString());
                    return;
                }

                for (int i = 0, countProv = ProvList.Count; i < countProv; i++)
                {
                    if (ProvList[i] == tvItemProv.cl.TVText)
                    {
                        ProvInit = ProvInitList[i];
                        break;
                    }
                }

                switch (ProvInit)
                {
                    case "NB":
                    case "NL":
                    case "NS":
                    case "PE":
                        {
                            BasePath = BasePathForExportToArcGIS.Replace("CSSP_NAT", "CSSP_ATL");
                        }
                        break;
                    case "BC":
                        {
                            BasePath = BasePathForExportToArcGIS.Replace("CSSP_NAT", "CSSP_NAT"); // will have to change to "CSSP_PYR"
                        }
                        break;
                    case "QC":
                        {
                            BasePath = BasePathForExportToArcGIS.Replace("CSSP_NAT", "CSSP_QC");
                        }
                        break;
                    default:
                        break;
                }

                NotUsed = string.Format(TaskRunnerServiceRes.Creating_, BasePath + @"WWTPs_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv");
                List<TextLanguage> TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", BasePath + @"WWTPs_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv");

                _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);

                var tvItemMuniList = (from t in db.TVItems
                                      from tl in db.TVItemLanguages
                                      where t.TVItemID == tl.TVItemID
                                      && tl.Language == (int)LanguageEnum.en
                                      && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                      && t.TVType == (int)TVTypeEnum.Municipality
                                      && t.IsActive == Active
                                      orderby tl.TVText
                                      select new { t, tl }).ToList();

                var WWTPList = (from c in db.TVItems
                                from cl in db.TVItemLanguages
                                let inf = (from a in db.Infrastructures
                                           where a.InfrastructureTVItemID == c.TVItemID
                                           && a.InfrastructureType == (int)InfrastructureTypeEnum.WWTP
                                           select a).FirstOrDefault()
                                let infLang = (from b in db.InfrastructureLanguages
                                               where b.InfrastructureID == inf.InfrastructureID
                                               && b.Language == (int)LanguageEnum.en
                                               select b).FirstOrDefault()
                                let latLng = (from mi in db.MapInfos
                                              from mip in db.MapInfoPoints
                                              where mi.TVItemID == c.TVItemID
                                              && mi.MapInfoID == mip.MapInfoID
                                              && mi.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Point
                                              && mi.TVType == (int)TVTypeEnum.WasteWaterTreatmentPlant
                                              select mip).FirstOrDefault()
                                let latLngOut = (from mi in db.MapInfos
                                                 from mip in db.MapInfoPoints
                                                 where mi.TVItemID == c.TVItemID
                                                 && mi.MapInfoID == mip.MapInfoID
                                                 && mi.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Point
                                                 && mi.TVType == (int)TVTypeEnum.Outfall
                                                 select mip).FirstOrDefault()
                                let address = (from a in db.Addresses
                                               let muni = (from cl in db.TVItemLanguages where cl.TVItemID == a.MunicipalityTVItemID && cl.Language == (int)LanguageEnum.en select cl.TVText).FirstOrDefault<string>()
                                               let add = a.StreetNumber + " " + a.StreetName
                                               where a.AddressTVItemID == inf.CivicAddressTVItemID
                                               select new { add }).FirstOrDefault()
                                where c.TVItemID == cl.TVItemID
                                && c.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                && c.TVType == (int)TVTypeEnum.Infrastructure
                                && c.IsActive == Active
                                && cl.Language == (int)LanguageEnum.en
                                && inf.InfrastructureType == (int)InfrastructureTypeEnum.WWTP
                                orderby cl.TVText
                                select new { c, cl, inf, infLang, latLng, latLngOut, address.add }).ToList();


                int TotalCount2 = tvItemMuniList.Count;
                int Count2 = 0;
                foreach (var tvItemMuni in tvItemMuniList)
                {
                    if (Count2 % 20 == 0)
                    {
                        _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)((100.0f * ((float)count - 0.5f + ((float)Count2 / (float)TotalCount2) - 0.5f)) / (float)TotalCount));

                        NotUsed = string.Format(TaskRunnerServiceRes.Creating_, BasePath + @"WWTPs_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv --- doing " + tvItemMuni.tl.TVText + "");
                        TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", BasePath + @"WWTPs_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv --- doing " + tvItemMuni.tl.TVText + "");

                        _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);
                    }

                    Count2 += 1;

                    foreach (var wwtp in WWTPList.Where(c => c.c.ParentID == tvItemMuni.t.TVItemID))
                    {
                        string MN = tvItemMuni.tl.TVText.Replace(",", "_");
                        string WN = wwtp.cl.TVText.Replace(",", "_");
                        string Lat = (wwtp.latLng != null ? wwtp.latLng.Lat.ToString("F5") : "");
                        string Lng = (wwtp.latLng != null ? wwtp.latLng.Lng.ToString("F5") : "");
                        string LatOut = (wwtp.latLngOut != null ? wwtp.latLngOut.Lat.ToString("F5") : "");
                        string LngOut = (wwtp.latLngOut != null ? wwtp.latLngOut.Lng.ToString("F5") : "");
                        string CO = (wwtp.inf.CanOverflow != null ? ((bool)wwtp.inf.CanOverflow).ToString() : "");
                        string AS = (wwtp.inf.AlarmSystemType != null && wwtp.inf.AlarmSystemType > 0 ? (_BaseEnumService.GetEnumText_AlarmSystemTypeEnum((AlarmSystemTypeEnum)wwtp.inf.AlarmSystemType)) : "");
                        string FT = (wwtp.inf.FacilityType != null && wwtp.inf.FacilityType > 0 ? (_BaseEnumService.GetEnumText_FacilityTypeEnum((FacilityTypeEnum)wwtp.inf.FacilityType)) : "");
                        string AT = (wwtp.inf.AerationType != null && wwtp.inf.AerationType > 0 ? (_BaseEnumService.GetEnumText_AerationTypeEnum((AerationTypeEnum)wwtp.inf.AerationType)) : "");
                        string DT = (wwtp.inf.DisinfectionType != null && wwtp.inf.DisinfectionType > 0 ? (_BaseEnumService.GetEnumText_DisinfectionTypeEnum((DisinfectionTypeEnum)wwtp.inf.DisinfectionType)) : "");
                        string CST = (wwtp.inf.CollectionSystemType != null && wwtp.inf.CollectionSystemType > 0 ? (_BaseEnumService.GetEnumText_CollectionSystemTypeEnum((CollectionSystemTypeEnum)wwtp.inf.CollectionSystemType)) : "");
                        string AF = (wwtp.inf.AverageFlow_m3_day != null ? ((double)wwtp.inf.AverageFlow_m3_day).ToString("F3") : "");
                        string PF = (wwtp.inf.PeakFlow_m3_day != null ? ((double)wwtp.inf.PeakFlow_m3_day).ToString("F3") : "");
                        string AD = (!string.IsNullOrWhiteSpace(wwtp.add) ? wwtp.add.Replace(",", "_") + " --- " + MN : "");
                        string URL = @"http://wmon01dtchlebl2/csspwebtools/en-CA/#!View/a|||" + tvItemMuni.t.TVItemID.ToString() + @"|||30010100004000000000000000000000";
                        string C = (wwtp.infLang != null && !string.IsNullOrWhiteSpace(wwtp.infLang.Comment) ? wwtp.infLang.Comment.Replace(",", "_") : "");
                        sb.AppendLine($"{MN},{WN},{Lat},{Lng},{LatOut},{LngOut},{CO},{AS},{FT},{AT},{DT},{CST},{AF},{PF},{AD},{URL},{C}".Replace("\r", "   ").Replace("\n", "").Replace("empty", "").Replace("Empty", ""));
                    }
                }
            }

            FileInfo fi = new FileInfo(BasePath + @"WWTPs_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv");

            if (fi.Exists)
            {
                try
                {
                    fi.Delete();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDeleteFile_Error_, fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDeleteFile_Error_", fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    return;
                }
            }

            fi = new FileInfo(BasePath + @"WWTPs_" + ProvInit + "_" + (Active ? "Active" : "Inactive") + ".csv");

            try
            {
                File.WriteAllText(fi.FullName, sb.ToString(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateFile_Error_, fi.FullName, ex.Message);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateFile_Error_", fi.FullName, ex.Message);
                return;
            }
        }

        public void ExportEmailDistributionLists()
        {
            string NotUsed = "";

            string Parameters = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters;
            string[] ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            int CountryTVItemID = 0;
            foreach (string s in ParamValueList)
            {
                string[] ParamValue = s.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                if (ParamValue.Length != 2)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, TaskRunnerServiceRes.Parameters);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", TaskRunnerServiceRes.Parameters);
                    return;
                }

                if (ParamValue[0] == "CountryTVItemID")
                {
                    CountryTVItemID = int.Parse(ParamValue[1]);
                }
                else
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind__, TaskRunnerServiceRes.Parameter, ParamValue[0]);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind__", TaskRunnerServiceRes.Parameter, ParamValue[0].ToString());
                    return;
                }
            }

            if (CountryTVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.CountryTVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.CountryTVItemID);
                return;
            }

            TVItemModel tvItemModelCountry = _TVItemService.GetTVItemModelWithTVItemIDDB(CountryTVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelCountry.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.CountryTVItemID, CountryTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.CountryTVItemID, CountryTVItemID.ToString());
                return;
            }

            TVItemModel tvItemModelRoot = _TVItemService.GetRootTVItemModelDB();
            if (!string.IsNullOrWhiteSpace(tvItemModelRoot.Error))
            {
                NotUsed = TaskRunnerServiceRes.CouldNotFindTVItemRoot;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("CouldNotFindTVItemRoot");
                return;
            }

            string ServerTemplateFilePath = _TVFileService.GetServerFilePath(tvItemModelRoot.TVItemID);

            DirectoryInfo di = new DirectoryInfo(ServerTemplateFilePath);
            if (!di.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.Directory_DoesNotExist, di.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Directory_DoesNotExist", di.FullName);
                return;
            }

            FileInfo fiTemplate = new FileInfo(ServerTemplateFilePath + @"Template_EmailDistributionList.xlsx");

            if (!fiTemplate.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.File_DoesNotExist, fiTemplate.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("File_DoesNotExist", fiTemplate.FullName);
                return;
            }
            DateTime CD = DateTime.Now;

            string DateText = "_" + CD.Year.ToString() +
                "_" + (CD.Month > 9 ? CD.Month.ToString() : "0" + CD.Month.ToString()) +
                "_" + (CD.Day > 9 ? CD.Day.ToString() : "0" + CD.Day.ToString()) +
                "_" + (CD.Hour > 9 ? CD.Hour.ToString() : "0" + CD.Hour.ToString()) +
                "_" + (CD.Minute > 9 ? CD.Minute.ToString() : "0" + CD.Minute.ToString());


            string ServerNewFilePath = _TVFileService.GetServerFilePath(CountryTVItemID);

            di = new DirectoryInfo(ServerNewFilePath);
            if (!di.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.Directory_DoesNotExist, di.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Directory_DoesNotExist", di.FullName);
                return;
            }

            FileInfo fi = new FileInfo(ServerNewFilePath + fiTemplate.Name.Replace("Template_", "").Replace(".xlsx", DateText + ".xlsx"));

            File.Copy(fiTemplate.FullName, fi.FullName, true);
            if (!fi.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.File_DoesNotExist, fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("File_DoesNotExist", fi.FullName);
                return;
            }

            Microsoft.Office.Interop.Excel.Application appExcel = new Microsoft.Office.Interop.Excel.Application();
            appExcel.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = appExcel.Workbooks.Open(fi.FullName);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];

            int currentRow = 5;

            List<EmailDistributionListModel> emailDistributionListModelList = _EmailDistributionListService.GetEmailDistributionListModelWithCountryTVItemIDDB(CountryTVItemID);

            foreach (EmailDistributionListModel emailDistributionListModel in emailDistributionListModelList)
            {

                xlWorksheet.Cells[currentRow, 1] = emailDistributionListModel.RegionName;

                List<EmailDistributionListContactModel> emailDistributionListContactModelList = _EmailDistributionListContactService.GetEmailDistributionListContactModelListWithEmailDistributionListIDDB(emailDistributionListModel.EmailDistributionListID);

                foreach (EmailDistributionListContactModel emailDistributionListContactModel in emailDistributionListContactModelList)
                {
                    xlWorksheet.Cells[currentRow, 2] = (emailDistributionListContactModel.IsCC == true ? "CC" : "To");
                    xlWorksheet.Cells[currentRow, 3] = emailDistributionListContactModel.Agency;
                    xlWorksheet.Cells[currentRow, 4] = emailDistributionListContactModel.Name;
                    xlWorksheet.Cells[currentRow, 5] = emailDistributionListContactModel.Email;
                    xlWorksheet.Cells[currentRow, 6] = (emailDistributionListContactModel.CMPRainfallSeasonal == true ? "Y" : "");
                    xlWorksheet.Cells[currentRow, 7] = (emailDistributionListContactModel.CMPWastewater == true ? "Y" : "");
                    xlWorksheet.Cells[currentRow, 8] = (emailDistributionListContactModel.EmergencyWeather == true ? "Y" : "");
                    xlWorksheet.Cells[currentRow, 9] = (emailDistributionListContactModel.EmergencyWastewater == true ? "Y" : "");
                    xlWorksheet.Cells[currentRow, 10] = (emailDistributionListContactModel.ReopeningAllTypes == true ? "Y" : "");

                    currentRow += 1;
                }

                currentRow += 2;
            }

            xlWorksheet.Cells[1, 6] = DateTime.Now;

            xlWorkbook.Save();
            xlWorkbook.Close();
            appExcel.Quit();

            fi = new FileInfo(fi.FullName);

            TVItemModel tvItemModelFile = _TVItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi.Name, TVTypeEnum.File);
            if (!string.IsNullOrEmpty(tvItemModelFile.Error))
            {
                tvItemModelFile = _TVItemService.PostAddChildTVItemDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi.Name, TVTypeEnum.File);
                if (!string.IsNullOrEmpty(tvItemModelFile.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_For_With_Equal_, TaskRunnerServiceRes.TVItem, TVTypeEnum.File.ToString(), TaskRunnerServiceRes.TVItemID + ", " + TaskRunnerServiceRes.TVText, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID + ", " + fi.Name);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat4List("CouldNotCreate_For_With_Equal_", TaskRunnerServiceRes.TVItem, TVTypeEnum.File.ToString(), TaskRunnerServiceRes.TVItemID + ", " + TaskRunnerServiceRes.TVText, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID + ", " + fi.Name);
                    return;
                }
            }

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.FileGeneratedFromTemplate, FilePurposeEnum.TemplateGenerated);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

        }
        public void CreateXlsxPDF()
        {
            string NotUsed = "";
            int TVItemID = 0;
            int TVFileTVItemID = 0;

            string Parameters = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters;
            string[] ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

            if (_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.TVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.TVItemID);
            }

            if (ParamValueList.Count() != 2)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.ParameterCount_NotEqual_, ParamValueList.Count().ToString(), "2");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("ParameterCount_NotEqual_", ParamValueList.Count().ToString(), "2");
            }

            foreach (string s in ParamValueList)
            {
                string[] ParamValue = s.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                if (ParamValue.Length != 2)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, TaskRunnerServiceRes.Parameters);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", TaskRunnerServiceRes.Parameters);
                    return;
                }

                if (ParamValue[0] == "TVItemID")
                {
                    TVItemID = int.Parse(ParamValue[1]);
                }
                else if (ParamValue[0] == "TVFileTVItemID")
                {
                    TVFileTVItemID = int.Parse(ParamValue[1]);
                }
                else
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, ParamValue[0]);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", ParamValue[0].ToString());
                    return;
                }
            }

            TVItemModel tvItemModelParent = _TVItemService.GetTVItemModelWithTVItemIDDB(TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelParent.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                return;
            }

            TVItemModel tvItemModelFile = _TVItemService.GetTVItemModelWithTVItemIDDB(TVFileTVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelFile.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVFileTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVFileTVItemID.ToString());
                return;
            }

            TVFileModel tvFileModel = _TVFileService.GetTVFileModelWithTVFileTVItemIDDB(TVFileTVItemID);
            if (!string.IsNullOrWhiteSpace(tvFileModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVFileTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVFileTVItemID.ToString());
                return;
            }

            FileInfo fiXlsx = new FileInfo(tvFileModel.ServerFilePath + tvFileModel.ServerFileName);
            if (!fiXlsx.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiXlsx.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiXlsx.FullName);
                return;
            }

            string NewPDFFileName = fiXlsx.FullName.Replace(".xlsx", "_xlsx.pdf");
            FileInfo fiPDF = new FileInfo(NewPDFFileName);
            try
            {
                // Doing PDF
                Microsoft.Office.Interop.Excel.Application _Excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook _Workbook = _Excel.Workbooks.Open(fiXlsx.FullName);
                _Workbook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, NewPDFFileName);
                _Workbook.Close();

                _Excel.Quit();
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotExportFile_ToPDFError_, fiXlsx.FullName, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotExportFile_ToPDFError_", fiXlsx.FullName, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                return;
            }

            TVItemModel tvItemModelFilePDF = _TaskRunnerBaseService.CreateFileTVItem(fiPDF);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fiPDF, tvItemModelFilePDF, tvFileModel.FileDescription, tvFileModel.FilePurpose);
        }
        public void GeneratePDF(FileInfo fi, TVTypeEnum tvType)
        {
            // Doing PDF
            Microsoft.Office.Interop.Excel.Application _Excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook _Workbook = _Excel.Workbooks.Open(fi.FullName);
            string NewPDFFileName = fi.FullName.Replace(".xlsx", "_xlsx.pdf");
            fi = new FileInfo(NewPDFFileName);
            _Workbook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, NewPDFFileName);
            _Workbook.Close();

            _Excel.Quit();

            TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.RootFileAutoGenerate, FilePurposeEnum.TemplateGenerated);
        }
        public string ParseXlsx(FileInfo fi)
        {
            return "Need to finalize ParseXlsx";
        }
        #endregion Function public

        #region Functions private
        private int GetLastClassificationColor(MWQMSiteLatestClassificationEnum? mwqmSiteLatestClassification)
        {
            if (mwqmSiteLatestClassification == null)
            {
                return 16777215;
            }

            switch (mwqmSiteLatestClassification)
            {
                case MWQMSiteLatestClassificationEnum.Approved:
                    return 5287936;
                case MWQMSiteLatestClassificationEnum.ConditionallyApproved:
                    return 5287936;
                case MWQMSiteLatestClassificationEnum.ConditionallyRestricted:
                    return 0;
                case MWQMSiteLatestClassificationEnum.Prohibited:
                    return 0;
                case MWQMSiteLatestClassificationEnum.Restricted:
                    return 255;
                case MWQMSiteLatestClassificationEnum.Unclassified:
                    return 16777215;
                default:
                    return 16777215;
            }
        }
        private string GetLastClassificationInitial(MWQMSiteLatestClassificationEnum? mwqmSiteLatestClassification)
        {
            if (mwqmSiteLatestClassification == null)
            {
                return "";
            }

            switch (mwqmSiteLatestClassification)
            {
                case MWQMSiteLatestClassificationEnum.Approved:
                    return (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr ? "A" : "A");
                case MWQMSiteLatestClassificationEnum.ConditionallyApproved:
                    return (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr ? "CA" : "AC");
                case MWQMSiteLatestClassificationEnum.ConditionallyRestricted:
                    return (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr ? "CR" : "RC");
                case MWQMSiteLatestClassificationEnum.Prohibited:
                    return (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr ? "P" : "P");
                case MWQMSiteLatestClassificationEnum.Restricted:
                    return (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr ? "R" : "R");
                case MWQMSiteLatestClassificationEnum.Unclassified:
                    return (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr ? "" : "");
                default:
                    return "";
            }
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
        private void SetupParametersAndBasicTextOnSheet1(Microsoft.Office.Interop.Excel.Application xlApp, Microsoft.Office.Interop.Excel.Workbook wb, MWQMAnalysisReportParameterModel mwqmAnalysisReportParameterModel)
        {
            string NotUsed = "";

            BaseEnumService _BaseEnumService = new BaseEnumService(LanguageEnum.en);
            if (wb.Worksheets.Count < 1)
            {
                wb.Worksheets.Add();
            }
            if (wb.Worksheets.Count < 2)
            {
                wb.Worksheets.Add();
            }
            if (wb.Worksheets.Count < 2)
            {
                NotUsed = TaskRunnerServiceRes.ExcelCouldNotCreateThe2WorksheetsRequired;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("ExcelCouldNotCreateThe2WorksheetsRequired");
                return;
            }
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];
            if (ws == null)
            {
                NotUsed = TaskRunnerServiceRes.ExcelCouldNotBeStartedCheckOfficeInstallation;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("ExcelCouldNotBeStartedCheckOfficeInstallation");
                return;
            }

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 10);

            try
            {
                Microsoft.Office.Interop.Excel.Range range = ws.get_Range("A1:A1");
                ws.Activate();
                ws.Name = "Stat and Data";
                range = xlApp.get_Range("A1:A1");
                range.Value = "Parameters";

                range = xlApp.get_Range("A1:J1");
                range.Select();
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                range.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                range.Merge();
                xlApp.Selection.Borders().LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone;
                xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;

                List<string> textList = new List<string>() { "", "Run Date\n\nRain Day", "Run Day (0)", "0-24h (-1)", "24-48h (-2)", "48-72h (-3)",
                "72-96h (-4)", "(-5)", "(-6)", "(-7)", "(-8)", "(-9)", "(-10)", "Start Tide", "End Tide" };

                for (int i = 1; i < 15; i++)
                {
                    range = xlApp.get_Range("K" + i + ":L" + i);
                    range.Select();
                    range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight;
                    range.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    range.Merge();

                    xlApp.Selection.Borders().LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone;
                    xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    range.Value = textList[i];
                }

                ws.Columns["A:A"].ColumnWidth = 4.89;
                ws.Columns["B:B"].ColumnWidth = 2.11;
                ws.Columns["C:C"].ColumnWidth = 6.33;
                ws.Columns["D:D"].ColumnWidth = 7.33;
                ws.Columns["E:E"].ColumnWidth = 5.22;
                ws.Columns["F:F"].ColumnWidth = 5.44;
                ws.Columns["G:G"].ColumnWidth = 5.22;
                ws.Columns["H:H"].ColumnWidth = 4.78;
                ws.Columns["I:I"].ColumnWidth = 3.89;
                ws.Columns["J:J"].ColumnWidth = 5.33;
                ws.Columns["K:K"].ColumnWidth = 5.67;
                ws.Columns["L:L"].ColumnWidth = 1.22;
                ws.Rows["1:1"].RowHeight = 43;

                textList = new List<string>() { "", "", "Between", "And", "Select Full Year", "Runs", "Sal", "Short Range", "Mid Range", "Calculation" };
                for (int i = 2; i < 10; i++)
                {
                    range = xlApp.get_Range("D" + i + ":E" + i);
                    range.Select();
                    range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight;
                    range.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    range.Merge();
                    range.Value = textList[i];
                }

                textList = new List<string>() { "", "", "'" + mwqmAnalysisReportParameterModel.StartDate.ToString("yyyy MMM dd"),
                "'" + mwqmAnalysisReportParameterModel.EndDate.ToString("yyyy MMM dd"),
                "'" + mwqmAnalysisReportParameterModel.FullYear.ToString(),
                "'" + mwqmAnalysisReportParameterModel.NumberOfRuns.ToString(),
                "'" + mwqmAnalysisReportParameterModel.SalinityHighlightDeviationFromAverage.ToString(),
                "'" + mwqmAnalysisReportParameterModel.ShortRangeNumberOfDays.ToString(),
                "'" + mwqmAnalysisReportParameterModel.MidRangeNumberOfDays.ToString(),
                "'" + _BaseEnumService.GetEnumText_AnalysisCalculationTypeEnum(mwqmAnalysisReportParameterModel.AnalysisCalculationType) };
                for (int i = 2; i < 10; i++)
                {
                    range = xlApp.get_Range("F" + i + ":G" + i);
                    range.Select();
                    range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight;
                    range.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    range.Merge();
                    range.Value = textList[i];
                }

                range = xlApp.get_Range("D2:G9");
                range.Select();
                xlApp.Selection.Borders().LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                xlApp.Selection.Borders().Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                for (int i = 12; i < 14; i++)
                {
                    range = xlApp.get_Range("C" + i + ":C" + i);
                    range.Select();
                    range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight;
                    range.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    range.Merge();
                    range.Value = (i == 12 ? "Dry" : "Wet");
                }

                for (int i = 11; i < 14; i++)
                {
                    range = xlApp.get_Range("D" + i + ":D" + i);
                    range.Select();
                    range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    range.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    range.Merge();
                    range.Value = (i == 11 ? "0-24h" : (i == 12 ? mwqmAnalysisReportParameterModel.DryLimit24h.ToString() : mwqmAnalysisReportParameterModel.WetLimit24h.ToString()));
                }

                for (int i = 11; i < 14; i++)
                {
                    range = xlApp.get_Range("E" + i + ":E" + i);
                    range.Select();
                    range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    range.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    range.Merge();
                    range.Value = (i == 11 ? "0-48h" : (i == 12 ? mwqmAnalysisReportParameterModel.DryLimit48h.ToString() : mwqmAnalysisReportParameterModel.WetLimit48h.ToString()));
                }

                for (int i = 11; i < 14; i++)
                {
                    range = xlApp.get_Range("F" + i + ":F" + i);
                    range.Select();
                    range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    range.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    range.Merge();
                    range.Value = (i == 11 ? "0-72h" : (i == 12 ? mwqmAnalysisReportParameterModel.DryLimit72h.ToString() : mwqmAnalysisReportParameterModel.WetLimit72h.ToString()));
                }

                for (int i = 11; i < 14; i++)
                {
                    range = xlApp.get_Range("G" + i + ":G" + i);
                    range.Select();
                    range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    range.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    range.Merge();
                    range.Value = (i == 11 ? "0-96h" : (i == 12 ? mwqmAnalysisReportParameterModel.DryLimit96h.ToString() : mwqmAnalysisReportParameterModel.WetLimit96h.ToString()));
                }

                range = xlApp.get_Range("D11:G11");
                range.Select();
                xlApp.Selection.Borders().LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                xlApp.Selection.Borders().Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                xlApp.Selection.Font.Bold = true;

                range = xlApp.get_Range("C12:G13");
                range.Select();
                xlApp.Selection.Borders().LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                xlApp.Selection.Borders().Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                range = xlApp.get_Range("C12:C13");
                range.Select();
                xlApp.Selection.Font.Bold = true;

                textList = new List<string>() { "Site", "Samples", "Period", "Min FC", "Max FC", "GMean", "Median", "P90", "% > 43", "% > 260" };
                List<string> LetterList = new List<string>() { "A", "C", "D", "E", "F", "G", "H", "I", "J", "K" };
                for (int i = 0; i < 10; i++)
                {
                    range = xlApp.get_Range(LetterList[i] + "15:" + LetterList[i] + "15");
                    range.Select();
                    range.Value = textList[i];
                }

                range = xlApp.get_Range("A15:K15");
                range.Select();
                xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                range = xlApp.get_Range("M15:M15");
                range.Select();
                List<string> showDataTypeTextList = mwqmAnalysisReportParameterModel.ShowDataTypes.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();
                string M15Text = "      ";
                foreach (string s in showDataTypeTextList)
                {
                    M15Text = M15Text + _BaseEnumService.GetEnumText_ExcelExportShowDataTypeEnum(((ExcelExportShowDataTypeEnum)int.Parse(s))) + ", ";
                }
                range.Value = M15Text;
                xlApp.Selection.WrapText = false;
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft;

                ws.Cells.Select();
                xlApp.Selection.Font.Size = 10;

                range = xlApp.get_Range("A1:A1");
                range.Select();

                _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 15);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.ErrorWhileCreatingExcelDocument_, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ErrorWhileCreatingExcelDocument_", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                return;
            }
        }
        private void SetupParametersAndBasicTextOnSheet2(Microsoft.Office.Interop.Excel.Application xlApp, Microsoft.Office.Interop.Excel.Workbook wb, MWQMAnalysisReportParameterModel mwqmAnalysisReportParameterModel)
        {
            string NotUsed = "";
            CSSPEnumsDLL.Services.BaseEnumService _BaseEnumService = new CSSPEnumsDLL.Services.BaseEnumService(LanguageEnum.en);
            if (wb.Worksheets.Count < 2)
            {
                wb.Worksheets.Add();
            }
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[2];
            if (ws == null)
            {
                NotUsed = TaskRunnerServiceRes.ExcelCouldNotBeStartedCheckOfficeInstallation;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("ExcelCouldNotBeStartedCheckOfficeInstallation");
                return;
            }

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 5);

            try
            {
                Microsoft.Office.Interop.Excel.Range range = ws.get_Range("A1:A1");
                ws.Activate();
                ws.Name = "Help";
                range = xlApp.get_Range("A1:G1");
                range.Select();
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                range.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                range.Merge();
                xlApp.Selection.Borders().LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone;
                xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                range.Value = "Color and letter schema";

                List<string> LetterList = new List<string>()
                {
                    "F","E","D","C","B","A","F","E","D","C","B","A","F","E","D","C","B","A",
                };
                List<string> RangeList = new List<string>()
                {
                    "GM > 181.33 or Med > 181.33 or P90 > 460.0 or % > 260 > 18.33",
                    "GM > 162.67 or Med > 162.67 or P90 > 420.0 or % > 260 > 16.67",
                    "GM > 144.0 or Med > 144.0 or P90 > 380.0 or % > 260 > 15.0",
                    "GM > 125.33 or Med > 125.33 or P90 > 340.0 or % > 260 > 13.33",
                    "GM > 106.67 or Med > 106.67 or P90 > 300.0 or % > 260 > 11.67",
                    "GM > 88 or Med > 88 or P90 > 260 or % > 260 > 10",
                    "GM > 75.67 or Med > 75.67 or P90 > 223.83 or % > 43 > 26.67",
                    "GM > 63.33 or Med > 63.33 or P90 > 187.67 or % > 43 > 23.33",
                    "GM > 51.0 or Med > 51.0 or P90 > 151.5 or % > 43 > 20.0",
                    "GM > 38.67 or Med > 38.67 or P90 > 115.33 or % > 43 > 16.67",
                    "GM > 26.33 or Med > 26.33 or P90 > 79.17 or % > 43 > 13.33",
                    "GM > 14 or Med > 14 or P90 > 43 or % > 43 > 10",
                    "GM > 11.67 or Med > 11.67 or P90 > 35.83 or % > 43 > 8.33",
                    "GM > 9.33 or Med > 9.33 or P90 > 28.67 or % > 43 > 6.67",
                    "GM > 7.0 or Med > 7.0 or P90 > 21.5 or % > 43 > 5.0",
                    "GM > 4.67 or Med > 4.67 or P90 > 14.33 or % > 43 > 3.33",
                    "GM > 2.33 or Med > 2.33 or P90 > 7.17 or % > 43 > 1.67",
                    "Everything else",
                };

                List<string> BGColorList = new List<string>
                {
                    "16746632",
                    "16751001",
                    "16755370",
                    "16759739",
                    "16764108",
                    "16768477",
                    "170",
                    "204",
                    "1118718",
                    "4474111",
                    "10066431",
                    "13421823",
                    "13434828",
                    "10092441",
                    "4521796",
                    "1179409",
                    "47872",
                    "39168",
                };

                for (int i = 0, count = LetterList.Count; i < count; i++)
                {
                    range = xlApp.get_Range("A" + (i + 3).ToString() + ":A" + (i + 3).ToString());
                    range.Select();
                    range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    range.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    xlApp.Selection.Borders().LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone;
                    xlApp.Selection.Borders().LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    xlApp.Selection.Borders().Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    range.Value = LetterList[i];

                    xlApp.Selection.Interior.Color = int.Parse(BGColorList[i]);

                    range = xlApp.get_Range("B" + (i + 3).ToString() + ":B" + (i + 3).ToString());
                    range.Select();
                    range.Value = RangeList[i];
                }

                ws.Columns["A:A"].ColumnWidth = 2.11;
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.ErrorWhileCreatingExcelDocument_, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ErrorWhileCreatingExcelDocument_", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                return;
            }

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 7);
        }
        private void SetupStatOnSheet1(Microsoft.Office.Interop.Excel.Application xlApp, Microsoft.Office.Interop.Excel.Workbook wb, MWQMAnalysisReportParameterModel mwqmAnalysisReportParameterModel)
        {
            string NotUsed = "";
            int LatestYear = 0;
            List<int> RunUsedColNumberList = new List<int>();
            List<RunDateColNumber> runDateColNumberList = new List<RunDateColNumber>();
            List<RowAndType> rowAndTypeList = new List<RowAndType>();

            List<SiteRowNumber> siteRowNumberList = new List<SiteRowNumber>();
            List<ExcelExportShowDataTypeEnum> showDataTypeList = new List<ExcelExportShowDataTypeEnum>();
            List<int> MWQMRunTVItemIDToOmitList = new List<int>();

            string[] showDataTypeTextList = mwqmAnalysisReportParameterModel.ShowDataTypes.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            foreach (string s in showDataTypeTextList)
            {
                showDataTypeList.Add((ExcelExportShowDataTypeEnum)int.Parse(s));
            }

            string[] runDateTextList = mwqmAnalysisReportParameterModel.RunsToOmit.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            foreach (string s in runDateTextList)
            {
                MWQMRunTVItemIDToOmitList.Add(int.Parse(s));
            }

            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];
            if (ws == null)
            {
                NotUsed = TaskRunnerServiceRes.ExcelCouldNotBeStartedCheckOfficeInstallation;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("ExcelCouldNotBeStartedCheckOfficeInstallation");
                return;
            }

            try
            {
                Microsoft.Office.Interop.Excel.Range range = ws.get_Range("A1:A1");
                ws.Activate();

                MWQMSubsectorAnalysisModel mwqmSubsectorAnalysisModel = _MWQMSubsectorService.GetMWQMSubsectorAnalysisModel(mwqmAnalysisReportParameterModel.SubsectorTVItemID);

                foreach (MWQMSampleAnalysisModel mwqmSampleAnalysisModel in mwqmSubsectorAnalysisModel.MWQMSampleAnalysisModelList)
                {
                    mwqmSampleAnalysisModel.SampleDateTime_Local = new DateTime(mwqmSampleAnalysisModel.SampleDateTime_Local.Year, mwqmSampleAnalysisModel.SampleDateTime_Local.Month, mwqmSampleAnalysisModel.SampleDateTime_Local.Day);
                }

                foreach (MWQMRunAnalysisModel mwqmRunAnalysisModel in mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList)
                {
                    mwqmRunAnalysisModel.DateTime_Local = new DateTime(mwqmRunAnalysisModel.DateTime_Local.Year, mwqmRunAnalysisModel.DateTime_Local.Month, mwqmRunAnalysisModel.DateTime_Local.Day);
                }

                _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 20);

                int CountRun = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList.Count();
                for (int i = 0, count = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList.Count(); i < count; i++)
                {
                    if (i % 20 == 0)
                    {
                        int Percent = (int)(20.0D + (30.0D * ((double)i / (double)CountRun)));
                        _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);
                    }

                    if (i == 0)
                    {
                        LatestYear = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].DateTime_Local.Year;
                    }
                    ws.Cells[1, 13 + i] = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].DateTime_Local.ToString("yyyy\nMMM\ndd");
                    ws.Cells[2, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay0_mm == null
                        ? "--" : Convert.ToInt32(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay0_mm).ToString());
                    ws.Cells[3, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay1_mm == null
                        ? "--" : Convert.ToInt32(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay1_mm).ToString());
                    ws.Cells[4, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay2_mm == null
                        ? "--" : Convert.ToInt32(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay2_mm).ToString());
                    ws.Cells[5, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay3_mm == null
                        ? "--" : Convert.ToInt32(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay3_mm).ToString());
                    ws.Cells[6, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay4_mm == null
                        ? "--" : Convert.ToInt32(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay4_mm).ToString());
                    ws.Cells[7, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay5_mm == null
                        ? "--" : Convert.ToInt32(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay5_mm).ToString());
                    ws.Cells[8, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay6_mm == null
                        ? "--" : Convert.ToInt32(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay6_mm).ToString());
                    ws.Cells[9, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay7_mm == null
                        ? "--" : Convert.ToInt32(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay7_mm).ToString());
                    ws.Cells[10, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay8_mm == null
                        ? "--" : Convert.ToInt32(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay8_mm).ToString());
                    ws.Cells[11, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay9_mm == null
                        ? "--" : Convert.ToInt32(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay9_mm).ToString());
                    ws.Cells[12, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay10_mm == null
                        ? "--" : Convert.ToInt32(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay10_mm).ToString());
                    ws.Cells[13, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].Tide_Start == null
                        ? "--" : GetTideInitial(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].Tide_Start));
                    ws.Cells[14, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].Tide_End == null
                        ? "--" : GetTideInitial(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].Tide_End));

                    ws.Columns[13 + i].ColumnWidth = 4.33;
                    range = ws.Columns[13 + i];
                    range.Select();
                    xlApp.Selection.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;

                    runDateColNumberList.Add(new RunDateColNumber()
                    {
                        MWQMRunTVItemID = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].MWQMRunTVItemID,
                        RunDate = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].DateTime_Local,
                        ColNumber = 13 + i,
                        Used = false,
                    });

                    if (LatestYear != mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].DateTime_Local.Year)
                    {
                        ws.Columns[13 + i].Select();
                        xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).ColorIndex = 0;
                        xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        LatestYear = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].DateTime_Local.Year;
                    }
                }

                ws.Columns["A:L"].Select();
                xlApp.Selection.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;

                ws.Range["M15"].Select();
                xlApp.Selection.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft;

                int RowCount = 16;
                List<MWQMSiteAnalysisModel> mwqmSiteAnalysisModelListAll = mwqmSubsectorAnalysisModel.MWQMSiteAnalysisModelList.Where(c => c.IsActive == true).OrderBy(c => c.MWQMSiteTVText)
                                                                            .Concat(mwqmSubsectorAnalysisModel.MWQMSiteAnalysisModelList.Where(c => c.IsActive == false).OrderBy(c => c.MWQMSiteTVText)).ToList();
                int CountSite = 0;
                int CountSiteTotal = mwqmSiteAnalysisModelListAll.Count();
                foreach (MWQMSiteAnalysisModel mwqmSiteAnalysisModel in mwqmSiteAnalysisModelListAll)
                {
                    CountSite += 1;
                    if (CountSite % 10 == 0)
                    {
                        int Percent = (int)(50.0D + (50.0D * ((double)CountSite / (double)CountSiteTotal)));
                        _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);
                    }

                    double? P90 = null;
                    double? GeoMean = null;
                    double? Median = null;
                    double? PercOver43 = null;
                    double? PercOver260 = null;

                    SiteRowNumber siteRowNumber = new SiteRowNumber() { MWQMSiteTVItemID = mwqmSiteAnalysisModel.MWQMSiteTVItemID, SiteName = mwqmSiteAnalysisModel.MWQMSiteTVText, RowNumber = RowCount };
                    siteRowNumberList.Add(siteRowNumber);

                    range = ws.Cells[RowCount, 1];
                    string classification = GetLastClassificationInitial(mwqmSiteAnalysisModel.MWQMSiteLatestClassification);
                    range.Value = "'" + mwqmSiteAnalysisModel.MWQMSiteTVText;
                    range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    range.Select();
                    if (mwqmSiteAnalysisModel.IsActive == true)
                    {
                        xlApp.Selection.Font.Color = 5287936; // green
                    }
                    else
                    {
                        xlApp.Selection.Font.Color = 255; // red
                    }

                    range = ws.Cells[RowCount, 2];
                    range.Value = "'" + (string.IsNullOrWhiteSpace(classification) ? "" : classification);
                    range.Select();
                    xlApp.Selection.Interior.Color = GetLastClassificationColor(mwqmSiteAnalysisModel.MWQMSiteLatestClassification);
                    xlApp.Selection.Font.Color = (classification == "P" ? 16777215 : 0);

                    // loading all site sample and doing the stats
                    List<MWQMSampleAnalysisModel> mwqmSampleAnalysisForSiteModelList = mwqmSubsectorAnalysisModel.MWQMSampleAnalysisModelList.Where(c => c.MWQMSiteTVItemID == siteRowNumber.MWQMSiteTVItemID).OrderByDescending(c => c.SampleDateTime_Local).ToList();
                    List<MWQMSampleAnalysisModel> mwqmSampleAnalysisForSiteModelToUseList = new List<MWQMSampleAnalysisModel>();
                    foreach (MWQMSampleAnalysisModel mwqmSampleAnalysisModel in mwqmSampleAnalysisForSiteModelList)
                    {
                        if (!MWQMRunTVItemIDToOmitList.Contains(mwqmSampleAnalysisModel.MWQMRunTVItemID))
                        {
                            if ((mwqmSampleAnalysisModel.SampleDateTime_Local <= mwqmAnalysisReportParameterModel.StartDate) && (mwqmSampleAnalysisModel.SampleDateTime_Local >= mwqmAnalysisReportParameterModel.EndDate))
                            {
                                if (mwqmSampleAnalysisForSiteModelToUseList.Count < mwqmAnalysisReportParameterModel.NumberOfRuns)
                                {
                                    if (mwqmAnalysisReportParameterModel.AnalysisCalculationType == AnalysisCalculationTypeEnum.WetAllAll)
                                    {
                                        MWQMRunAnalysisModel mwqmRunAnalysisModel = (from c in mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList
                                                                                     where c.DateTime_Local.Year == mwqmSampleAnalysisModel.SampleDateTime_Local.Year
                                                                                     && c.DateTime_Local.Month == mwqmSampleAnalysisModel.SampleDateTime_Local.Month
                                                                                     && c.DateTime_Local.Day == mwqmSampleAnalysisModel.SampleDateTime_Local.Day
                                                                                     select c).FirstOrDefault();

                                        List<int> RainData = new List<int>()
                                        {
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay0_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay1_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay2_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay3_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay4_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay5_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay6_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay7_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay8_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay9_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay10_mm),
                                        };

                                        int ShortRange = Math.Abs(mwqmAnalysisReportParameterModel.ShortRangeNumberOfDays);
                                        int MidRange = Math.Abs(mwqmAnalysisReportParameterModel.MidRangeNumberOfDays);
                                        int TotalRain = 0;
                                        bool AlreadyUsed = false;
                                        for (int i = 1; i < 11; i++)
                                        {
                                            TotalRain = TotalRain + RainData[i];
                                            if (i <= ShortRange)
                                            {
                                                if (i == 1)
                                                {
                                                    if (mwqmAnalysisReportParameterModel.WetLimit24h <= TotalRain)
                                                    {
                                                        int Col = 0;
                                                        for (int j = 0, count = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList.Count(); j < count; j++)
                                                        {
                                                            if (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[j].MWQMRunID == mwqmRunAnalysisModel.MWQMRunID)
                                                            {
                                                                Col = j + 13;
                                                                break;
                                                            }
                                                        }
                                                        ws.Cells[3, Col].Select();
                                                        xlApp.Selection.Interior.Color = 16772300;

                                                        if (!AlreadyUsed)
                                                        {
                                                            mwqmSampleAnalysisForSiteModelToUseList.Add(mwqmSampleAnalysisModel);
                                                            AlreadyUsed = true;
                                                        }
                                                    }
                                                }
                                                else if (i == 2)
                                                {
                                                    if (mwqmAnalysisReportParameterModel.WetLimit48h <= TotalRain)
                                                    {
                                                        int Col = 0;
                                                        for (int j = 0, count = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList.Count(); j < count; j++)
                                                        {
                                                            if (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[j].MWQMRunID == mwqmRunAnalysisModel.MWQMRunID)
                                                            {
                                                                Col = j + 13;
                                                                break;
                                                            }
                                                        }
                                                        ws.Cells[4, Col].Select();
                                                        xlApp.Selection.Interior.Color = 16772300;

                                                        if (!AlreadyUsed)
                                                        {
                                                            mwqmSampleAnalysisForSiteModelToUseList.Add(mwqmSampleAnalysisModel);
                                                            AlreadyUsed = true;
                                                        }
                                                    }
                                                }
                                                else if (i == 3)
                                                {
                                                    if (mwqmAnalysisReportParameterModel.WetLimit72h <= TotalRain)
                                                    {
                                                        int Col = 0;
                                                        for (int j = 0, count = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList.Count(); j < count; j++)
                                                        {
                                                            if (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[j].MWQMRunID == mwqmRunAnalysisModel.MWQMRunID)
                                                            {
                                                                Col = j + 13;
                                                                break;
                                                            }
                                                        }
                                                        ws.Cells[5, Col].Select();
                                                        xlApp.Selection.Interior.Color = 16772300;

                                                        if (!AlreadyUsed)
                                                        {
                                                            mwqmSampleAnalysisForSiteModelToUseList.Add(mwqmSampleAnalysisModel);
                                                            AlreadyUsed = true;
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    if (mwqmAnalysisReportParameterModel.WetLimit96h <= TotalRain)
                                                    {
                                                        int Col = 0;
                                                        for (int j = 0, count = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList.Count(); j < count; j++)
                                                        {
                                                            if (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[j].MWQMRunID == mwqmRunAnalysisModel.MWQMRunID)
                                                            {
                                                                Col = j + 13;
                                                                break;
                                                            }
                                                        }
                                                        ws.Cells[6, Col].Select();
                                                        xlApp.Selection.Interior.Color = 16772300;

                                                        if (!AlreadyUsed)
                                                        {
                                                            mwqmSampleAnalysisForSiteModelToUseList.Add(mwqmSampleAnalysisModel);
                                                            AlreadyUsed = true;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else if (mwqmAnalysisReportParameterModel.AnalysisCalculationType == AnalysisCalculationTypeEnum.DryAllAll)
                                    {
                                        MWQMRunAnalysisModel mwqmRunAnalysisModel = (from c in mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList
                                                                                     where c.DateTime_Local.Year == mwqmSampleAnalysisModel.SampleDateTime_Local.Year
                                                                                     && c.DateTime_Local.Month == mwqmSampleAnalysisModel.SampleDateTime_Local.Month
                                                                                     && c.DateTime_Local.Day == mwqmSampleAnalysisModel.SampleDateTime_Local.Day
                                                                                     select c).FirstOrDefault();

                                        List<int> RainData = new List<int>()
                                        {
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay0_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay1_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay2_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay3_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay4_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay5_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay6_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay7_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay8_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay9_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay10_mm),
                                        };

                                        int ShortRange = Math.Abs(mwqmAnalysisReportParameterModel.ShortRangeNumberOfDays);
                                        int MidRange = Math.Abs(mwqmAnalysisReportParameterModel.MidRangeNumberOfDays);
                                        int TotalRain = 0;
                                        bool CanUsed = true;
                                        for (int i = 1; i < 11; i++)
                                        {
                                            TotalRain = TotalRain + RainData[i];
                                            if (i <= ShortRange)
                                            {
                                                if (i == 1)
                                                {
                                                    if (mwqmAnalysisReportParameterModel.DryLimit24h < TotalRain)
                                                    {
                                                        int Col = 0;
                                                        for (int j = 0, count = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList.Count(); j < count; j++)
                                                        {
                                                            if (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[j].MWQMRunID == mwqmRunAnalysisModel.MWQMRunID)
                                                            {
                                                                Col = j + 13;
                                                                break;
                                                            }
                                                        }
                                                        ws.Cells[3, Col].Select();
                                                        xlApp.Selection.Interior.Color = 10079487;

                                                        CanUsed = false;
                                                    }
                                                }
                                                else if (i == 2)
                                                {
                                                    if (mwqmAnalysisReportParameterModel.DryLimit48h < TotalRain)
                                                    {
                                                        int Col = 0;
                                                        for (int j = 0, count = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList.Count(); j < count; j++)
                                                        {
                                                            if (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[j].MWQMRunID == mwqmRunAnalysisModel.MWQMRunID)
                                                            {
                                                                Col = j + 13;
                                                                break;
                                                            }
                                                        }
                                                        ws.Cells[4, Col].Select();
                                                        xlApp.Selection.Interior.Color = 10079487;

                                                        CanUsed = false;
                                                    }
                                                }
                                                else if (i == 3)
                                                {
                                                    if (mwqmAnalysisReportParameterModel.DryLimit72h < TotalRain)
                                                    {
                                                        int Col = 0;
                                                        for (int j = 0, count = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList.Count(); j < count; j++)
                                                        {
                                                            if (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[j].MWQMRunID == mwqmRunAnalysisModel.MWQMRunID)
                                                            {
                                                                Col = j + 13;
                                                                break;
                                                            }
                                                        }
                                                        ws.Cells[5, Col].Select();
                                                        xlApp.Selection.Interior.Color = 10079487;

                                                        CanUsed = false;
                                                    }
                                                }
                                                else
                                                {
                                                    if (mwqmAnalysisReportParameterModel.DryLimit96h < TotalRain)
                                                    {
                                                        int Col = 0;
                                                        for (int j = 0, count = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList.Count(); j < count; j++)
                                                        {
                                                            if (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[j].MWQMRunID == mwqmRunAnalysisModel.MWQMRunID)
                                                            {
                                                                Col = j + 13;
                                                                break;
                                                            }
                                                        }
                                                        ws.Cells[6, Col].Select();
                                                        xlApp.Selection.Interior.Color = 10079487;

                                                        CanUsed = false;
                                                    }
                                                }
                                            }
                                        }
                                        if (CanUsed)
                                        {
                                            mwqmSampleAnalysisForSiteModelToUseList.Add(mwqmSampleAnalysisModel);
                                        }
                                    }
                                    else
                                    {
                                        mwqmSampleAnalysisForSiteModelToUseList.Add(mwqmSampleAnalysisModel);
                                    }
                                }
                            }
                        }
                    }

                    if (mwqmAnalysisReportParameterModel.AnalysisCalculationType == AnalysisCalculationTypeEnum.AllAllAll)
                    {
                        if (mwqmSampleAnalysisForSiteModelToUseList.Count > 0 && mwqmAnalysisReportParameterModel.FullYear)
                        {
                            int FirstYear = mwqmSampleAnalysisForSiteModelToUseList[0].SampleDateTime_Local.Year;
                            int LastYear = mwqmSampleAnalysisForSiteModelToUseList[mwqmSampleAnalysisForSiteModelToUseList.Count - 1].SampleDateTime_Local.Year;

                            List<MWQMSampleAnalysisModel> mwqmSampleAnalysisMore = (from c in mwqmSampleAnalysisForSiteModelList
                                                                                    where c.SampleDateTime_Local.Year == FirstYear
                                                                                    && c.MWQMSiteTVItemID == mwqmSiteAnalysisModel.MWQMSiteTVItemID
                                                                                    select c).Concat((from c in mwqmSampleAnalysisForSiteModelList
                                                                                                      where c.SampleDateTime_Local.Year == LastYear
                                                                                                      && c.MWQMSiteTVItemID == mwqmSiteAnalysisModel.MWQMSiteTVItemID
                                                                                                      select c)).ToList();

                            List<MWQMSampleAnalysisModel> mwqmSampleAnalysisMore2 = new List<MWQMSampleAnalysisModel>();
                            foreach (MWQMSampleAnalysisModel mwqmSampleAnalysisModel in mwqmSampleAnalysisMore)
                            {
                                if (!MWQMRunTVItemIDToOmitList.Contains(mwqmSampleAnalysisModel.MWQMRunTVItemID))
                                {
                                    mwqmSampleAnalysisMore2.Add(mwqmSampleAnalysisModel);
                                }
                            }

                            mwqmSampleAnalysisForSiteModelToUseList = mwqmSampleAnalysisForSiteModelToUseList.Concat(mwqmSampleAnalysisMore2).Distinct().ToList();

                            mwqmSampleAnalysisForSiteModelToUseList = mwqmSampleAnalysisForSiteModelToUseList.OrderByDescending(c => c.SampleDateTime_Local).ToList();
                        }
                    }

                    int Coloring = 0;
                    string Letter = "";
                    if (mwqmSampleAnalysisForSiteModelToUseList.Count < 10)
                    {
                        range = ws.Cells[RowCount, 3];
                        range.Value = "--";

                        range = ws.Cells[RowCount, 4];
                        range.Value = "--";

                        range = ws.Cells[RowCount, 5];
                        range.Value = "--";

                        range = ws.Cells[RowCount, 6];
                        range.Value = "--";

                        range = ws.Cells[RowCount, 7];
                        range.Value = "--";

                        range = ws.Cells[RowCount, 8];
                        range.Value = "--";

                        range = ws.Cells[RowCount, 9];
                        range.Value = "--";

                        range = ws.Cells[RowCount, 10];
                        range.Value = "--";

                        range = ws.Cells[RowCount, 11];
                        range.Value = "--";

                        Letter = mwqmSampleAnalysisForSiteModelToUseList.Count.ToString();
                        Coloring = 16764057;

                        range = ws.Cells[RowCount, 12];
                        range.Value = "'" + Letter;
                        range.Select();
                        xlApp.Selection.Interior.Color = Coloring;

                    }
                    if (mwqmSampleAnalysisForSiteModelToUseList.Count >= 4)
                    {
                        int MWQMSampleCount = mwqmSampleAnalysisForSiteModelToUseList.Count;
                        int? MaxYear = mwqmSampleAnalysisForSiteModelToUseList[0].SampleDateTime_Local.Year;
                        int? MinYear = mwqmSampleAnalysisForSiteModelToUseList[mwqmSampleAnalysisForSiteModelToUseList.Count - 1].SampleDateTime_Local.Year;
                        int? MinFC = (from c in mwqmSampleAnalysisForSiteModelToUseList select c.FecCol_MPN_100ml).Min();
                        int? MaxFC = (from c in mwqmSampleAnalysisForSiteModelToUseList select c.FecCol_MPN_100ml).Max();

                        List<double> SampleList = (from c in mwqmSampleAnalysisForSiteModelToUseList
                                                   select (c.FecCol_MPN_100ml == 1 ? 1.9D : (double)c.FecCol_MPN_100ml)).ToList<double>();

                        P90 = _TVItemService.GetP90(SampleList);
                        GeoMean = _TVItemService.GeometricMean(SampleList);
                        Median = _TVItemService.GetMedian(SampleList);
                        PercOver43 = ((((double)SampleList.Where(c => c > 43).Count()) / (double)SampleList.Count()) * 100.0D);
                        PercOver260 = ((((double)SampleList.Where(c => c > 260).Count()) / (double)SampleList.Count()) * 100.0D);

                        int P90Int = (int)Math.Round((double)P90, 0);
                        int GeoMeanInt = (int)Math.Round((double)GeoMean, 0);
                        int MedianInt = (int)Math.Round((double)Median, 0);
                        int PercOver43Int = (int)Math.Round((double)PercOver43, 0);
                        int PercOver260Int = (int)Math.Round((double)PercOver260, 0);

                        if ((GeoMeanInt > 88) || (MedianInt > 88) || (P90Int > 260) || (PercOver260Int > 10))
                        {
                            if ((GeoMeanInt > 181) || (MedianInt > 181) || (P90Int > 460) || (PercOver260Int > 18))
                            {
                                Coloring = 16746632;
                                Letter = "F";
                            }
                            else if ((GeoMeanInt > 163) || (MedianInt > 163) || (P90Int > 420) || (PercOver260Int > 17))
                            {
                                Coloring = 16751001;
                                Letter = "E";
                            }
                            else if ((GeoMeanInt > 144) || (MedianInt > 144) || (P90Int > 380) || (PercOver260Int > 15))
                            {
                                Coloring = 16755370;
                                Letter = "D";
                            }
                            else if ((GeoMeanInt > 125) || (MedianInt > 125) || (P90Int > 340) || (PercOver260Int > 13))
                            {
                                Coloring = 16759739;
                                Letter = "C";
                            }
                            else if ((GeoMeanInt > 107) || (MedianInt > 107) || (P90Int > 300) || (PercOver260Int > 12))
                            {
                                Coloring = 16764108;
                                Letter = "B";
                            }
                            else
                            {
                                Coloring = 16768477;
                                Letter = "A";
                            }
                        }
                        else if ((GeoMeanInt > 14) || (MedianInt > 14) || (P90Int > 43) || (PercOver43Int > 10))
                        {
                            if ((GeoMeanInt > 76) || (MedianInt > 76) || (P90Int > 224) || (PercOver43Int > 27))
                            {
                                Coloring = 170;
                                Letter = "F";
                            }
                            else if ((GeoMeanInt > 63) || (MedianInt > 63) || (P90Int > 188) || (PercOver43Int > 23))
                            {
                                Coloring = 204;
                                Letter = "E";
                            }
                            else if ((GeoMeanInt > 51) || (MedianInt > 51) || (P90Int > 152) || (PercOver43Int > 20))
                            {
                                Coloring = 1118718;
                                Letter = "D";
                            }
                            else if ((GeoMeanInt > 39) || (MedianInt > 39) || (P90Int > 115) || (PercOver43Int > 17))
                            {
                                Coloring = 4474111;
                                Letter = "C";
                            }
                            else if ((GeoMeanInt > 26) || (MedianInt > 26) || (P90Int > 79) || (PercOver43Int > 13))
                            {
                                Coloring = 10066431;
                                Letter = "B";
                            }
                            else
                            {
                                Coloring = 13421823;
                                Letter = "A";
                            }
                        }
                        else
                        {
                            if ((GeoMeanInt > 12) || (MedianInt > 12) || (P90Int > 36) || (PercOver43Int > 8))
                            {
                                Coloring = 13434828;
                                Letter = "F";
                            }
                            else if ((GeoMeanInt > 9) || (MedianInt > 9) || (P90Int > 29) || (PercOver43Int > 7))
                            {
                                Coloring = 10092441;
                                Letter = "E";
                            }
                            else if ((GeoMeanInt > 7) || (MedianInt > 7) || (P90Int > 22) || (PercOver43Int > 5))
                            {
                                Coloring = 4521796;
                                Letter = "D";
                            }
                            else if ((GeoMeanInt > 5) || (MedianInt > 5) || (P90Int > 14) || (PercOver43Int > 3))
                            {
                                Coloring = 1179409;
                                Letter = "C";
                            }
                            else if ((GeoMeanInt > 2) || (MedianInt > 2) || (P90Int > 7) || (PercOver43Int > 2))
                            {
                                Coloring = 1179409;
                                Letter = "B";
                            }
                            else
                            {
                                Coloring = 39168;
                                Letter = "A";
                            }
                        }

                        range = ws.Cells[RowCount, 3];
                        range.Value = "'" + (mwqmSiteAnalysisModel.IsActive ? MWQMSampleCount.ToString() : "--");

                        range = ws.Cells[RowCount, 4];
                        range.Value = "'" + (mwqmSiteAnalysisModel.IsActive ? (MaxYear != null ? (MaxYear.ToString() + "-" + MinYear.ToString()) : "--") : "--");

                        range = ws.Cells[RowCount, 5];
                        range.Value = "'" + (mwqmSiteAnalysisModel.IsActive ? (MinFC != null ? (MinFC < 2 ? "< 2" : (MinFC.ToString())) : "--") : "--");

                        range = ws.Cells[RowCount, 6];
                        range.Value = "'" + (mwqmSiteAnalysisModel.IsActive ? (MaxFC != null ? (MaxFC < 2 ? "< 2" : (MaxFC.ToString())) : "--") : "--");

                        range = ws.Cells[RowCount, 7];
                        range.Value = "'" + (mwqmSiteAnalysisModel.IsActive ? (GeoMean != null ? ((double)GeoMean < 2.0D ? "< 2" : ((double)GeoMean).ToString("F0")) : "--") : "--");
                        if (GeoMean > 14)
                        {
                            if (mwqmSiteAnalysisModel.IsActive)
                            {
                                range.Interior.Color = 65535;
                            }
                        }

                        range = ws.Cells[RowCount, 8];
                        range.Value = "'" + (mwqmSiteAnalysisModel.IsActive ? (Median != null ? ((double)Median < 2.0D ? "< 2" : ((double)Median).ToString("F0")) : "--") : "--");
                        if (Median > 14)
                        {
                            if (mwqmSiteAnalysisModel.IsActive)
                            {
                                range.Interior.Color = 65535;
                            }
                        }

                        range = ws.Cells[RowCount, 9];
                        range.Value = "'" + (mwqmSiteAnalysisModel.IsActive ? (P90 != null ? ((double)P90 < 2.0D ? "< 2" : ((double)P90).ToString("F0")) : "--") : "--");
                        if (P90 > 43)
                        {
                            if (mwqmSiteAnalysisModel.IsActive)
                            {
                                range.Interior.Color = 65535;
                            }
                        }
                        range = ws.Cells[RowCount, 10];
                        range.Value = "'" + (mwqmSiteAnalysisModel.IsActive ? (PercOver43 != null ? ((double)PercOver43).ToString("F0") : "--") : "--");
                        if (PercOver43 > 10)
                        {
                            if (mwqmSiteAnalysisModel.IsActive)
                            {
                                range.Interior.Color = 65535;
                            }
                        }

                        range = ws.Cells[RowCount, 11];
                        range.Value = "'" + (mwqmSiteAnalysisModel.IsActive ? (PercOver260 != null ? ((double)PercOver260).ToString("F0") : "--") : "--");
                        if (PercOver260 > 10)
                        {
                            if (mwqmSiteAnalysisModel.IsActive)
                            {
                                range.Interior.Color = 65535;
                            }
                        }

                        if (mwqmSiteAnalysisModel.IsActive)
                        {
                            range = ws.Cells[RowCount, 12];
                            range.Value = "'" + Letter;
                            range.Select();
                            xlApp.Selection.Interior.Color = Coloring;
                        }
                    }

                    int AddedRows = 0;
                    foreach (MWQMSampleAnalysisModel mwqmSampleAnalysisModel in mwqmSampleAnalysisForSiteModelList)
                    {
                        AddedRows = 0;
                        int colNumber = (from c in runDateColNumberList
                                         where c.MWQMRunTVItemID == mwqmSampleAnalysisModel.MWQMRunTVItemID
                                         select c.ColNumber).FirstOrDefault();

                        if (mwqmSampleAnalysisForSiteModelToUseList.Count >= 4)
                        {
                            List<double> SampleList = (from c in mwqmSampleAnalysisForSiteModelToUseList
                                                       select (c.FecCol_MPN_100ml == 1 ? 1.9D : (double)c.FecCol_MPN_100ml)).ToList<double>();

                            P90 = _TVItemService.GetP90(SampleList);
                            GeoMean = _TVItemService.GeometricMean(SampleList);
                            Median = _TVItemService.GetMedian(SampleList);
                            PercOver43 = ((((double)SampleList.Where(c => c > 43).Count()) / (double)SampleList.Count()) * 100.0D);
                            PercOver260 = ((((double)SampleList.Where(c => c > 260).Count()) / (double)SampleList.Count()) * 100.0D);

                        }
                        else
                        {
                            P90 = null;
                            GeoMean = null;
                            Median = null;
                            PercOver43 = null;
                            PercOver260 = null;
                        }

                        range = ws.Cells[RowCount, colNumber];

                        range.Value = (mwqmSampleAnalysisModel.FecCol_MPN_100ml < 2 ? "< 2" : mwqmSampleAnalysisModel.FecCol_MPN_100ml.ToString());
                        if (mwqmSampleAnalysisModel.FecCol_MPN_100ml > 500)
                        {
                            range.Interior.Color = 255;
                        }
                        else if (mwqmSampleAnalysisModel.FecCol_MPN_100ml > 43)
                        {
                            range.Interior.Color = 65535;
                        }
                        if (mwqmSiteAnalysisModel.IsActive == true)
                        {
                            if (mwqmSampleAnalysisForSiteModelToUseList.Contains(mwqmSampleAnalysisModel))
                            {
                                range.Select();
                                xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Color = -11489280;
                                if (!RunUsedColNumberList.Contains(colNumber))
                                {
                                    RunUsedColNumberList.Add(colNumber);
                                }
                            }
                            else
                            {
                                if (MWQMRunTVItemIDToOmitList.Contains(mwqmSampleAnalysisModel.MWQMRunTVItemID))
                                {
                                    RunUsedColNumberList.Add(colNumber);
                                }
                            }
                        }
                        if (showDataTypeList.Contains(ExcelExportShowDataTypeEnum.Temperature))
                        {
                            AddedRows += 1;
                            range = ws.Cells[RowCount + AddedRows, colNumber];
                            range.Value = (mwqmSampleAnalysisModel.WaterTemp_C != null ? ((double)mwqmSampleAnalysisModel.WaterTemp_C).ToString("F1") : "--");
                            if (!(rowAndTypeList.Where(c => c.RowNumber == (RowCount + AddedRows)).Any()))
                            {
                                rowAndTypeList.Add(new RowAndType() { RowNumber = RowCount + AddedRows, ExcelExportShowDataType = ExcelExportShowDataTypeEnum.Temperature });
                            }
                        }
                        if (showDataTypeList.Contains(ExcelExportShowDataTypeEnum.Salinity))
                        {
                            AddedRows += 1;
                            range = ws.Cells[RowCount + AddedRows, colNumber];
                            range.Value = (mwqmSampleAnalysisModel.Salinity_PPT != null ? ((double)mwqmSampleAnalysisModel.Salinity_PPT).ToString("F1") : "--");

                            if (mwqmSampleAnalysisModel.Salinity_PPT != null)
                            {
                                double? avgSal = (from c in mwqmSampleAnalysisForSiteModelList
                                                  where c.Salinity_PPT != null
                                                  select c.Salinity_PPT).Average();

                                if (Math.Abs(((double)mwqmSampleAnalysisModel.Salinity_PPT) - ((double)avgSal)) >= mwqmAnalysisReportParameterModel.SalinityHighlightDeviationFromAverage)
                                {
                                    range.Interior.Color = 65535;
                                }
                            }
                            if (!(rowAndTypeList.Where(c => c.RowNumber == (RowCount + AddedRows)).Any()))
                            {
                                rowAndTypeList.Add(new RowAndType() { RowNumber = RowCount + AddedRows, ExcelExportShowDataType = ExcelExportShowDataTypeEnum.Salinity });
                            }

                        }
                        if (showDataTypeList.Contains(ExcelExportShowDataTypeEnum.P90))
                        {
                            AddedRows += 1;
                            range = ws.Cells[RowCount + AddedRows, colNumber];

                            range.Value = (P90 != null ? ((double)P90).ToString("F1") : "--");
                            if (P90 > 43)
                            {
                                range.Interior.Color = 65535;
                            }
                            if (!(rowAndTypeList.Where(c => c.RowNumber == (RowCount + AddedRows)).Any()))
                            {
                                rowAndTypeList.Add(new RowAndType() { RowNumber = RowCount + AddedRows, ExcelExportShowDataType = ExcelExportShowDataTypeEnum.P90 });
                            }
                        }
                        if (showDataTypeList.Contains(ExcelExportShowDataTypeEnum.GemetricMean))
                        {
                            AddedRows += 1;
                            range = ws.Cells[RowCount + AddedRows, colNumber];

                            range.Value = (GeoMean != null ? ((double)GeoMean).ToString("F0") : "--");
                            if (GeoMean > 14)
                            {
                                range.Interior.Color = 65535;
                            }
                            if (!(rowAndTypeList.Where(c => c.RowNumber == (RowCount + AddedRows)).Any()))
                            {
                                rowAndTypeList.Add(new RowAndType() { RowNumber = RowCount + AddedRows, ExcelExportShowDataType = ExcelExportShowDataTypeEnum.GemetricMean });
                            }
                        }
                        if (showDataTypeList.Contains(ExcelExportShowDataTypeEnum.Median))
                        {
                            AddedRows += 1;
                            range = ws.Cells[RowCount + AddedRows, colNumber];

                            range.Value = (Median != null ? ((double)Median).ToString("F0") : "--");
                            if (Median > 14)
                            {
                                range.Interior.Color = 65535;
                            }
                            if (!(rowAndTypeList.Where(c => c.RowNumber == (RowCount + AddedRows)).Any()))
                            {
                                rowAndTypeList.Add(new RowAndType() { RowNumber = RowCount + AddedRows, ExcelExportShowDataType = ExcelExportShowDataTypeEnum.Median });
                            }
                        }
                        if (showDataTypeList.Contains(ExcelExportShowDataTypeEnum.PercOfP90Over43))
                        {
                            AddedRows += 1;
                            range = ws.Cells[RowCount + AddedRows, colNumber];

                            range.Value = (PercOver43 != null ? ((double)PercOver43).ToString("F1") : "--");
                            if (PercOver43 > 20)
                            {
                                range.Interior.Color = 65535;
                            }
                            else if (PercOver43 > 10)
                            {
                                range.Interior.Color = 65535;
                            }
                            if (!(rowAndTypeList.Where(c => c.RowNumber == (RowCount + AddedRows)).Any()))
                            {
                                rowAndTypeList.Add(new RowAndType() { RowNumber = RowCount + AddedRows, ExcelExportShowDataType = ExcelExportShowDataTypeEnum.PercOfP90Over43 });
                            }
                        }
                        if (showDataTypeList.Contains(ExcelExportShowDataTypeEnum.PercOfP90Over260))
                        {
                            AddedRows += 1;
                            range = ws.Cells[RowCount + AddedRows, colNumber];

                            range.Value = (PercOver260 != null ? ((double)PercOver260).ToString("F1") : "--");
                            if (PercOver260 > 10)
                            {
                                range.Interior.Color = 65535;
                            }
                            if (!(rowAndTypeList.Where(c => c.RowNumber == (RowCount + AddedRows)).Any()))
                            {
                                rowAndTypeList.Add(new RowAndType() { RowNumber = RowCount + AddedRows, ExcelExportShowDataType = ExcelExportShowDataTypeEnum.PercOfP90Over260 });
                            }
                        }

                        if (mwqmSampleAnalysisForSiteModelToUseList.Contains(mwqmSampleAnalysisModel))
                        {
                            mwqmSampleAnalysisForSiteModelToUseList = mwqmSampleAnalysisForSiteModelToUseList.Skip(1).ToList();
                        }

                    }

                    RowCount += (1 + AddedRows);
                }

                foreach (RunDateColNumber runDateColNumber in runDateColNumberList)
                {
                    if (MWQMRunTVItemIDToOmitList.Contains(runDateColNumber.MWQMRunTVItemID))
                    {
                        if (RunUsedColNumberList.Contains(runDateColNumber.ColNumber))
                        {
                            range = ws.Cells[1, runDateColNumber.ColNumber];
                            range.Select();
                            xlApp.Selection.Interior.Color = 255;
                        }
                    }
                    else
                    {
                        if (RunUsedColNumberList.Contains(runDateColNumber.ColNumber))
                        {
                            range = ws.Cells[1, runDateColNumber.ColNumber];
                            range.Select();
                            xlApp.Selection.Interior.Color = 5296274;
                        }
                    }
                }

                xlApp.Range[ws.Cells[Math.Abs(mwqmAnalysisReportParameterModel.ShortRangeNumberOfDays) + 2, 13], ws.Cells[Math.Abs(mwqmAnalysisReportParameterModel.ShortRangeNumberOfDays) + 2, 13 + runDateColNumberList.Count() - 1]].Select();
                xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Color = -11489280;
                xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                xlApp.Range[ws.Cells[Math.Abs(mwqmAnalysisReportParameterModel.MidRangeNumberOfDays) + 2, 13], ws.Cells[Math.Abs(mwqmAnalysisReportParameterModel.MidRangeNumberOfDays) + 2, 13 + runDateColNumberList.Count() - 1]].Select();
                xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Color = -11489280;
                xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                foreach (RowAndType rowAndType in rowAndTypeList)
                {
                    xlApp.Range["A" + rowAndType.RowNumber.ToString() + ":L" + rowAndType.RowNumber.ToString()].Select();
                    xlApp.Selection.Merge();
                    xlApp.Selection.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight;
                    switch (rowAndType.ExcelExportShowDataType)
                    {
                        case ExcelExportShowDataTypeEnum.FecalColiform:
                        case ExcelExportShowDataTypeEnum.Temperature:
                        case ExcelExportShowDataTypeEnum.Salinity:
                        case ExcelExportShowDataTypeEnum.P90:
                        case ExcelExportShowDataTypeEnum.GemetricMean:
                        case ExcelExportShowDataTypeEnum.Median:
                        case ExcelExportShowDataTypeEnum.PercOfP90Over43:
                        case ExcelExportShowDataTypeEnum.PercOfP90Over260:
                            {
                                xlApp.Selection.Value = _BaseEnumService.GetEnumText_ExcelExportShowDataTypeEnum(rowAndType.ExcelExportShowDataType);
                            }
                            break;
                        default:
                            {
                                xlApp.Selection.Value = "Error";
                            }
                            break;
                    }
                }

                xlApp.Range["M16"].Select();
                xlApp.ActiveWindow.FreezePanes = true;

                ws.Range["A1"].Select();
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.ErrorWhileCreatingExcelDocument_, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ErrorWhileCreatingExcelDocument_", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                return;
            }
        }
        #endregion Function private
    }

    #region Other Classes
    public class RunDateColNumber
    {
        public RunDateColNumber()
        {

        }
        public int MWQMRunTVItemID { get; set; }
        public DateTime RunDate { get; set; }
        public int ColNumber { get; set; }
        public bool Used { get; set; }
    }
    public class StatRunSite
    {
        public StatRunSite()
        {

        }

        public int? SampleCount { get; set; }
        public int? PeriodStart { get; set; }
        public int? PeriodEnd { get; set; }
        public int? MinFC { get; set; }
        public int? MaxFC { get; set; }
        public double? GM { get; set; }
        public double? Med { get; set; }
        public double? P90 { get; set; }
        public double? P43 { get; set; }
        public double? P260 { get; set; }
        public string Letter { get; set; }
        public int Color { get; set; }

    }
    public class SiteRowNumber
    {
        public SiteRowNumber()
        {

        }
        public int MWQMSiteTVItemID { get; set; }
        public string SiteName { get; set; }
        public int RowNumber { get; set; }
    }
    public class RowAndType
    {
        public RowAndType()
        {

        }
        public int RowNumber { get; set; }
        public ExcelExportShowDataTypeEnum ExcelExportShowDataType { get; set; }
    }
    #endregion OtherClasses
}