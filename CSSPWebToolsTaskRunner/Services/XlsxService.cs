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
    public class XlsxService
    {
        #region Variables
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

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.FileGeneratedFromTemplate, FilePurposeEnum.Generated);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;
        }
        //public void GenerateRootXLSX()
        //{
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Root, FileGeneratorTypeEnum.Excel);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    // update percentage
        //    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 5);

        //    XlsxServiceRoot _XlsxService_Root = new XlsxServiceRoot(_TaskRunnerBaseService);
        //    _XlsxService_Root.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.RootFileAutoGenerate, FilePurposeEnum.Generated);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 90);
        //}
        //public void GenerateRootXLSXAndPDF()
        //{
        //    string NotUsed = "";

        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Root, FileGeneratorTypeEnum.ExcelAndPDF);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    XlsxServiceRoot xlsxServiceRoot = new XlsxServiceRoot(_TaskRunnerBaseService);
        //    xlsxServiceRoot.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.RootFileAutoGenerate, FilePurposeEnum.Generated);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    NotUsed = TaskRunnerServiceRes.ConvertingFileToPDF;
        //    _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("ConvertingFileToPDF"));

        //    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 50);

        //    GeneratePDF(fi, TVTypeEnum.Root);
        //}
        //public void GenerateCountryXLSX()
        //{
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Country, FileGeneratorTypeEnum.Excel);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    XlsxServiceCountry xlsxServiceCountry = new XlsxServiceCountry(_TaskRunnerBaseService);
        //    xlsxServiceCountry.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.ProvinceFileAutoGenerate, FilePurposeEnum.Generated);
        //}
        //public void GenerateCountryXLSXAndPDF()
        //{
        //    string NotUsed = "";

        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Country, FileGeneratorTypeEnum.ExcelAndPDF);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    XlsxServiceCountry xlsxServiceCountry = new XlsxServiceCountry(_TaskRunnerBaseService);
        //    xlsxServiceCountry.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.ProvinceFileAutoGenerate, FilePurposeEnum.Generated);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    NotUsed = TaskRunnerServiceRes.ConvertingFileToPDF;
        //    _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("ConvertingFileToPDF"));

        //    // update percentage
        //    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 50);

        //    GeneratePDF(fi, TVTypeEnum.Country);
        //}
        //public void GenerateProvinceXLSX()
        //{
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Province, FileGeneratorTypeEnum.Excel);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    XlsxServiceProvince xlsxServiceProvince = new XlsxServiceProvince(_TaskRunnerBaseService);
        //    xlsxServiceProvince.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.ProvinceFileAutoGenerate, FilePurposeEnum.Generated);
        //}
        //public void GenerateProvinceXLSXAndPDF()
        //{
        //    string NotUsed = "";

        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Province, FileGeneratorTypeEnum.ExcelAndPDF);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    XlsxServiceProvince _XlsxService_Province = new XlsxServiceProvince(_TaskRunnerBaseService);
        //    _XlsxService_Province.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.ProvinceFileAutoGenerate, FilePurposeEnum.Generated);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    NotUsed = TaskRunnerServiceRes.ConvertingFileToPDF;
        //    _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("ConvertingFileToPDF"));

        //    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 50);

        //    GeneratePDF(fi, TVTypeEnum.Province);
        //}
        //public void GenerateAreaXLSX()
        //{
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Area, FileGeneratorTypeEnum.Excel);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    XlsxServiceArea xlsxServiceArea = new XlsxServiceArea(_TaskRunnerBaseService);
        //    xlsxServiceArea.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.AreaFileAutoGenerate, FilePurposeEnum.Generated);
        //}
        //public void GenerateAreaXLSXAndPDF()
        //{
        //    string NotUsed = "";

        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Area, FileGeneratorTypeEnum.ExcelAndPDF);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    XlsxServiceArea xlsxServiceArea = new XlsxServiceArea(_TaskRunnerBaseService);
        //    xlsxServiceArea.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.AreaFileAutoGenerate, FilePurposeEnum.Generated);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    NotUsed = TaskRunnerServiceRes.ConvertingFileToPDF;
        //    _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("ConvertingFileToPDF"));

        //    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 50);

        //    GeneratePDF(fi, TVTypeEnum.Area);
        //}
        //public void GenerateSectorXLSX()
        //{
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Sector, FileGeneratorTypeEnum.Excel);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    XlsxServiceSector xlsxServiceSector = new XlsxServiceSector(_TaskRunnerBaseService);
        //    xlsxServiceSector.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.SectorFileAutoGenerate, FilePurposeEnum.Generated);
        //}
        //public void GenerateSectorXLSXAndPDF()
        //{
        //    string NotUsed = "";

        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Sector, FileGeneratorTypeEnum.ExcelAndPDF);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    XlsxServiceSector xlsxServiceSector = new XlsxServiceSector(_TaskRunnerBaseService);
        //    xlsxServiceSector.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.SectorFileAutoGenerate, FilePurposeEnum.Generated);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    NotUsed = TaskRunnerServiceRes.ConvertingFileToPDF;
        //    _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("ConvertingFileToPDF"));

        //    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 50);

        //    GeneratePDF(fi, TVTypeEnum.Sector);
        //}
        //public void GenerateSubsectorXLSX()
        //{
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Subsector, FileGeneratorTypeEnum.Excel);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    XlsxServiceSubsector xlsxServiceSubsector = new XlsxServiceSubsector(_TaskRunnerBaseService);
        //    xlsxServiceSubsector.GenerateSubsector(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.SubsectorFileAutoGenerate, FilePurposeEnum.Generated);
        //}
        //public void GenerateSubsectorPollutionSourceFieldSheetXLSX()
        //{
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.SubsectorPollutionSourceFieldSheet, FileGeneratorTypeEnum.Excel);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    XlsxServiceSubsector xlsxServiceSubsector = new XlsxServiceSubsector(_TaskRunnerBaseService);
        //    xlsxServiceSubsector.GenerateSubsectorPollutionSourceFieldSheet(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.SubsectorFileAutoGenerate, FilePurposeEnum.Generated);
        //}
        //public void GenerateSubsectorXLSXAndPDF()
        //{
        //    string NotUsed = "";

        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Subsector, FileGeneratorTypeEnum.ExcelAndPDF);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    XlsxServiceSubsector xlsxServiceSubsector = new XlsxServiceSubsector(_TaskRunnerBaseService);
        //    xlsxServiceSubsector.GenerateSubsector(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.SubsectorFileAutoGenerate, FilePurposeEnum.Generated);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    NotUsed = TaskRunnerServiceRes.ConvertingFileToPDF;
        //    _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("ConvertingFileToPDF"));

        //    // update percentage
        //    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 50);

        //    GeneratePDF(fi, TVTypeEnum.Subsector);
        //}
        //public void GenerateMunicipalityXLSX()
        //{
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Municipality, FileGeneratorTypeEnum.Excel);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    XlsxServiceMunicipality xlsxServiceMunicipality = new XlsxServiceMunicipality(_TaskRunnerBaseService);
        //    xlsxServiceMunicipality.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.MunicipalityFileAutoGenerate, FilePurposeEnum.Generated);
        //}
        //public void GenerateMunicipalityXLSXAndPDF()
        //{
        //    string NotUsed = "";

        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Municipality, FileGeneratorTypeEnum.ExcelAndPDF);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    XlsxServiceMunicipality xlsxServiceMunicipality = new XlsxServiceMunicipality(_TaskRunnerBaseService);
        //    xlsxServiceMunicipality.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.MunicipalityFileAutoGenerate, FilePurposeEnum.Generated);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    NotUsed = TaskRunnerServiceRes.ConvertingFileToPDF;
        //    _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("ConvertingFileToPDF"));

        //    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 50);

        //    GeneratePDF(fi, TVTypeEnum.Municipality);
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
        //public void GenerateTideSiteXLSX()
        //{
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Root, FileGeneratorTypeEnum.Excel);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    // update percentage
        //    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 5);

        //    XlsxServiceRoot _XlsxService_Root = new XlsxServiceRoot(_TaskRunnerBaseService);
        //    _XlsxService_Root.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.RootFileAutoGenerate, FilePurposeEnum.Generated);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 90);
        //}

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

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.RootFileAutoGenerate, FilePurposeEnum.Generated);
        }
        public string ParseXlsx(FileInfo fi)
        {
            return "Need to finalize ParseXlsx";
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

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.FileGeneratedFromTemplate, FilePurposeEnum.Generated);
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

            string FilePath = _TVFileService.GetServerFilePath(SubsectorTVItemID);

            string dateText = DateTime.Now.ToString("_yyyy_MM_dd_HH_mm_") + (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr ? "fr" : "en");
            FileInfo fi = new FileInfo(FilePath + mwqmAnalysisReportParameterModel.AnalysisName + dateText + ".xlsx");

            TVItemModel tvItemModelTVFile = _TVItemService.PostAddChildTVItemDB(SubsectorTVItemID, mwqmAnalysisReportParameterModel.AnalysisName, TVTypeEnum.File);
            if (!string.IsNullOrWhiteSpace(tvItemModelTVFile.Error))
            {
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(tvItemModelTVFile.Error);
                return;
            }

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

            wb.Close();
            xlApp.Quit();

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

            CSSPEnumsDLL.Services.BaseEnumService _BaseEnumService = new CSSPEnumsDLL.Services.BaseEnumService(LanguageEnum.en);
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

            CSSPEnumsDLL.Services.BaseEnumService _BaseEnumService = new CSSPEnumsDLL.Services.BaseEnumService(LanguageEnum.en);
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
                        ? "--" : ((double)mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay0_mm).ToString("F0"));
                    ws.Cells[3, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay1_mm == null
                        ? "--" : ((double)mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay1_mm).ToString("F0"));
                    ws.Cells[4, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay2_mm == null
                        ? "--" : ((double)mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay2_mm).ToString("F0"));
                    ws.Cells[5, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay3_mm == null
                        ? "--" : ((double)mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay3_mm).ToString("F0"));
                    ws.Cells[6, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay4_mm == null
                        ? "--" : ((double)mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay4_mm).ToString("F0"));
                    ws.Cells[7, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay5_mm == null
                        ? "--" : ((double)mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay5_mm).ToString("F0"));
                    ws.Cells[8, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay6_mm == null
                        ? "--" : ((double)mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay6_mm).ToString("F0"));
                    ws.Cells[9, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay7_mm == null
                        ? "--" : ((double)mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay7_mm).ToString("F0"));
                    ws.Cells[10, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay8_mm == null
                        ? "--" : ((double)mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay8_mm).ToString("F0"));
                    ws.Cells[11, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay9_mm == null
                        ? "--" : ((double)mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay9_mm).ToString("F0"));
                    ws.Cells[12, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay10_mm == null
                        ? "--" : ((double)mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay10_mm).ToString("F0"));
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
                                    mwqmSampleAnalysisForSiteModelToUseList.Add(mwqmSampleAnalysisModel);
                                }
                            }
                        }
                    }

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
                    if (mwqmSampleAnalysisForSiteModelToUseList.Count >= 10)
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
                        if ((GeoMean > 88) || (Median > 88) || (P90 > 260) || (PercOver260 > 10))
                        {
                            if ((GeoMean > 181.33) || (Median > 181.33) || (P90 > 460.0) || (PercOver260 > 18.33))
                            {
                                Coloring = 16746632;
                                Letter = "F";
                            }
                            else if ((GeoMean > 162.67) || (Median > 162.67) || (P90 > 420.0) || (PercOver260 > 16.67))
                            {
                                Coloring = 16751001;
                                Letter = "E";
                            }
                            else if ((GeoMean > 144.0) || (Median > 144.0) || (P90 > 380.0) || (PercOver260 > 15.0))
                            {
                                Coloring = 16755370;
                                Letter = "D";
                            }
                            else if ((GeoMean > 125.33) || (Median > 125.33) || (P90 > 340.0) || (PercOver260 > 13.33))
                            {
                                Coloring = 16759739;
                                Letter = "C";
                            }
                            else if ((GeoMean > 106.67) || (Median > 106.67) || (P90 > 300.0) || (PercOver260 > 11.67))
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
                        else if ((GeoMean > 14) || (Median > 14) || (P90 > 43) || (PercOver43 > 10))
                        {
                            if ((GeoMean > 75.67) || (Median > 75.67) || (P90 > 223.83) || (PercOver43 > 26.67))
                            {
                                Coloring = 170;
                                Letter = "F";
                            }
                            else if ((GeoMean > 63.33) || (Median > 63.33) || (P90 > 187.67) || (PercOver43 > 23.33))
                            {
                                Coloring = 204;
                                Letter = "E";
                            }
                            else if ((GeoMean > 51.0) || (Median > 51.0) || (P90 > 151.5) || (PercOver43 > 20.0))
                            {
                                Coloring = 1118718;
                                Letter = "D";
                            }
                            else if ((GeoMean > 38.67) || (Median > 38.67) || (P90 > 115.33) || (PercOver43 > 16.67))
                            {
                                Coloring = 4474111;
                                Letter = "C";
                            }
                            else if ((GeoMean > 26.33) || (Median > 26.33) || (P90 > 79.17) || (PercOver43 > 13.33))
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
                            if ((GeoMean > 11.67) || (Median > 11.67) || (P90 > 35.83) || (PercOver43 > 8.33))
                            {
                                Coloring = 13434828;
                                Letter = "F";
                            }
                            else if ((GeoMean > 9.33) || (Median > 9.33) || (P90 > 28.67) || (PercOver43 > 6.67))
                            {
                                Coloring = 10092441;
                                Letter = "E";
                            }
                            else if ((GeoMean > 7.0) || (Median > 7.0) || (P90 > 21.5) || (PercOver43 > 5.0))
                            {
                                Coloring = 4521796;
                                Letter = "D";
                            }
                            else if ((GeoMean > 4.67) || (Median > 4.67) || (P90 > 14.33) || (PercOver43 > 3.33))
                            {
                                Coloring = 1179409;
                                Letter = "C";
                            }
                            else if ((GeoMean > 2.33) || (Median > 2.33) || (P90 > 7.17) || (PercOver43 > 1.67))
                            {
                                Coloring = 47872;
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
                            range.Interior.Color = 65535;
                        }

                        range = ws.Cells[RowCount, 8];
                        range.Value = "'" + (mwqmSiteAnalysisModel.IsActive ? (Median != null ? ((double)Median < 2.0D ? "< 2" : ((double)Median).ToString("F0")) : "--") : "--");
                        if (Median > 14)
                        {
                            range.Interior.Color = 65535;
                        }

                        range = ws.Cells[RowCount, 9];
                        range.Value = "'" + (mwqmSiteAnalysisModel.IsActive ? (P90 != null ? ((double)P90 < 2.0D ? "< 2" : ((double)P90).ToString("F0")) : "--") : "--");
                        if (P90 > 43)
                        {
                            range.Interior.Color = 65535;
                        }
                        range = ws.Cells[RowCount, 10];
                        range.Value = "'" + (mwqmSiteAnalysisModel.IsActive ? (PercOver43 != null ? ((double)PercOver43).ToString("F0") : "--") : "--");
                        if (PercOver43 > 10)
                        {
                            range.Interior.Color = 65535;
                        }

                        range = ws.Cells[RowCount, 11];
                        range.Value = "'" + (mwqmSiteAnalysisModel.IsActive ? (PercOver260 != null ? ((double)PercOver260).ToString("F0") : "--") : "--");
                        if (PercOver260 > 10)
                        {
                            range.Interior.Color = 65535;
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

                        if (mwqmSampleAnalysisForSiteModelToUseList.Count >= 10)
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
                            RowCount += 1;
                            range = ws.Cells[RowCount, colNumber];
                            range.Value = (mwqmSampleAnalysisModel.WaterTemp_C != null ? ((double)mwqmSampleAnalysisModel.WaterTemp_C).ToString("F0") : "--");
                            if (!(rowAndTypeList.Where(c => c.RowNumber == (RowCount + AddedRows)).Any()))
                            {
                                rowAndTypeList.Add(new RowAndType() { RowNumber = RowCount + AddedRows, ExcelExportShowDataType = ExcelExportShowDataTypeEnum.Temperature });
                            }
                        }
                        if (showDataTypeList.Contains(ExcelExportShowDataTypeEnum.Salinity))
                        {
                            AddedRows += 1;
                            range = ws.Cells[RowCount + AddedRows, colNumber];
                            range.Value = (mwqmSampleAnalysisModel.Salinity_PPT != null ? ((double)mwqmSampleAnalysisModel.Salinity_PPT).ToString("F0") : "--");

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

                            range.Value = (P90 != null ? ((double)P90).ToString("F0") : "--");
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

                            range.Value = (PercOver43 != null ? ((double)PercOver43).ToString("F0") : "--");
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

                            range.Value = (PercOver260 != null ? ((double)PercOver260).ToString("F0") : "--");
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

                xlApp.Range["M16:M16"].Select();
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