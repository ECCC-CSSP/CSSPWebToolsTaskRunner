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
        #endregion Function public

        #region Functions private
        #endregion
    }
}