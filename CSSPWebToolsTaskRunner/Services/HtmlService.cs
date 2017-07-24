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
    public class HtmlService
    {
        #region Variables
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        #endregion Variables

        #region Constructors
        public HtmlService(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
        }
        #endregion Constructors

        #region Functions public
        public void GenerateRootHTML()
        {
            _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Root, FileGeneratorTypeEnum.HTML);

            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

            FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

            TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            HtmlServiceRoot htmlServiceRoot = new HtmlServiceRoot(_TaskRunnerBaseService);
            htmlServiceRoot.Generate(fi);

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.RootFileAutoGenerate, FilePurposeEnum.Generated);
        }
        public void GenerateCountryHTML()
        {
            _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Country, FileGeneratorTypeEnum.HTML);

            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

            FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

            TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            HtmlServiceCountry htmlServiceCountry = new HtmlServiceCountry(_TaskRunnerBaseService);
            htmlServiceCountry.Generate(fi);

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.CountryFileAutoGenerate, FilePurposeEnum.Generated);
        }
        public void GenerateProvinceHTML()
        {
            _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Province, FileGeneratorTypeEnum.HTML);

            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

            FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

            TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            HtmlServiceProvince htmlServiceProvince = new HtmlServiceProvince(_TaskRunnerBaseService);
            htmlServiceProvince.Generate(fi);

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.ProvinceFileAutoGenerate, FilePurposeEnum.Generated);
        }
        public void GenerateAreaHTML()
        {
            _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Area, FileGeneratorTypeEnum.HTML);

            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

            FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

            TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            HtmlServiceArea htmlServiceArea = new HtmlServiceArea(_TaskRunnerBaseService);
            htmlServiceArea.Generate(fi);

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.AreaFileAutoGenerate, FilePurposeEnum.Generated);
        }
        public void GenerateSectorHTML()
        {
            _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Sector, FileGeneratorTypeEnum.HTML);

            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

            FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

            TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            HtmlServiceSector htmlServiceSector = new HtmlServiceSector(_TaskRunnerBaseService);
            htmlServiceSector.Generate(fi);

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.SectorFileAutoGenerate, FilePurposeEnum.Generated);
        }
        public void GenerateSubsectorHTML()
        {
            _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Subsector, FileGeneratorTypeEnum.HTML);

            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

            FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

            TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            HtmlServiceSubsector htmlServiceSubsector = new HtmlServiceSubsector(_TaskRunnerBaseService);
            htmlServiceSubsector.Generate(fi);

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.SubsectorFileAutoGenerate, FilePurposeEnum.Generated);
        }
        public void GenerateMunicipalityHTML()
        {
            _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Municipality, FileGeneratorTypeEnum.HTML);

            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

            FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

            TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            HtmlServiceMunicipality htmlServiceMunicipality = new HtmlServiceMunicipality(_TaskRunnerBaseService);
            htmlServiceMunicipality.Generate(fi);

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.MunicipalityFileAutoGenerate, FilePurposeEnum.Generated);
        }
        #endregion Functions public

        #region Functions private
        #endregion Functions private
    }
}
