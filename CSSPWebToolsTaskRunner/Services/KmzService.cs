using DHI.Generic.MikeZero.DFS.dfsu;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CSSPWebToolsTaskRunner.Services;
using CSSPWebToolsTaskRunner.Services.Resources;
using System.Windows.Forms;
using System.Drawing;
using System.Diagnostics;
using System.Transactions;
using System.Xml;
using System.Threading;
using DHI.Generic.MikeZero.DFS;
using DHI.Generic.MikeZero;
using CSSPWebToolsDBDLL.Models;
using CSSPWebToolsDBDLL;
using CSSPWebToolsDBDLL.Services;
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;

namespace CSSPWebToolsTaskRunner.Services
{
    public class KmzService
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
        public MikeScenarioService _MikeScenarioService { get; private set; }
        #endregion Properties

        #region Constructors
        public KmzService(TaskRunnerBaseService taskRunnerBaseService)
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
            _MikeScenarioService = new MikeScenarioService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        }
        #endregion Constructors

        #region Functions public
        public void CreateDocumentFromTemplateKml(DocTemplateModel docTemplateModel)
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

            string retStr = ParseKml(fi);
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
        //public void GenerateRootKMZ()
        //{
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Root, FileGeneratorTypeEnum.KMZ);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    KmzServiceRoot kmzServiceRoot = new KmzServiceRoot(_TaskRunnerBaseService); 
        //    kmzServiceRoot.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.RootFileAutoGenerate, FilePurposeEnum.Generated);
        //}
        //public void GenerateCountryKMZ()
        //{
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Country, FileGeneratorTypeEnum.KMZ);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);

        //    KmzServiceCountry kmzServiceCountry = new KmzServiceCountry(_TaskRunnerBaseService);
        //    kmzServiceCountry.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.CountryFileAutoGenerate, FilePurposeEnum.Generated);
        //}
        //public void GenerateProvinceKMZ()
        //{
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Province, FileGeneratorTypeEnum.KMZ);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    KmzServiceProvince kmzServiceProvince = new KmzServiceProvince(_TaskRunnerBaseService);
        //    kmzServiceProvince.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.ProvinceFileAutoGenerate, FilePurposeEnum.Generated);
        //}
        //public void GenerateAreaKMZ()
        //{
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Area, FileGeneratorTypeEnum.KMZ);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    KmzServiceArea kmzServiceArea = new KmzServiceArea(_TaskRunnerBaseService);
        //    kmzServiceArea.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.AreaFileAutoGenerate, FilePurposeEnum.Generated);
        //}
        //public void GenerateSectorKMZ()
        //{
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Sector, FileGeneratorTypeEnum.KMZ);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    KmzServiceSector kmzServiceSector = new KmzServiceSector(_TaskRunnerBaseService);
        //    kmzServiceSector.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.SectorFileAutoGenerate, FilePurposeEnum.Generated);
        //}
        //public void GenerateSubsectorKMZ()
        //{
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Subsector, FileGeneratorTypeEnum.KMZ);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    KmzServiceSubsector kmzServiceSubsector = new KmzServiceSubsector(_TaskRunnerBaseService);
        //    kmzServiceSubsector.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.SubsectorFileAutoGenerate, FilePurposeEnum.Generated);
        //}
        //public void GenerateMunicipalityKMZ()
        //{
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Municipality, FileGeneratorTypeEnum.KMZ);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    KmzServiceMunicipality kmzServiceMunicipality = new KmzServiceMunicipality(_TaskRunnerBaseService);
        //    kmzServiceMunicipality.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.MunicipalityFileAutoGenerate, FilePurposeEnum.Generated);
        //}
        //public void GenerateMikeScenarioBoundaryConditionsKMZ()
        //{
        //    string NotUsed = "";
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.MIKEScenarioBoundaryConditions, FileGeneratorTypeEnum.KMZ);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    KmzServiceMikeScenario kmzServiceMIKEScenario = new KmzServiceMikeScenario(_TaskRunnerBaseService);
        //    kmzServiceMIKEScenario.GenerateMikeScenarioBoundaryConditions(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.MikeScenarioBoundaryConditionFileAutoGenerate, FilePurposeEnum.Generated);

        //    MikeScenarioModel mikeScenarioModel = _MikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //    if (!string.IsNullOrWhiteSpace(mikeScenarioModel.Error))
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        return;
        //    }
        //}
        //public void GenerateMikeScenarioConcentrationAnimationKMZ()
        //{
        //    string NotUsed = "";
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.MIKEScenarioConcentrationAnimation, FileGeneratorTypeEnum.KMZ);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    KmzServiceMikeScenario kmzServiceMIKEScenario = new KmzServiceMikeScenario(_TaskRunnerBaseService);
        //    kmzServiceMIKEScenario.GenerateMikeScenarioConcentrationAnimation(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.MikeScenarioConcentrationAnimation, FilePurposeEnum.Generated);

        //    MikeScenarioModel mikeScenarioModel = _MikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //    if (!string.IsNullOrWhiteSpace(mikeScenarioModel.Error))
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        return;
        //    }

        //}
        //public void GenerateMikeScenarioConcentrationLimitsKMZ()
        //{
        //    string NotUsed = "";
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.MIKEScenarioConcentrationLimits, FileGeneratorTypeEnum.KMZ);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    KmzServiceMikeScenario kmzServiceMIKEScenario = new KmzServiceMikeScenario(_TaskRunnerBaseService);
        //    kmzServiceMIKEScenario.GenerateMikeScenarioConcentrationLimits(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.MikeScenarioConcentrationLimits, FilePurposeEnum.Generated);

        //    MikeScenarioModel mikeScenarioModel = _MikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //    if (!string.IsNullOrWhiteSpace(mikeScenarioModel.Error))
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        return;
        //    }

        //}
        //public void GenerateMikeScenarioCurrentAnimationKMZ()
        //{
        //    string NotUsed = "";
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.MIKEScenarioCurrentAnimation, FileGeneratorTypeEnum.KMZ);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    KmzServiceMikeScenario kmzServiceMIKEScenario = new KmzServiceMikeScenario(_TaskRunnerBaseService);
        //    kmzServiceMIKEScenario.GenerateMikeScenarioCurrentAnimation(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.MikeScenarioCurrentAnimation, FilePurposeEnum.Generated);

        //    MikeScenarioModel mikeScenarioModel = _MikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //    if (!string.IsNullOrWhiteSpace(mikeScenarioModel.Error))
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        return;
        //    }

        //}
        //public void GenerateMikeScenarioCurrentMaximumKMZ()
        //{
        //    string NotUsed = "";
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.MIKEScenarioCurrentMaximum, FileGeneratorTypeEnum.KMZ);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    KmzServiceMikeScenario kmzServiceMIKEScenario = new KmzServiceMikeScenario(_TaskRunnerBaseService);
        //    kmzServiceMIKEScenario.GenerateMikeScenarioCurrentMaximum(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.MikeScenarioCurrentMaximum, FilePurposeEnum.Generated);

        //    MikeScenarioModel mikeScenarioModel = _MikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //    if (!string.IsNullOrWhiteSpace(mikeScenarioModel.Error))
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        return;
        //    }

        //}
        //public void GenerateMikeScenarioMeshKMZ()
        //{
        //    string NotUsed = "";
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.MIKEScenarioMesh, FileGeneratorTypeEnum.KMZ);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    KmzServiceMikeScenario kmzServiceMIKEScenario = new KmzServiceMikeScenario(_TaskRunnerBaseService);
        //    kmzServiceMIKEScenario.GenerateMikeScenarioMesh(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.MikeScenarioMesh, FilePurposeEnum.Generated);

        //    MikeScenarioModel mikeScenarioModel = _MikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //    if (!string.IsNullOrWhiteSpace(mikeScenarioModel.Error))
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        return;
        //    }

        //}
        //public void GenerateMikeScenarioStudyAreaKMZ()
        //{
        //    string NotUsed = "";
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.MIKEScenarioStudyArea, FileGeneratorTypeEnum.KMZ);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    KmzServiceMikeScenario kmzServiceMIKEScenario = new KmzServiceMikeScenario(_TaskRunnerBaseService);
        //    kmzServiceMIKEScenario.GenerateMikeScenarioStudyArea(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.MikeScenarioStudyArea, FilePurposeEnum.Generated);

        //    MikeScenarioModel mikeScenarioModel = _MikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //    if (!string.IsNullOrWhiteSpace(mikeScenarioModel.Error))
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        return;
        //    }

        //}
        public string ParseKml(FileInfo fi)
        {

            return "Need to finalize ParseKML";
        }
        #endregion Functions public

        #region Functions private
        #endregion Functions private

    }
}
