using CSSPWebToolsDBDLL.Models;
using CSSPWebToolsDBDLL.Services;
using CSSPWebToolsTaskRunner.Services.Resources;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Transactions;
using System.Windows.Forms;
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;
using WQMS.RTNotifiers;
using CSSPWebToolsDBDLL;
using CSSPFCFormWriterDLL.Services;
using CSSPReportWriterHelperDLL.Services;

namespace CSSPWebToolsTaskRunner.Services
{
    // Class Word and Excel
    public class UsedHyperlink
    {
        public UsedHyperlink()
        {
        }

        public string Id { get; set; }
        public string URL { get; set; }
    }

    public class TaskRunnerBaseService
    {
        #region Variables
        #endregion Variables

        #region Properties
        public IPrincipal _User { get; set; }
        public BWObj _BWObj { get; set; }
        public List<BWObj> _BWObjList { get; private set; }
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        public RichTextBox _RichTextBoxStatus { get; set; }
        #endregion Properties

        #region Constructors
        public TaskRunnerBaseService(List<BWObj> bwObjList)
        {
            _TaskRunnerBaseService = this;
            _BWObjList = bwObjList;
            _User = new GenericPrincipal(new GenericIdentity("Charles.LeBlanc2@Canada.ca", "Forms"), null);
        }
        #endregion Constructors

        #region Functions public
        public string CleanText(string TextToClean)
        {
            // got to make sure there are not \t, \r, \n, in the text
            TextToClean = TextToClean.Replace("/", "-").Replace("\t", " ").Replace("\r", " ").Replace("\n", " ").Replace("&", "AND").Replace("#", " ");

            // removing all double space
            for (int i = 0; i < 20; i++)
            {
                TextToClean = TextToClean.Replace("  ", " ");
            }

            return TextToClean.Trim();
        }
        public TVItemModel CreateFileTVItem(FileInfo fi)
        {
            string NotUsed = "";

            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemModel tvItemModelFile = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi.Name, TVTypeEnum.File);
            if (!string.IsNullOrEmpty(tvItemModelFile.Error))
            {
                tvItemModelFile = tvItemService.PostAddChildTVItemDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi.Name, TVTypeEnum.File);
                if (!string.IsNullOrEmpty(tvItemModelFile.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_For_With_Equal_, TaskRunnerServiceRes.TVItem, TVTypeEnum.File.ToString(), TaskRunnerServiceRes.TVItemID + ", " + TaskRunnerServiceRes.TVText, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID + ", " + fi.Name);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat4List("CouldNotCreate_For_With_Equal_", TaskRunnerServiceRes.TVItem, TVTypeEnum.File.ToString(), TaskRunnerServiceRes.TVItemID + ", " + TaskRunnerServiceRes.TVText, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID + ", " + fi.Name);
                    return new TVItemModel() { Error = _TaskRunnerBaseService._BWObj.TextLanguageList.Where(c => c.Language == _TaskRunnerBaseService._BWObj.appTaskModel.Language).First().Text };
                }
            }

            return tvItemModelFile;
        }
        public void ExecuteTask()
        {
            switch (_BWObj.appTaskModel.AppTaskCommand)
            {
                case AppTaskCommandEnum.OpenDataCSVOfMWQMSites:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        TxtService txtService = new TxtService(_TaskRunnerBaseService);
                        txtService.CreateCSVOfMWQMSites();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.OpenDataCSVNationalOfMWQMSites:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        TxtService txtService = new TxtService(_TaskRunnerBaseService);
                        txtService.CreateCSVNationalOfMWQMSites();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.OpenDataKMZOfMWQMSites:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        KmzService kmzService = new KmzService(_TaskRunnerBaseService);
                        kmzService.CreateKMZOfMWQMSites();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.OpenDataCSVOfMWQMSamples:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        TxtService txtService = new TxtService(_TaskRunnerBaseService);
                        txtService.CreateCSVOfMWQMSamples();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.OpenDataCSVNationalOfMWQMSamples:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        TxtService txtService = new TxtService(_TaskRunnerBaseService);
                        txtService.CreateCSVNationalOfMWQMSamples();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.OpenDataXlsxOfMWQMSitesAndSamples:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        XlsxService xlsxService = new XlsxService(_TaskRunnerBaseService);
                        xlsxService.CreateXlsxOfMWQMSitesAndSamples();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.ProvinceToolsCreateClassificationInputsKML:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        KmzService kmzService = new KmzService(_TaskRunnerBaseService);
                        kmzService.ProvinceToolsCreateClassificationInputsKML();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.ProvinceToolsCreateGroupingInputsKML:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        KmzService kmzService = new KmzService(_TaskRunnerBaseService);
                        kmzService.ProvinceToolsCreateGroupingInputsKML();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.ProvinceToolsCreateMWQMSitesAndPolSourceSitesKML:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        KmzService kmzService = new KmzService(_TaskRunnerBaseService);
                        kmzService.ProvinceToolsCreateMWQMSitesAndPolSourceSitesKML();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.CreateWebTideDataWLAtFirstNode:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        MikeScenarioFileService mikeScenarioFileService = new MikeScenarioFileService(_TaskRunnerBaseService);
                        mikeScenarioFileService.CreateWebTideDataWLAtFirstNode();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.ExportEmailDistributionLists:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        XlsxService xlsxService = new XlsxService(_TaskRunnerBaseService);
                        xlsxService.ExportEmailDistributionLists();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.GenerateWebTide:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        TidesAndCurrentsService tidesAndCurrentsService = new TidesAndCurrentsService(_TaskRunnerBaseService);
                        tidesAndCurrentsService.GenerateWebTideNodes();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.MikeScenarioAskToRun:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        MikeScenarioFileService mikeScenarioFileService = new MikeScenarioFileService(_TaskRunnerBaseService);
                        mikeScenarioFileService.MikeScenarioAskToRun();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            //appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.MikeScenarioWaitingToRun:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        MikeScenarioFileService mikeScenarioFileService = new MikeScenarioFileService(_TaskRunnerBaseService);
                        //mikeScenarioFileService.MikeScenarioAskToRun();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            //appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.MikeScenarioImport:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        MikeScenarioFileService mikeScenarioFileService = new MikeScenarioFileService(_TaskRunnerBaseService);
                        mikeScenarioFileService.MikeScenarioImportDB();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.MikeScenarioOtherFileImport:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        MikeScenarioFileService mikeScenarioFileService = new MikeScenarioFileService(_TaskRunnerBaseService);
                        mikeScenarioFileService.MikeScenarioOtherFileImportDB();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.MikeScenarioRunning:
                    {
                        // nothing
                    }
                    break;
                case AppTaskCommandEnum.SetupWebTide:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        TidesAndCurrentsService tidesAndCurrentsService = new TidesAndCurrentsService(_TaskRunnerBaseService);
                        tidesAndCurrentsService.SetupWebTide();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.ExportToArcGIS:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        XlsxService xlsxService = new XlsxService(_TaskRunnerBaseService);
                        xlsxService.ExportToArcGIS();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.GenerateClassificationForCSSPWebToolsVisualization:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        ClassificationGenerateService classificationGenerateService = new ClassificationGenerateService(_TaskRunnerBaseService);
                        classificationGenerateService.GenerateClassificationForCSSPWebToolsVisualization();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.GenerateLinksBetweenMWQMSitesAndPolSourceSitesForCSSPWebToolsVisualization:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        GroupingMWQMSitesAndPolSourceSitesGenerateService groupingMWQMSitesAndPolSourceSitesGenerateService = new GroupingMWQMSitesAndPolSourceSitesGenerateService(_TaskRunnerBaseService);
                        groupingMWQMSitesAndPolSourceSitesGenerateService.GenerateLinksBetweenMWQMSitesAndPolSourceSitesForCSSPWebToolsVisualization();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.GetAllPrecipitationForYear:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        ClimateService climateService = new ClimateService(_TaskRunnerBaseService);
                        climateService.GetAllPrecipitationForYear();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.FillRunPrecipByClimateSitePriorityForYear:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        ClimateService climateService = new ClimateService(_TaskRunnerBaseService);
                        climateService.FillRunPrecipByClimateSitePriorityForYear();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.FindMissingPrecipForProvince:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        ClimateService climateService = new ClimateService(_TaskRunnerBaseService);
                        climateService.FindMissingPrecipForProvince();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.GetClimateSitesDataForRunsOfYear:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        ClimateService climateService = new ClimateService(_TaskRunnerBaseService);
                        climateService.GetClimateSitesDataForRunsOfYear();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.UpdateClimateSiteInformation:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        ClimateService climateService = new ClimateService(_TaskRunnerBaseService);
                        climateService.UpdateClimateSitesInformationForCountryTVItemID();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.UpdateHydrometricSiteInformation:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        HydrometricService hydrometricService = new HydrometricService(_TaskRunnerBaseService);
                        hydrometricService.UpdateHydrometricSitesInformationForProvinceTVItemID();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.CreateFCForm:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        DocxService docxService = new DocxService(_TaskRunnerBaseService);

                        string Parameters = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters;
                        int LabSheetID = int.Parse(appTaskService.GetAppTaskParamStr(Parameters, "LabSheetID"));

                        LabSheetService labSheetService = new LabSheetService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        LabSheetModel labSheetModel = labSheetService.GetLabSheetModelWithLabSheetIDDB(LabSheetID);

                        SamplingPlanService samplingPlanService = new SamplingPlanService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        SamplingPlanModel samplingPlanModel = new SamplingPlanModel();
                        samplingPlanModel = samplingPlanService.GetSamplingPlanModelWithSamplingPlanIDDB(labSheetModel.SamplingPlanID);
                        //if (samplingPlanModel.IncludeLaboratoryQAQC)
                        //{
                        docxService.GenerateFCFormDOCX();
                        //}
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.CreateSamplingPlanSamplingPlanTextFile:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        TxtService txtService = new TxtService(_TaskRunnerBaseService);
                        txtService.CreateSamplingPlanSamplingPlan();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.ExportAnalysisToExcel:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        XlsxService xlsxService = new XlsxService(_TaskRunnerBaseService);
                        xlsxService.CreateExcelFileForAnalysisReportParameter();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.CreateDocxPDF:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        DocxService docxService = new DocxService(_TaskRunnerBaseService);
                        docxService.CreateDocxPDF();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.CreateXlsxPDF:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        XlsxService xlsxService = new XlsxService(_TaskRunnerBaseService);
                        xlsxService.CreateXlsxPDF();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.CreateDocumentFromParameters:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        ParametersService parametersService = new ParametersService(_TaskRunnerBaseService);
                        parametersService.Generate();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.CreateDocumentFromTemplate:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        string[] ParamValueList = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                        if (ParamValueList.Count() != 2)
                        {
                            string NotUsed = string.Format(TaskRunnerServiceRes.ParameterCount_NotEqual_, ParamValueList.Count().ToString(), "2");
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("ParameterCount_NotEqual_", ParamValueList.Count().ToString(), "4");
                            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                            {
                                appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                            }
                            else
                            {
                                SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                            }
                            return;
                        }

                        int TVItemID = (int.Parse(appTaskService.GetAppTaskParamStr(_TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "TVItemID")));

                        if (TVItemID == 0)
                        {
                            string NotUsed = string.Format(TaskRunnerServiceRes.Parameter_NotFound, "TVItemID");
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Parameter_NotFound", "TVItemID");
                            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                            {
                                appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                            }
                            else
                            {
                                SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                            }
                            return;
                        }

                        int DocTemplateID = (int.Parse(appTaskService.GetAppTaskParamStr(_TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "DocTemplateID")));

                        if (DocTemplateID == 0)
                        {
                            string NotUsed = string.Format(TaskRunnerServiceRes.Parameter_NotFound, "DocTemplateID");
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Parameter_NotFound", "DocTemplateID");
                            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                            {
                                appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                            }
                            else
                            {
                                SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                            }
                            return;
                        }

                        DocTemplateService docTemplateService = new DocTemplateService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        DocTemplateModel docTemplateModel = docTemplateService.GetDocTemplateModelWithDocTemplateIDDB(DocTemplateID);
                        if (!string.IsNullOrWhiteSpace(docTemplateModel.Error))
                        {
                            string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateDocumentFromTemplateError_, docTemplateModel.Error);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotCreateDocumentFromTemplateError_", docTemplateModel.Error);
                            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                            {
                                appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                            }
                            else
                            {
                                SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                            }
                            return;
                        }

                        if (string.IsNullOrWhiteSpace(docTemplateModel.FileName))
                        {
                            string NotUsed = string.Format(TaskRunnerServiceRes.FileNameForDocTemplateID_IsEmpty, DocTemplateID.ToString());
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("FileNameForDocTemplateID_IsEmpty", DocTemplateID.ToString());
                            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                            {
                                appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                            }
                            else
                            {
                                SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                            }
                            return;
                        }

                        TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        TVFileModel tvFileModel = tvFileService.GetTVFileModelWithTVFileTVItemIDDB(docTemplateModel.TVFileTVItemID);
                        if (!string.IsNullOrWhiteSpace(tvFileModel.Error))
                        {
                            string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateDocumentFromTemplateError_, tvFileModel.Error);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotCreateDocumentFromTemplateError_", tvFileModel.Error);
                            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                            {
                                appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                            }
                            else
                            {
                                SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                            }
                            return;
                        }

                        FileInfo fiTemplate = new FileInfo(tvFileModel.ServerFilePath + tvFileModel.ServerFileName);

                        if (!fiTemplate.Exists)
                        {
                            string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiTemplate.FullName);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiTemplate.FullName);
                            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                            {
                                appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                            }
                            else
                            {
                                SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                            }
                            return;
                        }

                        string NewFileName = fiTemplate.FullName.Replace("Template_", "");
                        string DateText = "_" + DateTime.Now.Year.ToString() + "_" + DateTime.Now.Month.ToString() + "_" + DateTime.Now.Day.ToString() + "_" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString();

                        switch (fiTemplate.Extension.ToLower())
                        {
                            case ".csv":
                                NewFileName = NewFileName.Replace(".csv", DateText + ".csv");
                                break;
                            case ".docx":
                                NewFileName = NewFileName.Replace(".docx", DateText + ".docx");
                                break;
                            case ".xlsx":
                                NewFileName = NewFileName.Replace(".xlsx", DateText + ".xlsx");
                                break;
                            case ".kml":
                                NewFileName = NewFileName.Replace(".kml", DateText + ".kml");
                                break;
                            default:
                                break;
                        }

                        FileInfo fi = new FileInfo(NewFileName);

                        // need to create the file on the server
                        try
                        {
                            File.Copy(fiTemplate.FullName, fi.FullName);
                        }
                        catch (Exception ex)
                        {
                            string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCopyFile_To_Error_, fiTemplate.FullName, fi.FullName, ex.Message + " --- Inner: " + (ex.InnerException == null ? "" : ex.InnerException.Message));
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotCopyFile_To_Error_", fiTemplate.FullName, fi.FullName, ex.Message + " --- Inner: " + (ex.InnerException == null ? "" : ex.InnerException.Message));
                            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                            {
                                appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                            }
                            else
                            {
                                SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                            }
                            return;
                        }

                        fi = new FileInfo(fi.FullName);
                        if (!fi.Exists)
                        {
                            string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fi.FullName);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fi.FullName);
                            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                            {
                                appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                            }
                            else
                            {
                                SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                            }
                            return;
                        }

                        switch (fiTemplate.Extension.ToLower())
                        {
                            case ".csv":
                                {
                                    ReportBaseService reportBaseService = new ReportBaseService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, new TreeView(), _TaskRunnerBaseService._User);
                                    string retStr = reportBaseService.GenerateReportFromTemplateCSV(fi, TVItemID, 99999999, _TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                                    if (!string.IsNullOrWhiteSpace(retStr))
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateDocumentFromTemplateError_, retStr);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotCreateDocumentFromTemplateError_", retStr);
                                        SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);

                                        try
                                        {
                                            fi.Delete();
                                        }
                                        catch (Exception)
                                        {
                                            // nothing there is already a error message in the AppTask Table
                                        }
                                        return;
                                    }

                                    TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

                                    TVItemModel tvItemModelExist = tvItemService.GetChildTVItemModelWithParentIDAndTVTextAndTVTypeDB(TVItemID, fi.Name, TVTypeEnum.File);
                                    if (string.IsNullOrWhiteSpace(tvItemModelExist.Error))
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes.FileAlreadyExist_, fi.Name);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("FileAlreadyExist_", fi.Name);
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        return;
                                    }

                                    TVItemModel tvItemModelNew = tvItemService.PostAddChildTVItemDB(TVItemID, fi.Name, TVTypeEnum.File);
                                    if (!string.IsNullOrWhiteSpace(tvItemModelNew.Error))
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVItem, tvItemModelNew.Error);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVItem, tvItemModelNew.Error);
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        return;
                                    }

                                    fi = new FileInfo(fi.FullName);

                                    TVFileModel tvFileModelNew = new TVFileModel()
                                    {
                                        TVFileTVItemID = tvItemModelNew.TVItemID,
                                        FilePurpose = FilePurposeEnum.TemplateGenerated,
                                        FileDescription = "Generated File",
                                        FileType = tvFileService.GetFileType(fi.Extension),
                                        FileSize_kb = Math.Max((int)fi.Length / 1024, 1),
                                        FileInfo = "Generated File",
                                        FileCreatedDate_UTC = DateTime.Now,
                                        FromWater = false,
                                        ClientFilePath = fi.Name,
                                        ServerFileName = fi.Name,
                                        ServerFilePath = fi.DirectoryName + "\\",
                                        Language = tvFileModel.Language,
                                        Year = DateTime.Now.Year,
                                    };

                                    TVFile tvFileExist = tvFileService.GetTVFileExistDB(tvFileModelNew);
                                    if (tvFileExist != null)
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes._AlreadyExist, TaskRunnerServiceRes.TVFile);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVItem, tvFileModel.Error);
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        return;
                                    }

                                    TVFileModel tvFileModelRet = tvFileService.PostAddTVFileDB(tvFileModelNew);
                                    if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVFile, tvFileModel.Error);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVFile, tvFileModel.Error);
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        return;
                                    }

                                    appTaskService.PostDeleteAppTaskDB(_BWObj.appTaskModel.AppTaskID);

                                }
                                break;
                            case ".docx":
                                {
                                    bool KeepAppTask = false;
                                    ReportBaseService reportBaseService = new ReportBaseService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, new TreeView(), _TaskRunnerBaseService._User);
                                    string retStr = reportBaseService.GenerateReportFromTemplateWord(fi, TVItemID, 99999999, _TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                                    if (!string.IsNullOrWhiteSpace(retStr))
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes.ErrorsWhileCreatingDocumentDeleteTaskAndCheckCreatedDocument, fi.Name);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("ErrorsWhileCreatingDocumentDeleteTaskAndCheckCreatedDocument");
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        KeepAppTask = true;
                                    }

                                    TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

                                    TVItemModel tvItemModelExist = tvItemService.GetChildTVItemModelWithParentIDAndTVTextAndTVTypeDB(TVItemID, fi.Name, TVTypeEnum.File);
                                    if (string.IsNullOrWhiteSpace(tvItemModelExist.Error))
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes.FileAlreadyExist_, fi.Name);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("FileAlreadyExist_", fi.Name);
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        return;
                                    }

                                    TVItemModel tvItemModelNew = tvItemService.PostAddChildTVItemDB(TVItemID, fi.Name, TVTypeEnum.File);
                                    if (!string.IsNullOrWhiteSpace(tvItemModelNew.Error))
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVItem, tvItemModelNew.Error);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVItem, tvItemModelNew.Error);
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        return;
                                    }

                                    fi = new FileInfo(fi.FullName);

                                    TVFileModel tvFileModelNew = new TVFileModel()
                                    {
                                        TVFileTVItemID = tvItemModelNew.TVItemID,
                                        FilePurpose = FilePurposeEnum.TemplateGenerated,
                                        FileDescription = "Generated File",
                                        FileType = tvFileService.GetFileType(fi.Extension),
                                        FileSize_kb = Math.Max((int)fi.Length / 1024, 1),
                                        FileInfo = "Generated File",
                                        FileCreatedDate_UTC = DateTime.Now,
                                        FromWater = false,
                                        ClientFilePath = fi.Name,
                                        ServerFileName = fi.Name,
                                        ServerFilePath = fi.DirectoryName + "\\",
                                        Language = tvFileModel.Language,
                                        Year = DateTime.Now.Year,
                                    };

                                    TVFile tvFileExist = tvFileService.GetTVFileExistDB(tvFileModelNew);
                                    if (tvFileExist != null)
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes._AlreadyExist, TaskRunnerServiceRes.TVFile);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVItem, tvFileModel.Error);
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        return;
                                    }

                                    TVFileModel tvFileModelRet = tvFileService.PostAddTVFileDB(tvFileModelNew);
                                    if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVFile, tvFileModel.Error);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVFile, tvFileModel.Error);
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        return;
                                    }

                                    if (!KeepAppTask)
                                        appTaskService.PostDeleteAppTaskDB(_BWObj.appTaskModel.AppTaskID);

                                }
                                break;
                            case ".xlsx":
                                {
                                    //XlsxService xlsxService = new XlsxService(_TaskRunnerBaseService);
                                    //xlsxService.CreateDocumentFromTemplateXlsx(docTemplateModel);
                                    //UpdateAppTaskWithDeleteOrError();
                                    string NotUsed = string.Format(TaskRunnerServiceRes._NotImplemented, _BWObj.appTaskModel.AppTaskCommand.ToString());
                                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_NotImplemented", _BWObj.appTaskModel.AppTaskCommand.ToString());
                                    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                    {
                                        appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                    }
                                    else
                                    {
                                        SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                    }
                                }
                                break;
                            case ".kml":
                                {
                                    bool KeepAppTask = false;
                                    ReportBaseService reportBaseService = new ReportBaseService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, new TreeView(), _TaskRunnerBaseService._User);
                                    string retStr = reportBaseService.GenerateReportFromTemplateKML(fi, TVItemID, 99999999, _TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                                    if (!string.IsNullOrWhiteSpace(retStr))
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateDocumentFromTemplateError_, retStr);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotCreateDocumentFromTemplateError_", retStr);
                                        SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);

                                        try
                                        {
                                            fi.Delete();
                                        }
                                        catch (Exception)
                                        {
                                            // nothing there is already a error message in the AppTask Table
                                        }
                                        return;
                                    }

                                    TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

                                    TVItemModel tvItemModelExist = tvItemService.GetChildTVItemModelWithParentIDAndTVTextAndTVTypeDB(TVItemID, fi.Name, TVTypeEnum.File);
                                    if (string.IsNullOrWhiteSpace(tvItemModelExist.Error))
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes.FileAlreadyExist_, fi.Name);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("FileAlreadyExist_", fi.Name);
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        return;
                                    }

                                    TVItemModel tvItemModelNew = tvItemService.PostAddChildTVItemDB(TVItemID, fi.Name, TVTypeEnum.File);
                                    if (!string.IsNullOrWhiteSpace(tvItemModelNew.Error))
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVItem, tvItemModelNew.Error);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVItem, tvItemModelNew.Error);
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        return;
                                    }

                                    fi = new FileInfo(fi.FullName);

                                    TVFileModel tvFileModelNew = new TVFileModel()
                                    {
                                        TVFileTVItemID = tvItemModelNew.TVItemID,
                                        FilePurpose = FilePurposeEnum.TemplateGenerated,
                                        FileDescription = "Generated File",
                                        FileType = tvFileService.GetFileType(fi.Extension),
                                        FileSize_kb = Math.Max((int)fi.Length / 1024, 1),
                                        FileInfo = "Generated File",
                                        FileCreatedDate_UTC = DateTime.Now,
                                        FromWater = false,
                                        ClientFilePath = fi.Name,
                                        ServerFileName = fi.Name,
                                        ServerFilePath = fi.DirectoryName + "\\",
                                        Language = tvFileModel.Language,
                                        Year = DateTime.Now.Year,
                                    };

                                    TVFile tvFileExist = tvFileService.GetTVFileExistDB(tvFileModelNew);
                                    if (tvFileExist != null)
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes._AlreadyExist, TaskRunnerServiceRes.TVFile);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVItem, tvFileModel.Error);
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        return;
                                    }

                                    TVFileModel tvFileModelRet = tvFileService.PostAddTVFileDB(tvFileModelNew);
                                    if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVFile, tvFileModel.Error);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVFile, tvFileModel.Error);
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        return;
                                    }

                                    if (!KeepAppTask)
                                        appTaskService.PostDeleteAppTaskDB(_BWObj.appTaskModel.AppTaskID);

                                }
                                break;
                            default:
                                break;
                        }

                        // need to add the file in TVFile and TVItemID

                    }
                    break;
                default:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        string NotUsed = string.Format(TaskRunnerServiceRes._NotImplemented, _BWObj.appTaskModel.AppTaskCommand.ToString());
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_NotImplemented", _BWObj.appTaskModel.AppTaskCommand.ToString());
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
            }
        }
        public string GetCap(string str)
        {
            string CapStr = "";
            foreach (char c in str)
            {
                if (Char.IsUpper(c))
                {
                    CapStr += c.ToString();
                }
            }

            return CapStr;
        }
        public List<TextLanguage> GetTextLanguageList()
        {
            List<TextLanguage> TextLanguageList = new List<TextLanguage>();
            AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            foreach (LanguageEnum lang in appTaskService.LanguageListAllowable)
            {
                TextLanguageList.Add(new TextLanguage() { Language = lang, Text = "" });
            }
            return TextLanguageList;
        }
        public List<TextLanguage> GetTextLanguageList(string ResString)
        {
            List<TextLanguage> TextLanguageList = new List<TextLanguage>();
            AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            foreach (LanguageEnum lang in appTaskService.LanguageListAllowable)
            {
                string TempText = TaskRunnerServiceRes.ResourceManager.GetString(ResString, new System.Globalization.CultureInfo(lang.ToString()));
                TextLanguageList.Add(new TextLanguage() { Language = lang, Text = TempText });
            }
            return TextLanguageList;
        }
        public List<TextLanguage> GetTextLanguageFormat1List(string ResString, string Param1)
        {
            List<TextLanguage> TextLanguageList = new List<TextLanguage>();
            AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            foreach (LanguageEnum lang in appTaskService.LanguageListAllowable)
            {
                string TempText = string.Format(TaskRunnerServiceRes.ResourceManager.GetString(ResString, new System.Globalization.CultureInfo(lang.ToString())), Param1);
                TextLanguageList.Add(new TextLanguage() { Language = lang, Text = TempText });
            }
            return TextLanguageList;
        }
        public List<TextLanguage> GetTextLanguageFormat2List(string ResString, string Param1, string Param2)
        {
            List<TextLanguage> TextLanguageList = new List<TextLanguage>();
            AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            foreach (LanguageEnum lang in appTaskService.LanguageListAllowable)
            {
                string TempText = string.Format(TaskRunnerServiceRes.ResourceManager.GetString(ResString, new System.Globalization.CultureInfo(lang.ToString())), Param1, Param2);
                TextLanguageList.Add(new TextLanguage() { Language = lang, Text = TempText });
            }
            return TextLanguageList;
        }
        public List<TextLanguage> GetTextLanguageFormat3List(string ResString, string Param1, string Param2, string Param3)
        {
            List<TextLanguage> TextLanguageList = new List<TextLanguage>();
            AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            foreach (LanguageEnum lang in appTaskService.LanguageListAllowable)
            {
                string TempText = string.Format(TaskRunnerServiceRes.ResourceManager.GetString(ResString, new System.Globalization.CultureInfo(lang.ToString())), Param1, Param2, Param3);
                TextLanguageList.Add(new TextLanguage() { Language = lang, Text = TempText });
            }
            return TextLanguageList;
        }
        public List<TextLanguage> GetTextLanguageFormat4List(string ResString, string Param1, string Param2, string Param3, string Param4)
        {
            List<TextLanguage> TextLanguageList = new List<TextLanguage>();
            AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            foreach (LanguageEnum lang in appTaskService.LanguageListAllowable)
            {
                string TempText = string.Format(TaskRunnerServiceRes.ResourceManager.GetString(ResString, new System.Globalization.CultureInfo(lang.ToString())), Param1, Param2, Param3, Param4);
                TextLanguageList.Add(new TextLanguage() { Language = lang, Text = TempText });
            }
            return TextLanguageList;
        }
        public List<TextLanguage> GetTextLanguageFormat5List(string ResString, string Param1, string Param2, string Param3, string Param4, string Param5)
        {
            List<TextLanguage> TextLanguageList = new List<TextLanguage>();
            AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            foreach (LanguageEnum lang in appTaskService.LanguageListAllowable)
            {
                string TempText = string.Format(TaskRunnerServiceRes.ResourceManager.GetString(ResString, new System.Globalization.CultureInfo(lang.ToString())), Param1, Param2, Param3, Param4, Param5);
                TextLanguageList.Add(new TextLanguage() { Language = lang, Text = TempText });
            }
            return TextLanguageList;
        }
        public string GetUrlFromTVItem(TVItemModel tvItemModel)
        {
            return "http://wmon01dtchlebl2/csspwebtools/" + _BWObj.appTaskModel.Language + "-CA/#!View/" + tvItemModel.TVText + "/A|||" + tvItemModel.TVItemID + "/1|||000000000000000000000000000000|||20101970"; ;
        }
        public string RunTaskForShawnDonohue()
        {
            RTStationMonitor rtStationMonitor = new RTStationMonitor();

            return rtStationMonitor.CheckAutmatedStationsAndNotify();
        }
        public void SendErrorTextToDB(List<TextLanguage> textLanguageList)
        {
            AppTaskModel appTaskModel = new AppTaskModel();
            AppTaskLanguageModel appTaskLanguageModel = new AppTaskLanguageModel();

            AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_BWObj.appTaskModel.AppTaskID);

            foreach (TextLanguage textLanguage in textLanguageList)
            {
                AppTaskLanguageService appTaskLanguageService = new AppTaskLanguageService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                appTaskLanguageModel = appTaskLanguageService.GetAppTaskLanguageModelWithAppTaskIDAndLanguageDB(_BWObj.appTaskModel.AppTaskID, textLanguage.Language);
                appTaskLanguageModel.ErrorText = textLanguage.Text;
                AppTaskLanguageModel appTaskLanguageModelRet = appTaskLanguageService.PostUpdateAppTaskLanguageDB(appTaskLanguageModel);
            }

        }
        public void SendErrorTextToDB(int AppTaskID, List<TextLanguage> textLanguageList)
        {
            AppTaskModel appTaskModel = new AppTaskModel();
            AppTaskLanguageModel appTaskLanguageModel = new AppTaskLanguageModel();

            AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(AppTaskID);

            foreach (TextLanguage textLanguage in textLanguageList)
            {
                AppTaskLanguageService appTaskLanguageService = new AppTaskLanguageService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                appTaskLanguageModel = appTaskLanguageService.GetAppTaskLanguageModelWithAppTaskIDAndLanguageDB(AppTaskID, textLanguage.Language);
                appTaskLanguageModel.ErrorText = textLanguage.Text;
                AppTaskLanguageModel appTaskLanguageModelRet = appTaskLanguageService.PostUpdateAppTaskLanguageDB(appTaskLanguageModel);
            }

        }
        public void SendStatusTextToDB(List<TextLanguage> textLanguageList)
        {
            AppTaskModel appTaskModel = new AppTaskModel();
            AppTaskLanguageModel appTaskLanguageModel = new AppTaskLanguageModel();

            AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_BWObj.appTaskModel.AppTaskID);

            foreach (TextLanguage textLanguage in textLanguageList)
            {
                AppTaskLanguageService appTaskLanguageService = new AppTaskLanguageService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                appTaskLanguageModel = appTaskLanguageService.GetAppTaskLanguageModelWithAppTaskIDAndLanguageDB(_BWObj.appTaskModel.AppTaskID, textLanguage.Language);
                appTaskLanguageModel.StatusText = textLanguage.Text;
                AppTaskLanguageModel appTaskLanguageModelRet = appTaskLanguageService.PostUpdateAppTaskLanguageDB(appTaskLanguageModel);
            }
        }
        public void SendParametersToDB(int AppTaskID, string parameter)
        {
            AppTaskModel appTaskModel = new AppTaskModel();
            AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(AppTaskID);
            appTaskModel.Parameters = parameter;
            appTaskService.PostUpdateAppTask(appTaskModel);
        }
        public void SendStatusToDB(AppTaskStatusEnum appTaskStatus)
        {
            AppTaskModel appTaskModel = new AppTaskModel();
            AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_BWObj.appTaskModel.AppTaskID);
            appTaskModel.AppTaskStatus = appTaskStatus;
            appTaskService.PostUpdateAppTask(appTaskModel);
        }
        public void SendStatusToDB(int AppTaskID, AppTaskStatusEnum appTaskStatus)
        {
            AppTaskModel appTaskModel = new AppTaskModel();
            AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(AppTaskID);
            appTaskModel.AppTaskStatus = appTaskStatus;
            appTaskService.PostUpdateAppTask(appTaskModel);
        }
        public void SendPercentToDB(int AppTaskID, int percent)
        {
            AppTaskModel appTaskModel = new AppTaskModel();
            AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(AppTaskID);
            appTaskModel.PercentCompleted = percent;
            appTaskService.PostUpdateAppTask(appTaskModel);
        }
        public void SendTaskCommandToDB(int AppTaskID, AppTaskCommandEnum AppTaskCommand)
        {
            AppTaskModel appTaskModel = new AppTaskModel();
            AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(AppTaskID);
            appTaskModel.AppTaskCommand = AppTaskCommand;
            appTaskService.PostUpdateAppTask(appTaskModel);
        }
        public void UpdateOrCreateTVFile(int ParentTVItemID, FileInfo fi, TVItemModel tvItemModelNew, string Description, FilePurposeEnum FilePurpose)
        {
            string NotUsed = "";
            fi = new FileInfo(fi.FullName);

            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVFileModel tvFileModel = tvFileService.GetTVFileModelWithServerFilePathAndServerFileNameDB(fi.Directory + @"\", fi.Name);
            if (!string.IsNullOrWhiteSpace(tvFileModel.Error))
            {
                TVFileModel tvFileModelNew = new TVFileModel();
                tvFileModelNew.TVFileTVItemID = tvItemModelNew.TVItemID;
                tvFileModelNew.TemplateTVType = TVTypeEnum.Error;
                tvFileModelNew.ReportTypeID = null;
                tvFileModelNew.Parameters = null;
                tvFileModelNew.ServerFileName = fi.Name;
                tvFileModelNew.FilePurpose = FilePurpose;
                tvFileModelNew.Language = _BWObj.appTaskModel.Language;
                tvFileModelNew.Year = DateTime.Now.Year;
                tvFileModelNew.FileDescription = Description;
                tvFileModelNew.FileType = tvFileService.GetFileType(fi.Extension);
                tvFileModelNew.FileSize_kb = (((int)fi.Length / 1024) == 0 ? 1 : (int)fi.Length / 1024);
                tvFileModelNew.FileInfo = TaskRunnerServiceRes.FileName + "[" + fi.Name + "]\r\n" + TaskRunnerServiceRes.FileType + "[" + fi.Extension + "]\r\n";
                tvFileModelNew.FileCreatedDate_UTC = fi.LastWriteTimeUtc;
                tvFileModelNew.ServerFilePath = (fi.DirectoryName + @"\").Replace(@"C:\", @"E:\");
                tvFileModelNew.LastUpdateDate_UTC = DateTime.UtcNow;
                tvFileModelNew.LastUpdateContactTVItemID = _BWObj.appTaskModel.LastUpdateContactTVItemID;

                TVFileModel tvFileModelRet = tvFileService.PostAddTVFileDB(tvFileModelNew);
                if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                    return;
                }
            }
            else
            {
                tvFileModel.FileCreatedDate_UTC = fi.LastWriteTimeUtc;
                tvFileModel.LastUpdateDate_UTC = DateTime.UtcNow;
                tvFileModel.LastUpdateContactTVItemID = _BWObj.appTaskModel.LastUpdateContactTVItemID;

                TVFileModel tvFileModelRet = tvFileService.PostUpdateTVFileDB(tvFileModel);
                if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                    return;
                }
            }

            return;
        }
        public void UpdateStatusOfRunningScenarios()
        {
            string NotUsed = "";
            List<int> appTaskTVItemListRemoved = new List<int>();

            AppTaskService appTaskService = new AppTaskService(LanguageEnum.en, _TaskRunnerBaseService._User);
            List<AppTaskModel> appTaskModelList = appTaskService.GetAppTaskModelListOfRunningMikeScenariosDB();

            foreach (AppTaskModel appTaskModel in appTaskModelList)
            {
                string Parameters = appTaskModel.Parameters;
                string[] ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                int MikeScenarioTVItemID = 0;
                int Id = 0;
                string ProjFileName = "";
                string HydroFileName = "";
                string TransFileName = "";
                foreach (string s in ParamValueList)
                {
                    string[] ParamValue = s.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                    if (ParamValue.Length != 2)
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, TaskRunnerServiceRes.Parameters);
                        SendErrorTextToDB(appTaskModel.AppTaskID, GetTextLanguageFormat1List("CouldNotParse_Properly", TaskRunnerServiceRes.Parameters));
                        return;
                    }

                    if (ParamValue[0] == "MikeScenarioTVItemID")
                    {
                        MikeScenarioTVItemID = int.Parse(ParamValue[1]);
                    }
                    else if (ParamValue[0] == "Id")
                    {
                        Id = int.Parse(ParamValue[1]);
                    }
                    else if (ParamValue[0] == "Proj")
                    {
                        ProjFileName = ParamValue[1];
                    }
                    else if (ParamValue[0] == "Hydro")
                    {
                        HydroFileName = ParamValue[1];
                    }
                    else if (ParamValue[0] == "Trans")
                    {
                        TransFileName = ParamValue[1];
                    }
                    else
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, ParamValue[0]);
                        SendErrorTextToDB(appTaskModel.AppTaskID, GetTextLanguageFormat1List("CouldNotFind_", ParamValue[0].ToString()));
                        return;
                    }
                }

                if (MikeScenarioTVItemID == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBeNullOrEmpty, "MikeScenarioTVItemID");
                    SendErrorTextToDB(appTaskModel.AppTaskID, GetTextLanguageFormat1List("_ShouldNotBeNullOrEmpty", "MikeScenarioTVItemID"));
                    return;
                }

                if (Id == 0)
                {
                    //NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBe0, "Id");
                    //appTaskService.AppTaskUpdateText(at.AppTaskID, AppTaskTextEnum.Error, GetTextLanguageFormat1List("_ShouldNotBe0", "Id"));
                    return;
                }

                if (appTaskModel.AppTaskCommand == AppTaskCommandEnum.MikeScenarioToCancel)
                {
                    try
                    {
                        Process p = Process.GetProcessById(Id);
                        p.Kill();
                    }
                    catch (Exception)
                    {
                        // nothing
                    }

                    SendTaskCommandToDB(appTaskModel.AppTaskID, AppTaskCommandEnum.MikeScenarioRunning);
                }

                if (string.IsNullOrEmpty(ProjFileName))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBeNullOrEmpty, "ProjFileName");
                    SendErrorTextToDB(appTaskModel.AppTaskID, GetTextLanguageFormat1List("_ShouldNotBeNullOrEmpty", "ProjFileName"));
                    return;
                }

                if (string.IsNullOrEmpty(HydroFileName) && string.IsNullOrEmpty(TransFileName))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBeNullOrEmpty, "Both HydroFileName and TransFileName");
                    SendErrorTextToDB(appTaskModel.AppTaskID, GetTextLanguageFormat1List("_ShouldNotBeNullOrEmpty", "Both HydroFileName and TransFileName"));
                    return;
                }

                MikeScenarioService mikeScenarioService = new MikeScenarioService(LanguageEnum.en, _TaskRunnerBaseService._User);
                MikeScenarioModel mikeScenarioModel = mikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(MikeScenarioTVItemID);
                if (!string.IsNullOrWhiteSpace(mikeScenarioModel.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, "MikeScenario", "MikeScenarioTVItemID", MikeScenarioTVItemID.ToString());
                    SendErrorTextToDB(appTaskModel.AppTaskID, GetTextLanguageFormat3List("CouldNotFind_With_Equal_", "MikeScenario", "MikeScenarioTVItemID", MikeScenarioTVItemID.ToString()));
                    return;
                }

                int TransFileSize = (int)mikeScenarioModel.EstimatedTransFileSize;
                int HydroFileSize = (int)mikeScenarioModel.EstimatedHydroFileSize;


                if (!string.IsNullOrEmpty(HydroFileName))
                {
                    if (HydroFileSize != 0)
                    {
                        FileInfo fi = new FileInfo(HydroFileName);
                        if (!fi.Exists)
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fi.FullName);
                            SendErrorTextToDB(appTaskModel.AppTaskID, GetTextLanguageFormat1List("CouldNotFindFile_", fi.FullName));
                            return;
                        }
                        int Percent = (int)(fi.Length * 100 / HydroFileSize);
                        Percent = (Percent < 100 ? Percent : 99);
                        SendPercentToDB(appTaskModel.AppTaskID, Percent);
                    }
                }
                else
                {
                    if (TransFileSize != 0)
                    {
                        FileInfo fi = new FileInfo(TransFileName);
                        if (!fi.Exists)
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fi.FullName);
                            SendErrorTextToDB(appTaskModel.AppTaskID, GetTextLanguageFormat1List("CouldNotFindFile_", fi.FullName));
                            return;
                        }
                        int Percent = (int)(fi.Length * 100 / TransFileSize);
                        Percent = (Percent < 100 ? Percent : 99);
                        SendPercentToDB(appTaskModel.AppTaskID, Percent);
                    }
                }

                List<Process> processList = Process.GetProcesses().ToList<Process>();

                bool IdFound = false;
                foreach (Process p in processList.Where(c => c.ProcessName == "MzLaunch"))
                {
                    if (p.Id == Id)
                    {
                        IdFound = true;
                    }
                }

                if (!IdFound)
                {
                    AppTaskModel appTaskModelToRemove = appTaskService.GetAppTaskModelWithAppTaskIDDB(appTaskModel.AppTaskID);
                    if (string.IsNullOrWhiteSpace(appTaskModelToRemove.Error))
                    {
                        // should Check if the log file shows the normal completion line as the last line
                        string LogFileName = ProjFileName.Substring(0, ProjFileName.LastIndexOf(".")) + ".log";

                        FileInfo fiLog = new FileInfo(LogFileName);

                        if (!fiLog.Exists)
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiLog.FullName);
                            SendErrorTextToDB(appTaskModel.AppTaskID, GetTextLanguageFormat1List("CouldNotFindFile_", fiLog.FullName));
                            return;
                        }

                        StringBuilder sb = new StringBuilder();
                        try
                        {
                            // file might not be available
                            StreamReader sr = fiLog.OpenText();
                            sb.Append(sr.ReadToEnd());
                            sr.Close();
                        }
                        catch (Exception)
                        {
                            return;
                        }

                        if (sb.ToString().Contains("Run cancelled by user"))
                        {
                            mikeScenarioModel.ScenarioStatus = ScenarioStatusEnum.Error;
                            NotUsed = TaskRunnerServiceRes.RunCancelledByUser;
                            SendErrorTextToDB(appTaskModel.AppTaskID, GetTextLanguageList("RunCancelledByUser"));
                            mikeScenarioModel.ErrorInfo = TaskRunnerServiceRes.RunCancelledByUser;
                        }
                        else if (sb.ToString().Contains("Abnormal run completion"))
                        {
                            mikeScenarioModel.ScenarioStatus = ScenarioStatusEnum.Error;
                            NotUsed = TaskRunnerServiceRes.AbnormalRunCompletion;
                            SendErrorTextToDB(appTaskModel.AppTaskID, GetTextLanguageList("AbnormalRunCompletion"));
                            mikeScenarioModel.ErrorInfo = TaskRunnerServiceRes.AbnormalRunCompletion;
                        }
                        else if (sb.ToString().Contains("Normal run completion"))
                        {
                            mikeScenarioModel.ScenarioStatus = ScenarioStatusEnum.Completed;
                            mikeScenarioModel.ErrorInfo = TaskRunnerServiceRes.NormalRunCompletion;
                        }
                        else
                        {
                            mikeScenarioModel.ScenarioStatus = ScenarioStatusEnum.Error;
                            NotUsed = TaskRunnerServiceRes.UnknownErrorPleaseCheckLogFile;
                            SendErrorTextToDB(appTaskModel.AppTaskID, GetTextLanguageList("UnknownErrorPleaseCheckLogFile"));
                            mikeScenarioModel.ErrorInfo = TaskRunnerServiceRes.UnknownErrorPleaseCheckLogFile;
                        }

                        MikeScenarioModel mikeScenarioModelRet = mikeScenarioService.PostUpdateMikeScenarioDB(mikeScenarioModel);
                        if (!string.IsNullOrWhiteSpace(mikeScenarioModelRet.Error))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.MikeScenario, mikeScenarioModelRet.Error);
                            SendErrorTextToDB(appTaskModel.AppTaskID, GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.MikeScenario, mikeScenarioModelRet.Error));
                            return;
                        }

                        AppTaskModel appTaskModel2 = appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        if (!string.IsNullOrWhiteSpace(appTaskModel2.Error))
                        {
                            // do something
                        }
                    }
                }

            }

        }
        #endregion Functions public

        #region Functions private
        #endregion Functions private

    }
}
