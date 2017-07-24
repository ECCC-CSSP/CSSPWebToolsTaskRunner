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
using System.Net.Mail;
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;
using CSSPLabSheetParserDLL.Services;
using CSSPFCFormWriterDLL.Services;

namespace CSSPWebToolsTaskRunner.Services
{
    public class DocxService
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
        #endregion Properties

        #region Constructors
        public DocxService(TaskRunnerBaseService taskRunnerBaseService)
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
        }
        #endregion Constructors

        #region Functions public
        public void CreateDocumentFromTemplateDocx(DocTemplateModel docTemplateModel)
        {
            string NotUsed = "";

            TVFileModel tvFileModelTemplate = _TVFileService.GetTVFileModelWithTVFileTVItemIDDB(docTemplateModel.TVFileTVItemID);
            if (!string.IsNullOrWhiteSpace(tvFileModelTemplate.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.TVFileTVItemID, docTemplateModel.TVFileTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.TVFileTVItemID, docTemplateModel.TVFileTVItemID.ToString());
                return;
            }

            string ServerNewFilePath = _TVFileService.GetServerFilePath(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);

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

            string retStr = ParseDocx(fi);
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
        public void GenerateFCFormDOCX()
        {
            string NotUsed = "";
            string Parameters = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters;
            int LabSheetID = int.Parse(_AppTaskService.GetAppTaskParamStr(Parameters, "LabSheetID"));

            string ServerPath = _TVFileService.GetServerFilePath(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);

            FileInfo fi = new FileInfo(_TVFileService.ChoseEDriveOrCDrive(ServerPath + _AppTaskService.GetAppTaskParamStr(Parameters, "FileName")));

            LabSheetModel labSheetModel = _LabSheetService.GetLabSheetModelWithLabSheetIDDB(LabSheetID);
            if (!string.IsNullOrWhiteSpace(labSheetModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.LabSheet, TaskRunnerServiceRes.LabSheetID, LabSheetID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.LabSheet, TaskRunnerServiceRes.LabSheetID, LabSheetID.ToString());
                return;
            }

            TVItemModel tvItemModelFile = new TVItemModel();
            TVFileModel tvFileModel = _TVFileService.GetTVFileModelWithServerFilePathAndServerFileNameDB(fi.Directory + @"\", fi.Name);
            if (string.IsNullOrWhiteSpace(tvFileModel.Error))
            {
                // Exist
                tvItemModelFile = _TVItemService.GetTVItemModelWithTVItemIDDB(tvFileModel.TVFileTVItemID);
                if (!string.IsNullOrWhiteSpace(tvItemModelFile.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, tvFileModel.TVFileTVItemID.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, tvFileModel.TVFileTVItemID.ToString());
                    return;
                }
            }
            else
            {
                tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                    return;
            }

            string FullLabSheetText = labSheetModel.FileContent;

            CSSPLabSheetParser csspLabSheetParser = new CSSPLabSheetParser();

            LabSheetA1Sheet labSheetA1Sheet = csspLabSheetParser.ParseLabSheetA1(FullLabSheetText);
            if (!string.IsNullOrWhiteSpace(labSheetA1Sheet.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParseLabSheetWithLabSheetID_, LabSheetID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List(TaskRunnerServiceRes.CouldNotParseLabSheetWithLabSheetID_, fi.Name);
                return;
            }

            if (labSheetA1Sheet.IncludeLaboratoryQAQC)
            {
                CSSPFCFormWriter csspFCFormWriter = new CSSPFCFormWriter(_TaskRunnerBaseService._BWObj.appTaskModel.Language, FullLabSheetText);
                string retStr = csspFCFormWriter.CreateFCForm(fi.FullName);
                if (!string.IsNullOrWhiteSpace(retStr))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateFile_, fi.Name);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List(TaskRunnerServiceRes.CouldNotCreateFile_, fi.Name);
                    return;
                }

                fi = new FileInfo(fi.FullName);

                _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.FCFormAutoGenerate, FilePurposeEnum.GeneratedFCForm);
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                    return;

            }

            DateTime RunDate = new DateTime(int.Parse(labSheetA1Sheet.RunYear), int.Parse(labSheetA1Sheet.RunMonth), int.Parse(labSheetA1Sheet.RunDay));
            DateTime SalinityReadDate = new DateTime(int.Parse(labSheetA1Sheet.SalinitiesReadYear), int.Parse(labSheetA1Sheet.SalinitiesReadMonth), int.Parse(labSheetA1Sheet.SalinitiesReadDay));
            DateTime ResultsReadDate = new DateTime(int.Parse(labSheetA1Sheet.ResultsReadYear), int.Parse(labSheetA1Sheet.ResultsReadMonth), int.Parse(labSheetA1Sheet.ResultsReadDay));
            DateTime ResultsRecordedDate = new DateTime(int.Parse(labSheetA1Sheet.ResultsRecordedYear), int.Parse(labSheetA1Sheet.ResultsRecordedMonth), int.Parse(labSheetA1Sheet.ResultsRecordedDay));

            float TCFirst = 0.0f;
            try
            {
                float.Parse(labSheetA1Sheet.TCFirst);
            }
            catch (Exception)
            {
                TCFirst = (float)(from c in labSheetA1Sheet.LabSheetA1MeasurementList
                                  where c.Temperature != null
                                  select c.Temperature).FirstOrDefault();

            }

            float TCAverage = 0.0f;
            try
            {
                float.Parse(labSheetA1Sheet.TCAverage);
            }
            catch (Exception)
            {
                TCAverage = (float)(from c in labSheetA1Sheet.LabSheetA1MeasurementList
                                    where c.Temperature != null
                                    select c.Temperature).Average();

            }
            LabSheetDetailModel labSheetDetailModelNew = new LabSheetDetailModel();
            labSheetDetailModelNew.LabSheetID = LabSheetID;
            labSheetDetailModelNew.SamplingPlanID = labSheetModel.SamplingPlanID;
            labSheetDetailModelNew.SubsectorTVItemID = labSheetModel.SubsectorTVItemID;
            labSheetDetailModelNew.Version = labSheetA1Sheet.Version;
            labSheetDetailModelNew.RunDate = RunDate;
            labSheetDetailModelNew.Tides = labSheetA1Sheet.Tides;
            labSheetDetailModelNew.SampleCrewInitials = labSheetA1Sheet.SampleCrewInitials;
            labSheetDetailModelNew.WaterBathCount = labSheetA1Sheet.WaterBathCount;
            if (labSheetA1Sheet.IncludeLaboratoryQAQC)
            {
                if (labSheetDetailModelNew.WaterBathCount > 0)
                {
                    labSheetDetailModelNew.IncubationBath1StartTime = RunDate.AddHours(int.Parse(labSheetA1Sheet.IncubationBath1StartTime.Substring(0, 2))).AddMinutes(int.Parse(labSheetA1Sheet.IncubationBath1StartTime.Substring(3, 2)));
                    labSheetDetailModelNew.IncubationBath1EndTime = RunDate.AddHours(int.Parse(labSheetA1Sheet.IncubationBath1EndTime.Substring(0, 2))).AddMinutes(int.Parse(labSheetA1Sheet.IncubationBath1EndTime.Substring(3, 2)));
                    labSheetDetailModelNew.IncubationBath1TimeCalculated_minutes = int.Parse(labSheetA1Sheet.IncubationBath1TimeCalculated.Substring(0, 2)) * 60 + int.Parse(labSheetA1Sheet.IncubationBath1TimeCalculated.Substring(3, 2));
                    labSheetDetailModelNew.WaterBath1 = labSheetA1Sheet.WaterBath1;
                }
                if (labSheetDetailModelNew.WaterBathCount > 1)
                {
                    labSheetDetailModelNew.IncubationBath2StartTime = RunDate.AddHours(int.Parse(labSheetA1Sheet.IncubationBath2StartTime.Substring(0, 2))).AddMinutes(int.Parse(labSheetA1Sheet.IncubationBath2StartTime.Substring(3, 2)));
                    labSheetDetailModelNew.IncubationBath2EndTime = RunDate.AddHours(int.Parse(labSheetA1Sheet.IncubationBath2EndTime.Substring(0, 2))).AddMinutes(int.Parse(labSheetA1Sheet.IncubationBath2EndTime.Substring(3, 2)));
                    labSheetDetailModelNew.IncubationBath2TimeCalculated_minutes = int.Parse(labSheetA1Sheet.IncubationBath2TimeCalculated.Substring(0, 2)) * 60 + int.Parse(labSheetA1Sheet.IncubationBath2TimeCalculated.Substring(3, 2));
                    labSheetDetailModelNew.WaterBath2 = labSheetA1Sheet.WaterBath2;
                }
                if (labSheetDetailModelNew.WaterBathCount > 2)
                {
                    labSheetDetailModelNew.IncubationBath3StartTime = RunDate.AddHours(int.Parse(labSheetA1Sheet.IncubationBath3StartTime.Substring(0, 2))).AddMinutes(int.Parse(labSheetA1Sheet.IncubationBath3StartTime.Substring(3, 2)));
                    labSheetDetailModelNew.IncubationBath3EndTime = RunDate.AddHours(int.Parse(labSheetA1Sheet.IncubationBath3EndTime.Substring(0, 2))).AddMinutes(int.Parse(labSheetA1Sheet.IncubationBath3EndTime.Substring(3, 2)));
                    labSheetDetailModelNew.IncubationBath3TimeCalculated_minutes = int.Parse(labSheetA1Sheet.IncubationBath3TimeCalculated.Substring(0, 2)) * 60 + int.Parse(labSheetA1Sheet.IncubationBath3TimeCalculated.Substring(3, 2));
                    labSheetDetailModelNew.WaterBath3 = labSheetA1Sheet.WaterBath3;
                }
                if (!string.IsNullOrWhiteSpace(labSheetA1Sheet.TCField1))
                {
                    labSheetDetailModelNew.TCField1 = float.Parse(labSheetA1Sheet.TCField1);
                }
                else
                {
                    labSheetDetailModelNew.TCField1 = null;
                }
                if (!string.IsNullOrWhiteSpace(labSheetA1Sheet.TCLab1))
                {
                    labSheetDetailModelNew.TCLab1 = float.Parse(labSheetA1Sheet.TCLab1);
                }
                else
                {
                    labSheetDetailModelNew.TCLab1 = null;
                }
                if (!string.IsNullOrWhiteSpace(labSheetA1Sheet.TCField2))
                {
                    labSheetDetailModelNew.TCField2 = float.Parse(labSheetA1Sheet.TCField2);
                }
                else
                {
                    labSheetDetailModelNew.TCField2 = null;
                }
                if (!string.IsNullOrWhiteSpace(labSheetA1Sheet.TCLab2))
                {
                    labSheetDetailModelNew.TCLab2 = float.Parse(labSheetA1Sheet.TCLab2);
                }
                else
                {
                    labSheetDetailModelNew.TCLab2 = null;
                }
                if (!string.IsNullOrWhiteSpace(labSheetA1Sheet.TCFirst))
                {
                    labSheetDetailModelNew.TCFirst = float.Parse(labSheetA1Sheet.TCFirst);
                }
                else
                {
                    labSheetDetailModelNew.TCFirst = null;
                }
                if (!string.IsNullOrWhiteSpace(labSheetA1Sheet.TCAverage))
                {
                    labSheetDetailModelNew.TCAverage = float.Parse(labSheetA1Sheet.TCAverage);
                }
                else
                {
                    labSheetDetailModelNew.TCAverage = null;
                }
                labSheetDetailModelNew.ControlLot = labSheetA1Sheet.ControlLot;
                labSheetDetailModelNew.Positive35 = labSheetA1Sheet.Positive35;
                labSheetDetailModelNew.NonTarget35 = labSheetA1Sheet.NonTarget35;
                labSheetDetailModelNew.Negative35 = labSheetA1Sheet.Negative35;
                if (labSheetDetailModelNew.WaterBathCount > 0)
                {
                    labSheetDetailModelNew.Bath1Positive44_5 = labSheetA1Sheet.Bath1Positive44_5;
                    labSheetDetailModelNew.Bath1NonTarget44_5 = labSheetA1Sheet.Bath1NonTarget44_5;
                    labSheetDetailModelNew.Bath1Negative44_5 = labSheetA1Sheet.Bath1Negative44_5;
                    labSheetDetailModelNew.Bath1Blank44_5 = labSheetA1Sheet.Bath1Blank44_5;
                }
                if (labSheetDetailModelNew.WaterBathCount > 1)
                {
                    labSheetDetailModelNew.Bath2Positive44_5 = labSheetA1Sheet.Bath2Positive44_5;
                    labSheetDetailModelNew.Bath2NonTarget44_5 = labSheetA1Sheet.Bath2NonTarget44_5;
                    labSheetDetailModelNew.Bath2Negative44_5 = labSheetA1Sheet.Bath2Negative44_5;
                    labSheetDetailModelNew.Bath2Blank44_5 = labSheetA1Sheet.Bath2Blank44_5;
                }
                if (labSheetDetailModelNew.WaterBathCount > 2)
                {
                    labSheetDetailModelNew.Bath3Positive44_5 = labSheetA1Sheet.Bath3Positive44_5;
                    labSheetDetailModelNew.Bath3NonTarget44_5 = labSheetA1Sheet.Bath3NonTarget44_5;
                    labSheetDetailModelNew.Bath3Negative44_5 = labSheetA1Sheet.Bath3Negative44_5;
                    labSheetDetailModelNew.Bath3Blank44_5 = labSheetA1Sheet.Bath3Blank44_5;
                }
                labSheetDetailModelNew.Blank35 = labSheetA1Sheet.Blank35;
                labSheetDetailModelNew.Lot35 = labSheetA1Sheet.Lot35;
                labSheetDetailModelNew.Lot44_5 = labSheetA1Sheet.Lot44_5;

                labSheetDetailModelNew.SampleBottleLotNumber = labSheetA1Sheet.SampleBottleLotNumber;
                labSheetDetailModelNew.SalinitiesReadBy = labSheetA1Sheet.SalinitiesReadBy;
                labSheetDetailModelNew.SalinitiesReadDate = SalinityReadDate;
                labSheetDetailModelNew.ResultsReadBy = labSheetA1Sheet.ResultsReadBy;
                labSheetDetailModelNew.ResultsReadDate = ResultsReadDate;
                labSheetDetailModelNew.ResultsRecordedBy = labSheetA1Sheet.ResultsRecordedBy;
                labSheetDetailModelNew.ResultsRecordedDate = ResultsRecordedDate;
                float tempFloat = 0.0f;
                if (float.TryParse(labSheetA1Sheet.DailyDuplicateRLog, out tempFloat))
                {
                    labSheetDetailModelNew.DailyDuplicateRLog = float.Parse(labSheetA1Sheet.DailyDuplicateRLog);
                }
                else
                {
                    labSheetDetailModelNew.DailyDuplicateRLog = null;
                }
                labSheetDetailModelNew.DailyDuplicatePrecisionCriteria = float.Parse(labSheetA1Sheet.DailyDuplicatePrecisionCriteria);
                labSheetDetailModelNew.DailyDuplicateAcceptable = (labSheetA1Sheet.DailyDuplicateAcceptableOrUnacceptable == "Acceptable" ? true : false);
                tempFloat = 0.0f;
                if (float.TryParse(labSheetA1Sheet.IntertechDuplicateRLog, out tempFloat))
                {
                    labSheetDetailModelNew.IntertechDuplicateRLog = float.Parse(labSheetA1Sheet.IntertechDuplicateRLog);
                }
                else
                {
                    labSheetDetailModelNew.IntertechDuplicateRLog = null;
                }
                labSheetDetailModelNew.IntertechDuplicatePrecisionCriteria = float.Parse(labSheetA1Sheet.IntertechDuplicatePrecisionCriteria);
                labSheetDetailModelNew.IntertechDuplicateAcceptable = (labSheetA1Sheet.IntertechDuplicateAcceptableOrUnacceptable == "Acceptable" ? true : false);
                labSheetDetailModelNew.IntertechReadAcceptable = (labSheetA1Sheet.IntertechReadAcceptableOrUnacceptable == "Acceptable" ? true : false);
            }

            labSheetDetailModelNew.RunComment = labSheetA1Sheet.RunComment;
            labSheetDetailModelNew.RunWeatherComment = labSheetA1Sheet.RunWeatherComment;

            LabSheetDetailModel labSheetDetailModelExist = _LabSheetDetailService.GetLabSheetDetailModelExistDB(labSheetDetailModelNew);
            if (!string.IsNullOrWhiteSpace(labSheetDetailModelExist.Error))
            {
                labSheetDetailModelExist = _LabSheetDetailService.PostAddLabSheetDetailDB(labSheetDetailModelNew);
                if (!string.IsNullOrWhiteSpace(labSheetDetailModelExist.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.LabSheetDetail, labSheetDetailModelExist.Error);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.LabSheetDetail, labSheetDetailModelExist.Error);
                    return;
                }
            }
            else
            {
                labSheetDetailModelNew.LabSheetDetailID = labSheetDetailModelExist.LabSheetDetailID;
                labSheetDetailModelExist = _LabSheetDetailService.PostUpdateLabSheetDetailDB(labSheetDetailModelNew);
                if (!string.IsNullOrWhiteSpace(labSheetDetailModelExist.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.LabSheetDetail, labSheetDetailModelExist.Error);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.LabSheetDetail, labSheetDetailModelExist.Error);
                    return;
                }
            }

            foreach (LabSheetA1Measurement LabSheetA1Measurement in labSheetA1Sheet.LabSheetA1MeasurementList)
            {
                if (LabSheetA1Measurement.TVItemID == 0)
                    continue;

                if (LabSheetA1Measurement.MPN == null)
                    continue;

                if (LabSheetA1Measurement.Time == null)
                    continue;

                LabSheetTubeMPNDetailModel labSheetTubeMPNDetailModelNew = new LabSheetTubeMPNDetailModel();
                labSheetTubeMPNDetailModelNew.LabSheetDetailID = labSheetDetailModelExist.LabSheetDetailID;
                labSheetTubeMPNDetailModelNew.MWQMSiteTVItemID = LabSheetA1Measurement.TVItemID;
                labSheetTubeMPNDetailModelNew.SampleDateTime = (DateTime)LabSheetA1Measurement.Time;
                labSheetTubeMPNDetailModelNew.MPN = (int)LabSheetA1Measurement.MPN;
                labSheetTubeMPNDetailModelNew.Tube10 = (int)LabSheetA1Measurement.Tube10;
                labSheetTubeMPNDetailModelNew.Tube1_0 = (int)LabSheetA1Measurement.Tube1_0;
                labSheetTubeMPNDetailModelNew.Tube0_1 = (int)LabSheetA1Measurement.Tube0_1;
                labSheetTubeMPNDetailModelNew.Salinity = (float)LabSheetA1Measurement.Salinity;
                labSheetTubeMPNDetailModelNew.Temperature = (float)LabSheetA1Measurement.Temperature;
                labSheetTubeMPNDetailModelNew.ProcessedBy = LabSheetA1Measurement.ProcessedBy;
                labSheetTubeMPNDetailModelNew.SampleType = (SampleTypeEnum)LabSheetA1Measurement.SampleType;
                labSheetTubeMPNDetailModelNew.SiteComment = LabSheetA1Measurement.SiteComment;

                LabSheetTubeMPNDetailModel labSheetTubeMPNDetailModelExist = _LabSheetTubeMPNDetailService.GetLabSheetTubeMPNDetailModelExistDB(labSheetTubeMPNDetailModelNew);
                if (!string.IsNullOrWhiteSpace(labSheetTubeMPNDetailModelExist.Error))
                {
                    labSheetTubeMPNDetailModelExist = _LabSheetTubeMPNDetailService.PostAddLabSheetTubeMPNDetailDB(labSheetTubeMPNDetailModelNew);
                    if (!string.IsNullOrWhiteSpace(labSheetTubeMPNDetailModelExist.Error))
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.LabSheetTubeMPNDetail, labSheetTubeMPNDetailModelExist.Error);
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.LabSheetTubeMPNDetail, labSheetTubeMPNDetailModelExist.Error);
                        return;
                    }
                }
                else
                {
                    labSheetTubeMPNDetailModelNew.LabSheetTubeMPNDetailID = labSheetTubeMPNDetailModelExist.LabSheetTubeMPNDetailID;
                    labSheetTubeMPNDetailModelExist = _LabSheetTubeMPNDetailService.PostUpdateLabSheetTubeMPNDetailDB(labSheetTubeMPNDetailModelNew);
                    if (!string.IsNullOrWhiteSpace(labSheetTubeMPNDetailModelExist.Error))
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.LabSheetTubeMPNDetail, labSheetTubeMPNDetailModelExist.Error);
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.LabSheetTubeMPNDetail, labSheetTubeMPNDetailModelExist.Error);
                        return;
                    }
                }
            }

            SendLabSheetAcceptedEmail(labSheetModel.LabSheetID);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

        }
        //public void GenerateRootDOCX() 
        //{
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Root, FileGeneratorTypeEnum.Word);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    DocxServiceRoot docxServiceRoot = new DocxServiceRoot(_TaskRunnerBaseService);
        //    docxServiceRoot.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.RootFileAutoGenerate, FilePurposeEnum.Generated);
        //}
        //public void GenerateRootDOCXAndPDF()
        //{
        //    string NotUsed = "";

        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Root, FileGeneratorTypeEnum.WordAndPDF);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    DocxServiceRoot docxServiceRoot = new DocxServiceRoot(_TaskRunnerBaseService);
        //    docxServiceRoot.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.RootFileAutoGenerate, FilePurposeEnum.Generated);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    NotUsed = TaskRunnerServiceRes.ConvertingFileToPDF;
        //    _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("ConvertingFileToPDF"));

        //    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 50);

        //    GeneratePDF(fi, TVTypeEnum.Root);
        //}
        //public void GenerateCountryDOCX()
        //{
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Country, FileGeneratorTypeEnum.Word);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    DocxServiceCountry docxServiceCountry = new DocxServiceCountry(_TaskRunnerBaseService);
        //    docxServiceCountry.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.ProvinceFileAutoGenerate, FilePurposeEnum.Generated);
        //}
        //public void GenerateCountryDOCXAndPDF()
        //{
        //    string NotUsed = "";

        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Country, FileGeneratorTypeEnum.WordAndPDF);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    DocxServiceCountry docxServiceCountry = new DocxServiceCountry(_TaskRunnerBaseService);
        //    docxServiceCountry.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.ProvinceFileAutoGenerate, FilePurposeEnum.Generated);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    NotUsed = TaskRunnerServiceRes.ConvertingFileToPDF;
        //    _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("ConvertingFileToPDF"));

        //    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 50);

        //    GeneratePDF(fi, TVTypeEnum.Country);
        //}
        //public void GenerateProvinceDOCX()
        //{
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Province, FileGeneratorTypeEnum.Word);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    DocxServiceProvince docxServiceProvince = new DocxServiceProvince(_TaskRunnerBaseService);
        //    docxServiceProvince.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.ProvinceFileAutoGenerate, FilePurposeEnum.Generated);
        //}
        //public void GenerateProvinceDOCXAndPDF()
        //{
        //    string NotUsed = "";

        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Province, FileGeneratorTypeEnum.WordAndPDF);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    DocxServiceProvince docxServiceProvince = new DocxServiceProvince(_TaskRunnerBaseService);
        //    docxServiceProvince.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.ProvinceFileAutoGenerate, FilePurposeEnum.Generated);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    NotUsed = TaskRunnerServiceRes.ConvertingFileToPDF;
        //    _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("ConvertingFileToPDF"));

        //    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 50);

        //    GeneratePDF(fi, TVTypeEnum.Province);
        //}
        //public void GenerateAreaDOCX()
        //{
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Area, FileGeneratorTypeEnum.Word);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    DocxServiceArea _DocxService_Area = new DocxServiceArea(_TaskRunnerBaseService);
        //    _DocxService_Area.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.AreaFileAutoGenerate, FilePurposeEnum.Generated);
        //}
        //public void GenerateAreaDOCXAndPDF()
        //{
        //    string NotUsed = "";

        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Area, FileGeneratorTypeEnum.WordAndPDF);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    DocxServiceArea docxServiceArea = new DocxServiceArea(_TaskRunnerBaseService);
        //    docxServiceArea.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.AreaFileAutoGenerate, FilePurposeEnum.Generated);

        //    NotUsed = TaskRunnerServiceRes.ConvertingFileToPDF;
        //    _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("ConvertingFileToPDF"));

        //    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 50);

        //    GeneratePDF(fi, TVTypeEnum.Area);
        //}
        //public void GenerateSectorDOCX()
        //{
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Sector, FileGeneratorTypeEnum.Word);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    DocxServiceSector docxServiceSector = new DocxServiceSector(_TaskRunnerBaseService);
        //    docxServiceSector.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.SectorFileAutoGenerate, FilePurposeEnum.Generated);
        //}
        //public void GenerateSectorDOCXAndPDF()
        //{
        //    string NotUsed = "";

        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Sector, FileGeneratorTypeEnum.WordAndPDF);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    DocxServiceSector docxServiceSector = new DocxServiceSector(_TaskRunnerBaseService);
        //    docxServiceSector.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.SectorFileAutoGenerate, FilePurposeEnum.Generated);

        //    NotUsed = TaskRunnerServiceRes.ConvertingFileToPDF;
        //    _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("ConvertingFileToPDF"));

        //    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 50);

        //    GeneratePDF(fi, TVTypeEnum.Sector);
        //}
        //public void GenerateSubsectorDOCX()
        //{
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Subsector, FileGeneratorTypeEnum.Word);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    DocxServiceSubsector docxServiceSubsector = new DocxServiceSubsector(_TaskRunnerBaseService);
        //    docxServiceSubsector.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.SubsectorFileAutoGenerate, FilePurposeEnum.Generated);
        //}
        //public void GenerateSubsectorDOCXAndPDF()
        //{
        //    string NotUsed = "";

        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Subsector, FileGeneratorTypeEnum.WordAndPDF);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    DocxServiceSubsector docxServiceSubsector = new DocxServiceSubsector(_TaskRunnerBaseService);
        //    docxServiceSubsector.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.SubsectorFileAutoGenerate, FilePurposeEnum.Generated);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    NotUsed = TaskRunnerServiceRes.ConvertingFileToPDF;
        //    _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("ConvertingFileToPDF"));

        //    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 50);

        //    GeneratePDF(fi, TVTypeEnum.Subsector);
        //}
        //public void GenerateMunicipalityDOCX()
        //{
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Municipality, FileGeneratorTypeEnum.Word);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    DocxServiceMunicipality docxServiceMunicipality = new DocxServiceMunicipality(_TaskRunnerBaseService);
        //    docxServiceMunicipality.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.MunicipalityFileAutoGenerate, FilePurposeEnum.Generated);
        //}
        //public void GenerateMunicipalityDOCXAndPDF()
        //{
        //    string NotUsed = "";

        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.Municipality, FileGeneratorTypeEnum.WordAndPDF);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    DocxServiceMunicipality docxServiceMunicipality = new DocxServiceMunicipality(_TaskRunnerBaseService);
        //    docxServiceMunicipality.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.MunicipalityFileAutoGenerate, FilePurposeEnum.Generated);

        //    NotUsed = TaskRunnerServiceRes.ConvertingFileToPDF;
        //    _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("ConvertingFileToPDF"));

        //    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 50);

        //    GeneratePDF(fi, TVTypeEnum.Municipality);
        //}
        //public void GenerateSubsectorFaecalColiformSummaryStatDOCX()
        //{
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.SubsectorFaecalColiformSummaryStat, FileGeneratorTypeEnum.Word);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    DocxServiceSubsectorFaecalColiformSummaryStat docxServiceSubsectorFaecalColiformSummaryStat = new DocxServiceSubsectorFaecalColiformSummaryStat(_TaskRunnerBaseService);
        //    docxServiceSubsectorFaecalColiformSummaryStat.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.SubsectorFileAutoGenerate, FilePurposeEnum.Generated);
        //}
        //public void GenerateSubsectorFaecalColiformDensitiesDOCX()
        //{
        //    _TaskRunnerBaseService.generateDocParams = _TaskRunnerBaseService.CheckGenerateModelOK(FileGeneratorEnum.SubsectorFaecalColiformDensities, FileGeneratorTypeEnum.Word);

        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService.generateDocParams.TVItemID);

        //    FileInfo fi = _TaskRunnerBaseService.GetFileInfo(_TaskRunnerBaseService.generateDocParams);

        //    TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    DocxServiceSubsectorFaecalColiformDensities docxServiceSubsectorFaecalColiformDensities = new DocxServiceSubsectorFaecalColiformDensities(_TaskRunnerBaseService);
        //    docxServiceSubsectorFaecalColiformDensities.Generate(fi);

        //    _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.SubsectorFileAutoGenerate, FilePurposeEnum.Generated);
        //}

        public void GeneratePDF(FileInfo fi, TVTypeEnum tvType)
        {
            // Doing PDF
            Microsoft.Office.Interop.Word.Application _Word = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document _Document = _Word.Documents.Open(fi.FullName);
            string NewPDFFileName = fi.FullName.Replace(".docx", "_docx.pdf");
            fi = new FileInfo(NewPDFFileName);
            _Document.ExportAsFixedFormat(NewPDFFileName, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
            _Document.Close();

            _Word.Quit();

            TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.RootFileAutoGenerate, FilePurposeEnum.Generated);
        }

        public string ParseDocx(FileInfo fi)
        {

            return "Need to finalize ParseDocx";
        }
        public void SendLabSheetAcceptedEmail(int LabSheetID)
        {
            string NotUsed = "";
            LabSheetModel labSheetModel = _LabSheetService.GetLabSheetModelWithLabSheetIDDB(LabSheetID);
            if (!string.IsNullOrWhiteSpace(labSheetModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.LabSheet, TaskRunnerServiceRes.LabSheetID, LabSheetID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.LabSheet, TaskRunnerServiceRes.LabSheetID, LabSheetID.ToString());
                return;
            }

            TVItemModel tvItemModelSubsector = _TVItemService.GetTVItemModelWithTVItemIDDB(labSheetModel.SubsectorTVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelSubsector.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_For_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, labSheetModel.SubsectorTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotCreate_For_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, labSheetModel.SubsectorTVItemID.ToString());
                return;
            }

            List<TVItemModel> tvItemModelParents = _TVItemService.GetParentsTVItemModelList(tvItemModelSubsector.TVPath);

            TVItemModel tvItemModelProvince = tvItemModelParents.Where(c => c.TVType == TVTypeEnum.Province).FirstOrDefault();
            if (tvItemModelProvince == null)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, TaskRunnerServiceRes.Province);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", TaskRunnerServiceRes.Province);
                return;
            }

            string hrefSubsector = "http://wmon01dtchlebl2/csspwebtools/en-CA/#!View/" + (tvItemModelProvince.TVText + "-" + tvItemModelSubsector.TVText).Replace(" ", "-") + "|||" + tvItemModelSubsector.TVItemID.ToString() + "|||30010000001000000000000000001000";

            List<ContactModel> contactModelList = _ContactService.GetContactModelWithSamplingPlanner_ProvincesTVItemIDDB(tvItemModelProvince.TVItemID);

            MailMessage mail = new MailMessage();

            //mail.To.Add("Test1.User@ssctest.itsso.gc.ca");

            foreach (ContactModel contactModel in contactModelList)
            {
                mail.To.Add(contactModel.LoginEmail.ToLower());
            }

            mail.From = new MailAddress("ec.pccsm-cssp.ec@canada.ca");
            mail.IsBodyHtml = true;

            SmtpClient myClient = new System.Net.Mail.SmtpClient();

            //myClient.Host = "ssctest.itsso.gc.ca";
            myClient.Host = "atlantic-exgate.Atlantic.int.ec.gc.ca";

            DateTime RunDate = new DateTime(labSheetModel.Year, labSheetModel.Month, labSheetModel.Day);

            string subject = "ACCEPTED: Lab sheet dated " + RunDate.ToString("yyyy MMMM dd") + " for subsector " + tvItemModelSubsector.TVText;

            StringBuilder msg = new StringBuilder();
            msg.AppendLine(@"<h2>" + tvItemModelProvince.TVText + "</h2>");
            msg.AppendLine(@"<h4>Subsector: <a href=""" + hrefSubsector + @""">" + tvItemModelSubsector.TVText + "</a></h4>");
            msg.AppendLine(@"<p><b>Date of run:</b> " + RunDate.ToString("yyyy MMMM dd") + "</p>");
            msg.AppendLine(@"<p><b>Accepted by:</b> " + labSheetModel.AcceptedOrRejectedByContactTVText + "</p>");
            msg.AppendLine(@"<p>The FC form was also automatically created.</p>");
            msg.AppendLine(@"<p></p>");
            msg.AppendLine(@"Auto email from CSSPWebTools.");

            mail.Subject = subject;
            mail.Body = msg.ToString();
            myClient.Send(mail);
        }

        #endregion Function public

        #region Functions private
        #endregion
    }
}