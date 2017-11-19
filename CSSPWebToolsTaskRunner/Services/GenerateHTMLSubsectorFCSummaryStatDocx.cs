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
using System.Windows.Forms;
using System.Web.Helpers;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        private bool GenerateHTMLSubsectorFCSummaryStatDocx(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            string NotUsed = "";
            int TVItemID = 0;
            int Year = 0;
            string HideAllAllAll = "";
            string HideWetAllAll = "";
            string HideDryAllAll = "";
            string HideMaxFCColumn = "";

            Random random = new Random();
            string FileNameExtra = "";
            for (int i = 0; i < 10; i++)
            {
                FileNameExtra = FileNameExtra + (char)random.Next(97, 122);
            }

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 3);

            if (!GetTopHTML(sbHTML))
            {
                return false;
            }

            List<string> ParamValueList = parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            if (!int.TryParse(GetParameters("TVItemID", ParamValueList), out TVItemID))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind__, TaskRunnerServiceRes.Parameter, TaskRunnerServiceRes.TVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind__", TaskRunnerServiceRes.Parameter, TaskRunnerServiceRes.TVItemID);
                return false;
            }

            if (!int.TryParse(GetParameters("Year", ParamValueList), out Year))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind__, TaskRunnerServiceRes.Parameter, TaskRunnerServiceRes.Year);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind__", TaskRunnerServiceRes.Parameter, TaskRunnerServiceRes.Year);
                return false;
            }

            HideAllAllAll = GetParameters("HideAllAllAll", ParamValueList);
            HideWetAllAll = GetParameters("HideWetAllAll", ParamValueList);
            HideDryAllAll = GetParameters("HideDryAllAll", ParamValueList);
            HideMaxFCColumn = GetParameters("HideMaxFCColumn", ParamValueList);

            bool MissingRainData = _MWQMRunService.IsRainDataMissingWithSubsectorTVItemID(TVItemID);

            List<MWQMAnalysisReportParameterModel> mwqmAnalysisReportParameterModelList = _MWQMAnalysisReportParameterService.GetMWQMAnalysisReportParameterModelListWithSubsectorTVItemIDDB(TVItemID).Where(c => c.AnalysisReportYear == Year).ToList();

            if (!mwqmAnalysisReportParameterModelList.Where(c => c.AnalysisReportYear == Year && c.AnalysisCalculationType == AnalysisCalculationTypeEnum.AllAllAll).Any())
            {
                MWQMAnalysisReportParameterModel mwqmAnalysisReportParameterModel = new MWQMAnalysisReportParameterModel()
                {
                    AnalysisCalculationType = AnalysisCalculationTypeEnum.AllAllAll,
                    AnalysisName = "",
                    AnalysisReportYear = Year,
                    Command = AnalysisReportExportCommandEnum.Report,
                    DryLimit24h = 4,
                    DryLimit48h = 8,
                    DryLimit72h = 12,
                    DryLimit96h = 16,
                    EndDate = new DateTime(1900, 1, 1),
                    ExcelTVFileTVItemID = null,
                    FullYear = true,
                    LastUpdateContactTVItemID = 2, // Charles LeBlanc
                    LastUpdateDate_UTC = DateTime.Now,
                    MidRangeNumberOfDays = -6,
                    MWQMAnalysisReportParameterID = 0,
                    NumberOfRuns = 30,
                    RunsToOmit = ",",
                    SalinityHighlightDeviationFromAverage = 8,
                    ShortRangeNumberOfDays = -3,
                    ShowDataTypes = ExcelExportShowDataTypeEnum.FecalColiform.ToString() + ",",
                    SubsectorTVItemID = TVItemID,
                    StartDate = new DateTime(1900, 1, 1),
                    WetLimit24h = 12,
                    WetLimit48h = 25,
                    WetLimit72h = 37,
                    WetLimit96h = 50,
                };

                mwqmAnalysisReportParameterModelList.Add(mwqmAnalysisReportParameterModel);
            }
            if (!mwqmAnalysisReportParameterModelList.Where(c => c.AnalysisReportYear == Year && c.AnalysisCalculationType == AnalysisCalculationTypeEnum.WetAllAll).Any())
            {
                MWQMAnalysisReportParameterModel mwqmAnalysisReportParameterModel = new MWQMAnalysisReportParameterModel()
                {
                    AnalysisCalculationType = AnalysisCalculationTypeEnum.WetAllAll,
                    AnalysisName = "",
                    AnalysisReportYear = Year,
                    Command = AnalysisReportExportCommandEnum.Report,
                    DryLimit24h = 4,
                    DryLimit48h = 8,
                    DryLimit72h = 12,
                    DryLimit96h = 16,
                    EndDate = new DateTime(1900, 1, 1),
                    ExcelTVFileTVItemID = null,
                    FullYear = false,
                    LastUpdateContactTVItemID = 2, // Charles LeBlanc
                    LastUpdateDate_UTC = DateTime.Now,
                    MidRangeNumberOfDays = -6,
                    MWQMAnalysisReportParameterID = 0,
                    NumberOfRuns = 30,
                    RunsToOmit = ",",
                    SalinityHighlightDeviationFromAverage = 8,
                    ShortRangeNumberOfDays = -3,
                    ShowDataTypes = ExcelExportShowDataTypeEnum.FecalColiform.ToString() + ",",
                    SubsectorTVItemID = TVItemID,
                    StartDate = new DateTime(1900, 1, 1),
                    WetLimit24h = 12,
                    WetLimit48h = 25,
                    WetLimit72h = 37,
                    WetLimit96h = 50,
                };

                mwqmAnalysisReportParameterModelList.Add(mwqmAnalysisReportParameterModel);
            }
            if (!mwqmAnalysisReportParameterModelList.Where(c => c.AnalysisReportYear == Year && c.AnalysisCalculationType == AnalysisCalculationTypeEnum.DryAllAll).Any())
            {
                MWQMAnalysisReportParameterModel mwqmAnalysisReportParameterModel = new MWQMAnalysisReportParameterModel()
                {
                    AnalysisCalculationType = AnalysisCalculationTypeEnum.DryAllAll,
                    AnalysisName = "",
                    AnalysisReportYear = Year,
                    Command = AnalysisReportExportCommandEnum.Report,
                    DryLimit24h = 4,
                    DryLimit48h = 8,
                    DryLimit72h = 12,
                    DryLimit96h = 16,
                    EndDate = new DateTime(1900, 1, 1),
                    ExcelTVFileTVItemID = null,
                    FullYear = false,
                    LastUpdateContactTVItemID = 2, // Charles LeBlanc
                    LastUpdateDate_UTC = DateTime.Now,
                    MidRangeNumberOfDays = -6,
                    MWQMAnalysisReportParameterID = 0,
                    NumberOfRuns = 30,
                    RunsToOmit = ",",
                    SalinityHighlightDeviationFromAverage = 8,
                    ShortRangeNumberOfDays = -3,
                    ShowDataTypes = ExcelExportShowDataTypeEnum.FecalColiform.ToString() + ",",
                    SubsectorTVItemID = TVItemID,
                    StartDate = new DateTime(1900, 1, 1),
                    WetLimit24h = 12,
                    WetLimit48h = 25,
                    WetLimit72h = 37,
                    WetLimit96h = 50,
                };

                mwqmAnalysisReportParameterModelList.Add(mwqmAnalysisReportParameterModel);
            }

            string SubsectorTVText = _MWQMSubsectorService.GetMWQMSubsectorModelWithMWQMSubsectorTVItemIDDB(TVItemID).MWQMSubsectorTVText;

            sbHTML.AppendLine($@"<h2>{ SubsectorTVText }</h2>");
            sbHTML.AppendLine($@"<br />");

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 5);

            MWQMSubsectorAnalysisModel mwqmSubsectorAnalysisModelFirst = _MWQMSubsectorService.GetMWQMSubsectorAnalysisModel(TVItemID);
            List<MWQMSiteAnalysisModel> mwqmSiteAnalysisModelListFirst = mwqmSubsectorAnalysisModelFirst.MWQMSiteAnalysisModelList.Where(c => c.IsActive == true).OrderBy(c => c.MWQMSiteTVText).ToList();

            int CountAllSiteTotal = mwqmSiteAnalysisModelListFirst.Count() * mwqmAnalysisReportParameterModelList.Count();
            int CountAllSite = 0;
            int CountAnalysisReportParameterModel = 0;
            foreach (MWQMAnalysisReportParameterModel mwqmAnalysisReportParameterModel in mwqmAnalysisReportParameterModelList)
            {
                List<DateTime> DateTimeList = new List<DateTime>();

                MWQMSubsectorAnalysisModel mwqmSubsectorAnalysisModel = _MWQMSubsectorService.GetMWQMSubsectorAnalysisModel(TVItemID);

                if (mwqmAnalysisReportParameterModel.StartDate.Year == 1900)
                {
                    mwqmAnalysisReportParameterModel.StartDate = mwqmSubsectorAnalysisModel.MWQMSampleAnalysisModelList.OrderByDescending(c => c.SampleDateTime_Local).FirstOrDefault().SampleDateTime_Local;
                }

                if (mwqmAnalysisReportParameterModel.EndDate.Year == 1900)
                {
                    mwqmAnalysisReportParameterModel.EndDate = mwqmSubsectorAnalysisModel.MWQMSampleAnalysisModelList.OrderBy(c => c.SampleDateTime_Local).FirstOrDefault().SampleDateTime_Local;
                }

                foreach (MWQMSampleAnalysisModel mwqmSampleAnalysisModel in mwqmSubsectorAnalysisModel.MWQMSampleAnalysisModelList)
                {
                    mwqmSampleAnalysisModel.SampleDateTime_Local = new DateTime(mwqmSampleAnalysisModel.SampleDateTime_Local.Year, mwqmSampleAnalysisModel.SampleDateTime_Local.Month, mwqmSampleAnalysisModel.SampleDateTime_Local.Day);
                }

                foreach (MWQMRunAnalysisModel mwqmRunAnalysisModel in mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList)
                {
                    mwqmRunAnalysisModel.DateTime_Local = new DateTime(mwqmRunAnalysisModel.DateTime_Local.Year, mwqmRunAnalysisModel.DateTime_Local.Month, mwqmRunAnalysisModel.DateTime_Local.Day);
                }

                _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 3);

                List<MWQMSiteAnalysisModel> mwqmSiteAnalysisModelList = mwqmSubsectorAnalysisModel.MWQMSiteAnalysisModelList.Where(c => c.IsActive == true).OrderBy(c => c.MWQMSiteTVText).ToList();

                if (mwqmAnalysisReportParameterModel.AnalysisCalculationType == AnalysisCalculationTypeEnum.AllAllAll && !string.IsNullOrWhiteSpace(HideAllAllAll))
                {
                    sbHTML.AppendLine($@"<h4>{ TaskRunnerServiceRes.All } { TaskRunnerServiceRes.NotShown }</h4>");
                    continue;
                }
                if (mwqmAnalysisReportParameterModel.AnalysisCalculationType == AnalysisCalculationTypeEnum.WetAllAll && !string.IsNullOrWhiteSpace(HideWetAllAll))
                {
                    sbHTML.AppendLine($@"<h4>{ TaskRunnerServiceRes.Wet } { TaskRunnerServiceRes.NotShown }</h4>");
                    continue;
                }
                if (mwqmAnalysisReportParameterModel.AnalysisCalculationType == AnalysisCalculationTypeEnum.DryAllAll && !string.IsNullOrWhiteSpace(HideDryAllAll))
                {
                    sbHTML.AppendLine($@"<h4>{ TaskRunnerServiceRes.Dry } { TaskRunnerServiceRes.NotShown }</h4>");
                    continue;
                }
                if (mwqmAnalysisReportParameterModel.AnalysisCalculationType == AnalysisCalculationTypeEnum.WetAllAll && MissingRainData)
                {
                    sbHTML.AppendLine($@"<h4>{ TaskRunnerServiceRes.Wet } { TaskRunnerServiceRes.NotShown } { TaskRunnerServiceRes.BecauseOfMissingRainData }</h4>");
                    continue;
                }
                if (mwqmAnalysisReportParameterModel.AnalysisCalculationType == AnalysisCalculationTypeEnum.DryAllAll && MissingRainData)
                {
                    sbHTML.AppendLine($@"<h4>{ TaskRunnerServiceRes.Dry } { TaskRunnerServiceRes.NotShown } { TaskRunnerServiceRes.BecauseOfMissingRainData }</h4>");
                    continue;
                }

                string AllWetDry = TaskRunnerServiceRes.All;
                if (mwqmAnalysisReportParameterModel.AnalysisCalculationType == AnalysisCalculationTypeEnum.WetAllAll)
                {
                    AllWetDry = TaskRunnerServiceRes.Wet;
                }
                if (mwqmAnalysisReportParameterModel.AnalysisCalculationType == AnalysisCalculationTypeEnum.DryAllAll)
                {
                    AllWetDry = TaskRunnerServiceRes.Dry;
                }

                sbHTML.Append($@"<h4>{ TaskRunnerServiceRes.SummaryStatisticsOfFCDensities } ({ TaskRunnerServiceRes.MPN }/100 mL) ");
                sbHTML.Append($@" --- { AllWetDry } ---");
                sbHTML.Append($@" ({ Year })");
                sbHTML.AppendLine(@"</h4>");
                sbHTML.AppendLine(@"<table>");
                sbHTML.AppendLine(@"    <thead>");
                sbHTML.AppendLine(@"        <tr>");
                sbHTML.AppendLine($@"            <th>{ TaskRunnerServiceRes.Site }</th>");
                sbHTML.AppendLine($@"            <th></th>");
                sbHTML.AppendLine($@"            <th>{ TaskRunnerServiceRes.Samples }</th>");
                sbHTML.AppendLine($@"            <th>{ TaskRunnerServiceRes.Period }</th>");
                sbHTML.AppendLine($@"            <th>{ TaskRunnerServiceRes.MinFC }</th>");
                if (string.IsNullOrWhiteSpace(HideMaxFCColumn))
                {
                    sbHTML.AppendLine($@"            <th>{ TaskRunnerServiceRes.MaxFC }</th>");
                }
                sbHTML.AppendLine($@"            <th>{ TaskRunnerServiceRes.GMean }</th>");
                sbHTML.AppendLine($@"            <th>{ TaskRunnerServiceRes.Median }</th> ");
                sbHTML.AppendLine($@"            <th>{ TaskRunnerServiceRes.P90 }</th>");
                sbHTML.AppendLine($@"            <th>% &gt; 43</th>");
                sbHTML.AppendLine($@"            <th>% &gt; 260</th>");
                sbHTML.AppendLine($@"            <th></th>");
                sbHTML.AppendLine(@"        </tr>");
                sbHTML.AppendLine(@"    </thead>");
                sbHTML.AppendLine(@"    <tbody>");

                int CountSite = 0;
                int CountSiteTotal = mwqmSiteAnalysisModelList.Count();
                foreach (MWQMSiteAnalysisModel mwqmSiteAnalysisModel in mwqmSiteAnalysisModelList)
                {
                    string classificationLetter = "";
                    string classificationColor = "";
                    List<int> MWQMRunTVItemIDToOmitList = new List<int>();

                    CountSite += 1;
                    CountAllSite += 1;
                    if (CountAllSite % 10 == 0)
                    {
                        int Percent = (int)(10.0D + (90.0D * ((double)CountAllSite / (double)CountAllSiteTotal)));
                        _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);
                    }

                    string[] runDateTextList = mwqmAnalysisReportParameterModel.RunsToOmit.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                    foreach (string s in runDateTextList)
                    {
                        MWQMRunTVItemIDToOmitList.Add(int.Parse(s));
                    }

                    double? P90 = null;
                    double? GeoMean = null;
                    double? Median = null;
                    double? PercOver43 = null;
                    double? PercOver260 = null;

                    classificationLetter = GetLastClassificationInitial(mwqmSiteAnalysisModel.MWQMSiteLatestClassification);
                    classificationColor = GetLastClassificationColor(mwqmSiteAnalysisModel.MWQMSiteLatestClassification);

                    // loading all site sample and doing the stats
                    List<MWQMSampleAnalysisModel> mwqmSampleAnalysisForSiteModelList = mwqmSubsectorAnalysisModel.MWQMSampleAnalysisModelList.Where(c => c.MWQMSiteTVItemID == mwqmSiteAnalysisModel.MWQMSiteTVItemID).OrderByDescending(c => c.SampleDateTime_Local).ToList();
                    List<MWQMSampleAnalysisModel> mwqmSampleAnalysisForSiteModelToUseList = new List<MWQMSampleAnalysisModel>();
                    foreach (MWQMSampleAnalysisModel mwqmSampleAnalysisModel in mwqmSampleAnalysisForSiteModelList)
                    {
                        if (mwqmSampleAnalysisModel.SampleDateTime_Local.Year <= Year)
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
                                                            CanUsed = false;
                                                        }
                                                    }
                                                    else if (i == 2)
                                                    {
                                                        if (mwqmAnalysisReportParameterModel.DryLimit48h < TotalRain)
                                                        {
                                                            CanUsed = false;
                                                        }
                                                    }
                                                    else if (i == 3)
                                                    {
                                                        if (mwqmAnalysisReportParameterModel.DryLimit72h < TotalRain)
                                                        {
                                                            CanUsed = false;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (mwqmAnalysisReportParameterModel.DryLimit96h < TotalRain)
                                                        {
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

                    string Coloring = "";
                    string Letter = "";
                    if (mwqmSampleAnalysisForSiteModelToUseList.Count < 10)
                    {
                        Letter = mwqmSampleAnalysisForSiteModelToUseList.Count.ToString();
                        sbHTML.AppendLine(@"        <tr>");
                        sbHTML.AppendLine($@"            <td>{ mwqmSiteAnalysisModel.MWQMSiteTVText }</td>");
                        sbHTML.AppendLine($@"            <td class=""{ classificationColor }"">{ classificationLetter }</td>");
                        sbHTML.AppendLine(@"            <td>--</td>");
                        sbHTML.AppendLine(@"            <td>--</td>");
                        sbHTML.AppendLine(@"            <td>--</td>");
                        if (string.IsNullOrWhiteSpace(HideMaxFCColumn))
                        {
                            sbHTML.AppendLine(@"            <td>--</td>");
                        }
                        sbHTML.AppendLine(@"            <td>--</td>");
                        sbHTML.AppendLine(@"            <td>--</td>");
                        sbHTML.AppendLine(@"            <td>--</td>");
                        sbHTML.AppendLine(@"            <td>--</td>");
                        sbHTML.AppendLine(@"            <td>--</td>");
                        sbHTML.AppendLine($@"            <td class=""bglightblue"">{ Letter }</td>");
                        sbHTML.AppendLine(@"        </tr>");
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
                                Coloring = "bgbluef";
                                Letter = "F";
                            }
                            else if ((GeoMean > 162.67) || (Median > 162.67) || (P90 > 420.0) || (PercOver260 > 16.67))
                            {
                                Coloring = "bgbluee";
                                Letter = "E";
                            }
                            else if ((GeoMean > 144.0) || (Median > 144.0) || (P90 > 380.0) || (PercOver260 > 15.0))
                            {
                                Coloring = "bgblued";
                                Letter = "D";
                            }
                            else if ((GeoMean > 125.33) || (Median > 125.33) || (P90 > 340.0) || (PercOver260 > 13.33))
                            {
                                Coloring = "bgbluec";
                                Letter = "C";
                            }
                            else if ((GeoMean > 106.67) || (Median > 106.67) || (P90 > 300.0) || (PercOver260 > 11.67))
                            {
                                Coloring = "bgblueb";
                                Letter = "B";
                            }
                            else
                            {
                                Coloring = "bgbluea";
                                Letter = "A";
                            }
                        }
                        else if ((GeoMean > 14) || (Median > 14) || (P90 > 43) || (PercOver43 > 10))
                        {
                            if ((GeoMean > 75.67) || (Median > 75.67) || (P90 > 223.83) || (PercOver43 > 26.67))
                            {
                                Coloring = "bgredf";
                                Letter = "F";
                            }
                            else if ((GeoMean > 63.33) || (Median > 63.33) || (P90 > 187.67) || (PercOver43 > 23.33))
                            {
                                Coloring = "bgrede";
                                Letter = "E";
                            }
                            else if ((GeoMean > 51.0) || (Median > 51.0) || (P90 > 151.5) || (PercOver43 > 20.0))
                            {
                                Coloring = "bgredd";
                                Letter = "D";
                            }
                            else if ((GeoMean > 38.67) || (Median > 38.67) || (P90 > 115.33) || (PercOver43 > 16.67))
                            {
                                Coloring = "bgredc";
                                Letter = "C";
                            }
                            else if ((GeoMean > 26.33) || (Median > 26.33) || (P90 > 79.17) || (PercOver43 > 13.33))
                            {
                                Coloring = "bgredb";
                                Letter = "B";
                            }
                            else
                            {
                                Coloring = "bgreda";
                                Letter = "A";
                            }
                        }
                        else
                        {
                            if ((GeoMean > 11.67) || (Median > 11.67) || (P90 > 35.83) || (PercOver43 > 8.33))
                            {
                                Coloring = "bggreenf";
                                Letter = "F";
                            }
                            else if ((GeoMean > 9.33) || (Median > 9.33) || (P90 > 28.67) || (PercOver43 > 6.67))
                            {
                                Coloring = "bggreene";
                                Letter = "E";
                            }
                            else if ((GeoMean > 7.0) || (Median > 7.0) || (P90 > 21.5) || (PercOver43 > 5.0))
                            {
                                Coloring = "bggreend";
                                Letter = "D";
                            }
                            else if ((GeoMean > 4.67) || (Median > 4.67) || (P90 > 14.33) || (PercOver43 > 3.33))
                            {
                                Coloring = "bggreenc";
                                Letter = "C";
                            }
                            else if ((GeoMean > 2.33) || (Median > 2.33) || (P90 > 7.17) || (PercOver43 > 1.67))
                            {
                                Coloring = "bggreenb";
                                Letter = "B";
                            }
                            else
                            {
                                Coloring = "bggreena";
                                Letter = "A";
                            }
                        }

                        sbHTML.AppendLine(@"        <tr>");
                        sbHTML.AppendLine($@"            <td>{ mwqmSiteAnalysisModel.MWQMSiteTVText }</td>");
                        sbHTML.AppendLine($@"            <td class=""{ classificationColor }"">{ classificationLetter }</td>");
                        sbHTML.AppendLine($@"            <td>{ MWQMSampleCount.ToString() }</td>");
                        sbHTML.AppendLine($@"            <td>{ (MaxYear != null ? (MaxYear.ToString() + "-" + MinYear.ToString()) : "--") }</td>");
                        sbHTML.AppendLine($@"            <td>{ (MinFC != null ? (MinFC < 2 ? "< 2" : (MinFC.ToString())) : "--") }</td>");
                        if (string.IsNullOrWhiteSpace(HideMaxFCColumn))
                        {
                            sbHTML.AppendLine($@"            <td>{ (MaxFC != null ? (MaxFC < 2 ? "< 2" : (MaxFC.ToString())) : "--") }</td>");
                        }
                        string bgClass = GeoMean != null && GeoMean > 14 ? "bgyellow" : "";
                        sbHTML.AppendLine($@"            <td class=""{ bgClass }"">{ (GeoMean != null ? ((double)GeoMean < 2.0D ? "< 2" : ((double)GeoMean).ToString("F0")) : "--") }</td>");
                        bgClass = Median != null && Median > 14 ? "bgyellow" : "";
                        sbHTML.AppendLine($@"            <td class=""{ bgClass }"">{ (Median != null ? ((double)Median < 2.0D ? "< 2" : ((double)Median).ToString("F0")) : "--") }</td>");
                        bgClass = P90 != null && P90 > 43 ? "bgyellow" : "";
                        sbHTML.AppendLine($@"            <td class=""{ bgClass }"">{ (P90 != null ? ((double)P90 < 2.0D ? "< 2" : ((double)P90).ToString("F0")) : "--") }</td>");
                        bgClass = PercOver43 != null && PercOver43 > 10 ? "bgyellow" : "";
                        sbHTML.AppendLine($@"            <td class=""{ bgClass }"">{ (PercOver43 != null ? ((double)PercOver43).ToString("F0") : "--") }</td>");
                        bgClass = PercOver260 != null && PercOver260 > 10 ? "bgyellow" : "";
                        sbHTML.AppendLine($@"            <td class=""{ bgClass }"">{ (PercOver260 != null ? ((double)PercOver260).ToString("F0") : "--") }</td>");
                        sbHTML.AppendLine($@"            <td class=""{ Coloring }"">{ Letter }</td>");
                        sbHTML.AppendLine(@"        </tr>");

                    }

                    DateTimeList = DateTimeList.Concat(mwqmSampleAnalysisForSiteModelToUseList.Select(c => c.SampleDateTime_Local)).Distinct().ToList();
                }
                sbHTML.AppendLine(@"    </tbody>");
                sbHTML.AppendLine(@"    <tfoot>");
                sbHTML.AppendLine(@"        <tr>");

                List<int> YearList = new List<int>();
                for (int year = 1980, maxYear = Math.Min(Year, DateTime.Now.Year) + 1; year < maxYear; year++)
                {
                    YearList.Add(year);
                }

                List<int> CountPerYear = new List<int>();
                foreach (int year in YearList)
                {
                    CountPerYear.Add(DateTimeList.Where(c => c.Year == year).Count());
                }

                Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook workbook = xlApp.Workbooks.Add();
                Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.Worksheets.get_Item(1);

                Microsoft.Office.Interop.Excel.ChartObjects xlCharts = (Microsoft.Office.Interop.Excel.ChartObjects)worksheet.ChartObjects();
                Microsoft.Office.Interop.Excel.ChartObject chart = xlCharts.Add(100, 100, 1000, 100);
                Microsoft.Office.Interop.Excel.Chart chartPage = chart.Chart;

                chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;

                Microsoft.Office.Interop.Excel.SeriesCollection seriesCollection = chartPage.SeriesCollection();
                Microsoft.Office.Interop.Excel.Series series = seriesCollection.NewSeries();

                series.XValues = YearList.ToArray();
                series.Values = CountPerYear.ToArray();

                chartPage.ApplyLayout(9, Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered);
                chartPage.ChartTitle.Select();
                xlApp.Selection.Delete();
                chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue).AxisTitle.Select();
                xlApp.Selection.Delete();
                chartPage.Legend.Select();
                xlApp.Selection.Delete();              

                chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = TaskRunnerServiceRes.YearsWithSamplesUsed;

                CountAnalysisReportParameterModel += 1;
                // need to save the file with a unique name under the TVItemID
                FileInfo fiImage = new FileInfo(fi.DirectoryName + @"\" + FileNameExtra + CountAnalysisReportParameterModel.ToString() + ".png");
                chartPage.Export(fiImage.FullName, "PNG", false);

                if (workbook != null)
                {
                    workbook.Close(false);
                }
                if (xlApp != null)
                {
                    xlApp.Quit();
                }

                if (string.IsNullOrWhiteSpace(HideMaxFCColumn))
                {
                    sbHTML.AppendLine($@"            <td colspan=""12"">");
                }
                else
                {
                    sbHTML.AppendLine($@"            <td colspan=""11"">");
                }
                sbHTML.AppendLine($@"|||Image|FileName,{ fiImage.FullName }|width,1000|height,100|||");
                sbHTML.AppendLine($@"|||FileNameExtra|Random,{ FileNameExtra }|||");
                sbHTML.AppendLine(@"            </td>");
                sbHTML.AppendLine(@"        </tr>");
                sbHTML.AppendLine(@"        <tr>");
                if (string.IsNullOrWhiteSpace(HideMaxFCColumn))
                {
                    sbHTML.AppendLine($@"            <td colspan=""12"">");
                }
                else
                {
                    sbHTML.AppendLine($@"            <td colspan=""11"">");
                }
                sbHTML.Append($@"<b>{ TaskRunnerServiceRes.AnalysisName }</b> : { (string.IsNullOrWhiteSpace(mwqmAnalysisReportParameterModel.AnalysisName) ? "---" : mwqmAnalysisReportParameterModel.AnalysisName) }&nbsp;&nbsp;&nbsp;");
                sbHTML.Append($@"<b>{ TaskRunnerServiceRes.CalculationType }</b> : { AllWetDry }&nbsp;&nbsp;&nbsp;");
                sbHTML.Append($@"<b>{ TaskRunnerServiceRes.ReportYear }</b> : { mwqmAnalysisReportParameterModel.AnalysisReportYear }&nbsp;&nbsp;&nbsp;");
                sbHTML.Append($@"<b>{ TaskRunnerServiceRes.StartDate }</b> : { mwqmAnalysisReportParameterModel.StartDate.ToString("yyyy MMM dd") }&nbsp;&nbsp;&nbsp;");
                sbHTML.Append($@"<b>{ TaskRunnerServiceRes.EndDate }</b> : { mwqmAnalysisReportParameterModel.EndDate.ToString("yyyy MMM dd") }&nbsp;&nbsp;&nbsp;");
                sbHTML.Append($@"<b>{ TaskRunnerServiceRes.NumberOfRuns }</b> : { mwqmAnalysisReportParameterModel.NumberOfRuns }&nbsp;&nbsp;&nbsp;");
                sbHTML.Append($@"<b>{ TaskRunnerServiceRes.FullYear }</b> : { (mwqmAnalysisReportParameterModel.FullYear ? TaskRunnerServiceRes.Yes : TaskRunnerServiceRes.No) }&nbsp;&nbsp;&nbsp;");

                if (mwqmAnalysisReportParameterModel.AnalysisCalculationType == AnalysisCalculationTypeEnum.DryAllAll)
                {
                    sbHTML.Append($@" <b>{ TaskRunnerServiceRes.ShortRangeNumberOfDays }</b> : { Math.Abs(mwqmAnalysisReportParameterModel.ShortRangeNumberOfDays) }&nbsp;&nbsp;&nbsp;");
                    sbHTML.Append($@"<b>{ TaskRunnerServiceRes.Dry24h }</b> : { mwqmAnalysisReportParameterModel.DryLimit24h } mm &nbsp;&nbsp;&nbsp;");
                    sbHTML.Append($@"<b>{ TaskRunnerServiceRes.Dry48h }</b> : { mwqmAnalysisReportParameterModel.DryLimit48h } mm &nbsp;&nbsp;&nbsp;");
                    sbHTML.Append($@"<b>{ TaskRunnerServiceRes.Dry72h }</b> : { mwqmAnalysisReportParameterModel.DryLimit72h } mm &nbsp;&nbsp;&nbsp;");
                    sbHTML.Append($@"<b>{ TaskRunnerServiceRes.Dry96h }</b> : { mwqmAnalysisReportParameterModel.DryLimit96h } mm &nbsp;&nbsp;&nbsp;");
                }
                if (mwqmAnalysisReportParameterModel.AnalysisCalculationType == AnalysisCalculationTypeEnum.WetAllAll)
                {
                    sbHTML.Append($@"<b>{ TaskRunnerServiceRes.ShortRangeNumberOfDays }</b> : { Math.Abs(mwqmAnalysisReportParameterModel.ShortRangeNumberOfDays) }&nbsp;&nbsp;&nbsp;");
                    sbHTML.Append($@"<b>{ TaskRunnerServiceRes.Wet24h }</b> : { mwqmAnalysisReportParameterModel.WetLimit24h } mm &nbsp;&nbsp;&nbsp;");
                    sbHTML.Append($@"<b>{ TaskRunnerServiceRes.Wet48h }</b> : { mwqmAnalysisReportParameterModel.WetLimit48h } mm &nbsp;&nbsp;&nbsp;");
                    sbHTML.Append($@"<b>{ TaskRunnerServiceRes.Wet72h }</b> : { mwqmAnalysisReportParameterModel.WetLimit72h } mm &nbsp;&nbsp;&nbsp;");
                    sbHTML.Append($@"<b>{ TaskRunnerServiceRes.Wet96h }</b> : { mwqmAnalysisReportParameterModel.WetLimit96h } mm &nbsp;&nbsp;&nbsp;");
                }
                List<int> MWQMRunTVItemIDList = mwqmAnalysisReportParameterModel.RunsToOmit.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Select(c => int.Parse(c)).ToList();

                if (MWQMRunTVItemIDList.Count > 0)
                {
                    sbHTML.Append($@"<b>{ TaskRunnerServiceRes.RunsOmitted }</b> : ");

                    foreach (int mwqmRunTVItemID in MWQMRunTVItemIDList)
                    {
                        MWQMRunModel mwqmRunModel = _MWQMRunService.GetMWQMRunModelWithMWQMRunTVItemIDDB(mwqmRunTVItemID);
                        if (!string.IsNullOrEmpty(mwqmRunModel.Error))
                        {
                            sbHTML.Append($@" [err - { mwqmRunTVItemID }]");
                        }
                        else
                        {
                            sbHTML.Append($@" [{ mwqmRunModel.DateTime_Local.ToString("yyyy MMM dd") }]");
                        }
                    }
                }
                sbHTML.AppendLine($@"");
                sbHTML.AppendLine(@"            </td>");
                sbHTML.AppendLine(@"        </tr>");
                sbHTML.AppendLine(@"    </tfoot>");
                sbHTML.AppendLine(@"</table>");

                sbHTML.AppendLine(@"<span>|||PageBreak|||</span>");
            }

            if (!GetBottomHTML(sbHTML, fi, parameters))
            {
                return false;
            }

            return true;
        }

        // for testing only can comment out when test is completed
        public bool PublicGenerateHTMLSubsectorFCSummaryStatDocx(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            bool retBool = GenerateHTMLSubsectorFCSummaryStatDocx(fi, sbHTML, parameters, reportTypeModel);

            StreamWriter sw = fi.CreateText();
            sw.Write(sbHTML.ToString());
            sw.Flush();
            sw.Close();

            return retBool;
        }
    }
}
