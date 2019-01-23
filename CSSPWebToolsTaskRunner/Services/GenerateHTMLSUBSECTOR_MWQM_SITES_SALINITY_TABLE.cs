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
using CSSPEnumsDLL.Services;
using CSSPModelsDLL.Models;
using System.Windows.Forms;
//using System.Web.Helpers;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        private bool GenerateHTMLSUBSECTOR_MWQM_SITES_SALINITY_TABLE(StringBuilder sbTemp)
        {
            List<string> Letters = new List<string>() { "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t" };
            List<int> ClimateSiteUsedList = new List<int>();
            int Percent = 10;
            string NotUsed = "";

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", ReportGenerateObjectsKeywordEnum.SUBSECTOR_MWQM_SITES_SALINITY_TABLE.ToString()));

            List<string> ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            // TVItemID and Year already loaded

            TVItemModel tvItemModelSubsector = _TVItemService.GetTVItemModelWithTVItemIDDB(TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelSubsector.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                return false;
            }

            string ServerPath = _TVFileService.GetServerFilePath(tvItemModelSubsector.TVItemID);

            List<TVItemModel> tvItemModelListMWQMSites = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.MWQMSite).Where(c => c.IsActive == true).ToList();
            List<MWQMSiteModel> mwqmSiteModelList = _MWQMSiteService.GetMWQMSiteModelListWithSubsectorTVItemIDDB(TVItemID);
            List<MWQMRunModel> mwqmRunModelList = _MWQMRunService.GetMWQMRunModelListWithSubsectorTVItemIDDB(TVItemID);
            List<MWQMSampleModel> mwqmSampleModelList = _MWQMSampleService.GetMWQMSampleModelListWithSubsectorTVItemIDDB(TVItemID);
            List<UseOfSiteModel> useOfSiteModelList = _UseOfSiteService.GetUseOfSiteModelListWithSubsectorTVItemIDDB(TVItemID).Where(c => c.TVType == TVTypeEnum.ClimateSite).ToList();

            List<ClimateSiteModel> climateSiteModelList = new List<ClimateSiteModel>();
            List<ClimateDataValueModel> climateDataValueModelList = new List<ClimateDataValueModel>();

            foreach (UseOfSiteModel useOfSiteModel in useOfSiteModelList)
            {
                ClimateSiteModel climateSiteModel = _ClimateSiteService.GetClimateSiteModelWithClimateSiteTVItemIDDB(useOfSiteModel.SiteTVItemID);
                if (climateSiteModel != null)
                {
                    climateSiteModelList.Add(climateSiteModel);

                    List<ClimateDataValueModel> climateDataValueModelOneSiteList = _ClimateDataValueService.GetClimateDataValueModelWithClimateSiteIDDB(climateSiteModel.ClimateSiteID);

                    climateDataValueModelList = climateDataValueModelList.Concat(climateDataValueModelOneSiteList).ToList();
                }
            }

            climateSiteModelList = climateSiteModelList.OrderBy(c => c.ClimateSiteName).ToList();
            List<MWQMSiteModel> mwqmSiteModelList2 = (from s in mwqmSiteModelList
                                                      from t in tvItemModelListMWQMSites
                                                      where s.MWQMSiteTVItemID == t.TVItemID
                                                      && t.IsActive == true
                                                      orderby s.MWQMSiteTVText
                                                      select s).ToList();

            int skip = 0;
            int take = 15;
            bool HasData = true;
            //int countRun = 0;
            int TableCount = 0;
            while (HasData)
            {
                Percent += 10;
                if (Percent > 100)
                {
                    Percent = 100;
                }
                _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);

                //countRun += 1;

                //if (countRun > 2)
                //{
                //    break;
                //}

                List<MWQMRunModel> mwqmRunModelList2 = (from r in mwqmRunModelList
                                                        from s in mwqmSiteModelList2
                                                        from sa in mwqmSampleModelList
                                                        where sa.MWQMRunTVItemID == r.MWQMRunTVItemID
                                                        && sa.MWQMSiteTVItemID == s.MWQMSiteTVItemID
                                                        && r.RunSampleType == SampleTypeEnum.Routine
                                                        && r.DateTime_Local.Year <= Year
                                                        orderby r.DateTime_Local descending
                                                        select r).Distinct().Skip(skip).Take(take).ToList();

                bool HasGreen = false;

                foreach (MWQMRunModel mwqmRunModel in mwqmRunModelList2)
                {
                    if (RunSiteInfoList.Where(c => c.RunTVItemID == mwqmRunModel.MWQMRunTVItemID).Any())
                    {
                        HasGreen = true;
                    }
                }

                if (mwqmRunModelList2.Count > 0 && HasGreen)
                {
                    sbTemp.AppendLine($@"|||TableCaption|Table { TaskRunnerServiceRes.Appendix } B 1.{ TableCount }: { TaskRunnerServiceRes.Salinity }|||");

                    TableCount += 1;

                    sbTemp.AppendLine($@" <table class=""FCSalTempDataTableClass"">");
                    sbTemp.AppendLine($@" <tr>");
                    sbTemp.AppendLine($@" <th class=""rightBottomBorder"">{ TaskRunnerServiceRes.Site }</th>");
                    foreach (MWQMRunModel mwqmRunModel in mwqmRunModelList2)
                    {
                        bool showGreenText = RunSiteInfoList.Where(c => c.RunTVItemID == mwqmRunModel.MWQMRunTVItemID).Any() ? true : false;

                        if (showGreenText)
                        {
                            sbTemp.AppendLine($@" <th class=""bottomBorderGreentext"">{ mwqmRunModel.DateTime_Local.ToString("yyyy") }<br />{ mwqmRunModel.DateTime_Local.ToString("MMM") }<br />{ mwqmRunModel.DateTime_Local.ToString("dd") }</th>");
                        }
                        else
                        {
                            sbTemp.AppendLine($@" <th class=""bottomBorder"">{ mwqmRunModel.DateTime_Local.ToString("yyyy") }<br />{ mwqmRunModel.DateTime_Local.ToString("MMM") }<br />{ mwqmRunModel.DateTime_Local.ToString("dd") }</th>");
                        }
                    }
                    sbTemp.AppendLine($@" </tr>");

                    foreach (MWQMSiteModel mwqmSiteModel in mwqmSiteModelList2)
                    {
                        string siteNameWithoutZeros = mwqmSiteModel.MWQMSiteTVText.Trim();
                        while (true)
                        {
                            if (siteNameWithoutZeros.StartsWith("0"))
                            {
                                siteNameWithoutZeros = siteNameWithoutZeros.Substring(1);
                            }
                            else
                            {
                                break;
                            }
                        }

                        sbTemp.AppendLine($@" <tr>");
                        sbTemp.AppendLine($@" <td class=""rightBorder"">{ siteNameWithoutZeros }</td>");
                        foreach (MWQMRunModel mwqmRunModel in mwqmRunModelList2)
                        {
                            bool showTextGreen = RunSiteInfoList.Where(c => c.RunTVItemID == mwqmRunModel.MWQMRunTVItemID && c.SiteTVItemID == mwqmSiteModel.MWQMSiteTVItemID).Any() ? true : false;

                            float? value = (float?)(from s in mwqmSampleModelList
                                            where s.MWQMRunTVItemID == mwqmRunModel.MWQMRunTVItemID
                                            && s.MWQMSiteTVItemID == mwqmSiteModel.MWQMSiteTVItemID
                                            select s.Salinity_PPT).FirstOrDefault();

                            string valueStr = value != null ? (((float)value).ToString("F1")) : "--";
                            if (showTextGreen)
                            {
                                sbTemp.AppendLine($@" <td class=""textGreen"">{ valueStr }</td>");
                            }
                            else
                            {
                                sbTemp.AppendLine($@" <td>{ valueStr }</td>");
                            }
                        }
                        sbTemp.AppendLine($@" </tr>");
                    }
                    sbTemp.AppendLine($@" <tr>");
                    sbTemp.AppendLine($@" <td class=""topRightBorder"">{ TaskRunnerServiceRes.StartTide }<br />{ TaskRunnerServiceRes.EndTide }</td>");
                    foreach (MWQMRunModel mwqmRunModel in mwqmRunModelList2)
                    {
                        string StartTide = GetTideInitial(mwqmRunModel.Tide_Start);
                        string EndTide = GetTideInitial(mwqmRunModel.Tide_End);
                        sbTemp.AppendLine($@" <td class=""topRightBorder"">{ StartTide }<br />{ EndTide }</td>");
                    }
                    sbTemp.AppendLine($@" </tr>");

                    sbTemp.AppendLine($@" <tr>");
                    sbTemp.AppendLine($@" <td class=""topRightBorder"">{ TaskRunnerServiceRes.Rain }(mm)<br />{ TaskRunnerServiceRes.Minus1Day }<br />{ TaskRunnerServiceRes.Minus2Day }<br />{ TaskRunnerServiceRes.Minus3Day }<br />{ TaskRunnerServiceRes.Minus4Day }<br />{ TaskRunnerServiceRes.Minus5Day }</td>");
                    foreach (MWQMRunModel mwqmRunModel in mwqmRunModelList2)
                    {
                        string sup1 = "";
                        string sup2 = "";
                        string sup3 = "";
                        string sup4 = "";
                        string sup5 = "";
                        DateTime Date1 = mwqmRunModel.DateTime_Local.AddDays(-1);
                        DateTime Date2 = mwqmRunModel.DateTime_Local.AddDays(-2);
                        DateTime Date3 = mwqmRunModel.DateTime_Local.AddDays(-3);
                        DateTime Date4 = mwqmRunModel.DateTime_Local.AddDays(-4);
                        DateTime Date5 = mwqmRunModel.DateTime_Local.AddDays(-5);

                        // RainDay1
                        for (int i = 0, count = climateSiteModelList.Count; i < count; i++)
                        {
                            if ((from c in climateDataValueModelList
                                 where c.DateTime_Local.Year == Date1.Year
                                 && c.DateTime_Local.Month == Date1.Month
                                 && c.DateTime_Local.Day == Date1.Day
                                 && c.TotalPrecip_mm_cm != null
                                 && mwqmRunModel.RainDay1_mm != null
                                 && (int)c.TotalPrecip_mm_cm == (int)mwqmRunModel.RainDay1_mm
                                 && c.ClimateSiteID == climateSiteModelList[i].ClimateSiteID
                                 select c).Any())
                            {
                                sup1 = Letters[i];
                                if (!ClimateSiteUsedList.Contains(i))
                                {
                                    ClimateSiteUsedList.Add(i);
                                }
                                break;
                            }
                        }

                        // RainDay2
                        for (int i = 0, count = climateSiteModelList.Count; i < count; i++)
                        {
                            if ((from c in climateDataValueModelList
                                 where c.DateTime_Local.Year == Date2.Year
                                 && c.DateTime_Local.Month == Date2.Month
                                 && c.DateTime_Local.Day == Date2.Day
                                 && c.TotalPrecip_mm_cm != null
                                 && mwqmRunModel.RainDay2_mm != null
                                 && (int)c.TotalPrecip_mm_cm == (int)mwqmRunModel.RainDay2_mm
                                 && c.ClimateSiteID == climateSiteModelList[i].ClimateSiteID
                                 select c).Any())
                            {
                                sup2 = Letters[i];
                                if (!ClimateSiteUsedList.Contains(i))
                                {
                                    ClimateSiteUsedList.Add(i);
                                }
                                break;
                            }
                        }

                        // RainDay3
                        for (int i = 0, count = climateSiteModelList.Count; i < count; i++)
                        {
                            if ((from c in climateDataValueModelList
                                 where c.DateTime_Local.Year == Date3.Year
                                 && c.DateTime_Local.Month == Date3.Month
                                 && c.DateTime_Local.Day == Date3.Day
                                 && c.TotalPrecip_mm_cm != null
                                 && mwqmRunModel.RainDay3_mm != null
                                 && (int)c.TotalPrecip_mm_cm == (int)mwqmRunModel.RainDay3_mm
                                 && c.ClimateSiteID == climateSiteModelList[i].ClimateSiteID
                                 select c).Any())
                            {
                                sup3 = Letters[i];
                                if (!ClimateSiteUsedList.Contains(i))
                                {
                                    ClimateSiteUsedList.Add(i);
                                }
                                break;
                            }
                        }

                        // RainDay4
                        for (int i = 0, count = climateSiteModelList.Count; i < count; i++)
                        {
                            if ((from c in climateDataValueModelList
                                 where c.DateTime_Local.Year == Date4.Year
                                 && c.DateTime_Local.Month == Date4.Month
                                 && c.DateTime_Local.Day == Date4.Day
                                 && c.TotalPrecip_mm_cm != null
                                 && mwqmRunModel.RainDay4_mm != null
                                 && (int)c.TotalPrecip_mm_cm == (int)mwqmRunModel.RainDay4_mm
                                 && c.ClimateSiteID == climateSiteModelList[i].ClimateSiteID
                                 select c).Any())
                            {
                                sup4 = Letters[i];
                                if (!ClimateSiteUsedList.Contains(i))
                                {
                                    ClimateSiteUsedList.Add(i);
                                }
                                break;
                            }
                        }

                        // RainDay5
                        for (int i = 0, count = climateSiteModelList.Count; i < count; i++)
                        {
                            if ((from c in climateDataValueModelList
                                 where c.DateTime_Local.Year == Date5.Year
                                 && c.DateTime_Local.Month == Date5.Month
                                 && c.DateTime_Local.Day == Date5.Day
                                 && c.TotalPrecip_mm_cm != null
                                 && mwqmRunModel.RainDay5_mm != null
                                 && (int)c.TotalPrecip_mm_cm == (int)mwqmRunModel.RainDay5_mm
                                 && c.ClimateSiteID == climateSiteModelList[i].ClimateSiteID
                                 select c).Any())
                            {
                                sup5 = Letters[i];
                                if (!ClimateSiteUsedList.Contains(i))
                                {
                                    ClimateSiteUsedList.Add(i);
                                }
                                break;
                            }
                        }

                        string RainDay1 = mwqmRunModel.RainDay1_mm != null ? ((double)mwqmRunModel.RainDay1_mm).ToString("F0") : "--";
                        string RainDay2 = mwqmRunModel.RainDay2_mm != null ? ((double)mwqmRunModel.RainDay2_mm).ToString("F0") : "--";
                        string RainDay3 = mwqmRunModel.RainDay3_mm != null ? ((double)mwqmRunModel.RainDay3_mm).ToString("F0") : "--";
                        string RainDay4 = mwqmRunModel.RainDay4_mm != null ? ((double)mwqmRunModel.RainDay4_mm).ToString("F0") : "--";
                        string RainDay5 = mwqmRunModel.RainDay5_mm != null ? ((double)mwqmRunModel.RainDay5_mm).ToString("F0") : "--";
                        sbTemp.AppendLine($@" <td class=""topRightBorder"">&nbsp;<br />{ RainDay1 }<sup>{ sup1 }</sup><br />{ RainDay2 }<sup>{ sup2 }</sup><br />{ RainDay3 }<sup>{ sup3 }</sup><br />{ RainDay4 }<sup>{ sup4 }</sup><br />{ RainDay5 }<sup>{ sup5 }</sup></td>");
                    }
                    sbTemp.AppendLine($@" </tr>");

                    sbTemp.AppendLine($@" </table>");
                    sbTemp.AppendLine($@" <p><span class=""textGreen"">({ TaskRunnerServiceRes.DataUsedForStatistics })</span>");
                    sbTemp.AppendLine($@" <span>&nbsp;&nbsp;&nbsp;{ TaskRunnerServiceRes.ClimateSite }</span>: ");
                    foreach (int i in ClimateSiteUsedList)
                    {
                        sbTemp.Append($@" <span>{ Letters[i] } - { climateSiteModelList[i].ClimateSiteName }<span>&nbsp;&nbsp;&nbsp;");
                    }
                    sbTemp.AppendLine($@" </p>");

                    sbTemp.AppendLine(@"<p>|||PAGE_BREAK|||</p>");

                    skip += take;
                }
                else
                {
                    HasData = false;
                }
            }


            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 80);

            return true;
        }

        // for testing only can comment out when test is completed
        public bool PublicGenerateHTMLSUBSECTOR_MWQM_SITES_SALINITY_TABLE(StringBuilder sbTemp)
        {
            bool retBool = GenerateHTMLSUBSECTOR_MWQM_SITES_SALINITY_TABLE(sbTemp);

            StreamWriter sw = fi.CreateText();
            sw.Write(sbTemp.ToString());
            sw.Flush();
            sw.Close();

            return retBool;
        }  
    }
}
