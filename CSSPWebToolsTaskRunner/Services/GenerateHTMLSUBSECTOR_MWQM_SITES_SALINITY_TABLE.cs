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
using CSSPEnumsDLL.Services;
using CSSPModelsDLL.Models;
using System.Windows.Forms;
using System.Web.Helpers;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        private bool GenerateHTMLSUBSECTOR_MWQM_SITES_SALINITY_TABLE(StringBuilder sbTemp)
        {
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

                List<MWQMRunModel> mwqmRunModelList2 = mwqmRunModelList.Where(c => c.RunSampleType == SampleTypeEnum.Routine && c.DateTime_Local.Year <= Year).OrderByDescending(c => c.DateTime_Local).Skip(skip).Take(take).ToList();
                if (mwqmRunModelList2.Count > 0)
                {
                    sbTemp.AppendLine($@"|||TableCaption|: { TaskRunnerServiceRes.ActiveMWQMSites }&nbsp;&nbsp;{ TaskRunnerServiceRes.Salinity }&nbsp;&nbsp;&nbsp;({ TaskRunnerServiceRes.Routine })&nbsp;&nbsp;&nbsp;{ mwqmRunModelList2[0].DateTime_Local.ToString("yyyy MMMM dd") } { TaskRunnerServiceRes.To } { mwqmRunModelList2[mwqmRunModelList2.Count - 1].DateTime_Local.ToString("yyyy MMMM dd") }|||");

                    sbTemp.AppendLine($@" <table class=""FCSalTempDataTableClass"">");
                    sbTemp.AppendLine($@" <tr>");
                    sbTemp.AppendLine($@" <th class=""rightBottomBorder"">{ TaskRunnerServiceRes.Site }</th>");
                    foreach (MWQMRunModel mwqmRunModel in mwqmRunModelList2)
                    {
                        sbTemp.AppendLine($@" <th class=""bottomBorder"">{ mwqmRunModel.DateTime_Local.ToString("yyyy") }<br />{ mwqmRunModel.DateTime_Local.ToString("MMM") }<br />{ mwqmRunModel.DateTime_Local.ToString("dd") }</th>");
                    }
                    sbTemp.AppendLine($@" </tr>");

                    foreach (MWQMSiteModel mwqmSiteModel in mwqmSiteModelList2)
                    {
                        sbTemp.AppendLine($@" <tr>");
                        sbTemp.AppendLine($@" <td class=""rightBorder"">{ mwqmSiteModel.MWQMSiteTVText }</td>");
                        foreach (MWQMRunModel mwqmRunModel in mwqmRunModelList2)
                        {
                            float? value = (float?)(from s in mwqmSampleModelList
                                            where s.MWQMRunTVItemID == mwqmRunModel.MWQMRunTVItemID
                                            && s.MWQMSiteTVItemID == mwqmSiteModel.MWQMSiteTVItemID
                                            select s.Salinity_PPT).FirstOrDefault();

                            string valueStr = value != null ? (value == 1 ? "< 2" : ((float)value).ToString("F0")) : "--";
                            sbTemp.AppendLine($@" <td>{ valueStr }</td>");
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
                        string RainDay1 = mwqmRunModel.RainDay1_mm != null ? ((double)mwqmRunModel.RainDay1_mm).ToString("F0") : "--";
                        string RainDay2 = mwqmRunModel.RainDay2_mm != null ? ((double)mwqmRunModel.RainDay2_mm).ToString("F0") : "--";
                        string RainDay3 = mwqmRunModel.RainDay3_mm != null ? ((double)mwqmRunModel.RainDay3_mm).ToString("F0") : "--";
                        string RainDay4 = mwqmRunModel.RainDay4_mm != null ? ((double)mwqmRunModel.RainDay4_mm).ToString("F0") : "--";
                        string RainDay5 = mwqmRunModel.RainDay5_mm != null ? ((double)mwqmRunModel.RainDay5_mm).ToString("F0") : "--";
                        sbTemp.AppendLine($@" <td class=""topRightBorder"">&nbsp;<br />{ RainDay1 }<br />{ RainDay2 }<br />{ RainDay3 }<br />{ RainDay4 }<br />{ RainDay5 }</td>");
                    }
                    sbTemp.AppendLine($@" </tr>");

                    sbTemp.AppendLine($@" </table>");

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
