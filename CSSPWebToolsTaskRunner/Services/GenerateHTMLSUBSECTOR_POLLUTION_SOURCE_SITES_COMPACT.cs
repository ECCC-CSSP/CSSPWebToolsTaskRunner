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
using CSSPDBDLL;
using System.Threading;
//using System.Web.Helpers;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        private bool GenerateHTMLSUBSECTOR_POLLUTION_SOURCE_SITES_COMPACT(StringBuilder sbTemp)
        {
            int Percent = 10;
            string NotUsed = "";
            LanguageEnum language = _TaskRunnerBaseService._BWObj.appTaskModel.Language;

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", ReportGenerateObjectsKeywordEnum.SUBSECTOR_POLLUTION_SOURCE_SITES_COMPACT.ToString()));

            TVItemModel tvItemModelSubsector = _TVItemService.GetTVItemModelWithTVItemIDDB(TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelSubsector.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                return false;
            }

            string ServerPath = _TVFileService.GetServerFilePath(tvItemModelSubsector.TVItemID);

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 5);

            List<TVItemModel> tvItemModelListPolSourceSite = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.PolSourceSite);
            List<PolSourceSiteModel> polSourceSiteModelList = _PolSourceSiteService.GetPolSourceSiteModelListWithSubsectorTVItemIDDB(TVItemID).OrderBy(c => c.Site).ToList();
            List<PolSourceObservationModel> polSourceObservationModelList = _PolSourceObservationService.GetPolSourceObservationModelListWithSubsectorTVItemIDDB(TVItemID);
            List<PolSourceObservationIssueModel> polSourceObservationIssueModelList = _PolSourceObservationIssueService.GetPolSourceObservationIssueModelListWithSubsectorTVItemIDDB(TVItemID);
            List<int> TVItemIDPolSourceSiteActiveList = tvItemModelListPolSourceSite.Where(c => c.IsActive == true).Select(c => c.TVItemID).ToList();

            List<MapInfo> mapInfoActiveList = new List<MapInfo>();
            List<MapInfoPoint> mapInfoPointActiveList = new List<MapInfoPoint>();

            using (CSSPDBEntities db2 = new CSSPDBEntities())
            {
                mapInfoActiveList = (from c in db2.MapInfos
                                     from a in TVItemIDPolSourceSiteActiveList
                                     where c.TVItemID == a
                                     select c).ToList();

                List<int> mapInfoIDActiveList = mapInfoActiveList.Select(c => c.MapInfoID).ToList();

                mapInfoPointActiveList = (from c in db2.MapInfoPoints
                                          from a in mapInfoIDActiveList
                                          where c.MapInfoID == a
                                          select c).ToList();

            }

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 10);

            sbTemp.AppendLine($@"<table class=""PolSourceSiteCompact"">");
            sbTemp.AppendLine($@"   <tr> ");
            sbTemp.AppendLine($@"       <th class=""allBordersNoWrap"">{ TaskRunnerServiceRes.Site }</th>");
            sbTemp.AppendLine($@"       <th class=""allBordersNoWrap"">{ TaskRunnerServiceRes.Lat }</th>");
            sbTemp.AppendLine($@"       <th class=""allBordersNoWrap"">{ TaskRunnerServiceRes.Long }</th>");
            sbTemp.AppendLine($@"       <th class=""allBorders"">{ TaskRunnerServiceRes.Description }</th>");
            sbTemp.AppendLine($@"   </tr>");

            foreach (TVItemModel tvItemModelPSSActive in tvItemModelListPolSourceSite.Where(c => c.IsActive == true))
            {
                string SiteName = "";
                string Lat = "";
                string Lng = "";
                string SourceType = "";
                string Structure = "";
                string Risk = "";
                string Desc = "";

                PolSourceSiteModel polSourceSiteModel = polSourceSiteModelList.Where(c => c.PolSourceSiteTVItemID == tvItemModelPSSActive.TVItemID).FirstOrDefault();

                SiteName = polSourceSiteModel.Site.ToString();

                if (polSourceSiteModel != null)
                {

                    MapInfo mapInfo = mapInfoActiveList.Where(c => c.TVItemID == tvItemModelPSSActive.TVItemID).FirstOrDefault();
                    if (mapInfo != null)
                    {
                        List<MapInfoPoint> mapInfoPointListCurrent = mapInfoPointActiveList.Where(c => c.MapInfoID == mapInfo.MapInfoID).ToList();
                        if (mapInfoPointListCurrent.Count > 0)
                        {
                            Lat = mapInfoPointListCurrent[0].Lat.ToString("F5");
                            Lng = mapInfoPointListCurrent[0].Lng.ToString("F5");
                        }
                    }

                    PolSourceObservationModel polSourceObservationModel = polSourceObservationModelList.Where(c => c.PolSourceSiteID == polSourceSiteModel.PolSourceSiteID).OrderByDescending(c => c.ObservationDate_Local).FirstOrDefault();

                    if (polSourceObservationModel != null)
                    {
                        foreach (PolSourceObservationIssueModel polSourceObservationIssueModel in polSourceObservationIssueModelList.Where(c => c.PolSourceObservationID == polSourceObservationModel.PolSourceObservationID).OrderBy(c => c.Ordinal).ToList())
                        {
                            List<string> polSourceObsInfoList = polSourceObservationIssueModel.ObservationInfo.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Where(c => c.Length == 5).ToList();
                            List<int> polSourceObsInfoIntList =  polSourceObsInfoList.Select(c => int.Parse(c)).ToList();

                            for (int i = 0, count = polSourceObsInfoList.Count; i < count; i++)
                            {
                                string StartTxt = polSourceObsInfoList[i].Substring(0, 3);

                                if (StartTxt.StartsWith("101"))
                                {
                                    SourceType = _BaseEnumService.GetEnumText_PolSourceObsInfoEnum((PolSourceObsInfoEnum)int.Parse(polSourceObsInfoList[i])).Trim();
                                    if (SourceType.Contains("|"))
                                    {
                                        SourceType = SourceType.Substring(0, SourceType.IndexOf("|"));
                                    }
                                    SourceType = SourceType.Trim();
                                }
                                if (StartTxt.StartsWith("143"))
                                {
                                    Structure = _BaseEnumService.GetEnumText_PolSourceObsInfoEnum((PolSourceObsInfoEnum)int.Parse(polSourceObsInfoList[i])).Trim();
                                    if (Structure.Contains("|"))
                                    {
                                        Structure = Structure.Substring(0, Structure.IndexOf("|"));
                                    }
                                    Structure = Structure.Trim();
                                }
                                if (StartTxt.StartsWith("910"))
                                {
                                    Risk = _BaseEnumService.GetEnumText_PolSourceObsInfoEnum((PolSourceObsInfoEnum)int.Parse(polSourceObsInfoList[i]));
                                    if (Risk.Contains("|"))
                                    {
                                        Risk = Risk.Substring(0, Risk.IndexOf("|"));
                                    }
                                    Risk = Risk.Trim();
                                }
                            }

                            if (!string.IsNullOrWhiteSpace(Risk))
                            {
                                Desc = $"{ SourceType } - { Structure } - { Risk }";
                            }
                            else
                            {
                                if (!string.IsNullOrWhiteSpace(polSourceObservationModel.Observation_ToBeDeleted))
                                {
                                    Desc = polSourceObservationModel.Observation_ToBeDeleted;
                                }
                                else
                                {
                                    PolSourceObservationModel polSourceObservationModel2 = polSourceObservationModelList.Where(c => c.PolSourceSiteID == polSourceSiteModel.PolSourceSiteID && c.Observation_ToBeDeleted.Length > 0).FirstOrDefault();
                                    if (polSourceObservationModel2 != null && !string.IsNullOrWhiteSpace(polSourceObservationModel2.Observation_ToBeDeleted))
                                    {
                                        Desc = polSourceObservationModel2.Observation_ToBeDeleted;
                                    }
                                }
                            }

                        }

                        sbTemp.AppendLine($@"   <tr>");
                        sbTemp.AppendLine($@"       <td class=""allBordersNoWrap"">{ SiteName }</td>");
                        sbTemp.AppendLine($@"       <td class=""allBordersNoWrap"">{ Lat }</td>");
                        sbTemp.AppendLine($@"       <td class=""allBordersNoWrap"">{ Lng }</td>");
                        sbTemp.AppendLine($@"       <td class=""allBorders"">{ Desc }</td>");
                        sbTemp.AppendLine($@"   </tr>");


                    }
                }
            }

            sbTemp.AppendLine($@"</table>");

            Percent = 98;
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);

            return true;
        }

        // for testing only can comment out when test is completed
        public bool PublicGenerateHTMLSUBSECTOR_POLLUTION_SOURCE_SITES_COMPACT(StringBuilder sbTemp)
        {
            bool retBool = GenerateHTMLSUBSECTOR_POLLUTION_SOURCE_SITES_COMPACT(sbTemp);

            StreamWriter sw = fi.CreateText();
            sw.Write(sbTemp.ToString());
            sw.Flush();
            sw.Close();

            return retBool;
        }
    }
}
