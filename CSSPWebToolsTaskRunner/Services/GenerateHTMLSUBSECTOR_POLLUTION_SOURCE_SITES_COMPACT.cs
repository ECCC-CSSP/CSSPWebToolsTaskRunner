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

            //using (CSSPDBEntities db2 = new CSSPDBEntities())
            //{
            //    mapInfoActiveList = (from c in db2.MapInfos
            //                         from a in TVItemIDPolSourceSiteActiveList
            //                         where c.TVItemID == a
            //                         select c).ToList();

            //    List<int> mapInfoIDActiveList = mapInfoActiveList.Select(c => c.MapInfoID).ToList();

            //    mapInfoPointActiveList = (from c in db2.MapInfoPoints
            //                              from a in mapInfoIDActiveList
            //                              where c.MapInfoID == a
            //                              select c).ToList();

            //}

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 10);

            sbTemp.AppendLine($@"<table class=""PolSourceSiteCompact"">");
            sbTemp.AppendLine($@"   <tr> ");
            sbTemp.AppendLine($@"       <th class=""allBordersNoWrap"">{ TaskRunnerServiceRes.Site }</th>");
            sbTemp.AppendLine($@"       <th class=""allBordersNoWrap"">{ TaskRunnerServiceRes.IN }</th>");
            sbTemp.AppendLine($@"       <th class=""allBordersNoWrap"">{ TaskRunnerServiceRes.Type }</th>");
            sbTemp.AppendLine($@"       <th class=""allBordersNoWrap"">{ TaskRunnerServiceRes.Path }</th>");
            sbTemp.AppendLine($@"       <th class=""allBordersNoWrap"">{ TaskRunnerServiceRes.Prob }</th>");
            sbTemp.AppendLine($@"       <th class=""allBordersNoWrap"">{ TaskRunnerServiceRes.Risk }</th>");
            sbTemp.AppendLine($@"       <th class=""allBordersNoWrap"">{ TaskRunnerServiceRes.Lat }</th>");
            sbTemp.AppendLine($@"       <th class=""allBordersNoWrap"">{ TaskRunnerServiceRes.Long }</th>");
            sbTemp.AppendLine($@"       <th class=""allBordersNoWrap"">{ TaskRunnerServiceRes.ObsDate }</th>");
            sbTemp.AppendLine($@"   </tr>");

            using (CSSPDBEntities db = new CSSPDBEntities())
            {

                var tvItemSS = (from t in db.TVItems
                                    from tl in db.TVItemLanguages
                                    where t.TVItemID == tl.TVItemID
                                    && tl.Language == (int)language
                                    && t.TVType == (int)TVTypeEnum.Subsector
                                    && t.IsActive == true
                                    && t.TVItemID == TVItemID
                                    select new { t, tl }).FirstOrDefault();

                var PollutionSourceSiteList = (from t in db.TVItems
                                               from tl in db.TVItemLanguages
                                               from mi in db.MapInfos
                                               from mip in db.MapInfoPoints
                                               from pss in db.PolSourceSites
                                               let address = (from a in db.Addresses
                                                              let muni = (from cl in db.TVItemLanguages where cl.TVItemID == a.MunicipalityTVItemID && cl.Language == (int)LanguageEnum.en select cl.TVText).FirstOrDefault<string>()
                                                              let add = a.StreetNumber + " " + a.StreetName + " --- " + muni
                                                              where a.AddressTVItemID == t.TVItemID
                                                              select new { add }).FirstOrDefault()
                                               let pso = (from pso in db.PolSourceObservations
                                                          where pso.PolSourceSiteID == pss.PolSourceSiteID
                                                          orderby pso.ObservationDate_Local descending
                                                          select new { pso }).FirstOrDefault()
                                               let psi = (from psi in db.PolSourceObservationIssues
                                                          where psi.PolSourceObservationID == pso.pso.PolSourceObservationID
                                                          orderby psi.Ordinal ascending
                                                          select new { psi }).ToList()
                                               where t.TVItemID == tl.TVItemID
                                               && mi.TVItemID == t.TVItemID
                                               && mip.MapInfoID == mi.MapInfoID
                                               && t.TVItemID == pss.PolSourceSiteTVItemID
                                               && tl.Language == (int)LanguageEnum.en
                                               && t.TVPath.StartsWith(tvItemSS.t.TVPath + "p")
                                               && t.TVType == (int)TVTypeEnum.PolSourceSite
                                               && t.IsActive == true
                                               && mi.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Point
                                               && mi.TVType == (int)TVTypeEnum.PolSourceSite
                                               orderby tl.TVText
                                               select new { t, tl, mip, address.add, pss, pso, psi }).ToList();

                foreach (var polSourceSite in PollutionSourceSiteList.Where(c => c.t.ParentID == tvItemSS.t.TVItemID))
                {
                    string SS = tvItemSS.tl.TVText.Replace(",", "_");
                    string Desc = "";
                    if (SS.Contains(" "))
                    {
                        Desc = SS.Substring(SS.IndexOf(" "));
                        Desc = Desc.Trim();
                        Desc = Desc.Replace("(", "").Replace(")", "");
                        SS = SS.Substring(0, SS.IndexOf(" "));
                    }
                    string PSS = "P" + (polSourceSite.pss != null && polSourceSite.pss.Site != null ? polSourceSite.pss.Site.ToString().Replace(",", "_") : "");
                    string OBSDate = (polSourceSite.pso != null && polSourceSite.pso.pso.ObservationDate_Local != null ? polSourceSite.pso.pso.ObservationDate_Local.ToString("yyyy-MM-dd") : "");
                    string PSTVT = polSourceSite.tl.TVText;
                    string[] PSArr = PSTVT.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToArray();
                    string PSType = "";
                    string PSPath = "";
                    string PSProb = "";
                    string PSRisk = "";
                    if (PSArr.Length > 0)
                    {
                        PSType = PSArr[0];
                        if (PSType.Contains(" - "))
                        {
                            PSType = PSType.Substring(PSType.IndexOf(" - ") + 3);
                        }
                    }
                    if (PSArr.Length > 1)
                    {
                        PSPath = PSArr[1];
                    }
                    if (PSArr.Length > 2)
                    {
                        PSRisk = PSArr[2];
                    }
                    string Lat = (polSourceSite.mip != null ? polSourceSite.mip.Lat.ToString("F5") : "");
                    string Lng = (polSourceSite.mip != null ? polSourceSite.mip.Lng.ToString("F5") : "");
                    string URL = @"http://wmon01dtchlebl2/csspwebtools/en-CA/#!View/a|||" + polSourceSite.t.TVItemID.ToString() + @"|||30010100004000000000000000000000";

                    string TVText = "";

                    int IN = 0;
                    foreach (var psi in polSourceSite.psi)
                    {
                        if (psi != null && psi.psi != null)
                        {
                            IN += 1;
                            List<string> ObservationInfoList = (string.IsNullOrWhiteSpace(psi.psi.ObservationInfo) ? new List<string>() : psi.psi.ObservationInfo.Trim().Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList());

                            for (int i = 0, countObs = ObservationInfoList.Count; i < countObs; i++)
                            {
                                string Temp = _BaseEnumService.GetEnumText_PolSourceObsInfoReportEnum((PolSourceObsInfoEnum)int.Parse(ObservationInfoList[i]));
                                switch (ObservationInfoList[i].Substring(0, 3))
                                {
                                    case "101":
                                        {
                                            Temp = Temp.Replace("Source", "     Source");
                                        }
                                        break;
                                    case "250":
                                        {
                                            Temp = Temp.Replace("Pathway", "     Pathway");
                                        }
                                        break;
                                    case "900":
                                        {
                                            Temp = Temp.Replace("Status", "     Status");
                                            if (!string.IsNullOrWhiteSpace(Temp))
                                            {
                                                PSProb = Temp.Replace("Status:", "");
                                                PSProb = PSProb.Trim();
                                                if (PSProb.Contains(" "))
                                                {
                                                    PSProb = PSProb.Substring(0, PSProb.IndexOf(" "));
                                                }
                                            }
                                        }
                                        break;
                                    case "910":
                                        {
                                            Temp = Temp.Replace("Risk", "     Risk");
                                        }
                                        break;
                                    default:
                                        break;
                                }
                                TVText = TVText + Temp;
                            }
                        }

                        string TVT = (polSourceSite.pso != null && polSourceSite.pso.pso.Observation_ToBeDeleted != null ? polSourceSite.pso.pso.Observation_ToBeDeleted : "");

                        string TempISS = (!string.IsNullOrWhiteSpace(TVT) ? TVT.Replace(",", "_") + " ----- " : "") + TVText;

                        string ISS = TempISS.Replace("\r", "   ").Replace("\n", "").Replace("empty", "").Replace("Empty", "").Replace("\r", "   ").Replace("\n", "");

                        if (SS.Length == 0)
                        {
                            SS = " ";
                        }
                        if (Desc.Length == 0)
                        {
                            Desc = " ";
                        }
                        if (PSS.Length == 0)
                        {
                            PSS = " ";
                        }
                        if (PSType.Length == 0)
                        {
                            PSType = " ";
                        }
                        if (PSPath.Length == 0)
                        {
                            PSPath = " ";
                        }
                        if (PSProb.Length == 0)
                        {
                            PSProb = " ";
                        }
                        if (Lat.Length == 0)
                        {
                            Lat = " ";
                        }
                        if (Lng.Length == 0)
                        {
                            Lng = " ";
                        }
                        if (OBSDate.Length == 0)
                        {
                            OBSDate = " ";
                        }
                        if (URL.Length == 0)
                        {
                            URL = " ";
                        }

                        sbTemp.AppendLine($@"   <tr>");
                        sbTemp.AppendLine($@"       <td class=""allBordersNoWrap"">{ PSS }</td>");
                        sbTemp.AppendLine($@"       <td class=""allBordersNoWrap"">{ IN }</td>");
                        sbTemp.AppendLine($@"       <td class=""allBordersNoWrap"">{ PSType }</td>");
                        sbTemp.AppendLine($@"       <td class=""allBordersNoWrap"">{ PSPath }</td>");
                        sbTemp.AppendLine($@"       <td class=""allBordersNoWrap"">{ PSProb }</td>");
                        sbTemp.AppendLine($@"       <td class=""allBordersNoWrap"">{ PSRisk }</td>");
                        sbTemp.AppendLine($@"       <td class=""allBordersNoWrap"">{ Lat }</td>");
                        sbTemp.AppendLine($@"       <td class=""allBordersNoWrap"">{ Lng }</td>");
                        sbTemp.AppendLine($@"       <td class=""allBordersNoWrap"">{ OBSDate }</td>");
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
