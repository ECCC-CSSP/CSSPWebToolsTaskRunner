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
//using System.Web.Helpers;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        private bool GenerateHTMLSUBSECTOR_POLLUTION_SOURCE_SITES(StringBuilder sbTemp)
        {
            int Percent = 10;
            string NotUsed = "";
            LanguageEnum language = _TaskRunnerBaseService._BWObj.appTaskModel.Language;

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", ReportGenerateObjectsKeywordEnum.SUBSECTOR_POLLUTION_SOURCE_SITES.ToString()));


            //sbTemp.AppendLine("<h2>SUBSECTOR_POLLUTION_SOURCE_SITES - Not implemented</h2>");

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

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 5);

            List<TVItemModel> tvItemModelListPolSourceSite = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.PolSourceSite);
            List<PolSourceSiteModel> polSourceSiteModelList = _PolSourceSiteService.GetPolSourceSiteModelListWithSubsectorTVItemIDDB(TVItemID).OrderBy(c => c.Site).ToList();
            List<PolSourceObservationModel> polSourceObservationModelList = _PolSourceObservationService.GetPolSourceObservationModelListWithSubsectorTVItemIDDB(TVItemID);
            List<PolSourceObservationIssueModel> polSourceObservationIssueModelList = _PolSourceObservationIssueService.GetPolSourceObservationIssueModelListWithSubsectorTVItemIDDB(TVItemID);
            List<int> TVItemIDPolSourceSiteActiveList = tvItemModelListPolSourceSite.Where(c => c.IsActive == true).Select(c => c.TVItemID).ToList();
            List<int> TVItemIDPolSourceSiteInactiveList = tvItemModelListPolSourceSite.Where(c => c.IsActive == false).Select(c => c.TVItemID).ToList();
            List<int> TVItemIDCivicAddressList = polSourceSiteModelList.Where(c => c.CivicAddressTVItemID != null && c.CivicAddressTVItemID != 0).Where(c => c != null).Select(c => (int)c.CivicAddressTVItemID).ToList();
            List<int> TVItemIDContactList = polSourceObservationModelList.Select(c => c.ContactTVItemID).ToList();

            List<Address> addressList = new List<Address>();
            List<MapInfo> mapInfoActiveList = new List<MapInfo>();
            List<MapInfo> mapInfoInactiveList = new List<MapInfo>();
            List<MapInfoPoint> mapInfoPointActiveList = new List<MapInfoPoint>();
            List<MapInfoPoint> mapInfoPointInactiveList = new List<MapInfoPoint>();
            List<TVItemLanguage> countryList = new List<TVItemLanguage>();
            List<TVItemLanguage> provinceList = new List<TVItemLanguage>();
            List<TVItemLanguage> municipalityList = new List<TVItemLanguage>();
            List<TVItemLanguage> contactList = new List<TVItemLanguage>();

            using (CSSPDBEntities db2 = new CSSPDBEntities())
            {
                addressList = (from c in db2.Addresses
                               from a in TVItemIDCivicAddressList
                               where c.AddressTVItemID == a
                               select c).ToList();

                List<int> countryTVItemIDList = addressList.Select(c => c.CountryTVItemID).ToList();
                List<int> provinceTVItemIDList = addressList.Select(c => c.ProvinceTVItemID).ToList();
                List<int> municipalityTVItemIDList = addressList.Select(c => c.MunicipalityTVItemID).ToList();

                countryList = (from c in db2.TVItemLanguages
                               from a in countryTVItemIDList
                               where c.TVItemID == a
                               && c.Language == (int)language
                               select c).ToList();

                provinceList = (from c in db2.TVItemLanguages
                                from a in provinceTVItemIDList
                                where c.TVItemID == a
                                && c.Language == (int)language
                                select c).ToList();

                municipalityList = (from c in db2.TVItemLanguages
                                    from a in municipalityTVItemIDList
                                    where c.TVItemID == a
                                    && c.Language == (int)language
                                    select c).ToList();

                contactList = (from c in db2.TVItemLanguages
                               from a in TVItemIDContactList
                               where c.TVItemID == a
                               && c.Language == (int)language
                               select c).ToList();


                mapInfoActiveList = (from c in db2.MapInfos
                                     from a in TVItemIDPolSourceSiteActiveList
                                     where c.TVItemID == a
                                     select c).ToList();

                mapInfoInactiveList = (from c in db2.MapInfos
                                       from a in TVItemIDPolSourceSiteInactiveList
                                       where c.TVItemID == a
                                       select c).ToList();

                List<int> mapInfoIDActiveList = mapInfoActiveList.Select(c => c.MapInfoID).ToList();

                mapInfoPointActiveList = (from c in db2.MapInfoPoints
                                          from a in mapInfoIDActiveList
                                          where c.MapInfoID == a
                                          select c).ToList();

                List<int> mapInfoIDInactiveList = mapInfoInactiveList.Select(c => c.MapInfoID).ToList();

                mapInfoPointInactiveList = (from c in db2.MapInfoPoints
                                            from a in mapInfoIDInactiveList
                                            where c.MapInfoID == a
                                            select c).ToList();

            }

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 10);

            //sbTemp.AppendLine(@"<p>|||PAGE_BREAK|||</p>");
            //sbTemp.AppendLine("");
            //sbTemp.AppendLine($@"<h1 style=""text-align: center"">{ tvItemModelSubsector.TVText }</h1>");
            //sbTemp.AppendLine($@"<h2 style=""text-align: center"">{ TaskRunnerServiceRes.ActivePollutionSourceSites }</h2>");
            //int PSSNumber = 0;
            foreach (TVItemModel tvItemModelPSSActive in tvItemModelListPolSourceSite.Where(c => c.IsActive == true))
            {
                //PSSNumber += 1;

                PolSourceSiteModel polSourceSiteModel = polSourceSiteModelList.Where(c => c.PolSourceSiteTVItemID == tvItemModelPSSActive.TVItemID).FirstOrDefault();

                if (polSourceSiteModel != null)
                {
                    string tvText = polSourceSiteModel.PolSourceSiteTVText;
                    if (tvText.Contains(" "))
                    {
                        tvText = tvText.Substring(0, tvText.IndexOf(" "));
                    }

                    sbTemp.AppendLine($@"<div class=""smalltext"">");
                    sbTemp.AppendLine($@"<p>");
                    sbTemp.AppendLine($@"<strong>{TaskRunnerServiceRes.Site} #</strong>: {tvText}&nbsp;&nbsp;&nbsp;&nbsp;");

                    MapInfo mapInfo = mapInfoActiveList.Where(c => c.TVItemID == tvItemModelPSSActive.TVItemID).FirstOrDefault();
                    if (mapInfo != null)
                    {
                        List<MapInfoPoint> mapInfoPointListCurrent = mapInfoPointActiveList.Where(c => c.MapInfoID == mapInfo.MapInfoID).ToList();
                        if (mapInfoPointListCurrent.Count > 0)
                        {
                            sbTemp.AppendLine($@"<span><strong>{TaskRunnerServiceRes.Lat} {TaskRunnerServiceRes.Long}</strong>: {mapInfoPointListCurrent[0].Lat.ToString("F5")} {mapInfoPointListCurrent[0].Lng.ToString("F5")}</span>");
                        }
                    }
                    else
                    {
                        sbTemp.AppendLine($@"<span><strong>{TaskRunnerServiceRes.Lat} {TaskRunnerServiceRes.Long}</strong>: --- ---</span>");
                    }
                    sbTemp.AppendLine($@"&nbsp;&nbsp;&nbsp;&nbsp;");

                    //sbTemp.AppendLine($@"</p>");


                    //if (polSourceSiteModel.CivicAddressTVItemID != null)
                    //{
                    //    if (polSourceSiteModel.CivicAddressTVItemID != 0)
                    //    {
                    //        Address address = addressList.Where(c => c.AddressTVItemID == ((int)polSourceSiteModel.CivicAddressTVItemID)).FirstOrDefault();
                    //        if (address != null)
                    //        {
                    //            sbTemp.AppendLine($@"<p>");
                    //            if (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr)
                    //            {
                    //                string CountryText = countryList.Where(c => c.TVItemID == address.CountryTVItemID).Select(c => c.TVText).FirstOrDefault();
                    //                string ProvinceText = provinceList.Where(c => c.TVItemID == address.ProvinceTVItemID).Select(c => c.TVText).FirstOrDefault();
                    //                string MunicipalityText = municipalityList.Where(c => c.TVItemID == address.MunicipalityTVItemID).Select(c => c.TVText).FirstOrDefault();
                    //                string StreetTypeText = _BaseEnumService.GetEnumText_StreetTypeEnum((StreetTypeEnum)address.StreetType);
                    //                string AddressText = $" { address.StreetNumber} { address.StreetName } { StreetTypeText }, { MunicipalityText }, { ProvinceText }, { CountryText }";
                    //                sbTemp.AppendLine($@"<strong>{ TaskRunnerServiceRes.CivicAddress }</strong>: { AddressText }");
                    //            }
                    //            else
                    //            {
                    //                string CountryText = countryList.Where(c => c.TVItemID == address.CountryTVItemID).Select(c => c.TVText).FirstOrDefault();
                    //                string ProvinceText = provinceList.Where(c => c.TVItemID == address.ProvinceTVItemID).Select(c => c.TVText).FirstOrDefault();
                    //                string MunicipalityText = municipalityList.Where(c => c.TVItemID == address.MunicipalityTVItemID).Select(c => c.TVText).FirstOrDefault();
                    //                string StreetTypeText = _BaseEnumService.GetEnumText_StreetTypeEnum((StreetTypeEnum)address.StreetType);
                    //                string AddressText = $" { address.StreetNumber}, { StreetTypeText } { address.StreetName }, { MunicipalityText }, { ProvinceText }, { CountryText }";
                    //                sbTemp.AppendLine($@"<strong>{ TaskRunnerServiceRes.CivicAddress }</strong>: { AddressText }");
                    //            }
                    //            sbTemp.AppendLine($@"</p>");
                    //        }
                    //    }
                    //}

                    PolSourceObservationModel polSourceObservationModel = polSourceObservationModelList.Where(c => c.PolSourceSiteID == polSourceSiteModel.PolSourceSiteID).OrderByDescending(c => c.ObservationDate_Local).FirstOrDefault();

                    if (polSourceObservationModel != null)
                    {
                        //sbTemp.AppendLine($@"<p>");
                        string ContactObsText = contactList.Where(c => c.TVItemID == polSourceObservationModel.ContactTVItemID).Select(c => c.TVText).FirstOrDefault();
                        sbTemp.AppendLine($@"<strong>{TaskRunnerServiceRes.LastObservationDate}</strong>: {polSourceObservationModel.ObservationDate_Local.ToString("yyyy MMMM dd")}");
                        //sbTemp.AppendLine($@"<strong>{TaskRunnerServiceRes.LastObservationDate}</strong>: {polSourceObservationModel.ObservationDate_Local.ToString("yyyy MMMM dd")} <strong>{TaskRunnerServiceRes.by}</strong>: {ContactObsText}");

                        int IssueNumber = 0;
                        foreach (PolSourceObservationIssueModel polSourceObservationIssueModel in polSourceObservationIssueModelList.Where(c => c.PolSourceObservationID == polSourceObservationModel.PolSourceObservationID).OrderBy(c => c.Ordinal).ToList())
                        {
                            IssueNumber += 1;

                            if (IssueNumber > 1) continue;

                            string TVText = "";
                            //sbTemp.AppendLine($@"<p><strong>{ TaskRunnerServiceRes.Issue }</strong>: { IssueNumber }</p>");
                            //sbTemp.AppendLine($@"<blockquote>");
                            List<string> ObservationInfoList = (string.IsNullOrWhiteSpace(polSourceObservationIssueModel.ObservationInfo) ? new List<string>() : polSourceObservationIssueModel.ObservationInfo.Trim().Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList());

                            if (ObservationInfoList.Count > 0)
                            {
                                for (int i = 0, countObs = ObservationInfoList.Count; i < countObs; i++)
                                {
                                    string Temp = _BaseEnumService.GetEnumText_PolSourceObsInfoReportEnum((PolSourceObsInfoEnum)int.Parse(ObservationInfoList[i]));
                                    switch (ObservationInfoList[i].Substring(0, 3))
                                    {
                                        case "101":
                                            {
                                                Temp = Temp.Replace("Source", "<br /><strong>Source</strong>");
                                            }
                                            break;
                                        //case "153":
                                        //    {
                                        //        Temp = Temp.Replace("Dilution Analyses", "     Dilution Analyses");
                                        //    }
                                        //    break;
                                        case "250":
                                            {
                                                Temp = Temp.Replace("Pathway", "<br /><strong>Pathway</strong>");
                                            }
                                            break;
                                        case "900":
                                            {
                                                Temp = Temp.Replace("Status", "<br /><strong>Status</strong>");
                                            }
                                            break;
                                        case "910":
                                            {
                                                Temp = Temp.Replace("Risk", "<br /><strong>Risk</strong>");
                                            }
                                            break;
                                        case "110":
                                        case "120":
                                        case "122":
                                        case "151":
                                        case "152":
                                        case "153":
                                        case "155":
                                        case "156":
                                        case "157":
                                        case "163":
                                        case "166":
                                        case "167":
                                        case "170":
                                        case "171":
                                        case "172":
                                        case "173":
                                        case "176":
                                        case "178":
                                        case "181":
                                        case "182":
                                        case "183":
                                        case "185":
                                        case "186":
                                        case "187":
                                        case "190":
                                        case "191":
                                        case "192":
                                        case "193":
                                        case "194":
                                        case "196":
                                        case "198":
                                        case "199":
                                        case "220":
                                        case "930":
                                            {
                                                //Temp = @"<span class=""hidden"">" + Temp + "</span>";
                                            }
                                            break;
                                        default:
                                            break;
                                    }
                                    TVText = TVText + Temp;
                                }

                                sbTemp.AppendLine($@"{TVText}");
                                sbTemp.AppendLine($@"</p>");

                                //if (polSourceObservationIssueModel.ExtraComment != null)
                                //{
                                //    if (polSourceObservationIssueModel.ExtraComment.Length > 0)
                                //    {
                                //        sbTemp.AppendLine($@"<p><strong>{TaskRunnerServiceRes.ExtraComment}</strong></p>");
                                //        sbTemp.AppendLine($@"<p>");
                                //        sbTemp.AppendLine($@"{polSourceObservationIssueModel.ExtraComment}");
                                //        sbTemp.AppendLine($@"</p>");
                                //    }
                                //}

                            }
                            //sbTemp.AppendLine($@"</blockquote>");
                        }

                    }
                    else
                    {
                        sbTemp.AppendLine($@"</p>");
                    }

                    sbTemp.AppendLine($@"</div>");
                    //sbTemp.AppendLine($@"<hr />");
                }
            }

            Percent = 98;
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);

            return true;
        }

        // for testing only can comment out when test is completed
        public bool PublicGenerateHTMLSUBSECTOR_POLLUTION_SOURCE_SITES(StringBuilder sbTemp)
        {
            bool retBool = GenerateHTMLSUBSECTOR_POLLUTION_SOURCE_SITES(sbTemp);

            StreamWriter sw = fi.CreateText();
            sw.Write(sbTemp.ToString());
            sw.Flush();
            sw.Close();

            return retBool;
        }
    }
}
