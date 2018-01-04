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
        private bool GenerateHTMLSUBSECTOR_MWQM_SITES_DATA_AVAILABILITY(StringBuilder sbTemp)
        {
            string NotUsed = "";

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 3);

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

            List<TVItemModel> tvItemModelListMWQMSites = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.MWQMSite).Where(c => c.IsActive == true).ToList();
            List<MWQMSiteModel> mwqmSiteModelList = _MWQMSiteService.GetMWQMSiteModelListWithSubsectorTVItemIDDB(TVItemID);
            List<MWQMRunModel> mwqmRunModelList = _MWQMRunService.GetMWQMRunModelListWithSubsectorTVItemIDDB(TVItemID);
            List<MWQMSampleModel> mwqmSampleModelList = _MWQMSampleService.GetMWQMSampleModelListWithSubsectorTVItemIDDB(TVItemID);

            sbTemp.AppendLine($@" <h3>{ TaskRunnerServiceRes.MWQMSiteSampleDataAvailability }</h3>");
            sbTemp.AppendLine($@" <table cellpadding=""5"">");
            sbTemp.AppendLine($@" <tr>");
            sbTemp.AppendLine($@" <th>{ TaskRunnerServiceRes.Site }</th>");
            bool FirstHit = false;
            for (int year = DateTime.Now.Year; year > 1975; year--)
            {
                if (year % 5 == 0)
                {
                    FirstHit = true;
                    int colSpan = 5;
                    if (year == 1980)
                    {
                        colSpan = 4;
                    }
                    sbTemp.AppendLine($@" <th class=""textAlignLeftAndLeftBorder"" colspan=""{ colSpan }"">{ year }</th>");
                }
                if (!FirstHit)
                {
                    sbTemp.AppendLine($@" <th>&nbsp;</th>");
                }
            }
            sbTemp.AppendLine($@" </tr>");
            int countSite = 0;
            foreach (MWQMSiteModel mwqmSiteModel in mwqmSiteModelList)
            {
                TVItemModel tvItemModel = tvItemModelListMWQMSites.Where(c => c.TVItemID == mwqmSiteModel.MWQMSiteTVItemID).FirstOrDefault();
                if (tvItemModel != null)
                {
                    if (tvItemModel.IsActive)
                    {
                        countSite += 1;
                        string bottomClass = "";
                        if (countSite % 5 == 0)
                        {
                            bottomClass = "bottomBorder";
                        }
                        sbTemp.AppendLine($@" <tr>");
                        sbTemp.AppendLine($@" <td class=""{ bottomClass }"">{ mwqmSiteModel.MWQMSiteTVText }</td>");
                        for (int year = DateTime.Now.Year; year > 1979; year--)
                        {
                            string leftClass = year % 5 == 0 ? "leftBorder" : "";
                            bool hasSamples = mwqmSampleModelList.Where(c => c.MWQMSiteTVItemID == mwqmSiteModel.MWQMSiteTVItemID && c.SampleDateTime_Local.Year == year && c.SampleTypesText.Contains(((int)SampleTypeEnum.Routine).ToString())).Any();
                            if (hasSamples)
                            {
                                if (leftClass != "")
                                {
                                    if (bottomClass != "")
                                    {
                                        sbTemp.AppendLine($@" <td class=""bggreenfLeftAndBottomBorder"">&nbsp;</td>");
                                    }
                                    else
                                    {
                                        sbTemp.AppendLine($@" <td class=""bggreenfLeftBorder"">&nbsp;</td>");
                                    }
                                }
                                else
                                {
                                    if (bottomClass != "")
                                    {
                                        sbTemp.AppendLine($@" <td class=""bggreenfBottomBorder"">&nbsp;</td>");
                                    }
                                    else
                                    {
                                        sbTemp.AppendLine($@" <td class=""bggreenf"">&nbsp;</td>");
                                    }
                                }
                            }
                            else
                            {
                                if (leftClass != "")
                                {
                                    if (bottomClass != "")
                                    {
                                        sbTemp.AppendLine($@" <td class=""leftAndBottomBorder"">&nbsp;</td>");
                                    }
                                    else
                                    {
                                        sbTemp.AppendLine($@" <td class=""leftBorder"">&nbsp;</td>");
                                    }
                                }
                                else
                                {
                                    if (bottomClass != "")
                                    {
                                        sbTemp.AppendLine($@" <td class=""bottomBorder"">&nbsp;</td>");
                                    }
                                    else
                                    {
                                        sbTemp.AppendLine($@" <td>&nbsp;</td>");
                                    }
                                }
                            }
                        }
                        sbTemp.AppendLine($@" </tr>");
                    }
                }
            }
            sbTemp.AppendLine($@" </table>");

            return true;
        }
        public bool PublicGenerateHTMLSUBSECTOR_MWQM_SITES_DATA_AVAILABILITY(StringBuilder sbTemp)
        {
            bool retBool = GenerateHTMLSUBSECTOR_MWQM_SITES_DATA_AVAILABILITY(sbTemp);

            StreamWriter sw = fi.CreateText();
            sw.Write(sbTemp.ToString());
            sw.Flush();
            sw.Close();

            return retBool;
        }
    }
}

