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
        private bool GenerateHTMLSubsectorMWQMSites(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            string NotUsed = "";
            int TVItemID = 0;

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

            sbHTML.AppendLine($@" <h3>{ TaskRunnerServiceRes.MWQMSiteSampleDataAvailability }</h3>");
            sbHTML.AppendLine($@" <table cellpadding=""5"">");
            sbHTML.AppendLine($@" <tr>");
            sbHTML.AppendLine($@" <th>{ TaskRunnerServiceRes.Site }</th>");
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
                    sbHTML.AppendLine($@" <th class=""textAlignLeftAndLeftBorder"" colspan=""{ colSpan }"">{ year }</th>");
                }
                if (!FirstHit)
                {
                    sbHTML.AppendLine($@" <th>&nbsp;</th>");
                }
            }
            sbHTML.AppendLine($@" </tr>");
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
                        sbHTML.AppendLine($@" <tr>");
                        sbHTML.AppendLine($@" <td class=""{ bottomClass }"">{ mwqmSiteModel.MWQMSiteTVText }</td>");
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
                                        sbHTML.AppendLine($@" <td class=""bggreenfLeftAndBottomBorder"">&nbsp;</td>");
                                    }
                                    else
                                    {
                                        sbHTML.AppendLine($@" <td class=""bggreenfLeftBorder"">&nbsp;</td>");
                                    }
                                }
                                else
                                {
                                    if (bottomClass != "")
                                    {
                                        sbHTML.AppendLine($@" <td class=""bggreenfBottomBorder"">&nbsp;</td>");
                                    }
                                    else
                                    {
                                        sbHTML.AppendLine($@" <td class=""bggreenf"">&nbsp;</td>");
                                    }
                                }
                            }
                            else
                            {
                                if (leftClass != "")
                                {
                                    if (bottomClass != "")
                                    {
                                        sbHTML.AppendLine($@" <td class=""leftAndBottomBorder"">&nbsp;</td>");
                                    }
                                    else
                                    {
                                        sbHTML.AppendLine($@" <td class=""leftBorder"">&nbsp;</td>");
                                    }
                                }
                                else
                                {
                                    if (bottomClass != "")
                                    {
                                        sbHTML.AppendLine($@" <td class=""bottomBorder"">&nbsp;</td>");
                                    }
                                    else
                                    {
                                        sbHTML.AppendLine($@" <td>&nbsp;</td>");
                                    }
                                }
                            }
                        }
                        sbHTML.AppendLine($@" </tr>");
                    }
                }
            }
            sbHTML.AppendLine($@" </table>");

            sbHTML.AppendLine(@"<span>|||PageBreak|||</span>");


            sbHTML.AppendLine($@" <h3>{ TaskRunnerServiceRes.MWQMSitesInformation }</h3>");
            sbHTML.AppendLine($@" <table cellpadding=""5"" class=""textAlignCenter"">");
            sbHTML.AppendLine($@" <tr>");
            sbHTML.AppendLine($@" <th colspan=""2"">{ TaskRunnerServiceRes.Site }</th>");
            sbHTML.AppendLine($@" <th>{ TaskRunnerServiceRes.Coordinates }</th>");
            sbHTML.AppendLine($@" <th>{ TaskRunnerServiceRes.Description }</th>");
            sbHTML.AppendLine($@" <th>{ TaskRunnerServiceRes.Photos }</th>");
            sbHTML.AppendLine($@" </tr>");
            foreach (MWQMSiteModel mwqmSiteModel in mwqmSiteModelList)
            {
                TVItemModel tvItemModel = tvItemModelListMWQMSites.Where(c => c.TVItemID == mwqmSiteModel.MWQMSiteTVItemID).FirstOrDefault();
                if (tvItemModel != null)
                {
                    if (tvItemModel.IsActive)
                    {
                        string classificationLetter = "";
                        string classificationColor = "";

                        classificationLetter = GetLastClassificationInitial(mwqmSiteModel.MWQMSiteLatestClassification);
                        classificationColor = GetLastClassificationColor(mwqmSiteModel.MWQMSiteLatestClassification);

                        sbHTML.AppendLine($@" <tr>");
                        sbHTML.AppendLine($@" <td>{ mwqmSiteModel.MWQMSiteTVText }</td>");
                        sbHTML.AppendLine($@" <td class=""{ classificationColor }"">{ classificationLetter }</td>");

                        List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(mwqmSiteModel.MWQMSiteTVItemID, TVTypeEnum.MWQMSite, MapInfoDrawTypeEnum.Point);
                        if (mapInfoPointModelList.Count > 0)
                        {
                            sbHTML.AppendLine($@" <td>{ mapInfoPointModelList[0].Lat.ToString("F5") } { mapInfoPointModelList[0].Lng.ToString("F5") }</td>");
                        }
                        else
                        {
                            sbHTML.AppendLine($@" <td>&nbsp;</td>");
                        }
                        sbHTML.AppendLine($@" <td class=""textAlignLeft"">{ mwqmSiteModel.MWQMSiteDescription }</td>");
                        sbHTML.AppendLine($@" <td>Photo</td>");
                        sbHTML.AppendLine($@" </tr>");
                    }
                }
            }
            sbHTML.AppendLine($@" </table>");

            sbHTML.AppendLine($@"|||FileNameExtra|Random,{ FileNameExtra }|||");

            sbHTML.AppendLine(@"<span>|||PageBreak|||</span>");

            if (!GetBottomHTML(sbHTML, fi, parameters))
            {
                return false;
            }

            return true;
        }

        // for testing only can comment out when test is completed
        public bool PublicGenerateHTMLSubsectorMWQMSites(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            bool retBool = GenerateHTMLSubsectorMWQMSites(fi, sbHTML, parameters, reportTypeModel);

            StreamWriter sw = fi.CreateText();
            sw.Write(sbHTML.ToString());
            sw.Flush();
            sw.Close();

            return retBool;
        }
    }
}
