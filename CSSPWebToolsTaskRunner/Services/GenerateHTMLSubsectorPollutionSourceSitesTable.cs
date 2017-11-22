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
        private bool GenerateHTMLSubsectorPollutionSourceSitesTable(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
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
                NotUsed = tvItemModelSubsector.Error;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(tvItemModelSubsector.Error);
                return false;
            }

            string ServerPath = _TVFileService.GetServerFilePath(tvItemModelSubsector.TVItemID);

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 5);

            List<TVItemModel> tvItemModelListPolSourceSite = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.PolSourceSite).Where(c => c.IsActive == true).ToList();
            List<PolSourceSiteModel> polSourceSiteModelList = _PolSourceSiteService.GetPolSourceSiteModelListWithSubsectorTVItemIDDB(TVItemID);
            List<PolSourceObservationModel> polSourceObservationModelList = _PolSourceObservationService.GetPolSourceObservationModelListWithSubsectorTVItemIDDB(TVItemID);
            List<PolSourceObservationIssueModel> polSourceObservationIssueModelList = _PolSourceObservationIssueService.GetPolSourceObservationIssueModelListWithSubsectorTVItemIDDB(TVItemID);

            List<PolSourceObsInfoEnumTextAndID> polSourceObsInfoEnumTextAndIDListPolSourceType = new List<PolSourceObsInfoEnumTextAndID>();

            foreach (int id in Enum.GetValues(typeof(PolSourceObsInfoEnum)))
            {
                if (id == 0)
                    continue;

                if (id.ToString().StartsWith("105") && !id.ToString().EndsWith("00"))
                {
                    polSourceObsInfoEnumTextAndIDListPolSourceType.Add(new PolSourceObsInfoEnumTextAndID() { Text = _BaseEnumService.GetEnumText_PolSourceObsInfoEnum((PolSourceObsInfoEnum)id), ID = id });
                }
            }

            polSourceObsInfoEnumTextAndIDListPolSourceType = polSourceObsInfoEnumTextAndIDListPolSourceType.OrderBy(c => c.Text).ToList();

            sbHTML.AppendLine($@" <h3>{ TaskRunnerServiceRes.LandBasePollutionSourceSiteObservationAndIssues }</h3>");
            foreach (PolSourceObsInfoEnumTextAndID PolSourceObsInfoEnumTextAndID in polSourceObsInfoEnumTextAndIDListPolSourceType)
            {
                sbHTML.AppendLine($@" <h4>{ PolSourceObsInfoEnumTextAndID.Text }</h4>");
                foreach (TVItemModel tvItemModel in tvItemModelListPolSourceSite)
                {
                    PolSourceSiteModel polSourceSiteModel = polSourceSiteModelList.Where(c => c.PolSourceSiteTVItemID == tvItemModel.TVItemID).FirstOrDefault();
                    if (polSourceSiteModel != null)
                    {
                        PolSourceObservationModel polSourceObservationModel = polSourceObservationModelList.Where(c => c.PolSourceSiteID == polSourceSiteModel.PolSourceSiteID).OrderByDescending(c => c.ObservationDate_Local).FirstOrDefault();
                        if (polSourceObservationModel != null)
                        {
                            List<PolSourceObservationIssueModel> polSourceObservationIssueModelList2 = polSourceObservationIssueModelList.Where(c => c.PolSourceObservationID == polSourceObservationModel.PolSourceObservationID).OrderBy(c => c.Ordinal).ToList();
                            if (polSourceObservationIssueModelList2.Count > 0)
                            {
                                if (polSourceObservationIssueModelList2[0].ObservationInfo.Contains(((int)PolSourceObsInfoEnum.LandBased).ToString() + ","))
                                {
                                    if (polSourceObservationModel == null)
                                    {
                                        sbHTML.AppendLine($@" <blockquote>");
                                        sbHTML.AppendLine($@" <span>{ TaskRunnerServiceRes.ID }: { polSourceSiteModel.Site }</span>");

                                        List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(polSourceSiteModel.PolSourceSiteTVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);
                                        if (mapInfoPointModelList.Count > 0)
                                        {
                                            sbHTML.AppendLine($@" <span>{ TaskRunnerServiceRes.Coord }: { mapInfoPointModelList[0].Lat }, { mapInfoPointModelList[0].Lng }</span>");
                                        }
                                        sbHTML.AppendLine($@"           <br />");
                                        sbHTML.AppendLine($@"           <span>{ TaskRunnerServiceRes.NoObservationForThisPollutionSourceSite }</span>");
                                        sbHTML.AppendLine($@" </blockquote>");
                                    }
                                    else
                                    {
                                        if (polSourceObservationIssueModelList2.Count == 0)
                                        {
                                            sbHTML.AppendLine($@" <blockquote>");
                                            sbHTML.AppendLine($@" <span>{ TaskRunnerServiceRes.ID }: { polSourceSiteModel.Site }</span>");

                                            List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(polSourceSiteModel.PolSourceSiteTVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);
                                            if (mapInfoPointModelList.Count > 0)
                                            {
                                                sbHTML.AppendLine($@" <span>{ TaskRunnerServiceRes.Coord }: { mapInfoPointModelList[0].Lat }, { mapInfoPointModelList[0].Lng }</span>");
                                            }
                                            sbHTML.AppendLine($@"           <br />");
                                            sbHTML.AppendLine($@"           <span>{ polSourceObservationModel.Observation_ToBeDeleted }</span>");
                                            sbHTML.AppendLine($@" </blockquote>");
                                        }
                                        else
                                        {
                                            bool IssueOfSourceTypeExist = false;
                                            foreach (PolSourceObservationIssueModel polSourceObservationIssueModel in polSourceObservationIssueModelList2)
                                            {
                                                if (polSourceObservationIssueModel.ObservationInfo.Contains(PolSourceObsInfoEnumTextAndID.ID.ToString()))
                                                {
                                                    IssueOfSourceTypeExist = true;
                                                }
                                            }

                                            if (IssueOfSourceTypeExist)
                                            {
                                                foreach (PolSourceObservationIssueModel polSourceObservationIssueModel in polSourceObservationIssueModelList2)
                                                {
                                                    List<int> obsInfoList = polSourceObservationIssueModel.ObservationInfo.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Where(c => !string.IsNullOrWhiteSpace(c)).Select(c => int.Parse(c)).ToList();
                                                    List<PolSourceObsInfoEnumTextAndID> polSourceObsInfoEnumTextAndIDList = new List<PolSourceObsInfoEnumTextAndID>();
                                                    foreach (int id in Enum.GetValues(typeof(PolSourceObsInfoEnum)))
                                                    {
                                                        if (id == 0)
                                                            continue;

                                                        polSourceObsInfoEnumTextAndIDList.Add(new PolSourceObsInfoEnumTextAndID() { Text = _BaseEnumService.GetEnumText_PolSourceObsInfoEnum((PolSourceObsInfoEnum)id), ID = id });
                                                    }
                                                    if (obsInfoList.Count > 1)
                                                    {
                                                        sbHTML.AppendLine($@"           <span>{ TaskRunnerServiceRes.Dist }: { polSourceObsInfoEnumTextAndIDList.Where(c => c.ID == (int)obsInfoList[1]).FirstOrDefault().Text }</span>");
                                                    }
                                                    if (obsInfoList.Count > 2)
                                                    {
                                                        sbHTML.AppendLine($@"           <span>{ TaskRunnerServiceRes.Slope }: { polSourceObsInfoEnumTextAndIDList.Where(c => c.ID == (int)obsInfoList[2]).FirstOrDefault().Text }</span>");
                                                    }

                                                    foreach (int obsInfo in obsInfoList.Skip(4))
                                                    {
                                                        sbHTML.AppendLine($@"           <span>{ polSourceObsInfoEnumTextAndIDList.Where(c => c.ID == obsInfo).FirstOrDefault().Text }</span>");
                                                    }
                                                    sbHTML.AppendLine($@"           <br />");
                                                    sbHTML.AppendLine($@"           <span>Photo</span>");

                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            sbHTML.AppendLine(@"<span>|||PageBreak|||</span>");

            if (!GetBottomHTML(sbHTML, fi, parameters))
            {
                return false;
            }

            return true;
        }

        // for testing only can comment out when test is completed
        public bool PublicGenerateHTMLSubsectorPollutionSourceSitesTable(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            bool retBool = GenerateHTMLSubsectorPollutionSourceSitesTable(fi, sbHTML, parameters, reportTypeModel);

            StreamWriter sw = fi.CreateText();
            sw.Write(sbHTML.ToString());
            sw.Flush();
            sw.Close();

            return retBool;
        }
    }
}
