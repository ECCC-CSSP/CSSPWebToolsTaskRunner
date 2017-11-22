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

            sbHTML.AppendLine($@" <h3>{ TaskRunnerServiceRes.LandBasePollutionSourceSiteObservationAndIssues }</h3>");
            sbHTML.AppendLine($@" <table border=""0"">");
            sbHTML.AppendLine($@"    <tr> ");
            sbHTML.AppendLine($@"       <th>{ TaskRunnerServiceRes.ID }</th>");
            sbHTML.AppendLine($@"       <th>{ TaskRunnerServiceRes.Dist }</th>");
            sbHTML.AppendLine($@"       <th>{ TaskRunnerServiceRes.Slope }</th>");
            sbHTML.AppendLine($@"       <th>{ TaskRunnerServiceRes.Type }</th>");
            sbHTML.AppendLine($@"       <th>{ TaskRunnerServiceRes.Issues }</th>");
            sbHTML.AppendLine($@"       <th>{ TaskRunnerServiceRes.Photos }</th>");
            sbHTML.AppendLine($@"    </tr> ");
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
                                sbHTML.AppendLine($@"    <tr> ");
                                sbHTML.AppendLine($@"       <td rowspan=""2""> ");
                                sbHTML.AppendLine($@"           <span>{ polSourceSiteModel.Site }</span>");
                                sbHTML.AppendLine($@"       </td> ");
                                if (polSourceObservationModel == null)
                                {
                                    sbHTML.AppendLine($@"       <td rowspan=""2"" colspan=""5""> ");
                                    sbHTML.AppendLine($@"           <span>{ TaskRunnerServiceRes.NoObservationForThisPollutionSourceSite }</span>");
                                    sbHTML.AppendLine($@"       </td> ");
                                }
                                else
                                {
                                    if (polSourceObservationIssueModelList2.Count == 0)
                                    {
                                        sbHTML.AppendLine($@"       <td rowspan=""2"" colspan=""5""> ");
                                        sbHTML.AppendLine($@"           <span>{ polSourceObservationModel.Observation_ToBeDeleted }</span>");
                                        sbHTML.AppendLine($@"       </td> ");
                                    }
                                    else
                                    {
                                        //sbHTML.AppendLine($@" <table border=""0"">");

                                        foreach (PolSourceObservationIssueModel polSourceObservationIssueModel in polSourceObservationIssueModelList2)
                                        {
                                            //sbHTML.AppendLine($@"    <tr> ");
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
                                                sbHTML.AppendLine($@"           <td><span>{ polSourceObsInfoEnumTextAndIDList.Where(c => c.ID == (int)obsInfoList[1]).FirstOrDefault().Text }</span></td>");
                                            }
                                            if (obsInfoList.Count > 2)
                                            {
                                                sbHTML.AppendLine($@"           <td><span>{ polSourceObsInfoEnumTextAndIDList.Where(c => c.ID == (int)obsInfoList[2]).FirstOrDefault().Text }</span></td>");
                                            }
                                            if (obsInfoList.Count > 3)
                                            {
                                                sbHTML.AppendLine($@"           <td><span>{ polSourceObsInfoEnumTextAndIDList.Where(c => c.ID == (int)obsInfoList[3]).FirstOrDefault().Text }</span></td>");
                                            }

                                            sbHTML.AppendLine($@"          <td>");
                                            foreach (int obsInfo in obsInfoList.Skip(4))
                                            {
                                                sbHTML.AppendLine($@"           <span>{ polSourceObsInfoEnumTextAndIDList.Where(c => c.ID == obsInfo).FirstOrDefault().Text }</span>");
                                            }
                                            sbHTML.AppendLine($@"           </td>");
                                            sbHTML.AppendLine($@"           <td><span>Photo</span></td>");

                                            //sbHTML.AppendLine($@"    <tr> ");
                                        }
                                        //sbHTML.AppendLine($@" </table>");
                                    }
                                }

                                sbHTML.AppendLine($@"    </tr>");
                            }
                        }
                    }
                }
            }
            sbHTML.AppendLine($@" </table>");


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
