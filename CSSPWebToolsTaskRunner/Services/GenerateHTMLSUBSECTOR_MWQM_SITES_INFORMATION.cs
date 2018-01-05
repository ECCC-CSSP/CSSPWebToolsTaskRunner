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
        private bool GenerateHTMLSUBSECTOR_MWQM_SITES_INFORMATION(StringBuilder sbTemp)
        {
            int Percent = 10;
            string NotUsed = "";

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", ReportGenerateObjectsKeywordEnum.SUBSECTOR_MWQM_SITES_INFORMATION.ToString()));

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


            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 40);

            sbTemp.AppendLine($@"|||TableCaption|: { TaskRunnerServiceRes.MWQMSitesInformation }|||");

            sbTemp.AppendLine($@" <table cellpadding=""5"" class=""textAlignCenter"">");
            sbTemp.AppendLine($@" <tr>");
            sbTemp.AppendLine($@" <th colspan=""2"">{ TaskRunnerServiceRes.Site }</th>");
            sbTemp.AppendLine($@" <th>{ TaskRunnerServiceRes.Coordinates }</th>");
            sbTemp.AppendLine($@" <th>{ TaskRunnerServiceRes.Description }</th>");
            sbTemp.AppendLine($@" <th>{ TaskRunnerServiceRes.Photos }</th>");
            sbTemp.AppendLine($@" </tr>");
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

                        sbTemp.AppendLine($@" <tr>");
                        sbTemp.AppendLine($@" <td>{ mwqmSiteModel.MWQMSiteTVText }</td>");
                        sbTemp.AppendLine($@" <td class=""{ classificationColor }"">{ classificationLetter }</td>");

                        List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(mwqmSiteModel.MWQMSiteTVItemID, TVTypeEnum.MWQMSite, MapInfoDrawTypeEnum.Point);
                        if (mapInfoPointModelList.Count > 0)
                        {
                            sbTemp.AppendLine($@" <td>{ mapInfoPointModelList[0].Lat.ToString("F5") } { mapInfoPointModelList[0].Lng.ToString("F5") }</td>");
                        }
                        else
                        {
                            sbTemp.AppendLine($@" <td>&nbsp;</td>");
                        }
                        sbTemp.AppendLine($@" <td class=""textAlignLeft"">{ mwqmSiteModel.MWQMSiteDescription }</td>");
                        sbTemp.AppendLine($@" <td>Photo</td>");
                        sbTemp.AppendLine($@" </tr>");
                    }
                }
            }
            sbTemp.AppendLine($@" </table>");

            Percent = 80;
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);

            return true;
        }

        // for testing only can comment out when test is completed
        public bool PublicGenerateHTMLSUBSECTOR_MWQM_SITES_INFORMATION(StringBuilder sbTemp)
        {
            bool retBool = GenerateHTMLSUBSECTOR_MWQM_SITES_INFORMATION(sbTemp);

            StreamWriter sw = fi.CreateText();
            sw.Write(sbTemp.ToString());
            sw.Flush();
            sw.Close();

            return retBool;
        }
    }
}
