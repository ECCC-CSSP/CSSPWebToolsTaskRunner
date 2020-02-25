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
        private bool GenerateHTMLSUBSECTOR_MUNICIPALITIES(StringBuilder sbTemp)
        {
            int Percent = 10;
            string NotUsed = "";

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", ReportGenerateObjectsKeywordEnum.SUBSECTOR_MWQM_SITES_FC_TABLE.ToString()));

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

            List<UseOfSiteModel> useOfSiteModelList = _UseOfSiteService.GetUseOfSiteModelListWithTVTypeAndSubsectorTVItemIDDB(TVTypeEnum.Municipality, tvItemModelSubsector.TVItemID);

            if (useOfSiteModelList.Count > 0)
            {
                foreach (UseOfSiteModel useOfSiteModel in useOfSiteModelList)
                {
                    TVItemModel tvItemModel = _TVItemService.GetTVItemModelWithTVItemIDDB(useOfSiteModel.SiteTVItemID);

                    if (string.IsNullOrWhiteSpace(tvItemModel.Error)) // municipality exist
                    {
                        sbTemp.AppendLine($@"<h3>{ tvItemModel.TVText }</h3>");

                        List<TVItemModelInfrastructureTypeTVItemLinkModel> tvItemModelInfrastructureTypeTVItemLinkModelList = _InfrastructureService.GetInfrastructureTVItemAndTVItemLinkAndInfrastructureTypeListWithMunicipalityTVItemIDDB(tvItemModel.TVItemID);

                        List<TVItemModelInfrastructureTypeTVItemLinkModel> tvItemModelInfrastructureTypeTVItemLinkModelListDone = new List<TVItemModelInfrastructureTypeTVItemLinkModel>();

                        foreach (TVItemModelInfrastructureTypeTVItemLinkModel tvItemModelInfrastructureTypeTVItemLinkModel in tvItemModelInfrastructureTypeTVItemLinkModelList.Where(c => c.InfrastructureType != InfrastructureTypeEnum.Other && c.InfrastructureType != InfrastructureTypeEnum.SeeOtherMunicipality && c.TVItemModelLinkList.Count == 0))
                        {
                            next(tvItemModelInfrastructureTypeTVItemLinkModel, tvItemModelInfrastructureTypeTVItemLinkModelList, tvItemModelInfrastructureTypeTVItemLinkModelListDone);
                        }

                        List<TVItemModelInfrastructureTypeTVItemLinkModel> tvItemModelInfrastructureTypeTVItemLinkModelListFloating = new List<TVItemModelInfrastructureTypeTVItemLinkModel>();

                        foreach (TVItemModelInfrastructureTypeTVItemLinkModel tvItemModelInfrastructureTypeTVItemLinkModel in tvItemModelInfrastructureTypeTVItemLinkModelList.Where(c => c.InfrastructureType != InfrastructureTypeEnum.Other && c.InfrastructureType != InfrastructureTypeEnum.SeeOtherMunicipality))
                        {
                            if (!tvItemModelInfrastructureTypeTVItemLinkModelListDone.Contains(tvItemModelInfrastructureTypeTVItemLinkModel))
                            {
                                tvItemModelInfrastructureTypeTVItemLinkModelListFloating.Add(tvItemModelInfrastructureTypeTVItemLinkModel);
                            }
                        }

                        if (tvItemModelInfrastructureTypeTVItemLinkModelList.Count == 0)
                        {
                            sbTemp.AppendLine($@"<h3>{ TaskRunnerServiceRes.NoInfrastructure }</span>");
                        }
                        else
                        {
                            sbTemp.AppendLine($@"<ul>");
                            sbTemp.AppendLine($@"    <li><label>{ TaskRunnerServiceRes.Legend }</label></li>");
                            sbTemp.AppendLine($@"    <li>W - { TaskRunnerServiceRes.WWTPs }</li>");
                            sbTemp.AppendLine($@"    <li>LS - { TaskRunnerServiceRes.LiftStations }</li>");
                            sbTemp.AppendLine($@"    <li>LO - { TaskRunnerServiceRes.LineOverflows }</li>");
                            sbTemp.AppendLine($@"</ul>");

                            foreach (TVItemModelInfrastructureTypeTVItemLinkModel tvItemModelInfrastructureTypeTVItemLinkModel in tvItemModelInfrastructureTypeTVItemLinkModelList.Where(c => c.InfrastructureType != InfrastructureTypeEnum.Other && c.InfrastructureType != InfrastructureTypeEnum.SeeOtherMunicipality && c.TVItemModelLinkList.Count == 0))
                            {
                                InfrastructureItem(sbTemp, tvItemModelInfrastructureTypeTVItemLinkModel, tvItemModelInfrastructureTypeTVItemLinkModelList, true);
                            }

                            foreach (TVItemModelInfrastructureTypeTVItemLinkModel tvItemModelInfrastructureTypeTVItemLinkModel in tvItemModelInfrastructureTypeTVItemLinkModelListFloating)
                            {
                                InfrastructureItem(sbTemp, tvItemModelInfrastructureTypeTVItemLinkModel, tvItemModelInfrastructureTypeTVItemLinkModelListFloating, false);
                            }

                            if (tvItemModelInfrastructureTypeTVItemLinkModelList.Where(c => c.InfrastructureType == InfrastructureTypeEnum.SeeOtherMunicipality).ToList().Count > 0)
                            {
                                sbTemp.AppendLine($@"<div>");
                                foreach (TVItemModelInfrastructureTypeTVItemLinkModel tvItemModelInfrastructureTypeTVItemLinkModel in tvItemModelInfrastructureTypeTVItemLinkModelList.Where(c => c.InfrastructureType == InfrastructureTypeEnum.SeeOtherMunicipality))
                                {
                                    if (tvItemModelInfrastructureTypeTVItemLinkModel.SeeOtherMunicipalityTVItemID != null)
                                    {
                                        InfrastructureItem(sbTemp, tvItemModelInfrastructureTypeTVItemLinkModel, tvItemModelInfrastructureTypeTVItemLinkModelList, true);
                                    }
                                }
                                sbTemp.AppendLine($@"</div>");
                            }
                        }
                    }
                }
            }
            else
            {
                sbTemp.AppendLine($@"<p style=""font-size: 2em;"">{ TaskRunnerServiceRes.NoMunicipalityWithinSubsector }</p>");
            }

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 80);

            return true;
        }

        // for testing only can comment out when test is completed
        public bool PublicGenerateHTMLSUBSECTOR_MUNICIPALITIES(StringBuilder sbTemp)
        {
            bool retBool = GenerateHTMLSUBSECTOR_MUNICIPALITIES(sbTemp);

            StreamWriter sw = fi.CreateText();
            sw.Write(sbTemp.ToString());
            sw.Flush();
            sw.Close();

            return retBool;
        }

        private void next(TVItemModelInfrastructureTypeTVItemLinkModel tvItemModelInfrastructureTypeTVItemLinkModel, List<TVItemModelInfrastructureTypeTVItemLinkModel> tvItemModelInfrastructureTypeTVItemLinkModelList, List<TVItemModelInfrastructureTypeTVItemLinkModel> tvItemModelInfrastructureTypeTVItemLinkModelListDone)
        {
            if (!tvItemModelInfrastructureTypeTVItemLinkModelListDone.Contains(tvItemModelInfrastructureTypeTVItemLinkModel))
            {
                tvItemModelInfrastructureTypeTVItemLinkModelListDone.Add(tvItemModelInfrastructureTypeTVItemLinkModel);
            }
            foreach (TVItemModelInfrastructureTypeTVItemLinkModel tvItemModelInfrastructureTypeTVItemLinkModelNext in tvItemModelInfrastructureTypeTVItemLinkModelList.Where(c => c.FlowTo == tvItemModelInfrastructureTypeTVItemLinkModel).ToList())
            {
                next(tvItemModelInfrastructureTypeTVItemLinkModelNext, tvItemModelInfrastructureTypeTVItemLinkModelList, tvItemModelInfrastructureTypeTVItemLinkModelListDone);
            }
        }
        private void InfrastructureItem(StringBuilder sbTemp, TVItemModelInfrastructureTypeTVItemLinkModel tvItemModelInfrastructureTypeTVItemLinkModel, List<TVItemModelInfrastructureTypeTVItemLinkModel> tvItemModelInfrastructureTypeTVItemLinkModelList, bool DoNext)
        {

            string BorderColor = "";
            bool HasOtherThanOutfall = false;

            if (tvItemModelInfrastructureTypeTVItemLinkModel.TVItemModelLinkList.Count > 0)
            {
                foreach (TVItemLinkModel tvItemLinkModel in tvItemModelInfrastructureTypeTVItemLinkModel.TVItemModelLinkList)
                {
                    if (tvItemModelInfrastructureTypeTVItemLinkModel.InfrastructureType != InfrastructureTypeEnum.WWTP)
                    {
                        HasOtherThanOutfall = true;
                    }
                }
            }
            switch (tvItemModelInfrastructureTypeTVItemLinkModel.InfrastructureType)
            {
                case InfrastructureTypeEnum.LiftStation:
                    BorderColor = "BorderLiftStation";
                    break;
                case InfrastructureTypeEnum.Other:
                    BorderColor = "BorderOtherInfrastructure";
                    break;
                case InfrastructureTypeEnum.LineOverflow:
                    BorderColor = "BorderLineOverflow";
                    break;
                case InfrastructureTypeEnum.WWTP:
                    BorderColor = "BorderWWTP";
                    break;
                default:
                    break;
            }

            sbTemp.AppendLine($@"<ul>");
            sbTemp.AppendLine($@"<li>{tvItemModelInfrastructureTypeTVItemLinkModel.TVItemModel.TVText}</li>");

            if (DoNext)
            {
                foreach (TVItemModelInfrastructureTypeTVItemLinkModel tvItemModelInfrastructureTypeTVItemLinkModelNext in tvItemModelInfrastructureTypeTVItemLinkModelList.Where(c => c.FlowTo == tvItemModelInfrastructureTypeTVItemLinkModel).ToList())
                {
                    InfrastructureItem(sbTemp, tvItemModelInfrastructureTypeTVItemLinkModelNext, tvItemModelInfrastructureTypeTVItemLinkModelList, true);
                }
            }
            sbTemp.AppendLine($@"</ul>");
        }
    }
}
