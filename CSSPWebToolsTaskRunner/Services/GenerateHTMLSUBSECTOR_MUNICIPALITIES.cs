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
using DocumentFormat.OpenXml.Wordprocessing;
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
                sbTemp.AppendLine($@"<table>");
                sbTemp.AppendLine($@"<tr>");
                sbTemp.AppendLine($@"<td colspan=""20"">");
                sbTemp.AppendLine($@"<b>{ TaskRunnerServiceRes.Legend }</b>");
                sbTemp.AppendLine($@"&nbsp;&nbsp;<b>W</b> - { TaskRunnerServiceRes.WWTPs }");
                sbTemp.AppendLine($@"&nbsp;&nbsp;<b>LS</b> - { TaskRunnerServiceRes.LiftStations }");
                sbTemp.AppendLine($@"&nbsp;&nbsp;<b>LO</b> - { TaskRunnerServiceRes.LineOverflows }");
                sbTemp.AppendLine($@"</td>");
                sbTemp.AppendLine($@"</tr>");
                sbTemp.AppendLine($@"<tr>");
                sbTemp.AppendLine($@"<td colspan=""20"">&nbsp;</td>");
                sbTemp.AppendLine($@"</tr>");

                foreach (UseOfSiteModel useOfSiteModel in useOfSiteModelList)
                {
                    TVItemModel tvItemModelMuni = _TVItemService.GetTVItemModelWithTVItemIDDB(useOfSiteModel.SiteTVItemID);

                    if (string.IsNullOrWhiteSpace(tvItemModelMuni.Error)) // municipality exist
                    {
                        sbTemp.AppendLine($@"<tr>");
                        sbTemp.AppendLine($@"<td colspan=""20"">");
                        sbTemp.AppendLine($@"<h3>{ tvItemModelMuni.TVText }</h3>");
                        sbTemp.AppendLine($@"</td>");
                        sbTemp.AppendLine($@"</tr>");
                        sbTemp.AppendLine($@"<tr>");
                        sbTemp.AppendLine($@"<td colspan=""20"">&nbsp;</td>");
                        sbTemp.AppendLine($@"</tr>");

                        List<TVItemLinkModel> tvItemLinkModelList = _TVItemLinkService.GetTVItemLinkModelListWithFromTVItemIDDB(tvItemModelMuni.TVItemID);

                        foreach (TVItemLinkModel tvItemLinkModel in tvItemLinkModelList)
                        {
                            if (tvItemLinkModel.ToTVType == TVTypeEnum.Contact)
                            {
                                ContactItem(sbTemp, tvItemLinkModel.ToTVItemID);
                            }
                        }

                        List<TVItemModelInfrastructureTypeTVItemLinkModel> tvItemModelInfrastructureTypeTVItemLinkModelList = _InfrastructureService.GetInfrastructureTVItemAndTVItemLinkAndInfrastructureTypeListWithMunicipalityTVItemIDDB(tvItemModelMuni.TVItemID);

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
                            sbTemp.AppendLine($@"<tr>");
                            sbTemp.AppendLine($@"<td colspan=""20"">");
                            sbTemp.AppendLine($@"<span style=""font-size: 2em;"">{ TaskRunnerServiceRes.NoInfrastructure }</span>");
                            sbTemp.AppendLine($@"</td>");
                            sbTemp.AppendLine($@"</tr>");
                            sbTemp.AppendLine($@"<tr>");
                            sbTemp.AppendLine($@"<td colspan=""20"">&nbsp;</td>");
                            sbTemp.AppendLine($@"</tr>");
                        }
                        else
                        {

                            foreach (TVItemModelInfrastructureTypeTVItemLinkModel tvItemModelInfrastructureTypeTVItemLinkModel in tvItemModelInfrastructureTypeTVItemLinkModelList.Where(c => c.InfrastructureType != InfrastructureTypeEnum.Other && c.InfrastructureType != InfrastructureTypeEnum.SeeOtherMunicipality && c.TVItemModelLinkList.Count == 0))
                            {
                                InfrastructureItem(sbTemp, tvItemModelInfrastructureTypeTVItemLinkModel, tvItemModelInfrastructureTypeTVItemLinkModelList, true, 0);
                            }

                            foreach (TVItemModelInfrastructureTypeTVItemLinkModel tvItemModelInfrastructureTypeTVItemLinkModel in tvItemModelInfrastructureTypeTVItemLinkModelListFloating)
                            {
                                InfrastructureItem(sbTemp, tvItemModelInfrastructureTypeTVItemLinkModel, tvItemModelInfrastructureTypeTVItemLinkModelListFloating, false, 0);
                            }

                            if (tvItemModelInfrastructureTypeTVItemLinkModelList.Where(c => c.InfrastructureType == InfrastructureTypeEnum.SeeOtherMunicipality).ToList().Count > 0)
                            {

                                foreach (TVItemModelInfrastructureTypeTVItemLinkModel tvItemModelInfrastructureTypeTVItemLinkModel in tvItemModelInfrastructureTypeTVItemLinkModelList.Where(c => c.InfrastructureType == InfrastructureTypeEnum.SeeOtherMunicipality))
                                {
                                    if (tvItemModelInfrastructureTypeTVItemLinkModel.SeeOtherMunicipalityTVItemID != null)
                                    {
                                        TVItemModel tvItemModelSeeOther = _TVItemService.GetTVItemModelWithTVItemIDDB((int)tvItemModelInfrastructureTypeTVItemLinkModel.SeeOtherMunicipalityTVItemID);
                                        if (string.IsNullOrWhiteSpace(tvItemModelSeeOther.Error))
                                        {
                                            sbTemp.AppendLine($@"<tr>");
                                            sbTemp.AppendLine($@"<td colspan=""20"">");
                                            sbTemp.AppendLine($@"<span style=""font-size: 2em;"">{ TaskRunnerServiceRes.NoInfrastructure }</span>");
                                            sbTemp.AppendLine($@"</td>");
                                            sbTemp.AppendLine($@"</tr>");
                                            sbTemp.AppendLine($@"<tr>");
                                            sbTemp.AppendLine($@"<td colspan=""20"">&nbsp;</td>");
                                            sbTemp.AppendLine($@"</tr>");
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                sbTemp.AppendLine($@"</table>");
            }
            else
            {
                sbTemp.AppendLine($@"<span style=""font-size: 2em;"">{ TaskRunnerServiceRes.NoMunicipalityWithinSubsector }</span>");
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
        private void ContactItem(StringBuilder sbTemp, int ContactTVItemID)
        {
            ContactModel contactModel = _ContactService.GetContactModelWithContactTVItemIDDB(ContactTVItemID);
            if (string.IsNullOrWhiteSpace(contactModel.Error))
            {
                string initial = string.IsNullOrWhiteSpace(contactModel.Initial) ? "" : $"{ contactModel.Initial}, ";
                string contactText = $"{contactModel.FirstName} {initial} {contactModel.LastName}";
                sbTemp.AppendLine($@"<tr>");
                sbTemp.AppendLine($@"<td>&nbsp;&nbsp;</td>");
                sbTemp.AppendLine($@"<td colspan=""19"">");
                sbTemp.AppendLine($@"<h4>{ contactText }</h4>");
                sbTemp.AppendLine($@"</td>");
                sbTemp.AppendLine($@"</tr>");

                List<TVItemLinkModel> tvItemLinkModelList = _TVItemLinkService.GetTVItemLinkModelListWithFromTVItemIDDB(ContactTVItemID);
                if (tvItemLinkModelList.Where(c => c.ToTVType == TVTypeEnum.Tel).Any())
                {
                    sbTemp.AppendLine($@"<tr>");
                    sbTemp.AppendLine($@"<td>&nbsp;&nbsp;</td><td>&nbsp;&nbsp;</td>");
                    sbTemp.AppendLine($@"<td colspan=""18"">");
                    foreach (TVItemLinkModel tvItemLinkModel in tvItemLinkModelList)
                    {
                        if (tvItemLinkModel.ToTVType == TVTypeEnum.Tel)
                        {
                            TelModel telModel = _TelService.GetTelModelWithTelTVItemIDDB(tvItemLinkModel.ToTVItemID);

                            sbTemp.AppendLine($@"<span style=""white-space: nowrap;"">{telModel.TelNumber} ({_BaseEnumService.GetEnumText_TelTypeEnum(telModel.TelType)})</span>");
                        }
                    }
                    sbTemp.AppendLine($@"</td>");
                    sbTemp.AppendLine($@"</tr>");
                }

                if (tvItemLinkModelList.Where(c => c.ToTVType == TVTypeEnum.Email).Any())
                {
                    sbTemp.AppendLine($@"<tr>");
                    sbTemp.AppendLine($@"<td>&nbsp;&nbsp;</td><td>&nbsp;&nbsp;</td>");
                    sbTemp.AppendLine($@"<td colspan=""18"">");
                    foreach (TVItemLinkModel tvItemLinkModel in tvItemLinkModelList)
                    {
                        if (tvItemLinkModel.ToTVType == TVTypeEnum.Email)
                        {
                            EmailModel emailModel = _EmailService.GetEmailModelWithEmailTVItemIDDB(tvItemLinkModel.ToTVItemID);

                            sbTemp.AppendLine($@"<span style=""white-space: nowrap;"">{emailModel.EmailAddress} ({_BaseEnumService.GetEnumText_EmailTypeEnum(emailModel.EmailType)})</span>");
                        }
                    }
                    sbTemp.AppendLine($@"</td>");
                    sbTemp.AppendLine($@"</tr>");
                }

                if (tvItemLinkModelList.Where(c => c.ToTVType == TVTypeEnum.Address).Any())
                {
                    sbTemp.AppendLine($@"<tr>");
                    sbTemp.AppendLine($@"<td>&nbsp;&nbsp;</td><td>&nbsp;&nbsp;</td>");
                    sbTemp.AppendLine($@"<td colspan=""18"">");
                    foreach (TVItemLinkModel tvItemLinkModel in tvItemLinkModelList)
                    {
                        if (tvItemLinkModel.ToTVType == TVTypeEnum.Address)
                        {
                            AddressModel addressModel = _AddressService.GetAddressModelWithAddressTVItemIDDB(tvItemLinkModel.ToTVItemID);

                            if (string.IsNullOrWhiteSpace(addressModel.Error))
                            {
                                string CivicAddress = "";
                                if (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr)
                                {
                                    CivicAddress = $"{addressModel.StreetNumber} {addressModel.StreetName} {_BaseEnumService.GetEnumText_StreetTypeEnum(addressModel.StreetType)} {addressModel.MunicipalityTVText} {addressModel.ProvinceTVText} {addressModel.CountryTVText} {addressModel.PostalCode}";
                                }
                                else
                                {
                                    CivicAddress = $"{addressModel.StreetNumber} {addressModel.StreetName} {_BaseEnumService.GetEnumText_StreetTypeEnum(addressModel.StreetType)} {addressModel.MunicipalityTVText} {addressModel.ProvinceTVText} {addressModel.CountryTVText} {addressModel.PostalCode}";
                                }
                                _AddressService.CreateTVText(addressModel);
                                sbTemp.AppendLine($@"<b>{ TaskRunnerServiceRes.CivicAddress }:</b>&nbsp;&nbsp;{CivicAddress}");
                            }
                        }
                    }
                    sbTemp.AppendLine($@"</td>");
                    sbTemp.AppendLine($@"</tr>");
                }

                sbTemp.AppendLine($@"<td colspan=""20"">&nbsp;</td>");
                sbTemp.AppendLine($@"</tr>");
            }
        }
        private void InfrastructureItem(StringBuilder sbTemp, TVItemModelInfrastructureTypeTVItemLinkModel tvItemModelInfrastructureTypeTVItemLinkModel, List<TVItemModelInfrastructureTypeTVItemLinkModel> tvItemModelInfrastructureTypeTVItemLinkModelList, bool DoNext, int Level)
        {

            string InfType = "";
            TVTypeEnum tvType = TVTypeEnum.Error;
            switch (tvItemModelInfrastructureTypeTVItemLinkModel.InfrastructureType)
            {
                case InfrastructureTypeEnum.LiftStation:
                    {
                        InfType = "(LS)";
                        tvType = TVTypeEnum.LiftStation;
                    }
                    break;
                case InfrastructureTypeEnum.LineOverflow:
                    {
                        InfType = "(LO)";
                        tvType = TVTypeEnum.LineOverflow;
                    }
                    break;
                case InfrastructureTypeEnum.WWTP:
                    {
                        InfType = "(W)";
                        tvType = TVTypeEnum.WasteWaterTreatmentPlant;
                    }
                    break;
                default:
                    break;
            }

            InfrastructureModel infrastructureModel = _InfrastructureService.GetInfrastructureModelWithInfrastructureTVItemIDDB(tvItemModelInfrastructureTypeTVItemLinkModel.TVItemModel.TVItemID);
            if (string.IsNullOrWhiteSpace(infrastructureModel.Error))
            {
                List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelInfrastructureTypeTVItemLinkModel.TVItemModel.TVItemID, tvType, MapInfoDrawTypeEnum.Point);
                List<MapInfoPointModel> mapInfoPointModelListOut = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelInfrastructureTypeTVItemLinkModel.TVItemModel.TVItemID, TVTypeEnum.Outfall, MapInfoDrawTypeEnum.Point);

                if (mapInfoPointModelList.Count > 0 && mapInfoPointModelListOut.Count > 0)
                {
                    sbTemp.AppendLine($@"<tr>");
                    for (int i = 0; i < Level + 1; i++)
                    {
                        sbTemp.AppendLine($@"<td>&nbsp;&nbsp;</td>");
                    }
                    sbTemp.AppendLine($@"<td colspan=""{19 - Level}"">");
                    sbTemp.AppendLine($@"<h4>{InfType} - <b>{tvItemModelInfrastructureTypeTVItemLinkModel.TVItemModel.TVText}&nbsp;{ TaskRunnerServiceRes.Lat }&nbsp;{ TaskRunnerServiceRes.Long }&nbsp;{mapInfoPointModelList[0].Lat}&nbsp;{mapInfoPointModelList[0].Lng}</b></h4>");
                    sbTemp.AppendLine($@"&nbsp;&nbsp;&nbsp;&nbsp;<b>{ TaskRunnerServiceRes.Outfall } { TaskRunnerServiceRes.Lat } { TaskRunnerServiceRes.Long }:</b>&nbsp;&nbsp;{mapInfoPointModelListOut[0].Lat} {mapInfoPointModelListOut[0].Lng}");
                    sbTemp.AppendLine($@"</tr>");
                    sbTemp.AppendLine($@"<td colspan=""20"">&nbsp;</td>");
                    sbTemp.AppendLine($@"</tr>");
                }
                else
                {
                    sbTemp.AppendLine($@"<tr>");
                    for (int i = 0; i < Level + 1; i++)
                    {
                        sbTemp.AppendLine($@"<td>&nbsp;&nbsp;</td>");
                    }
                    sbTemp.AppendLine($@"<td colspan=""{19 - Level}"">");
                    sbTemp.AppendLine($@"<h4>{InfType} - <b>{tvItemModelInfrastructureTypeTVItemLinkModel.TVItemModel.TVText}</b></h4>");
                    sbTemp.AppendLine($@"</td>");
                    sbTemp.AppendLine($@"</tr>");
                    sbTemp.AppendLine($@"<td colspan=""20"">&nbsp;</td>");
                    sbTemp.AppendLine($@"</tr>");
                }
                if (infrastructureModel.CivicAddressTVItemID != null)
                {
                    AddressModel addressModel = _AddressService.GetAddressModelWithAddressTVItemIDDB((int)infrastructureModel.CivicAddressTVItemID);
                    if (string.IsNullOrWhiteSpace(addressModel.Error))
                    {
                        string CivicAddress = "";
                        if (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr)
                        {
                            CivicAddress = $"{addressModel.StreetNumber} {addressModel.StreetName} {_BaseEnumService.GetEnumText_StreetTypeEnum(addressModel.StreetType)} {addressModel.MunicipalityTVText} {addressModel.ProvinceTVText} {addressModel.CountryTVText} {addressModel.PostalCode}";
                        }
                        else
                        {
                            CivicAddress = $"{addressModel.StreetNumber} {addressModel.StreetName} {_BaseEnumService.GetEnumText_StreetTypeEnum(addressModel.StreetType)} {addressModel.MunicipalityTVText} {addressModel.ProvinceTVText} {addressModel.CountryTVText} {addressModel.PostalCode}";
                        }
                        _AddressService.CreateTVText(addressModel);

                        sbTemp.AppendLine($@"<tr>");
                        for (int i = 0; i < Level + 2; i++)
                        {
                            sbTemp.AppendLine($@"<td>&nbsp;&nbsp;</td>");
                        }
                        sbTemp.AppendLine($@"<td colspan=""{18 - Level}"">");
                        sbTemp.AppendLine($@"<b>{ TaskRunnerServiceRes.CivicAddress }:</b>&nbsp;&nbsp;{CivicAddress}<br />");
                        sbTemp.AppendLine($@"</td>");
                        sbTemp.AppendLine($@"</tr>");
                    }
                }

                sbTemp.AppendLine($@"<tr>");
                for (int i = 0; i < Level + 2; i++)
                {
                    sbTemp.AppendLine($@"<td>&nbsp;&nbsp;</td>");
                }
                sbTemp.AppendLine($@"<td style=""border-left: solid black 1px"" colspan=""{18 - Level}"">");

                if (tvType == TVTypeEnum.WasteWaterTreatmentPlant && infrastructureModel.FacilityType == FacilityTypeEnum.Lagoon)
                {
                    sbTemp.AppendLine($@"<b>{ TaskRunnerServiceRes.FacilityType }:</b>&nbsp;&nbsp;{_BaseEnumService.GetEnumText_FacilityTypeEnum(infrastructureModel.FacilityType)}");
                    if (infrastructureModel.IsMechanicallyAerated != null)
                    {
                        sbTemp.AppendLine($@"&nbsp;&nbsp;&nbsp;<b>{ TaskRunnerServiceRes.IsMechanicallyAerated }:</b>&nbsp;&nbsp;{infrastructureModel.IsMechanicallyAerated}");
                        if ((bool)infrastructureModel.IsMechanicallyAerated)
                        {
                            sbTemp.AppendLine($@"&nbsp;&nbsp;&nbsp;<b>{ TaskRunnerServiceRes.AerationType }:</b>&nbsp;&nbsp;{_BaseEnumService.GetEnumText_AerationTypeEnum(infrastructureModel.AerationType)}");
                        }
                    }
                    sbTemp.AppendLine($@"&nbsp;&nbsp;&nbsp;<b>{ TaskRunnerServiceRes.NumberOfCells }:</b>&nbsp;&nbsp;{infrastructureModel.NumberOfCells}");
                    sbTemp.AppendLine($@"&nbsp;&nbsp;&nbsp;<b>{ TaskRunnerServiceRes.NumberOfAeratedCells }:</b>&nbsp;&nbsp;{infrastructureModel.NumberOfAeratedCells}");
                }
                else if (tvType == TVTypeEnum.WasteWaterTreatmentPlant && infrastructureModel.FacilityType == FacilityTypeEnum.Plant)
                {
                    sbTemp.AppendLine($@"<b>{ TaskRunnerServiceRes.FacilityType }:</b>&nbsp;&nbsp;{_BaseEnumService.GetEnumText_FacilityTypeEnum(infrastructureModel.FacilityType)}");
                    if (infrastructureModel.PreliminaryTreatmentType != null)
                    {
                        sbTemp.AppendLine($@"&nbsp;&nbsp;&nbsp;<b>{ TaskRunnerServiceRes.PreliminaryTreatmentType }:</b>&nbsp;&nbsp;{_BaseEnumService.GetEnumText_PreliminaryTreatmentTypeEnum(infrastructureModel.PreliminaryTreatmentType)}");
                    }
                    if (infrastructureModel.PrimaryTreatmentType != null)
                    {
                        sbTemp.AppendLine($@"&nbsp;&nbsp;&nbsp;<b>{ TaskRunnerServiceRes.PrimaryTreatmentType }:</b>&nbsp;&nbsp;{_BaseEnumService.GetEnumText_PrimaryTreatmentTypeEnum(infrastructureModel.PrimaryTreatmentType)}");
                    }
                    if (infrastructureModel.SecondaryTreatmentType != null)
                    {
                        sbTemp.AppendLine($@"&nbsp;&nbsp;&nbsp;<b>{ TaskRunnerServiceRes.SecondaryTreatmentType }:</b>&nbsp;&nbsp;{_BaseEnumService.GetEnumText_SecondaryTreatmentTypeEnum(infrastructureModel.SecondaryTreatmentType)}");
                    }
                    if (infrastructureModel.TertiaryTreatmentType != null)
                    {
                        sbTemp.AppendLine($@"&nbsp;&nbsp;&nbsp;<b>{ TaskRunnerServiceRes.TertiaryTreatmentType }:</b>&nbsp;&nbsp;{_BaseEnumService.GetEnumText_TertiaryTreatmentTypeEnum(infrastructureModel.TertiaryTreatmentType)}");
                    }
                }
                if (tvType == TVTypeEnum.WasteWaterTreatmentPlant)
                {
                    if (infrastructureModel.DisinfectionType != null)
                    {
                        sbTemp.AppendLine($@"&nbsp;&nbsp;&nbsp;<b>{ TaskRunnerServiceRes.DisinfectionType }:</b>&nbsp;&nbsp;{_BaseEnumService.GetEnumText_DisinfectionTypeEnum(infrastructureModel.DisinfectionType)}");
                    }
                    if (infrastructureModel.CollectionSystemType != null)
                    {
                        sbTemp.AppendLine($@"&nbsp;&nbsp;&nbsp;<b>{ TaskRunnerServiceRes.CollectionSystemType }:</b>&nbsp;&nbsp;{_BaseEnumService.GetEnumText_CollectionSystemTypeEnum(infrastructureModel.CollectionSystemType)}");
                    }
                    sbTemp.AppendLine($@"&nbsp;&nbsp;&nbsp;<b>{ TaskRunnerServiceRes.DesignFlow }:</b>&nbsp;&nbsp;{infrastructureModel.DesignFlow_m3_day} (m<sup>3</sup>/{TaskRunnerServiceRes.day})");
                    sbTemp.AppendLine($@"&nbsp;&nbsp;&nbsp;<b>{ TaskRunnerServiceRes.AverageFlow }:</b>&nbsp;&nbsp;{infrastructureModel.AverageFlow_m3_day} (m<sup>3</sup>/{TaskRunnerServiceRes.day})");
                    sbTemp.AppendLine($@"&nbsp;&nbsp;&nbsp;<b>{ TaskRunnerServiceRes.PeakFlow }:</b>&nbsp;&nbsp;{infrastructureModel.PeakFlow_m3_day} (m<sup>3</sup>/{TaskRunnerServiceRes.day})");
                    sbTemp.AppendLine($@"&nbsp;&nbsp;&nbsp;<b>{ TaskRunnerServiceRes.PopulationServed }:</b>&nbsp;&nbsp;{infrastructureModel.PopServed}");
                }
                sbTemp.AppendLine($@"&nbsp;&nbsp;&nbsp;<b>{ TaskRunnerServiceRes.PercentOfFlow }:</b>&nbsp;&nbsp;{infrastructureModel.PercFlowOfTotal} %");
                if (infrastructureModel.AlarmSystemType != null)
                {
                    sbTemp.AppendLine($@"&nbsp;&nbsp;&nbsp;<b>{ TaskRunnerServiceRes.AlarmSystemType }:</b>&nbsp;&nbsp;{_BaseEnumService.GetEnumText_AlarmSystemTypeEnum(infrastructureModel.AlarmSystemType)}");
                }
                if (infrastructureModel.CanOverflow != null)
                {
                    sbTemp.AppendLine($@"&nbsp;&nbsp;&nbsp;<b>{ TaskRunnerServiceRes.CanOverflow }:</b>&nbsp;&nbsp;{infrastructureModel.CanOverflow}");
                    if (infrastructureModel.ValveType != null)
                    {
                        sbTemp.AppendLine($@"&nbsp;&nbsp;&nbsp;<b>{ TaskRunnerServiceRes.ValveType }:</b>&nbsp;&nbsp;{_BaseEnumService.GetEnumText_ValveTypeEnum(infrastructureModel.ValveType)}");
                    }
                }
                if (infrastructureModel.HasBackupPower != null)
                {
                    sbTemp.AppendLine($@"&nbsp;&nbsp;&nbsp;<b>{ TaskRunnerServiceRes.HasBackupPower }:</b>&nbsp;&nbsp;{infrastructureModel.HasBackupPower}");
                }
                if (!string.IsNullOrWhiteSpace(infrastructureModel.Comment))
                {
                    sbTemp.AppendLine($@"<br /><h4>{ TaskRunnerServiceRes.Comments }</h4>{infrastructureModel.Comment}<br />");
                }

                sbTemp.AppendLine($@"<h4>Outfall</h4>");
                sbTemp.AppendLine($@"<b>{ TaskRunnerServiceRes.AverageDepth }:</b>&nbsp;&nbsp;{infrastructureModel.AverageDepth_m} (m)");
                sbTemp.AppendLine($@"<b>{ TaskRunnerServiceRes.DecayRate }:</b>&nbsp;&nbsp;{infrastructureModel.DecayRate_per_day} (/{ TaskRunnerServiceRes.day})");
                sbTemp.AppendLine($@"<b>{ TaskRunnerServiceRes.DistanceFromShore }:</b>&nbsp;&nbsp;{infrastructureModel.DistanceFromShore_m} (m)");
                sbTemp.AppendLine($@"<b>{ TaskRunnerServiceRes.FarFieldVelocity }:</b>&nbsp;&nbsp;{infrastructureModel.FarFieldVelocity_m_s} (m/s)");
                sbTemp.AppendLine($@"<b>{ TaskRunnerServiceRes.HorizontalAngle }:</b>&nbsp;&nbsp;{infrastructureModel.HorizontalAngle_deg} (deg)");
                sbTemp.AppendLine($@"<b>{ TaskRunnerServiceRes.NearFieldVelocity }:</b>&nbsp;&nbsp;{infrastructureModel.NearFieldVelocity_m_s} (m/s)");
                sbTemp.AppendLine($@"<b>{ TaskRunnerServiceRes.NumberOfPorts }:</b>&nbsp;&nbsp;{infrastructureModel.NumberOfPorts}");
                sbTemp.AppendLine($@"<b>{ TaskRunnerServiceRes.PortDiameter }:</b>&nbsp;&nbsp;{infrastructureModel.PortDiameter_m} (m)");
                sbTemp.AppendLine($@"<b>{ TaskRunnerServiceRes.PortElevation }:</b>&nbsp;&nbsp;{infrastructureModel.PortElevation_m} (m)");
                sbTemp.AppendLine($@"<b>{ TaskRunnerServiceRes.PortSpacing }:</b>&nbsp;&nbsp;{infrastructureModel.PortSpacing_m} (m)");
                sbTemp.AppendLine($@"<b>{ TaskRunnerServiceRes.ReceivingWater }:</b>&nbsp;&nbsp;{infrastructureModel.ReceivingWater_MPN_per_100ml} (MPN/100 mL)");
                sbTemp.AppendLine($@"<b>{ TaskRunnerServiceRes.ReceivingWaterSalinity }:</b>&nbsp;&nbsp;{infrastructureModel.ReceivingWaterSalinity_PSU} (PSU)");
                sbTemp.AppendLine($@"<b>{ TaskRunnerServiceRes.ReceivingWaterTemperature }:</b>&nbsp;&nbsp;{infrastructureModel.ReceivingWaterTemperature_C} (ºC)");
                sbTemp.AppendLine($@"<b>{ TaskRunnerServiceRes.VerticalAngle }:</b>&nbsp;&nbsp;{infrastructureModel.VerticalAngle_deg} (deg)");

                sbTemp.AppendLine($@"</td>");
                sbTemp.AppendLine($@"</tr>");
                sbTemp.AppendLine($@"<td colspan=""20"">&nbsp;</td>");
                sbTemp.AppendLine($@"</tr>");

            }

            if (DoNext)
            {
                foreach (TVItemModelInfrastructureTypeTVItemLinkModel tvItemModelInfrastructureTypeTVItemLinkModelNext in tvItemModelInfrastructureTypeTVItemLinkModelList.Where(c => c.FlowTo == tvItemModelInfrastructureTypeTVItemLinkModel).ToList())
                {
                    InfrastructureItem(sbTemp, tvItemModelInfrastructureTypeTVItemLinkModelNext, tvItemModelInfrastructureTypeTVItemLinkModelList, true, Level + 1);
                }
            }
        }
    }
}
