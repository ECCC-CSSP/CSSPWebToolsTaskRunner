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
using CSSPModelsDLL.Models;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        private bool GenerateHTML()
        {
            switch (reportTypeModel.TVType)
            {
                case TVTypeEnum.Root:
                    {
                        if (!GenerateHTMLRoot())
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.Area:
                    {
                        if (!GenerateHTMLArea())
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.Country:
                    {
                        if (!GenerateHTMLCountry())
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.Infrastructure:
                    {
                        if (!GenerateHTMLInfrastructure())
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.MikeScenario:
                    {
                        if (!GenerateHTMLMikeScenario())
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.MikeSource:
                    {
                        if (!GenerateHTMLMikeSource())
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.Municipality:
                    {
                        if (!GenerateHTMLMunicipality())
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.MWQMSite:
                    {
                        if (!GenerateHTMLMWQMSite())
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.PolSourceSite:
                    {
                        if (!GenerateHTMLPolSourceSite())
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.Province:
                    {
                        if (!GenerateHTMLProvince())
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.Sector:
                    {
                        if (!GenerateHTMLSector())
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.Subsector:
                    {
                        if (!GenerateHTMLSubsector())
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.BoxModel:
                    {
                        if (!GenerateHTMLBoxModel())
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.VisualPlumesScenario:
                    {
                        if (!GenerateHTMLVisualPlumesScenario())
                        {
                            return false;
                        }
                    }
                    break;
                default:
                    {
                        if (!GenerateHTMLNotImplemented())
                        {
                            return false;
                        }
                    }
                    break;
            }

            return true;
        }
        private bool GenerateHTMLNotImplemented()
        {
            if (!GetTopHTML())
            {
                return false;
            }

            sb.AppendLine(@"<h2>" + reportTypeModel.TVType.ToString() + " Not Implemented Yet</h2>");

            if (!GetBottomHTML())
            {
                return false;
            }

            return true;
        }
        private bool GetBottomHTML()
        {
            //sb.AppendLine(@"    <br />");
            //sb.AppendLine(@"    <div class=""small"">");
            //if (fi.Extension.ToLower() == ".html")
            //{
            //    sb.AppendLine(@"        <b>File name:</b> " + fi.FullName.Replace(".html", ".docx") + @"");
            //}
            //else
            //{
            //    sb.AppendLine(@"        <b>File name:</b> " + fi.FullName + @"");
            //}
            //sb.AppendLine(@"    </div>");
            sb.AppendLine(@"</body>");
            sb.AppendLine(@"</html>");

            return true;
        }
        private bool GenerateObjects()
        {

            // need to search the whole document for some keywords Ex: |||SUBSECTOR_RE_EVALUATION_COVER_PAGE|||
            List<ReportGenerateObjectsKeywordEnum> KeywordList = new List<ReportGenerateObjectsKeywordEnum>();

            for(int i = 1, count = Enum.GetNames(typeof(ReportGenerateObjectsKeywordEnum)).Length; i < count; i++)
            {
                KeywordList.Add((ReportGenerateObjectsKeywordEnum)i);
            };

            foreach (ReportGenerateObjectsKeywordEnum keyword in KeywordList)
            {
                bool Found = true;
                int StartPos = 0;
                int EndPos = 0;
                while (Found)
                {
                    string sbText = sb.ToString();

                    StartPos = sbText.IndexOf("|||" + keyword.ToString() + "|||");
                    EndPos = StartPos + ("|||" + keyword + "|||").ToString().Length;

                    StringBuilder sbTemp = new StringBuilder();
                    if (StartPos > 0)
                    {
                        switch (keyword)
                        {
                            case ReportGenerateObjectsKeywordEnum.SUBSECTOR_RE_EVALUATION_COVER_PAGE:
                                {
                                    GenerateHTMLSUBSECTOR_RE_EVALUATION_COVER_PAGE(sbTemp);

                                    sb.Remove(StartPos, EndPos - StartPos);
                                    sb.Insert(StartPos, sbTemp);
                                }
                                break;
                            case ReportGenerateObjectsKeywordEnum.SUBSECTOR_FC_SUMMARY_STAT_ALL:
                                {
                                    GenerateHTMLSUBSECTOR_FC_SUMMARY_STAT_ALL(sbTemp);

                                    sb.Remove(StartPos, EndPos - StartPos);
                                    sb.Insert(StartPos, sbTemp);
                                }
                                break;
                            case ReportGenerateObjectsKeywordEnum.SUBSECTOR_FC_SUMMARY_STAT_WET:
                                {
                                    GenerateHTMLSUBSECTOR_FC_SUMMARY_STAT_WET(sbTemp);

                                    sb.Remove(StartPos, EndPos - StartPos);
                                    sb.Insert(StartPos, sbTemp);
                                }
                                break;
                            case ReportGenerateObjectsKeywordEnum.SUBSECTOR_FC_SUMMARY_STAT_DRY:
                                {
                                    GenerateHTMLSUBSECTOR_FC_SUMMARY_STAT_DRY(sbTemp);

                                    sb.Remove(StartPos, EndPos - StartPos);
                                    sb.Insert(StartPos, sbTemp);
                                }
                                break;
                            case ReportGenerateObjectsKeywordEnum.SUBSECTOR_MWQM_SITES_DATA_AVAILABILITY:
                                {
                                    GenerateHTMLSUBSECTOR_MWQM_SITES_DATA_AVAILABILITY(sbTemp);

                                    sb.Remove(StartPos, EndPos - StartPos);
                                    sb.Insert(StartPos, sbTemp);
                                }
                                break;
                            case ReportGenerateObjectsKeywordEnum.SUBSECTOR_MWQM_SITES_NUMBER_OF_SITES_BY_YEAR:
                                {
                                    GenerateHTMLSUBSECTOR_MWQM_SITES_NUMBER_OF_SITES_BY_YEAR(sbTemp);

                                    sb.Remove(StartPos, EndPos - StartPos);
                                    sb.Insert(StartPos, sbTemp);
                                }
                                break;
                            case ReportGenerateObjectsKeywordEnum.SUBSECTOR_MWQM_SITES_NUMBER_OF_RUNS_BY_YEAR:
                                {
                                    GenerateHTMLSUBSECTOR_MWQM_SITES_NUMBER_OF_RUNS_BY_YEAR(sbTemp);

                                    sb.Remove(StartPos, EndPos - StartPos);
                                    sb.Insert(StartPos, sbTemp);
                                }
                                break;
                            case ReportGenerateObjectsKeywordEnum.SUBSECTOR_MWQM_SITES_NUMBER_OF_SAMPLES_BY_YEAR:
                                {
                                    GenerateHTMLSUBSECTOR_MWQM_SITES_NUMBER_OF_SAMPLES_BY_YEAR(sbTemp);

                                    sb.Remove(StartPos, EndPos - StartPos);
                                    sb.Insert(StartPos, sbTemp);
                                }
                                break;
                            case ReportGenerateObjectsKeywordEnum.SUBSECTOR_MWQM_SITES_FC_TABLE:
                                {
                                    GenerateHTMLSUBSECTOR_MWQM_SITES_FC_TABLE(sbTemp);

                                    sb.Remove(StartPos, EndPos - StartPos);
                                    sb.Insert(StartPos, sbTemp);
                                }
                                break;
                            case ReportGenerateObjectsKeywordEnum.SUBSECTOR_MWQM_SITES_SALINITY_TABLE:
                                {
                                    GenerateHTMLSUBSECTOR_MWQM_SITES_SALINITY_TABLE(sbTemp);

                                    sb.Remove(StartPos, EndPos - StartPos);
                                    sb.Insert(StartPos, sbTemp);
                                }
                                break;
                            case ReportGenerateObjectsKeywordEnum.SUBSECTOR_MWQM_SITES_TEMPERATURE_TABLE:
                                {
                                    GenerateHTMLSUBSECTOR_MWQM_SITES_TEMPERATURE_TABLE(sbTemp);

                                    sb.Remove(StartPos, EndPos - StartPos);
                                    sb.Insert(StartPos, sbTemp);
                                }
                                break;
                            case ReportGenerateObjectsKeywordEnum.SUBSECTOR_POLLUTION_SOURCE_SITES:
                                {
                                    GenerateHTMLSUBSECTOR_POLLUTION_SOURCE_SITES(sbTemp);

                                    sb.Remove(StartPos, EndPos - StartPos);
                                    sb.Insert(StartPos, sbTemp);
                                }
                                break;
                            case ReportGenerateObjectsKeywordEnum.SUBSECTOR_POLLUTION_SOURCE_SITES_COMPACT:
                                {
                                    GenerateHTMLSUBSECTOR_POLLUTION_SOURCE_SITES_COMPACT(sbTemp);

                                    sb.Remove(StartPos, EndPos - StartPos);
                                    sb.Insert(StartPos, sbTemp);
                                }
                                break;
                            case ReportGenerateObjectsKeywordEnum.SUBSECTOR_MUNICIPALITIES:
                                {
                                    GenerateHTMLSUBSECTOR_MUNICIPALITIES(sbTemp);

                                    sb.Remove(StartPos, EndPos - StartPos);
                                    sb.Insert(StartPos, sbTemp);
                                }
                                break;
                            case ReportGenerateObjectsKeywordEnum.SUBSECTOR_MUNICIPALITY_INFRASTRUCTURE_MAP:
                                {
                                    GenerateHTMLSUBSECTOR_MUNICIPALITY_INFRASTRUCTURE_MAP(sbTemp);

                                    sb.Remove(StartPos, EndPos - StartPos);
                                    sb.Insert(StartPos, sbTemp);
                                }
                                break;
                            case ReportGenerateObjectsKeywordEnum.SUBSECTOR_MUNICIPALITY_CONTACTS:
                                {
                                    GenerateHTMLSUBSECTOR_MUNICIPALITY_CONTACTS(sbTemp);

                                    sb.Remove(StartPos, EndPos - StartPos);
                                    sb.Insert(StartPos, sbTemp);
                                }
                                break;
                            case ReportGenerateObjectsKeywordEnum.SUBSECTOR_MUNICIPALITY_INFRASTRUCTURE_DETAIL:
                                {
                                    GenerateHTMLSUBSECTOR_MUNICIPALITY_INFRASTRUCTURE_DETAIL(sbTemp);

                                    sb.Remove(StartPos, EndPos - StartPos);
                                    sb.Insert(StartPos, sbTemp);
                                }
                                break;
                            case ReportGenerateObjectsKeywordEnum.SUBSECTOR_ECCC_AND_SWCP_LOGO:
                                {
                                    GenerateHTMLSUBSECTOR_ECCC_AND_SWCP_LOGO(sbTemp);

                                    sb.Remove(StartPos, EndPos - StartPos);
                                    sb.Insert(StartPos, sbTemp);
                                }
                                break;
                            case ReportGenerateObjectsKeywordEnum.SUBSECTOR_CSSP_LOGO:
                                {
                                    GenerateHTMLSUBSECTOR_CSSP_LOGO(sbTemp);

                                    sb.Remove(StartPos, EndPos - StartPos);
                                    sb.Insert(StartPos, sbTemp);
                                }
                                break;
                            case ReportGenerateObjectsKeywordEnum.SUBSECTOR_LOCATION_OF_SURVEY_AREA_MAP:
                                {
                                    GenerateHTMLSUBSECTOR_LOCATION_OF_SURVEY_AREA_MAP(sbTemp);

                                    sb.Remove(StartPos, EndPos - StartPos);
                                    sb.Insert(StartPos, sbTemp);
                                }
                                break;
                            case ReportGenerateObjectsKeywordEnum.SUBSECTOR_CURRENT_CLASSIFICATION_AND_SAMPLING_LOCATION_MAP:
                                {
                                    GenerateHTMLSUBSECTOR_CURRENT_CLASSIFICATION_AND_SAMPLING_LOCATION_MAP(sbTemp);

                                    sb.Remove(StartPos, EndPos - StartPos);
                                    sb.Insert(StartPos, sbTemp);
                                }
                                break;
                            case ReportGenerateObjectsKeywordEnum.SUBSECTOR_RECOMMENDED_CLASSIFICATION_MAP:
                                {
                                    GenerateHTMLSUBSECTOR_RECOMMENDED_CLASSIFICATION_MAP(sbTemp);

                                    sb.Remove(StartPos, EndPos - StartPos);
                                    sb.Insert(StartPos, sbTemp);
                                }
                                break;
                            case ReportGenerateObjectsKeywordEnum.SUBSECTOR_POLLUTION_SOURCE_LOCATIONS_MAP:
                                {
                                    GenerateHTMLSUBSECTOR_POLLUTION_SOURCE_LOCATIONS_MAP(sbTemp);

                                    sb.Remove(StartPos, EndPos - StartPos);
                                    sb.Insert(StartPos, sbTemp);
                                }
                                break;
                            default:
                                {
                                    sb.Insert(EndPos, "Tag not recognised");
                                }
                                break;
                        }
                    }
                    else
                    {
                        Found = false;
                    }
                }
            }

            return true;
        }
        private bool GetTopHTML()
        {
            sb.AppendLine(@"<html>");
            sb.AppendLine(@"<head>");
            sb.AppendLine(@"    <meta charset=""utf-8"" />");
            sb.AppendLine(@"    <title>This is the title of the html document</title>");
            sb.AppendLine(@"    <style>");
            sb.AppendLine(@"body {");
            sb.AppendLine(@"    font: normal 14px arial, helvetica, sans-serif;");
            sb.AppendLine(@"}");
            sb.AppendLine(@"table.FCStatTableClass {");
            sb.AppendLine(@"    font: normal 10px arial, helvetica, sans-serif;");
            sb.AppendLine(@"    border: 1px solid black;");
            sb.AppendLine(@"    text-align: center;");
            sb.AppendLine(@"    width: 100%;");
            sb.AppendLine(@"    border-collapse: collapse;");
            sb.AppendLine(@"}");
            sb.AppendLine(@"table.DataAvailabilityTableClass {");
            sb.AppendLine(@"    font: normal 10px arial, helvetica, sans-serif;");
            sb.AppendLine(@"    border: 1px solid black;");
            sb.AppendLine(@"    text-align: center;");
            sb.AppendLine(@"    width: 100%;");
            sb.AppendLine(@"    border-collapse: collapse;");
            sb.AppendLine(@"}");
            sb.AppendLine(@"table.FCSalTempDataTableClass {");
            sb.AppendLine(@"    font: normal 10px arial, helvetica, sans-serif;");
            sb.AppendLine(@"    border: 1px solid black;");
            sb.AppendLine(@"    text-align: center;");
            sb.AppendLine(@"    width: 100%;");
            sb.AppendLine(@"    border-collapse: collapse;");
            sb.AppendLine(@"}");
            sb.AppendLine(@"table.PolSourceSiteCompact {");
            sb.AppendLine(@"    font: normal 10px arial, helvetica, sans-serif;");
            sb.AppendLine(@"    border: 1px solid black;");
            sb.AppendLine(@"    text-align: left;");
            sb.AppendLine(@"    width: 100%;");
            sb.AppendLine(@"    border-collapse: collapse;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".textGreen {");
            sb.AppendLine(@"    color: green;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bggreenfBottomBorder {");
            sb.AppendLine(@"    text-align: left;");
            sb.AppendLine(@"    background-color: #ccffcc;");
            sb.AppendLine(@"    border-bottom: 1px solid black;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bggreenfLeftBorder {");
            sb.AppendLine(@"    text-align: left;");
            sb.AppendLine(@"    background-color: #ccffcc;");
            sb.AppendLine(@"    border-left: 1px solid black;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bggreenfLeftAndBottomBorder {");
            sb.AppendLine(@"    text-align: left;");
            sb.AppendLine(@"    background-color: #ccffcc;");
            sb.AppendLine(@"    border-left: 1px solid black;");
            sb.AppendLine(@"    border-bottom: 1px solid black;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".leftAndBottomBorder {");
            sb.AppendLine(@"    text-align: left;");
            sb.AppendLine(@"    border-left: 1px solid black;");
            sb.AppendLine(@"    border-bottom: 1px solid black;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".textAlignLeftAndLeftAndBottomBorder {");
            sb.AppendLine(@"    text-align: left;");
            sb.AppendLine(@"    border-left: 1px solid black;");
            sb.AppendLine(@"    border-bottom: 1px solid black;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".textAlignLeft {");
            sb.AppendLine(@"    text-align: left;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".textAlignCenter {");
            sb.AppendLine(@"    text-align: center;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".textAlignRight {");
            sb.AppendLine(@"    text-align: right;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".leftBorder {");
            sb.AppendLine(@"    border-left: 1px solid black;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".allBorders {");
            sb.AppendLine(@"    border: 1px solid black;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".allBordersNoWrap {");
            sb.AppendLine(@"    border: 1px solid black;");
            sb.AppendLine(@"    white-space: nowrap;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".rightBorder {");
            sb.AppendLine(@"    border-right: 1px solid black;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".rightBottomBorder {");
            sb.AppendLine(@"    border-right: 1px solid black;");
            sb.AppendLine(@"    border-bottom: 1px solid black;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bottomBorderGreenText {");
            sb.AppendLine(@"    color: green;");
            sb.AppendLine(@"    border-bottom: 1px solid black;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bottomBorder {");
            sb.AppendLine(@"    border-bottom: 1px solid black;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".topBorder {");
            sb.AppendLine(@"    border-top: 1px solid black;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".topRightBorder {");
            sb.AppendLine(@"    border-top: 1px solid black;");
            sb.AppendLine(@"    border-right: 1px solid black;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bglightblue {");
            sb.AppendLine(@"    background-color: lightblue;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bgyellow {");
            sb.AppendLine(@"    background-color: yellow;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bggreena {");
            sb.AppendLine(@"    background-color: #009900;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bggreenb {");
            sb.AppendLine(@"    background-color: #00bb00;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bggreenc {");
            sb.AppendLine(@"    background-color: #11ff11;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bggreend {");
            sb.AppendLine(@"    background-color: #44ff44;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bggreene {");
            sb.AppendLine(@"    background-color: #99ff99;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bggreenf {");
            sb.AppendLine(@"    background-color: #ccffcc;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bgreda {");
            sb.AppendLine(@"    background-color: #ffcccc;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bgredb {");
            sb.AppendLine(@"    background-color: #ff9999;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bgredc {");
            sb.AppendLine(@"    background-color: #ff4444;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bgredd {");
            sb.AppendLine(@"    background-color: #ff1111;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bgrede {");
            sb.AppendLine(@"    background-color: #cc0000;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bgredf {");
            sb.AppendLine(@"    background-color: #aa0000;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bgbluea {");
            sb.AppendLine(@"    background-color: #ddddff;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bgblueb {");
            sb.AppendLine(@"    background-color: #ccccff;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bgbluec {");
            sb.AppendLine(@"    background-color: #bbbbff;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bgblued {");
            sb.AppendLine(@"    background-color: #aaaaff;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bgbluee {");
            sb.AppendLine(@"    background-color: #9999ff;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bgbluef {");
            sb.AppendLine(@"    background-color: #8888ff;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".bgblack {");
            sb.AppendLine(@"    background-color: #000000;");
            sb.AppendLine(@"    color: #ffffff;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".listypestylenone {");
            sb.AppendLine(@"    list-style-type: none;");
            sb.AppendLine(@"}");
            sb.AppendLine(@".lifloatleft {");
            sb.AppendLine(@"    float: left;");
            sb.AppendLine(@"    display: block;");
            sb.AppendLine(@"}");
            sb.AppendLine(@"    </style>");
            sb.AppendLine(@"</head>");
            sb.AppendLine(@"<body>");
         
            return true;
        }
    }
}
