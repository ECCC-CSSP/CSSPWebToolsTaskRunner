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
using CSSPModelsDLL.Models;
using System.Windows.Forms;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        private bool GenerateHTMLSubsector(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            switch (reportTypeModel.UniqueCode)
            {
                case "FCSummaryStatDocx":
                    {
                        if (!GenerateHTMLSubsectorFCSummaryStatDocx(fi, sbHTML, parameters, reportTypeModel))
                        {
                            return false;
                        }
                    }
                    break;
                case "MapActivePolSourceSitesDocx":
                    {
                        if (!GenerateHTMLSubsectorMapActivePolSourceSitesDocx(fi, sbHTML, parameters, reportTypeModel))
                        {
                            return false;
                        }
                    }
                    break;
                case "MapActiveMWQMSitesDocx":
                    {
                        if (!GenerateHTMLSubsectorMapMWQMSitesDocx(fi, sbHTML, parameters, reportTypeModel))
                        {
                            return false;
                        }
                    }
                    break;
                case "FullReportCoverPage":
                    {
                        if (!GenerateHTMLSubsectorFullReportCoverPage(fi, sbHTML, parameters, reportTypeModel))
                        {
                            return false;
                        }
                    }
                    break;
                case "PollutionSourceSites":
                    {
                        if (!GenerateHTMLSubsectorPollutionSourceSites(fi, sbHTML, parameters, reportTypeModel))
                        {
                            return false;
                        }
                    }
                    break;
                case "FCSummaryStatXlsx":
                    {
                        if (!GenerateHTMLSubsector_SubsectorTestXlsx(fi, sbHTML, parameters, reportTypeModel))
                        {
                            return false;
                        }
                    }
                    break;
                case "SomethingElseAsUniqueCode":
                    {
                    }
                    break;
                default:
                    if (!GenerateHTMLSubsector_NotImplemented(fi, sbHTML, parameters, reportTypeModel))
                    {
                        return false;
                    }
                    break;
            }
            return true;
        }
        private bool GenerateHTMLSubsector_SubsectorTestXlsx(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            if (!GetTopHTML(sbHTML))
            {
                return false;
            }

            sbHTML.AppendLine(@"<h2>This will contain the Summary statistics FC densities (MPN/100mL)</h2>");

            if (!GetBottomHTML(sbHTML, fi, parameters))
            {
                return false;
            }

            return true;
        }
        private bool GenerateHTMLSubsector_NotImplemented(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            if (!GetTopHTML(sbHTML))
            {
                return false;
            }

            sbHTML.AppendLine(@"<h2>UniqueCode [" + reportTypeModel.UniqueCode + " is not implemented.</h2>");

            if (!GetBottomHTML(sbHTML, fi, parameters))
            {
                return false;
            }

            return true;
        }
        private string GetLastClassificationColor(MWQMSiteLatestClassificationEnum? mwqmSiteLatestClassification)
        {
            if (mwqmSiteLatestClassification == null)
            {
                return "";
            }

            switch (mwqmSiteLatestClassification)
            {
                case MWQMSiteLatestClassificationEnum.Approved:
                    return "bggreena";
                case MWQMSiteLatestClassificationEnum.ConditionallyApproved:
                    return "bggreenf";
                case MWQMSiteLatestClassificationEnum.ConditionallyRestricted:
                    return "bgredf";
                case MWQMSiteLatestClassificationEnum.Prohibited:
                    return "bgblack";
                case MWQMSiteLatestClassificationEnum.Restricted:
                    return "bgredf";
                case MWQMSiteLatestClassificationEnum.Unclassified:
                    return "";
                default:
                    return "";
            }
        }
        private string GetLastClassificationInitial(MWQMSiteLatestClassificationEnum? mwqmSiteLatestClassification)
        {
            if (mwqmSiteLatestClassification == null)
            {
                return "";
            }

            switch (mwqmSiteLatestClassification)
            {
                case MWQMSiteLatestClassificationEnum.Approved:
                    return (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr ? "A" : "A");
                case MWQMSiteLatestClassificationEnum.ConditionallyApproved:
                    return (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr ? "CA" : "AC");
                case MWQMSiteLatestClassificationEnum.ConditionallyRestricted:
                    return (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr ? "CR" : "RC");
                case MWQMSiteLatestClassificationEnum.Prohibited:
                    return (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr ? "P" : "P");
                case MWQMSiteLatestClassificationEnum.Restricted:
                    return (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr ? "R" : "R");
                case MWQMSiteLatestClassificationEnum.Unclassified:
                    return (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr ? "" : "");
                default:
                    return "";
            }
        }

        // for testing only can comment out when test is completed
    }
}
