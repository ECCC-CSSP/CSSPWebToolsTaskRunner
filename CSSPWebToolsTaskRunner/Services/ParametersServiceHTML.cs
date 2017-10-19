﻿using System;
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

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        private bool GenerateHTML(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            switch (reportTypeModel.TVType)
            {
                case TVTypeEnum.Root:
                    {
                        if (!GenerateHTMLRoot(fi, sbHTML, parameters, reportTypeModel))
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.Area:
                    {
                        if (!GenerateHTMLArea(fi, sbHTML, parameters, reportTypeModel))
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.Country:
                    {
                        if (!GenerateHTMLCountry(fi, sbHTML, parameters, reportTypeModel))
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.Infrastructure:
                    {
                        if (!GenerateHTMLInfrastructure(fi, sbHTML, parameters, reportTypeModel))
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.MikeScenario:
                    {
                        if (!GenerateHTMLMikeScenario(fi, sbHTML, parameters, reportTypeModel))
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.MikeSource:
                    {
                        if (!GenerateHTMLMikeSource(fi, sbHTML, parameters, reportTypeModel))
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.Municipality:
                    {
                        if (!GenerateHTMLMunicipality(fi, sbHTML, parameters, reportTypeModel))
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.MWQMSite:
                    {
                        if (!GenerateHTMLMWQMSite(fi, sbHTML, parameters, reportTypeModel))
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.PolSourceSite:
                    {
                        if (!GenerateHTMLPolSourceSite(fi, sbHTML, parameters, reportTypeModel))
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.Province:
                    {
                        if (!GenerateHTMLProvince(fi, sbHTML, parameters, reportTypeModel))
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.Sector:
                    {
                        if (!GenerateHTMLSector(fi, sbHTML, parameters, reportTypeModel))
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.Subsector:
                    {
                        if (!GenerateHTMLSubsector(fi, sbHTML, parameters, reportTypeModel))
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.BoxModel:
                    {
                        if (!GenerateHTMLBoxModel(fi, sbHTML, parameters, reportTypeModel))
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.VisualPlumesScenario:
                    {
                        if (!GenerateHTMLVisualPlumesScenario(fi, sbHTML, parameters, reportTypeModel))
                        {
                            return false;
                        }
                    }
                    break;
                default:
                    {
                        if (!GenerateHTMLNotImplemented(fi, sbHTML, parameters, reportTypeModel))
                        {
                            return false;
                        }
                    }
                    break;
            }

            return true;
        }
        private bool GenerateHTMLNotImplemented(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            if (!GetTopHTML(sbHTML))
            {
                return false;
            }

            sbHTML.AppendLine(@"<h2>" + reportTypeModel.TVType.ToString() + " Not Implemented Yet</h2>");

            if (!GetBottomHTML(sbHTML, fi, parameters))
            {
                return false;
            }

            return true;
        }
        private bool GetBottomHTML(StringBuilder sbHTML, FileInfo fi, string parameters)
        {
            sbHTML.AppendLine(@"    <div class=""small"">");
            sbHTML.AppendLine(@"        <b>File name:</b> " + fi.FullName + @"");
            sbHTML.AppendLine(@"    </div>");
            sbHTML.AppendLine(@"    <div class=""small"">");
            sbHTML.AppendLine(@"        <b>Parameters:</b> " + parameters + @"");
            sbHTML.AppendLine(@"    </div>");
            sbHTML.AppendLine(@"</body>");
            sbHTML.AppendLine(@"</html>");

            return true;
        }
        private bool GetTopHTML(StringBuilder sbHTML)
        {
            sbHTML.AppendLine(@"<html>");
            sbHTML.AppendLine(@"<head>");
            sbHTML.AppendLine(@"    <title>This is the title of the html document</title>");
            sbHTML.AppendLine(@"    <link href=""C:\Users\leblancc\Desktop\TestHTML\ReportType.css"" rel=""stylesheet"" />");
            sbHTML.AppendLine(@"</head>");
            sbHTML.AppendLine(@"<body>");

            return true;
        }
    }
}
