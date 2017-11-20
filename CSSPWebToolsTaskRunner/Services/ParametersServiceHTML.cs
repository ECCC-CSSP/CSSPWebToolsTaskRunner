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
            sbHTML.AppendLine(@"    <br />");
            sbHTML.AppendLine(@"    <div class=""small"">");
            if (fi.Extension.ToLower() == ".html")
            {
                sbHTML.AppendLine(@"        <b>File name:</b> " + fi.FullName.Replace(".html", ".docx") + @"");
            }
            else
            {
                sbHTML.AppendLine(@"        <b>File name:</b> " + fi.FullName + @"");
            }
            sbHTML.AppendLine(@"    </div>");
            sbHTML.AppendLine(@"</body>");
            sbHTML.AppendLine(@"</html>");

            return true;
        }
        private bool GetTopHTML(StringBuilder sbHTML)
        {
            sbHTML.AppendLine(@"<html>");
            sbHTML.AppendLine(@"<head>");
            sbHTML.AppendLine(@"    <meta charset=""utf-8"" />");
            sbHTML.AppendLine(@"    <title>This is the title of the html document</title>");
            sbHTML.AppendLine(@"    <style>");
            sbHTML.AppendLine(@"body {");
            sbHTML.AppendLine(@"    font: normal 8px arial, helvetica, sans-serif;");
            sbHTML.AppendLine(@"}");
            sbHTML.AppendLine(@"table {");
            sbHTML.AppendLine(@"    font: normal 8px arial, helvetica, sans-serif;");
            sbHTML.AppendLine(@"    border: 1px solid black;");
            sbHTML.AppendLine(@"    align-content: center;");
            sbHTML.AppendLine(@"}");
            sbHTML.AppendLine(@".bglightblue {");
            sbHTML.AppendLine(@"    background-color: lightblue;");
            sbHTML.AppendLine(@"}");
            sbHTML.AppendLine(@".bgyellow {");
            sbHTML.AppendLine(@"    background-color: yellow;");
            sbHTML.AppendLine(@"}");
            sbHTML.AppendLine(@".bggreena {");
            sbHTML.AppendLine(@"    background-color: #009900;");
            sbHTML.AppendLine(@"}");
            sbHTML.AppendLine(@".bggreenb {");
            sbHTML.AppendLine(@"    background-color: #00bb00;");
            sbHTML.AppendLine(@"}");
            sbHTML.AppendLine(@".bggreenc {");
            sbHTML.AppendLine(@"    background-color: #11ff11;");
            sbHTML.AppendLine(@"}");
            sbHTML.AppendLine(@".bggreend {");
            sbHTML.AppendLine(@"    background-color: #44ff44;");
            sbHTML.AppendLine(@"}");
            sbHTML.AppendLine(@".bggreene {");
            sbHTML.AppendLine(@"    background-color: #99ff99;");
            sbHTML.AppendLine(@"}");
            sbHTML.AppendLine(@".bggreenf {");
            sbHTML.AppendLine(@"    background-color: #ccffcc;");
            sbHTML.AppendLine(@"}");
            sbHTML.AppendLine(@".bgreda {");
            sbHTML.AppendLine(@"    background-color: #ffcccc;");
            sbHTML.AppendLine(@"}");
            sbHTML.AppendLine(@".bgredb {");
            sbHTML.AppendLine(@"    background-color: #ff9999;");
            sbHTML.AppendLine(@"}");
            sbHTML.AppendLine(@".bgredc {");
            sbHTML.AppendLine(@"    background-color: #ff4444;");
            sbHTML.AppendLine(@"}");
            sbHTML.AppendLine(@".bgredd {");
            sbHTML.AppendLine(@"    background-color: #ff1111;");
            sbHTML.AppendLine(@"}");
            sbHTML.AppendLine(@".bgrede {");
            sbHTML.AppendLine(@"    background-color: #cc0000;");
            sbHTML.AppendLine(@"}");
            sbHTML.AppendLine(@".bgredf {");
            sbHTML.AppendLine(@"    background-color: #aa0000;");
            sbHTML.AppendLine(@"}");
            sbHTML.AppendLine(@".bgbluea {");
            sbHTML.AppendLine(@"    background-color: #ddddff;");
            sbHTML.AppendLine(@"}");
            sbHTML.AppendLine(@".bgblueb {");
            sbHTML.AppendLine(@"    background-color: #ccccff;");
            sbHTML.AppendLine(@"}");
            sbHTML.AppendLine(@".bgbluec {");
            sbHTML.AppendLine(@"    background-color: #bbbbff;");
            sbHTML.AppendLine(@"}");
            sbHTML.AppendLine(@".bgblued {");
            sbHTML.AppendLine(@"    background-color: #aaaaff;");
            sbHTML.AppendLine(@"}");
            sbHTML.AppendLine(@".bgbluee {");
            sbHTML.AppendLine(@"    background-color: #9999ff;");
            sbHTML.AppendLine(@"}");
            sbHTML.AppendLine(@".bgbluef {");
            sbHTML.AppendLine(@"    background-color: #8888ff;");
            sbHTML.AppendLine(@"}");
            sbHTML.AppendLine(@".bgblack {");
            sbHTML.AppendLine(@"    background-color: #000000;");
            sbHTML.AppendLine(@"    color: #ffffff;");
            sbHTML.AppendLine(@"}");
            sbHTML.AppendLine(@"    </style>");
            sbHTML.AppendLine(@"</head>");
            sbHTML.AppendLine(@"<body>");

            return true;
        }
    }
}
