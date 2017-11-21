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
        private bool GenerateHTMLMunicipality(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            switch (reportTypeModel.UniqueCode)
            {
                case "MunicipalityTestDocx":
                    {
                        if (!GenerateHTMLMunicipality_MunicipalityTestDocx(fi, sbHTML, parameters, reportTypeModel))
                        {
                            return false;
                        }
                    }
                    break;
                case "MunicipalityTestExcel":
                    {
                        if (!GenerateHTMLMunicipality_MunicipalityTestXlsx(fi, sbHTML, parameters, reportTypeModel))
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
                    if (!GenerateHTMLMunicipality_NotImplemented(fi, sbHTML, parameters, reportTypeModel))
                    {
                        return false;
                    }
                    break;
            }
            return true;
        }
        private bool GenerateHTMLMunicipality_MunicipalityTestDocx(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            if (!GetTopHTML(sbHTML))
            {
                return false;
            }

            sbHTML.AppendLine(@"<h2>Bonjour</h2>");

            if (!GetBottomHTML(sbHTML, fi, parameters))
            {
                return false;
            }

            return true;
        }
        private bool GenerateHTMLMunicipality_MunicipalityTestXlsx(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            if (!GetTopHTML(sbHTML))
            {
                return false;
            }

            sbHTML.AppendLine(@"<h2>Bonjour 2 for xlsx</h2>");

            if (!GetBottomHTML(sbHTML, fi, parameters))
            {
                return false;
            }

            return true;
        }
        private bool GenerateHTMLMunicipality_NotImplemented(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
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
    }
}