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
        private bool GenerateHTMLMikeSource(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            switch (reportTypeModel.UniqueCode)
            {
                case "MikeSourceTestDocx":
                    {
                        if (!GenerateHTMLMikeSource_MikeSourceTestDocx(fi, sbHTML, parameters, reportTypeModel))
                        {
                            return false;
                        }
                    }
                    break;
                case "MikeSourceTestExcel":
                    {
                        if (!GenerateHTMLMikeSource_MikeSourceTestXlsx(fi, sbHTML, parameters, reportTypeModel))
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
                    if (!GenerateHTMLMikeSource_NotImplemented(fi, sbHTML, parameters, reportTypeModel))
                    {
                        return false;
                    }
                    break;
            }
            return true;
        }
        private bool GenerateHTMLMikeSource_MikeSourceTestDocx(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
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
        private bool GenerateHTMLMikeSource_MikeSourceTestXlsx(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
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
        private bool GenerateHTMLMikeSource_NotImplemented(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
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