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
        private bool GenerateHTMLMWQMSite()
        {
            switch (reportTypeModel.UniqueCode)
            {
                case "MWQMSiteTestFRDOCX":
                case "MWQMSiteTestENDOCX":
                    {
                        if (!GenerateHTMLMWQMSite_MWQMSiteTestDocx())
                        {
                            return false;
                        }
                    }
                    break;
                case "MWQMSiteTestExcelFRXLSX":
                case "MWQMSiteTestExcelENXLSX":
                    {
                        if (!GenerateHTMLMWQMSite_MWQMSiteTestXlsx())
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
                    if (!GenerateHTMLMWQMSite_NotImplemented())
                    {
                        return false;
                    }
                    break;
            }
            return true;
        }
        private bool GenerateHTMLMWQMSite_MWQMSiteTestDocx()
        {
            if (!GetTopHTML())
            {
                return false;
            }

            sb.AppendLine(@"<h2>Bonjour</h2>");

            if (!GetBottomHTML())
            {
                return false;
            }

            return true;
        }
        private bool GenerateHTMLMWQMSite_MWQMSiteTestXlsx()
        {
            if (!GetTopHTML())
            {
                return false;
            }

            sb.AppendLine(@"<h2>Bonjour 2 for xlsx</h2>");

            if (!GetBottomHTML())
            {
                return false;
            }

            return true;
        }
        private bool GenerateHTMLMWQMSite_NotImplemented()
        {
            if (!GetTopHTML())
            {
                return false;
            }

            sb.AppendLine(@"<h2>UniqueCode [" + reportTypeModel.UniqueCode + " is not implemented.</h2>");

            if (!GetBottomHTML())
            {
                return false;
            }

            return true;
        }
    }
}
