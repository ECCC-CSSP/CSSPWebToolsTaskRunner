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
        private bool GenerateHTMLInfrastructure()
        {
            switch (reportTypeModel.UniqueCode)
            {
                case "InfrastructureTestFRDOCX":
                case "InfrastructureTestENDOCX":
                    {
                        if (!GenerateHTMLInfrastructure_InfrastructureTestDocx())
                        {
                            return false;
                        }
                    }
                    break;
                case "InfrastructureTestExcelFRXLSX":
                case "InfrastructureTestExcelENXLSX":
                    {
                        if (!GenerateHTMLInfrastructure_InfrastructureTestXlsx())
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
                    if (!GenerateHTMLInfrastructure_NotImplemented())
                    {
                        return false;
                    }
                    break;
            }
            return true;
        }
        private bool GenerateHTMLInfrastructure_InfrastructureTestDocx()
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
        private bool GenerateHTMLInfrastructure_InfrastructureTestXlsx()
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
        private bool GenerateHTMLInfrastructure_NotImplemented()
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
