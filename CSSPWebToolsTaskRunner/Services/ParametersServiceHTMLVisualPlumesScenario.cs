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
        private bool GenerateHTMLVisualPlumesScenario()
        {
            switch (reportTypeModel.UniqueCode)
            {
                case "VisualPlumesScenarioTestFRDOCX":
                case "VisualPlumesScenarioTestENDOCX":
                    {
                        if (!GenerateHTMLVisualPlumesScenario_VisualPlumesScenarioTestDocx())
                        {
                            return false;
                        }
                    }
                    break;
                case "VisualPlumesScenarioTestExcelFRXLSX":
                case "VisualPlumesScenarioTestExcelENXLSX":
                    {
                        if (!GenerateHTMLVisualPlumesScenario_VisualPlumesScenarioTestXlsx())
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
                    if (!GenerateHTMLVisualPlumesScenario_NotImplemented())
                    {
                        return false;
                    }
                    break;
            }
            return true;
        }
        private bool GenerateHTMLVisualPlumesScenario_VisualPlumesScenarioTestDocx()
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
        private bool GenerateHTMLVisualPlumesScenario_VisualPlumesScenarioTestXlsx()
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
        private bool GenerateHTMLVisualPlumesScenario_NotImplemented()
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
