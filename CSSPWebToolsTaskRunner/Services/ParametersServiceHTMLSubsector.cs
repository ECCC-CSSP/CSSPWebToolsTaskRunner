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
        private bool GenerateHTMLSubsector(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            switch (reportTypeModel.UniqueCode)
            {
                case "FCSummaryStatDocx":
                    {
                        if (!GenerateHTMLSubsector_SubsectorTestDocx(fi, sbHTML, parameters, reportTypeModel))
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
        private bool GenerateHTMLSubsector_SubsectorTestDocx(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            if (!GetTopHTML(sbHTML))
            {
                return false;
            }

            sbHTML.AppendLine(@"<h2>This will contain the Summary statistics FC densities (MPN/100mL)</h2>");

            sbHTML.AppendLine(@"<table>");
            sbHTML.AppendLine(@"    <thead>");
            sbHTML.AppendLine(@"        <tr>");
            sbHTML.AppendLine(@"            <th>Station</th>");
            sbHTML.AppendLine(@"            <th>Samples</th>");
            sbHTML.AppendLine(@"            <th>Period</th>");
            sbHTML.AppendLine(@"            <th>Min FC</th>");
            sbHTML.AppendLine(@"            <th>Max FC</th>");
            sbHTML.AppendLine(@"            <th>GMean</th>");
            sbHTML.AppendLine(@"            <th>Median</th> ");
            sbHTML.AppendLine(@"            <th>P90</th>");
            sbHTML.AppendLine(@"            <th>% &gt; 43</th>");
            sbHTML.AppendLine(@"        </tr>");
            sbHTML.AppendLine(@"    </thead>");
            sbHTML.AppendLine(@"    <tbody>");
            sbHTML.AppendLine(@"        <tr>");
            sbHTML.AppendLine(@"            <td>1</td>");
            sbHTML.AppendLine(@"            <td>2</td>");
            sbHTML.AppendLine(@"            <td>3</td>");
            sbHTML.AppendLine(@"            <td>4</td>");
            sbHTML.AppendLine(@"            <td>5</td>");
            sbHTML.AppendLine(@"            <td>6</td>");
            sbHTML.AppendLine(@"            <td>7</td>");
            sbHTML.AppendLine(@"            <td>8</td>");
            sbHTML.AppendLine(@"            <td>9</td>");
            sbHTML.AppendLine(@"        </tr>");
            sbHTML.AppendLine(@"        <tr class=""alternate"">");
            sbHTML.AppendLine(@"            <td>1</td>");
            sbHTML.AppendLine(@"            <td>2</td>");
            sbHTML.AppendLine(@"            <td>3</td>");
            sbHTML.AppendLine(@"            <td>4</td>");
            sbHTML.AppendLine(@"            <td>5</td>");
            sbHTML.AppendLine(@"            <td>6</td>");
            sbHTML.AppendLine(@"            <td>7</td>");
            sbHTML.AppendLine(@"            <td>8</td>");
            sbHTML.AppendLine(@"            <td>9</td>");
            sbHTML.AppendLine(@"        </tr>");
            sbHTML.AppendLine(@"        <tr>");
            sbHTML.AppendLine(@"            <td>1</td>");
            sbHTML.AppendLine(@"            <td>2</td>");
            sbHTML.AppendLine(@"            <td>3</td>");
            sbHTML.AppendLine(@"            <td>4</td>");
            sbHTML.AppendLine(@"            <td>5</td>");
            sbHTML.AppendLine(@"            <td>6</td>");
            sbHTML.AppendLine(@"            <td>7</td>");
            sbHTML.AppendLine(@"            <td>8</td>");
            sbHTML.AppendLine(@"            <td>9</td>");
            sbHTML.AppendLine(@"        </tr>");
            sbHTML.AppendLine(@"        <tr class=""alternate"">");
            sbHTML.AppendLine(@"            <td>1</td>");
            sbHTML.AppendLine(@"            <td>2</td>");
            sbHTML.AppendLine(@"            <td>3</td>");
            sbHTML.AppendLine(@"            <td>4</td>");
            sbHTML.AppendLine(@"            <td>5</td>");
            sbHTML.AppendLine(@"            <td>6</td>");
            sbHTML.AppendLine(@"            <td>7</td>");
            sbHTML.AppendLine(@"            <td>8</td>");
            sbHTML.AppendLine(@"            <td>9</td>");
            sbHTML.AppendLine(@"        </tr>");
            sbHTML.AppendLine(@"    </tbody>");
            sbHTML.AppendLine(@"    <tfoot>");
            sbHTML.AppendLine(@"        <tr>");
            sbHTML.AppendLine(@"            <td colspan=""9"">This is the footer</td>");
            sbHTML.AppendLine(@"        </tr>");
            sbHTML.AppendLine(@"    </tfoot>");
            sbHTML.AppendLine(@"</table>");

            if (!GetBottomHTML(sbHTML, fi, parameters))
            {
                return false;
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

        // for testing only can comment out when test is completed
        public bool PublicGenerateHTMLSubsector_SubsectorTestDocx(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            return GenerateHTMLSubsector_SubsectorTestDocx(fi, sbHTML, parameters, reportTypeModel);
        }
    }
}
