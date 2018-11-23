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
using System.Windows.Forms;
//using System.Web.Helpers;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        private bool GenerateHTMLDocx()
        {
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 3);

            if (!GetTopHTML())
            {
                return false;
            }

            if (!GenerateSectionsRecursiveDocx(new ReportSectionModel()))
            {
                return false;
            }

            if (!GenerateObjects())
            {
                return false;
            }

            if (!GetBottomHTML())
            {
                return false;
            }

            return true;
        }

        private bool GenerateSectionsRecursiveDocx(ReportSectionModel reportSectionModel)
        {
            List<ReportSectionModel> reportSectionModelList = new List<ReportSectionModel>();
            if (reportSectionModel.ReportSectionID == 0)
            {
                reportSectionModelList = _ReportSectionService.GetReportSectionModelListWithReportTypeIDAndTVItemIDAndYearDB(reportTypeModel.ReportTypeID, null, null).Where(c => c.ParentReportSectionID == null).OrderBy(c => c.Ordinal).ToList();
            }
            else
            {
                reportSectionModelList = new List<ReportSectionModel>() { reportSectionModel };
            }

            if (reportSectionModelList.Count == 0)
            {
                sb.AppendLine(TaskRunnerServiceRes.NoReportSectionFound);
            }
            else
            {
                foreach (ReportSectionModel reportSectionModelTemp in reportSectionModelList)
                {
                    List<ReportSectionModel> reportSectionModelYearsList = _ReportSectionService.GetReportSectionModelListWithTemplateReportSectionIDDB(reportSectionModelTemp.ReportSectionID).ToList();

                    if (reportSectionModelYearsList.Count > 0)
                    {
                        ReportSectionModel reportSectionModelUnderOrEqualYear = reportSectionModelYearsList.Where(c => c.Year <= Year).OrderByDescending(c => c.Year).FirstOrDefault();
                        if (reportSectionModelUnderOrEqualYear != null)
                        {
                            sb.Append($@"<p class=""bgyellow"">{ TaskRunnerServiceRes.InformationBelowWasWrittenFor } { reportSectionModelUnderOrEqualYear.Year } { TaskRunnerServiceRes.Report }. { TaskRunnerServiceRes.PleaseModifyIfNeeded }</p>");
                            sb.Append(reportSectionModelUnderOrEqualYear.ReportSectionText);
                            sb.Append($@"<p class=""bgyellow"">{ TaskRunnerServiceRes.InformationAboveWasWrittenFor } { reportSectionModelUnderOrEqualYear.Year } { TaskRunnerServiceRes.Report }. { TaskRunnerServiceRes.PleaseModifyIfNeeded }</p>");
                        }
                        else
                        {
                            ReportSectionModel reportSectionModelOverYear = reportSectionModelYearsList.Where(c => c.Year > Year).OrderByDescending(c => c.Year).FirstOrDefault();
                            if (reportSectionModelOverYear != null)
                            {
                                sb.Append($@"<p class=""bgyellow"">{ TaskRunnerServiceRes.InformationBelowWasWrittenFor } { reportSectionModelOverYear.Year } { TaskRunnerServiceRes.Report }. { TaskRunnerServiceRes.PleaseModifyIfNeeded }</p>");
                                sb.Append(reportSectionModelOverYear.ReportSectionText);
                                sb.Append($@"<p class=""bgyellow"">{ TaskRunnerServiceRes.InformationAboveWasWrittenFor } { reportSectionModelOverYear.Year } { TaskRunnerServiceRes.Report }. { TaskRunnerServiceRes.PleaseModifyIfNeeded }</p>");
                            }
                            else
                            {

                                sb.AppendLine($"<p> { TaskRunnerServiceRes.CouldNotFindYearSpecificReportSection }</p>");
                            }
                        }
                    }
                    else
                    {
                        sb.Append(reportSectionModelTemp.ReportSectionText);
                    }

                    List<ReportSectionModel> reportSectionModelChildrenList = _ReportSectionService.GetReportSectionModelListWithParentReportSectionIDDB(reportSectionModelTemp.ReportSectionID).OrderBy(c => c.Ordinal).ToList();

                    foreach (ReportSectionModel reportSectionModelChild in reportSectionModelChildrenList)
                    {
                        GenerateSectionsRecursiveDocx(reportSectionModelChild);
                    }
                }
            }

            return true;
        }

        // for testing only can comment out when test is completed
        public bool PublicGenerateHTMLDocx()
        {
            bool retBool = GenerateHTMLDocx();

            StreamWriter sw = fi.CreateText();
            sw.Write(sb.ToString());
            sw.Flush();
            sw.Close();

            return retBool;
        }
    }
}
