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
                    if (reportSectionModelTemp.Year != null && reportSectionModelTemp.Year != Year)
                    {
                        sb.Append($@"<p class=""bgyellow"">{ TaskRunnerServiceRes.InformationBelowWrittenFor } { reportSectionModelTemp.Year } { TaskRunnerServiceRes.Report }. { TaskRunnerServiceRes.PleaseModifyIfNeeded }</p>");
                        sb.Append(reportSectionModelTemp.ReportSectionText);
                        sb.Append($@"<p class=""bgyellow"">{ TaskRunnerServiceRes.InformationAboveWrittenFor } { reportSectionModelTemp.Year } { TaskRunnerServiceRes.Report }. { TaskRunnerServiceRes.PleaseModifyIfNeeded }</p>");
                    }
                    else
                    {
                        sb.Append(reportSectionModelTemp.ReportSectionText);
                    }

                    List<ReportSectionModel> reportSectionModelAllChildrenList = _ReportSectionService.GetReportSectionModelListWithParentReportSectionIDDB(reportSectionModelTemp.ReportSectionID).OrderBy(c => c.Ordinal).ToList();

                    ReportSectionModel reportSectionModelNonStaticChild = (from c in reportSectionModelAllChildrenList
                                                                           where c.TVItemID == TVItemID
                                                                           && c.Year != null
                                                                           && c.Year <= Year
                                                                           orderby c.Year descending
                                                                           select c).FirstOrDefault();

                    List<ReportSectionModel> reportSectionModelAllStaticChildrenList = (from c in reportSectionModelAllChildrenList
                                                                                        where c.TVItemID == null
                                                                                        && c.Year == null
                                                                                        select c).ToList();

                    List<ReportSectionModel> reportSectionModelChildrenList = new List<ReportSectionModel>();

                    int Ordinal = 0;
                    if (reportSectionModelNonStaticChild != null)
                    {
                        Ordinal = reportSectionModelNonStaticChild.Ordinal;
                        reportSectionModelChildrenList.Add(reportSectionModelNonStaticChild);
                    }

                    foreach (ReportSectionModel reportSectionModelNoYear in reportSectionModelAllStaticChildrenList)
                    {
                        if (Ordinal > 0 && reportSectionModelNoYear.Ordinal == Ordinal)
                        {
                            continue;
                        }

                        reportSectionModelChildrenList.Add(reportSectionModelNoYear);
                    }

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
