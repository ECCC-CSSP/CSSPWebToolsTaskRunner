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
using CSSPEnumsDLL.Services;
using CSSPModelsDLL.Models;
using System.Windows.Forms;
using System.Web.Helpers;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        private bool GenerateHTMLSUBSECTOR_MWQM_SITES_NUMBER_OF_SITES_BY_YEAR(StringBuilder sbTemp)
        {
            string NotUsed = "";

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 3);

            List<string> ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            // TVItemID and Year already loaded

            TVItemModel tvItemModelSubsector = _TVItemService.GetTVItemModelWithTVItemIDDB(TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelSubsector.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                return false;
            }

            string ServerPath = _TVFileService.GetServerFilePath(tvItemModelSubsector.TVItemID);

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 5);

            List<TVItemModel> tvItemModelListMWQMSites = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.MWQMSite).Where(c => c.IsActive == true).ToList();
            List<MWQMSiteModel> mwqmSiteModelList = _MWQMSiteService.GetMWQMSiteModelListWithSubsectorTVItemIDDB(TVItemID);
            List<MWQMRunModel> mwqmRunModelList = _MWQMRunService.GetMWQMRunModelListWithSubsectorTVItemIDDB(TVItemID);
            List<MWQMSampleModel> mwqmSampleModelList = _MWQMSampleService.GetMWQMSampleModelListWithSubsectorTVItemIDDB(TVItemID);

            //------------------------------------------------------------------------------
            // doing number of sites by year
            //------------------------------------------------------------------------------

            List<int> YearList = new List<int>();
            List<int> CountPerYear = new List<int>();

            int MaxYear = Math.Min(DateTime.Now.Year, Year);

            for (int i = MaxYear; i > 1979; i--)
            {
                YearList.Add(i);
                int count = (from s in mwqmSiteModelList
                             from samp in mwqmSampleModelList
                             where s.MWQMSiteTVItemID == samp.MWQMSiteTVItemID
                             && samp.SampleDateTime_Local.Year == i
                             select s.MWQMSiteTVItemID).Distinct().Count();

                CountPerYear.Add(count);
            }

            Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = xlApp.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.Worksheets.get_Item(1);

            Microsoft.Office.Interop.Excel.ChartObjects xlCharts = (Microsoft.Office.Interop.Excel.ChartObjects)worksheet.ChartObjects();
            Microsoft.Office.Interop.Excel.ChartObject chart = xlCharts.Add(100, 100, 600, 200);
            Microsoft.Office.Interop.Excel.Chart chartPage = chart.Chart;

            chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;

            Microsoft.Office.Interop.Excel.SeriesCollection seriesCollection = chartPage.SeriesCollection();
            Microsoft.Office.Interop.Excel.Series series = seriesCollection.NewSeries();

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 40);

            series.XValues = YearList.ToArray();
            series.Values = CountPerYear.ToArray();

            chartPage.ApplyLayout(9, Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered);
            chartPage.ChartTitle.Select();
            xlApp.Selection.Delete();
            chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue).AxisTitle.Select();
            xlApp.Selection.Delete();
            chartPage.Legend.Select();
            xlApp.Selection.Delete();
            chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).MajorTickMark = Microsoft.Office.Interop.Excel.Constants.xlOutside;
            chartPage.Parent.RoundedCorners = true;

            chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = TaskRunnerServiceRes.NumberOfMWQMSitesByYear;

            // need to save the file with a unique name under the TVItemID
            FileInfo fiImageNumberOfSitesByYearStat = new FileInfo(fi.DirectoryName + @"\NumberOfSitesByYearStat" + FileNameExtra + ".png");

            DirectoryInfo di = new DirectoryInfo(fi.DirectoryName);

            if (!di.Exists)
            {
                try
                {
                    di.Create();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateDirectory__, di.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateDirectory__", di.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return false;
                }
            }

            chartPage.Export(fiImageNumberOfSitesByYearStat.FullName, "PNG", false);

            sbTemp.AppendLine($@"<div class=""textAlignCenter"">|||Image|FileName,{ fiImageNumberOfSitesByYearStat.FullName }|width,400|height,150|||</div>");
            sbTemp.AppendLine($@"|||FigureCaption|: { TaskRunnerServiceRes.NumberOfMWQMSitesByYear }|||");

            if (workbook != null)
            {
                workbook.Close(false);
            }
            if (xlApp != null)
            {
                xlApp.Quit();
            }

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 80);

            return true;
        }

        // for testing only can comment out when test is completed
        public bool PublicGenerateHTMLSUBSECTOR_MWQM_SITES_NUMBER_OF_SITES_BY_YEAR(StringBuilder sbTemp)
        {
            bool retBool = GenerateHTMLSUBSECTOR_MWQM_SITES_NUMBER_OF_SITES_BY_YEAR(sbTemp);

            StreamWriter sw = fi.CreateText();
            sw.Write(sbTemp.ToString());
            sw.Flush();
            sw.Close();

            return retBool;
        }

    }
}
