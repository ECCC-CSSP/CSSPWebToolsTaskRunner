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
        private bool GenerateHTMLSUBSECTOR_MWQM_SITES(StringBuilder sbTemp)
        {
            string NotUsed = "";

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 3);

            if (!GetTopHTML())
            {
                return false;
            }

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

            sbTemp.AppendLine($@" <h3>{ TaskRunnerServiceRes.MWQMSiteSampleDataAvailability }</h3>");
            sbTemp.AppendLine($@" <table cellpadding=""5"">");
            sbTemp.AppendLine($@" <tr>");
            sbTemp.AppendLine($@" <th>{ TaskRunnerServiceRes.Site }</th>");
            bool FirstHit = false;
            for (int year = DateTime.Now.Year; year > 1975; year--)
            {
                if (year % 5 == 0)
                {
                    FirstHit = true;
                    int colSpan = 5;
                    if (year == 1980)
                    {
                        colSpan = 4;
                    }
                    sbTemp.AppendLine($@" <th class=""textAlignLeftAndLeftBorder"" colspan=""{ colSpan }"">{ year }</th>");
                }
                if (!FirstHit)
                {
                    sbTemp.AppendLine($@" <th>&nbsp;</th>");
                }
            }
            sbTemp.AppendLine($@" </tr>");
            int countSite = 0;
            foreach (MWQMSiteModel mwqmSiteModel in mwqmSiteModelList)
            {
                TVItemModel tvItemModel = tvItemModelListMWQMSites.Where(c => c.TVItemID == mwqmSiteModel.MWQMSiteTVItemID).FirstOrDefault();
                if (tvItemModel != null)
                {
                    if (tvItemModel.IsActive)
                    {
                        countSite += 1;
                        string bottomClass = "";
                        if (countSite % 5 == 0)
                        {
                            bottomClass = "bottomBorder";
                        }
                        sbTemp.AppendLine($@" <tr>");
                        sbTemp.AppendLine($@" <td class=""{ bottomClass }"">{ mwqmSiteModel.MWQMSiteTVText }</td>");
                        for (int year = DateTime.Now.Year; year > 1979; year--)
                        {
                            string leftClass = year % 5 == 0 ? "leftBorder" : "";
                            bool hasSamples = mwqmSampleModelList.Where(c => c.MWQMSiteTVItemID == mwqmSiteModel.MWQMSiteTVItemID && c.SampleDateTime_Local.Year == year && c.SampleTypesText.Contains(((int)SampleTypeEnum.Routine).ToString())).Any();
                            if (hasSamples)
                            {
                                if (leftClass != "")
                                {
                                    if (bottomClass != "")
                                    {
                                        sbTemp.AppendLine($@" <td class=""bggreenfLeftAndBottomBorder"">&nbsp;</td>");
                                    }
                                    else
                                    {
                                        sbTemp.AppendLine($@" <td class=""bggreenfLeftBorder"">&nbsp;</td>");
                                    }
                                }
                                else
                                {
                                    if (bottomClass != "")
                                    {
                                        sbTemp.AppendLine($@" <td class=""bggreenfBottomBorder"">&nbsp;</td>");
                                    }
                                    else
                                    {
                                        sbTemp.AppendLine($@" <td class=""bggreenf"">&nbsp;</td>");
                                    }
                                }
                            }
                            else
                            {
                                if (leftClass != "")
                                {
                                    if (bottomClass != "")
                                    {
                                        sbTemp.AppendLine($@" <td class=""leftAndBottomBorder"">&nbsp;</td>");
                                    }
                                    else
                                    {
                                        sbTemp.AppendLine($@" <td class=""leftBorder"">&nbsp;</td>");
                                    }
                                }
                                else
                                {
                                    if (bottomClass != "")
                                    {
                                        sbTemp.AppendLine($@" <td class=""bottomBorder"">&nbsp;</td>");
                                    }
                                    else
                                    {
                                        sbTemp.AppendLine($@" <td>&nbsp;</td>");
                                    }
                                }
                            }
                        }
                        sbTemp.AppendLine($@" </tr>");
                    }
                }
            }
            sbTemp.AppendLine($@" </table>");

            sbTemp.AppendLine(@"<p>|||PAGE_BREAK|||</p>");

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 10);

            // ---------------------------------------------------------------------------------------------
            // MWQM Sites Summary
            // ---------------------------------------------------------------------------------------------

            //------------------------------------------------------------------------------
            // doing number of sites by year
            //------------------------------------------------------------------------------

            List<int> YearList = new List<int>();
            List<int> CountPerYear = new List<int>();

            for (int i = DateTime.Now.Year; i > 1979; i--)
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
            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).TickLabelSpacing = 5;
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


            //------------------------------------------------------------------------------
            // doing number of runs by year
            //------------------------------------------------------------------------------
            YearList = new List<int>();
            CountPerYear = new List<int>();

            for (int i = DateTime.Now.Year; i > 1979; i--)
            {
                YearList.Add(i);
                int count = (from r in mwqmRunModelList
                             from samp in mwqmSampleModelList
                             where r.MWQMRunTVItemID == samp.MWQMRunTVItemID
                             && samp.SampleDateTime_Local.Year == i
                             select r.MWQMRunTVItemID).Distinct().Count();

                CountPerYear.Add(count);
            }

            chart = xlCharts.Add(100, 100, 600, 200);
            chartPage = chart.Chart;

            chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;

            seriesCollection = chartPage.SeriesCollection();
            series = seriesCollection.NewSeries();

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
            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).TickLabelSpacing = 5;
            chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).MajorTickMark = Microsoft.Office.Interop.Excel.Constants.xlOutside;
            chartPage.Parent.RoundedCorners = true;

            chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = TaskRunnerServiceRes.NumberOfMWQMRunsByYear;

            // need to save the file with a unique name under the TVItemID
            FileInfo fiImageNumberOfRunsByYearStat = new FileInfo(fi.DirectoryName + @"\NumberOfRunsByYearStat" + FileNameExtra + ".png");

            di = new DirectoryInfo(fi.DirectoryName);

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

            chartPage.Export(fiImageNumberOfRunsByYearStat.FullName, "PNG", false);

            //------------------------------------------------------------------------------
            // doing number of samples by year
            //------------------------------------------------------------------------------
            YearList = new List<int>();
            CountPerYear = new List<int>();

            for (int i = DateTime.Now.Year; i > 1979; i--)
            {
                YearList.Add(i);
                int count = (from samp in mwqmSampleModelList
                             where samp.SampleDateTime_Local.Year == i
                             select samp.MWQMSampleID).Distinct().Count();

                CountPerYear.Add(count);
            }

            chart = xlCharts.Add(100, 100, 600, 200);
            chartPage = chart.Chart;

            chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;

            seriesCollection = chartPage.SeriesCollection();
            series = seriesCollection.NewSeries();

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
            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).TickLabelSpacing = 5;
            chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).MajorTickMark = Microsoft.Office.Interop.Excel.Constants.xlOutside;
            chartPage.Parent.RoundedCorners = true;

            chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = TaskRunnerServiceRes.NumberOfSamplesByYear;

            // need to save the file with a unique name under the TVItemID
            FileInfo fiImageNumberOfSamplesByYearStat = new FileInfo(fi.DirectoryName + @"\NumberOfSamplesByYearStat" + FileNameExtra + ".png");

            di = new DirectoryInfo(fi.DirectoryName);

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

            chartPage.Export(fiImageNumberOfSamplesByYearStat.FullName, "PNG", false);


            if (workbook != null)
            {
                workbook.Close(false);
            }
            if (xlApp != null)
            {
                xlApp.Quit();
            }

            sbTemp.AppendLine($@" <h3>{ TaskRunnerServiceRes.MWQMSitesSummary }</h3>");
            sbTemp.AppendLine($@" <br /");
            sbTemp.AppendLine($@" <h4>{ TaskRunnerServiceRes.NumberOfMWQMSitesByYear }</h4>");
            sbTemp.AppendLine($@"<div class=""textAlignCenter"">|||Image|FileName,{ fiImageNumberOfSitesByYearStat.FullName }|width,400|height,150|||</div>");
            sbTemp.AppendLine($@" <h4>{ TaskRunnerServiceRes.NumberOfMWQMRunsByYear }</h4>");
            sbTemp.AppendLine($@"<div class=""textAlignCenter"">|||Image|FileName,{ fiImageNumberOfRunsByYearStat.FullName }|width,400|height,150|||</div>");
            sbTemp.AppendLine($@" <h4>{ TaskRunnerServiceRes.NumberOfSamplesByYear }</h4>");
            sbTemp.AppendLine($@"<div class=""textAlignCenter"">|||Image|FileName,{ fiImageNumberOfSamplesByYearStat.FullName }|width,400|height,150|||</div>");

            sbTemp.AppendLine(@"<span>|||PageBreak|||</span>");

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 40);

            sbTemp.AppendLine($@" <h3>{ TaskRunnerServiceRes.MWQMSitesInformation }</h3>");
            sbTemp.AppendLine($@" <table cellpadding=""5"" class=""textAlignCenter"">");
            sbTemp.AppendLine($@" <tr>");
            sbTemp.AppendLine($@" <th colspan=""2"">{ TaskRunnerServiceRes.Site }</th>");
            sbTemp.AppendLine($@" <th>{ TaskRunnerServiceRes.Coordinates }</th>");
            sbTemp.AppendLine($@" <th>{ TaskRunnerServiceRes.Description }</th>");
            sbTemp.AppendLine($@" <th>{ TaskRunnerServiceRes.Photos }</th>");
            sbTemp.AppendLine($@" </tr>");
            foreach (MWQMSiteModel mwqmSiteModel in mwqmSiteModelList)
            {
                TVItemModel tvItemModel = tvItemModelListMWQMSites.Where(c => c.TVItemID == mwqmSiteModel.MWQMSiteTVItemID).FirstOrDefault();
                if (tvItemModel != null)
                {
                    if (tvItemModel.IsActive)
                    {
                        string classificationLetter = "";
                        string classificationColor = "";

                        classificationLetter = GetLastClassificationInitial(mwqmSiteModel.MWQMSiteLatestClassification);
                        classificationColor = GetLastClassificationColor(mwqmSiteModel.MWQMSiteLatestClassification);

                        sbTemp.AppendLine($@" <tr>");
                        sbTemp.AppendLine($@" <td>{ mwqmSiteModel.MWQMSiteTVText }</td>");
                        sbTemp.AppendLine($@" <td class=""{ classificationColor }"">{ classificationLetter }</td>");

                        List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(mwqmSiteModel.MWQMSiteTVItemID, TVTypeEnum.MWQMSite, MapInfoDrawTypeEnum.Point);
                        if (mapInfoPointModelList.Count > 0)
                        {
                            sbTemp.AppendLine($@" <td>{ mapInfoPointModelList[0].Lat.ToString("F5") } { mapInfoPointModelList[0].Lng.ToString("F5") }</td>");
                        }
                        else
                        {
                            sbTemp.AppendLine($@" <td>&nbsp;</td>");
                        }
                        sbTemp.AppendLine($@" <td class=""textAlignLeft"">{ mwqmSiteModel.MWQMSiteDescription }</td>");
                        sbTemp.AppendLine($@" <td>Photo</td>");
                        sbTemp.AppendLine($@" </tr>");
                    }
                }
            }
            sbTemp.AppendLine($@" </table>");


            sbTemp.AppendLine(@"<p>|||PAGE_BREAK|||</p>");

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 50);


            List<MWQMSiteModel> mwqmSiteModelList2 = (from s in mwqmSiteModelList
                                                      from t in tvItemModelListMWQMSites
                                                      where s.MWQMSiteTVItemID == t.TVItemID
                                                      && t.IsActive == true
                                                      orderby s.MWQMSiteTVText
                                                      select s).ToList();

            int skip = 0;
            int take = 15;
            bool HasData = true;
            int countRun = 0;
            while (HasData)
            {
                countRun += 1;

                if (countRun > 2)
                {
                    break;
                }

                List<MWQMRunModel> mwqmRunModelList2 = mwqmRunModelList.Where(c => c.RunSampleType == SampleTypeEnum.Routine).OrderByDescending(c => c.DateTime_Local).Skip(skip).Take(take).ToList();
                if (mwqmRunModelList2.Count > 0)
                {
                    sbTemp.AppendLine($@" <h4>{ TaskRunnerServiceRes.ActiveMWQMSites }&nbsp;&nbsp;{ TaskRunnerServiceRes.FCDensities }&nbsp;&nbsp;&nbsp;({ TaskRunnerServiceRes.Routine })&nbsp;&nbsp;&nbsp;{ mwqmRunModelList2[0].DateTime_Local.ToString("yyyy MMMM dd") } { TaskRunnerServiceRes.To } { mwqmRunModelList2[mwqmRunModelList2.Count -1].DateTime_Local.ToString("yyyy MMMM dd") }</h4>");
                    sbTemp.AppendLine($@" <table class=""FCSalTempDataTableClass"">");
                    sbTemp.AppendLine($@" <tr>");
                    sbTemp.AppendLine($@" <th class=""rightBottomBorder"">{ TaskRunnerServiceRes.Site }</th>");
                    foreach (MWQMRunModel mwqmRunModel in mwqmRunModelList2)
                    {
                        sbTemp.AppendLine($@" <th class=""bottomBorder"">{ mwqmRunModel.DateTime_Local.ToString("yyyy") }<br />{ mwqmRunModel.DateTime_Local.ToString("MMM") }<br />{ mwqmRunModel.DateTime_Local.ToString("dd") }</th>");
                    }
                    sbTemp.AppendLine($@" </tr>");

                    foreach (MWQMSiteModel mwqmSiteModel in mwqmSiteModelList2)
                    {
                        sbTemp.AppendLine($@" <tr>");
                        sbTemp.AppendLine($@" <td class=""rightBorder"">{ mwqmSiteModel.MWQMSiteTVText }</td>");
                        foreach (MWQMRunModel mwqmRunModel in mwqmRunModelList2)
                        {
                            float? value = (from s in mwqmSampleModelList
                                           where s.MWQMRunTVItemID == mwqmRunModel.MWQMRunTVItemID
                                           && s.MWQMSiteTVItemID == mwqmSiteModel.MWQMSiteTVItemID
                                           select s.FecCol_MPN_100ml).FirstOrDefault();

                            string valueStr = value != null ? (value == 1 ? "< 2" : ((float)value).ToString("F0")) : "--";
                            sbTemp.AppendLine($@" <td>{ valueStr }</td>");
                        }
                        sbTemp.AppendLine($@" </tr>");
                    }
                    sbTemp.AppendLine($@" <tr>");
                    sbTemp.AppendLine($@" <td class=""topRightBorder"">{ TaskRunnerServiceRes.StartTide }<br />{ TaskRunnerServiceRes.EndTide }</td>");
                    foreach (MWQMRunModel mwqmRunModel in mwqmRunModelList2)
                    {
                        string StartTide = GetTideInitial(mwqmRunModel.Tide_Start);
                        string EndTide = GetTideInitial(mwqmRunModel.Tide_End);
                        sbTemp.AppendLine($@" <td class=""topRightBorder"">{ StartTide }<br />{ EndTide }</td>");
                    }
                    sbTemp.AppendLine($@" </tr>");

                    sbTemp.AppendLine($@" <tr>");
                    sbTemp.AppendLine($@" <td class=""topRightBorder"">{ TaskRunnerServiceRes.Rain }(mm)<br />{ TaskRunnerServiceRes.Minus1Day }<br />{ TaskRunnerServiceRes.Minus2Day }<br />{ TaskRunnerServiceRes.Minus3Day }<br />{ TaskRunnerServiceRes.Minus4Day }<br />{ TaskRunnerServiceRes.Minus5Day }</td>");
                    foreach (MWQMRunModel mwqmRunModel in mwqmRunModelList2)
                    {
                        string RainDay1 = mwqmRunModel.RainDay1_mm != null ? ((double)mwqmRunModel.RainDay1_mm).ToString("F0") : "--";
                        string RainDay2 = mwqmRunModel.RainDay2_mm != null ? ((double)mwqmRunModel.RainDay2_mm).ToString("F0") : "--";
                        string RainDay3 = mwqmRunModel.RainDay3_mm != null ? ((double)mwqmRunModel.RainDay3_mm).ToString("F0") : "--";
                        string RainDay4 = mwqmRunModel.RainDay4_mm != null ? ((double)mwqmRunModel.RainDay4_mm).ToString("F0") : "--";
                        string RainDay5 = mwqmRunModel.RainDay5_mm != null ? ((double)mwqmRunModel.RainDay5_mm).ToString("F0") : "--";
                        sbTemp.AppendLine($@" <td class=""topRightBorder"">&nbsp;<br />{ RainDay1 }<br />{ RainDay2 }<br />{ RainDay3 }<br />{ RainDay4 }<br />{ RainDay5 }</td>");
                    }
                    sbTemp.AppendLine($@" </tr>");

                    sbTemp.AppendLine($@" </table>");

                    sbTemp.AppendLine(@"<p>|||PAGE_BREAK|||</p>");

                    skip += take;
                }
                else
                {
                    HasData = false;
                }
            }

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 60);


            skip = 0;
            take = 15;
            HasData = true;
            countRun = 0;
            while (HasData)
            {
                countRun += 1;

                if (countRun > 2)
                {
                    break;
                }

                List<MWQMRunModel> mwqmRunModelList2 = mwqmRunModelList.Where(c => c.RunSampleType == SampleTypeEnum.Routine).OrderByDescending(c => c.DateTime_Local).Skip(skip).Take(take).ToList();
                if (mwqmRunModelList2.Count > 0)
                {
                    sbTemp.AppendLine($@" <h4>{ TaskRunnerServiceRes.ActiveMWQMSites }&nbsp;&nbsp;{ TaskRunnerServiceRes.Salinity }&nbsp;&nbsp;&nbsp;({ TaskRunnerServiceRes.Routine })&nbsp;&nbsp;&nbsp;{ mwqmRunModelList2[0].DateTime_Local.ToString("yyyy MMMM dd") } { TaskRunnerServiceRes.To } { mwqmRunModelList2[mwqmRunModelList2.Count - 1].DateTime_Local.ToString("yyyy MMMM dd") }</h4>");
                    sbTemp.AppendLine($@" <table class=""FCSalTempDataTableClass"">");
                    sbTemp.AppendLine($@" <tr>");
                    sbTemp.AppendLine($@" <th class=""rightBottomBorder"">{ TaskRunnerServiceRes.Site }</th>");
                    foreach (MWQMRunModel mwqmRunModel in mwqmRunModelList2)
                    {
                        sbTemp.AppendLine($@" <th class=""bottomBorder"">{ mwqmRunModel.DateTime_Local.ToString("yyyy") }<br />{ mwqmRunModel.DateTime_Local.ToString("MMM") }<br />{ mwqmRunModel.DateTime_Local.ToString("dd") }</th>");
                    }
                    sbTemp.AppendLine($@" </tr>");

                    foreach (MWQMSiteModel mwqmSiteModel in mwqmSiteModelList2)
                    {
                        sbTemp.AppendLine($@" <tr>");
                        sbTemp.AppendLine($@" <td class=""rightBorder"">{ mwqmSiteModel.MWQMSiteTVText }</td>");
                        foreach (MWQMRunModel mwqmRunModel in mwqmRunModelList2)
                        {
                            float? value = (float?)(from s in mwqmSampleModelList
                                            where s.MWQMRunTVItemID == mwqmRunModel.MWQMRunTVItemID
                                            && s.MWQMSiteTVItemID == mwqmSiteModel.MWQMSiteTVItemID
                                            select s.Salinity_PPT).FirstOrDefault();

                            string valueStr = value != null ? (value == 1 ? "< 2" : ((float)value).ToString("F0")) : "--";
                            sbTemp.AppendLine($@" <td>{ valueStr }</td>");
                        }
                        sbTemp.AppendLine($@" </tr>");
                    }
                    sbTemp.AppendLine($@" <tr>");
                    sbTemp.AppendLine($@" <td class=""topRightBorder"">{ TaskRunnerServiceRes.StartTide }<br />{ TaskRunnerServiceRes.EndTide }</td>");
                    foreach (MWQMRunModel mwqmRunModel in mwqmRunModelList2)
                    {
                        string StartTide = GetTideInitial(mwqmRunModel.Tide_Start);
                        string EndTide = GetTideInitial(mwqmRunModel.Tide_End);
                        sbTemp.AppendLine($@" <td class=""topRightBorder"">{ StartTide }<br />{ EndTide }</td>");
                    }
                    sbTemp.AppendLine($@" </tr>");

                    sbTemp.AppendLine($@" <tr>");
                    sbTemp.AppendLine($@" <td class=""topRightBorder"">{ TaskRunnerServiceRes.Rain }(mm)<br />{ TaskRunnerServiceRes.Minus1Day }<br />{ TaskRunnerServiceRes.Minus2Day }<br />{ TaskRunnerServiceRes.Minus3Day }<br />{ TaskRunnerServiceRes.Minus4Day }<br />{ TaskRunnerServiceRes.Minus5Day }</td>");
                    foreach (MWQMRunModel mwqmRunModel in mwqmRunModelList2)
                    {
                        string RainDay1 = mwqmRunModel.RainDay1_mm != null ? ((double)mwqmRunModel.RainDay1_mm).ToString("F0") : "--";
                        string RainDay2 = mwqmRunModel.RainDay2_mm != null ? ((double)mwqmRunModel.RainDay2_mm).ToString("F0") : "--";
                        string RainDay3 = mwqmRunModel.RainDay3_mm != null ? ((double)mwqmRunModel.RainDay3_mm).ToString("F0") : "--";
                        string RainDay4 = mwqmRunModel.RainDay4_mm != null ? ((double)mwqmRunModel.RainDay4_mm).ToString("F0") : "--";
                        string RainDay5 = mwqmRunModel.RainDay5_mm != null ? ((double)mwqmRunModel.RainDay5_mm).ToString("F0") : "--";
                        sbTemp.AppendLine($@" <td class=""topRightBorder"">&nbsp;<br />{ RainDay1 }<br />{ RainDay2 }<br />{ RainDay3 }<br />{ RainDay4 }<br />{ RainDay5 }</td>");
                    }
                    sbTemp.AppendLine($@" </tr>");

                    sbTemp.AppendLine($@" </table>");

                    sbTemp.AppendLine(@"<p>|||PAGE_BREAK|||</p>");

                    skip += take;
                }
                else
                {
                    HasData = false;
                }
            }

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 70);


            skip = 0;
            take = 15;
            HasData = true;
            countRun = 0;
            while (HasData)
            {
                countRun += 1;

                if (countRun > 2)
                {
                    break;
                }

                List<MWQMRunModel> mwqmRunModelList2 = mwqmRunModelList.Where(c => c.RunSampleType == SampleTypeEnum.Routine).OrderByDescending(c => c.DateTime_Local).Skip(skip).Take(take).ToList();
                if (mwqmRunModelList2.Count > 0)
                {
                    sbTemp.AppendLine($@" <h4>{ TaskRunnerServiceRes.ActiveMWQMSites }&nbsp;&nbsp;{ TaskRunnerServiceRes.Temperature }&nbsp;&nbsp;&nbsp;({ TaskRunnerServiceRes.Routine })&nbsp;&nbsp;&nbsp;{ mwqmRunModelList2[0].DateTime_Local.ToString("yyyy MMMM dd") } { TaskRunnerServiceRes.To } { mwqmRunModelList2[mwqmRunModelList2.Count - 1].DateTime_Local.ToString("yyyy MMMM dd") }</h4>");
                    sbTemp.AppendLine($@" <table class=""FCSalTempDataTableClass"">");
                    sbTemp.AppendLine($@" <tr>");
                    sbTemp.AppendLine($@" <th class=""rightBottomBorder"">{ TaskRunnerServiceRes.Site }</th>");
                    foreach (MWQMRunModel mwqmRunModel in mwqmRunModelList2)
                    {
                        sbTemp.AppendLine($@" <th class=""bottomBorder"">{ mwqmRunModel.DateTime_Local.ToString("yyyy") }<br />{ mwqmRunModel.DateTime_Local.ToString("MMM") }<br />{ mwqmRunModel.DateTime_Local.ToString("dd") }</th>");
                    }
                    sbTemp.AppendLine($@" </tr>");

                    foreach (MWQMSiteModel mwqmSiteModel in mwqmSiteModelList2)
                    {
                        sbTemp.AppendLine($@" <tr>");
                        sbTemp.AppendLine($@" <td class=""rightBorder"">{ mwqmSiteModel.MWQMSiteTVText }</td>");
                        foreach (MWQMRunModel mwqmRunModel in mwqmRunModelList2)
                        {
                            float? value = (float?)(from s in mwqmSampleModelList
                                            where s.MWQMRunTVItemID == mwqmRunModel.MWQMRunTVItemID
                                            && s.MWQMSiteTVItemID == mwqmSiteModel.MWQMSiteTVItemID
                                            select s.WaterTemp_C).FirstOrDefault();

                            string valueStr = value != null ? (value == 1 ? "< 2" : ((float)value).ToString("F0")) : "--";
                            sbTemp.AppendLine($@" <td>{ valueStr }</td>");
                        }
                        sbTemp.AppendLine($@" </tr>");
                    }
                    sbTemp.AppendLine($@" <tr>");
                    sbTemp.AppendLine($@" <td class=""topRightBorder"">{ TaskRunnerServiceRes.StartTide }<br />{ TaskRunnerServiceRes.EndTide }</td>");
                    foreach (MWQMRunModel mwqmRunModel in mwqmRunModelList2)
                    {
                        string StartTide = GetTideInitial(mwqmRunModel.Tide_Start);
                        string EndTide = GetTideInitial(mwqmRunModel.Tide_End);
                        sbTemp.AppendLine($@" <td class=""topRightBorder"">{ StartTide }<br />{ EndTide }</td>");
                    }
                    sbTemp.AppendLine($@" </tr>");

                    sbTemp.AppendLine($@" <tr>");
                    sbTemp.AppendLine($@" <td class=""topRightBorder"">{ TaskRunnerServiceRes.Rain }(mm)<br />{ TaskRunnerServiceRes.Minus1Day }<br />{ TaskRunnerServiceRes.Minus2Day }<br />{ TaskRunnerServiceRes.Minus3Day }<br />{ TaskRunnerServiceRes.Minus4Day }<br />{ TaskRunnerServiceRes.Minus5Day }</td>");
                    foreach (MWQMRunModel mwqmRunModel in mwqmRunModelList2)
                    {
                        string RainDay1 = mwqmRunModel.RainDay1_mm != null ? ((double)mwqmRunModel.RainDay1_mm).ToString("F0") : "--";
                        string RainDay2 = mwqmRunModel.RainDay2_mm != null ? ((double)mwqmRunModel.RainDay2_mm).ToString("F0") : "--";
                        string RainDay3 = mwqmRunModel.RainDay3_mm != null ? ((double)mwqmRunModel.RainDay3_mm).ToString("F0") : "--";
                        string RainDay4 = mwqmRunModel.RainDay4_mm != null ? ((double)mwqmRunModel.RainDay4_mm).ToString("F0") : "--";
                        string RainDay5 = mwqmRunModel.RainDay5_mm != null ? ((double)mwqmRunModel.RainDay5_mm).ToString("F0") : "--";
                        sbTemp.AppendLine($@" <td class=""topRightBorder"">&nbsp;<br />{ RainDay1 }<br />{ RainDay2 }<br />{ RainDay3 }<br />{ RainDay4 }<br />{ RainDay5 }</td>");
                    }
                    sbTemp.AppendLine($@" </tr>");

                    sbTemp.AppendLine($@" </table>");

                    sbTemp.AppendLine(@"<p>|||PAGE_BREAK|||</p>");

                    skip += take;
                }
                else
                {
                    HasData = false;
                }
            }

            sbTemp.AppendLine(@"<p>|||PAGE_BREAK|||</p>");

            if (!GetBottomHTML())
            {
                return false;
            }

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 80);

            return true;
        }

        // for testing only can comment out when test is completed
        public bool PublicGenerateHTMLSUBSECTOR_MWQM_SITES(StringBuilder sbTemp)
        {
            bool retBool = GenerateHTMLSUBSECTOR_MWQM_SITES(sbTemp);

            StreamWriter sw = fi.CreateText();
            sw.Write(sbTemp.ToString());
            sw.Flush();
            sw.Close();

            return retBool;
        }     
    }
}
