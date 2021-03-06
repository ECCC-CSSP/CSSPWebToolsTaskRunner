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
using CSSPEnumsDLL.Services;
using CSSPModelsDLL.Models;
using System.Windows.Forms;
using System.Web.Helpers;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        private bool GenerateHTMLSubsectorPollutionSourceSites(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            string NotUsed = "";
            int TVItemID = 0;

            Random random = new Random();
            string FileNameExtra = "";
            for (int i = 0; i < 10; i++)
            {
                FileNameExtra = FileNameExtra + (char)random.Next(97, 122);
            }

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 3);

            if (!GetTopHTML(sbHTML))
            {
                return false;
            }

            List<string> ParamValueList = parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            if (!int.TryParse(GetParameters("TVItemID", ParamValueList), out TVItemID))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind__, TaskRunnerServiceRes.Parameter, TaskRunnerServiceRes.TVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind__", TaskRunnerServiceRes.Parameter, TaskRunnerServiceRes.TVItemID);
                return false;
            }

            TVItemModel tvItemModelSubsector = _TVItemService.GetTVItemModelWithTVItemIDDB(TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelSubsector.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                return false;
            }

            string ServerPath = _TVFileService.GetServerFilePath(tvItemModelSubsector.TVItemID);

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 5);

            List<TVItemModel> tvItemModelListPolSourceSite = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.PolSourceSite).Where(c => c.IsActive == true).ToList();
            List<PolSourceSiteModel> polSourceSiteModelList = _PolSourceSiteService.GetPolSourceSiteModelListWithSubsectorTVItemIDDB(TVItemID).OrderBy(c => c.Site).ToList();
            List<PolSourceObservationModel> polSourceObservationModelList = _PolSourceObservationService.GetPolSourceObservationModelListWithSubsectorTVItemIDDB(TVItemID);
            List<PolSourceObservationIssueModel> polSourceObservationIssueModelList = _PolSourceObservationIssueService.GetPolSourceObservationIssueModelListWithSubsectorTVItemIDDB(TVItemID);
            List<PolSourceObsInfoEnumTextAndID> polSourceObsInfoEnumTextAndIDList = new List<PolSourceObsInfoEnumTextAndID>();
            List<PolSourceObsInfoEnumTextAndID> polSourceObsInfoEnumTextAndIDListLandBasePolSourceType = new List<PolSourceObsInfoEnumTextAndID>();
            List<PolSourceObsInfoEnumTextAndID> polSourceObsInfoEnumTextAndIDListWaterBasePolSourceType = new List<PolSourceObsInfoEnumTextAndID>();

            List<PolSourceObsInfoEnumTextAndID> polSourceObsInfoEnumTextAndIDListLandBaseAgriculture = new List<PolSourceObsInfoEnumTextAndID>();
            List<PolSourceObsInfoEnumTextAndID> polSourceObsInfoEnumTextAndIDListLandBaseForested = new List<PolSourceObsInfoEnumTextAndID>();
            List<PolSourceObsInfoEnumTextAndID> polSourceObsInfoEnumTextAndIDListLandBaseIndustry = new List<PolSourceObsInfoEnumTextAndID>();
            List<PolSourceObsInfoEnumTextAndID> polSourceObsInfoEnumTextAndIDListLandBaseMarine = new List<PolSourceObsInfoEnumTextAndID>();
            List<PolSourceObsInfoEnumTextAndID> polSourceObsInfoEnumTextAndIDListLandBaseRecreational = new List<PolSourceObsInfoEnumTextAndID>();
            List<PolSourceObsInfoEnumTextAndID> polSourceObsInfoEnumTextAndIDListLandBaseUrban = new List<PolSourceObsInfoEnumTextAndID>();

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 10);

            foreach (int id in Enum.GetValues(typeof(PolSourceObsInfoEnum)))
            {
                if (id == 0)
                    continue;

                polSourceObsInfoEnumTextAndIDList.Add(new PolSourceObsInfoEnumTextAndID() { Text = _BaseEnumService.GetEnumText_PolSourceObsInfoEnum((PolSourceObsInfoEnum)id), ID = id });

                if (id.ToString().StartsWith("105") && !id.ToString().EndsWith("00"))
                {
                    polSourceObsInfoEnumTextAndIDListLandBasePolSourceType.Add(new PolSourceObsInfoEnumTextAndID() { Text = _BaseEnumService.GetEnumText_PolSourceObsInfoEnum((PolSourceObsInfoEnum)id), ID = id });
                }

                if (id.ToString().StartsWith("152") && !id.ToString().EndsWith("00"))
                {
                    polSourceObsInfoEnumTextAndIDListWaterBasePolSourceType.Add(new PolSourceObsInfoEnumTextAndID() { Text = _BaseEnumService.GetEnumText_PolSourceObsInfoEnum((PolSourceObsInfoEnum)id), ID = id });
                }

                if (id.ToString().StartsWith("106") && !id.ToString().EndsWith("00"))
                {
                    polSourceObsInfoEnumTextAndIDListLandBaseAgriculture.Add(new PolSourceObsInfoEnumTextAndID() { Text = _BaseEnumService.GetEnumText_PolSourceObsInfoEnum((PolSourceObsInfoEnum)id), ID = id });
                }

                if (id.ToString().StartsWith("111") && !id.ToString().EndsWith("00"))
                {
                    polSourceObsInfoEnumTextAndIDListLandBaseForested.Add(new PolSourceObsInfoEnumTextAndID() { Text = _BaseEnumService.GetEnumText_PolSourceObsInfoEnum((PolSourceObsInfoEnum)id), ID = id });
                }

                if (id.ToString().StartsWith("112") && !id.ToString().EndsWith("00"))
                {
                    polSourceObsInfoEnumTextAndIDListLandBaseIndustry.Add(new PolSourceObsInfoEnumTextAndID() { Text = _BaseEnumService.GetEnumText_PolSourceObsInfoEnum((PolSourceObsInfoEnum)id), ID = id });
                }

                if (id.ToString().StartsWith("114") && !id.ToString().EndsWith("00"))
                {
                    polSourceObsInfoEnumTextAndIDListLandBaseMarine.Add(new PolSourceObsInfoEnumTextAndID() { Text = _BaseEnumService.GetEnumText_PolSourceObsInfoEnum((PolSourceObsInfoEnum)id), ID = id });
                }

                if (id.ToString().StartsWith("117") && !id.ToString().EndsWith("00"))
                {
                    polSourceObsInfoEnumTextAndIDListLandBaseRecreational.Add(new PolSourceObsInfoEnumTextAndID() { Text = _BaseEnumService.GetEnumText_PolSourceObsInfoEnum((PolSourceObsInfoEnum)id), ID = id });
                }

                if (id.ToString().StartsWith("121") && !id.ToString().EndsWith("00"))
                {
                    polSourceObsInfoEnumTextAndIDListLandBaseUrban.Add(new PolSourceObsInfoEnumTextAndID() { Text = _BaseEnumService.GetEnumText_PolSourceObsInfoEnum((PolSourceObsInfoEnum)id), ID = id });
                }
            }

            polSourceObsInfoEnumTextAndIDListLandBasePolSourceType = polSourceObsInfoEnumTextAndIDListLandBasePolSourceType.OrderBy(c => c.Text).ToList();
            polSourceObsInfoEnumTextAndIDListWaterBasePolSourceType = polSourceObsInfoEnumTextAndIDListWaterBasePolSourceType.OrderBy(c => c.Text).ToList();
            polSourceObsInfoEnumTextAndIDListLandBaseAgriculture = polSourceObsInfoEnumTextAndIDListLandBaseAgriculture.OrderBy(c => c.Text).ToList();
            polSourceObsInfoEnumTextAndIDListLandBaseForested = polSourceObsInfoEnumTextAndIDListLandBaseForested.OrderBy(c => c.Text).ToList();
            polSourceObsInfoEnumTextAndIDListLandBaseIndustry = polSourceObsInfoEnumTextAndIDListLandBaseIndustry.OrderBy(c => c.Text).ToList();
            polSourceObsInfoEnumTextAndIDListLandBaseMarine = polSourceObsInfoEnumTextAndIDListLandBaseMarine.OrderBy(c => c.Text).ToList();
            polSourceObsInfoEnumTextAndIDListLandBaseRecreational = polSourceObsInfoEnumTextAndIDListLandBaseRecreational.OrderBy(c => c.Text).ToList();
            polSourceObsInfoEnumTextAndIDListLandBaseUrban = polSourceObsInfoEnumTextAndIDListLandBaseUrban.OrderBy(c => c.Text).ToList();

            // ---------------------------------------------------------------------------------------------------------------
            // Land Base graphic showing the number of polsource issues of different pollution source type (Agriculture, Forested, ...)
            // ---------------------------------------------------------------------------------------------------------------

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 15);

            List<string> PolSourceTypeList = new List<string>();
            List<int> CountOfPolSourceType = new List<int>();

            foreach (PolSourceObsInfoEnumTextAndID polSourceObsInfoEnumTextAndID in polSourceObsInfoEnumTextAndIDListLandBasePolSourceType)
            {
                List<PolSourceSiteModel> polSourceSiteModelListActive = (from pss in polSourceSiteModelList
                                                                         from t in tvItemModelListPolSourceSite
                                                                         where pss.PolSourceSiteTVItemID == t.TVItemID
                                                                         && t.IsActive == true
                                                                         select pss).ToList();

                List<PolSourceObservationModel> polSourceObservationModelLastOfActive = (from pss in polSourceSiteModelListActive
                                                                                         let psoLast = (from pso in polSourceObservationModelList
                                                                                                        where pss.PolSourceSiteID == pso.PolSourceSiteID
                                                                                                        orderby pso.ObservationDate_Local descending
                                                                                                        select pso).FirstOrDefault()
                                                                                         select psoLast).ToList();

                int count = (from pso in polSourceObservationModelLastOfActive
                             from psoi in polSourceObservationIssueModelList
                             where pso.PolSourceObservationID == psoi.PolSourceObservationID
                             && psoi.ObservationInfo.Contains("," + polSourceObsInfoEnumTextAndID.ID.ToString() + ",")
                             select psoi.PolSourceObservationIssueID).Count();

                PolSourceTypeList.Add(polSourceObsInfoEnumTextAndID.Text);
                CountOfPolSourceType.Add(count);
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

            series.XValues = PolSourceTypeList.ToArray();
            series.Values = CountOfPolSourceType.ToArray();

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

            chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = TaskRunnerServiceRes.NumberOfIssuesByLandBasePollutionType;

            // need to save the file with a unique name under the TVItemID
            FileInfo fiImageLand = new FileInfo(fi.DirectoryName + @"\LandBaseStat" + FileNameExtra + ".png");

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

            chartPage.Export(fiImageLand.FullName, "PNG", false);

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 25);

            // ---------------------------------------------------------------------------------------------------------------
            // Water Base graphic showing the number of polsource issues of different pollution source type (Aquaculture, Seaport ...)
            // ---------------------------------------------------------------------------------------------------------------

            PolSourceTypeList = new List<string>();
            CountOfPolSourceType = new List<int>();

            foreach (PolSourceObsInfoEnumTextAndID polSourceObsInfoEnumTextAndID in polSourceObsInfoEnumTextAndIDListWaterBasePolSourceType)
            {
                List<PolSourceSiteModel> polSourceSiteModelListActive = (from pss in polSourceSiteModelList
                                                                         from t in tvItemModelListPolSourceSite
                                                                         where pss.PolSourceSiteTVItemID == t.TVItemID
                                                                         && t.IsActive == true
                                                                         select pss).ToList();

                List<PolSourceObservationModel> polSourceObservationModelLastOfActive = (from pss in polSourceSiteModelListActive
                                                                                         let psoLast = (from pso in polSourceObservationModelList
                                                                                                        where pss.PolSourceSiteID == pso.PolSourceSiteID
                                                                                                        orderby pso.ObservationDate_Local descending
                                                                                                        select pso).FirstOrDefault()
                                                                                         select psoLast).ToList();

                int count = (from pso in polSourceObservationModelLastOfActive
                             from psoi in polSourceObservationIssueModelList
                             where pso.PolSourceObservationID == psoi.PolSourceObservationID
                             && psoi.ObservationInfo.Contains("," + polSourceObsInfoEnumTextAndID.ID.ToString() + ",")
                             select psoi.PolSourceObservationIssueID).Count();

                PolSourceTypeList.Add(polSourceObsInfoEnumTextAndID.Text);
                CountOfPolSourceType.Add(count);
            }

            chart = xlCharts.Add(100, 100, 600, 200);
            chartPage = chart.Chart;

            chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;

            seriesCollection = chartPage.SeriesCollection();
            series = seriesCollection.NewSeries();

            series.XValues = PolSourceTypeList.ToArray();
            series.Values = CountOfPolSourceType.ToArray();

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

            chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = TaskRunnerServiceRes.NumberOfIssuesByWaterBasePollutionType;

            // need to save the file with a unique name under the TVItemID
            FileInfo fiImageWater = new FileInfo(fi.DirectoryName + @"\WaterBaseStat" + FileNameExtra + ".png");

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

            chartPage.Export(fiImageWater.FullName, "PNG", false);

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 35);

            // ---------------------------------------------------------------------------------------------------------------
            // Agriculture graphic showing the number of polsource issues of different pollution source type
            // ---------------------------------------------------------------------------------------------------------------

            PolSourceTypeList = new List<string>();
            CountOfPolSourceType = new List<int>();

            foreach (PolSourceObsInfoEnumTextAndID polSourceObsInfoEnumTextAndID in polSourceObsInfoEnumTextAndIDListLandBaseAgriculture)
            {
                List<PolSourceSiteModel> polSourceSiteModelListActive = (from pss in polSourceSiteModelList
                                                                         from t in tvItemModelListPolSourceSite
                                                                         where pss.PolSourceSiteTVItemID == t.TVItemID
                                                                         && t.IsActive == true
                                                                         select pss).ToList();

                List<PolSourceObservationModel> polSourceObservationModelLastOfActive = (from pss in polSourceSiteModelListActive
                                                                                         let psoLast = (from pso in polSourceObservationModelList
                                                                                                        where pss.PolSourceSiteID == pso.PolSourceSiteID
                                                                                                        orderby pso.ObservationDate_Local descending
                                                                                                        select pso).FirstOrDefault()
                                                                                         select psoLast).ToList();

                int count = (from pso in polSourceObservationModelLastOfActive
                             from psoi in polSourceObservationIssueModelList
                             where pso.PolSourceObservationID == psoi.PolSourceObservationID
                             && psoi.ObservationInfo.Contains("," + polSourceObsInfoEnumTextAndID.ID.ToString() + ",")
                             select psoi.PolSourceObservationIssueID).Count();

                PolSourceTypeList.Add(polSourceObsInfoEnumTextAndID.Text);
                CountOfPolSourceType.Add(count);
            }

            chart = xlCharts.Add(100, 100, 600, 200);
            chartPage = chart.Chart;

            chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;

            seriesCollection = chartPage.SeriesCollection();
            series = seriesCollection.NewSeries();

            series.XValues = PolSourceTypeList.ToArray();
            series.Values = CountOfPolSourceType.ToArray();

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

            chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = TaskRunnerServiceRes.NumberOfIssuesByWaterBasePollutionType;

            // need to save the file with a unique name under the TVItemID
            FileInfo fiImageAgriculture = new FileInfo(fi.DirectoryName + @"\AgricultureStat" + FileNameExtra + ".png");

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

            chartPage.Export(fiImageAgriculture.FullName, "PNG", false);

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 45);

            // ---------------------------------------------------------------------------------------------------------------
            // Forested graphic showing the number of polsource issues of different pollution source type
            // ---------------------------------------------------------------------------------------------------------------

            PolSourceTypeList = new List<string>();
            CountOfPolSourceType = new List<int>();

            foreach (PolSourceObsInfoEnumTextAndID polSourceObsInfoEnumTextAndID in polSourceObsInfoEnumTextAndIDListLandBaseForested)
            {
                List<PolSourceSiteModel> polSourceSiteModelListActive = (from pss in polSourceSiteModelList
                                                                         from t in tvItemModelListPolSourceSite
                                                                         where pss.PolSourceSiteTVItemID == t.TVItemID
                                                                         && t.IsActive == true
                                                                         select pss).ToList();

                List<PolSourceObservationModel> polSourceObservationModelLastOfActive = (from pss in polSourceSiteModelListActive
                                                                                         let psoLast = (from pso in polSourceObservationModelList
                                                                                                        where pss.PolSourceSiteID == pso.PolSourceSiteID
                                                                                                        orderby pso.ObservationDate_Local descending
                                                                                                        select pso).FirstOrDefault()
                                                                                         select psoLast).ToList();

                int count = (from pso in polSourceObservationModelLastOfActive
                             from psoi in polSourceObservationIssueModelList
                             where pso.PolSourceObservationID == psoi.PolSourceObservationID
                             && psoi.ObservationInfo.Contains("," + polSourceObsInfoEnumTextAndID.ID.ToString() + ",")
                             select psoi.PolSourceObservationIssueID).Count();

                PolSourceTypeList.Add(polSourceObsInfoEnumTextAndID.Text);
                CountOfPolSourceType.Add(count);
            }

            chart = xlCharts.Add(100, 100, 600, 200);
            chartPage = chart.Chart;

            chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;

            seriesCollection = chartPage.SeriesCollection();
            series = seriesCollection.NewSeries();

            series.XValues = PolSourceTypeList.ToArray();
            series.Values = CountOfPolSourceType.ToArray();

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

            chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = TaskRunnerServiceRes.NumberOfIssuesByWaterBasePollutionType;

            // need to save the file with a unique name under the TVItemID
            FileInfo fiImageForested = new FileInfo(fi.DirectoryName + @"\ForestedStat" + FileNameExtra + ".png");

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

            chartPage.Export(fiImageForested.FullName, "PNG", false);

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 55);

            // ---------------------------------------------------------------------------------------------------------------
            // Industry graphic showing the number of polsource issues of different pollution source type
            // ---------------------------------------------------------------------------------------------------------------

            PolSourceTypeList = new List<string>();
            CountOfPolSourceType = new List<int>();

            foreach (PolSourceObsInfoEnumTextAndID polSourceObsInfoEnumTextAndID in polSourceObsInfoEnumTextAndIDListLandBaseIndustry)
            {
                List<PolSourceSiteModel> polSourceSiteModelListActive = (from pss in polSourceSiteModelList
                                                                         from t in tvItemModelListPolSourceSite
                                                                         where pss.PolSourceSiteTVItemID == t.TVItemID
                                                                         && t.IsActive == true
                                                                         select pss).ToList();

                List<PolSourceObservationModel> polSourceObservationModelLastOfActive = (from pss in polSourceSiteModelListActive
                                                                                         let psoLast = (from pso in polSourceObservationModelList
                                                                                                        where pss.PolSourceSiteID == pso.PolSourceSiteID
                                                                                                        orderby pso.ObservationDate_Local descending
                                                                                                        select pso).FirstOrDefault()
                                                                                         select psoLast).ToList();

                int count = (from pso in polSourceObservationModelLastOfActive
                             from psoi in polSourceObservationIssueModelList
                             where pso.PolSourceObservationID == psoi.PolSourceObservationID
                             && psoi.ObservationInfo.Contains("," + polSourceObsInfoEnumTextAndID.ID.ToString() + ",")
                             select psoi.PolSourceObservationIssueID).Count();

                PolSourceTypeList.Add(polSourceObsInfoEnumTextAndID.Text);
                CountOfPolSourceType.Add(count);
            }

            chart = xlCharts.Add(100, 100, 600, 200);
            chartPage = chart.Chart;

            chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;

            seriesCollection = chartPage.SeriesCollection();
            series = seriesCollection.NewSeries();

            series.XValues = PolSourceTypeList.ToArray();
            series.Values = CountOfPolSourceType.ToArray();

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

            chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = TaskRunnerServiceRes.NumberOfIssuesByWaterBasePollutionType;

            // need to save the file with a unique name under the TVItemID
            FileInfo fiImageIndustry = new FileInfo(fi.DirectoryName + @"\IndustryStat" + FileNameExtra + ".png");

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

            chartPage.Export(fiImageIndustry.FullName, "PNG", false);

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 65);

            // ---------------------------------------------------------------------------------------------------------------
            // Marine graphic showing the number of polsource issues of different pollution source type
            // ---------------------------------------------------------------------------------------------------------------

            PolSourceTypeList = new List<string>();
            CountOfPolSourceType = new List<int>();

            foreach (PolSourceObsInfoEnumTextAndID polSourceObsInfoEnumTextAndID in polSourceObsInfoEnumTextAndIDListLandBaseMarine)
            {
                List<PolSourceSiteModel> polSourceSiteModelListActive = (from pss in polSourceSiteModelList
                                                                         from t in tvItemModelListPolSourceSite
                                                                         where pss.PolSourceSiteTVItemID == t.TVItemID
                                                                         && t.IsActive == true
                                                                         select pss).ToList();

                List<PolSourceObservationModel> polSourceObservationModelLastOfActive = (from pss in polSourceSiteModelListActive
                                                                                         let psoLast = (from pso in polSourceObservationModelList
                                                                                                        where pss.PolSourceSiteID == pso.PolSourceSiteID
                                                                                                        orderby pso.ObservationDate_Local descending
                                                                                                        select pso).FirstOrDefault()
                                                                                         select psoLast).ToList();

                int count = (from pso in polSourceObservationModelLastOfActive
                             from psoi in polSourceObservationIssueModelList
                             where pso.PolSourceObservationID == psoi.PolSourceObservationID
                             && psoi.ObservationInfo.Contains("," + polSourceObsInfoEnumTextAndID.ID.ToString() + ",")
                             select psoi.PolSourceObservationIssueID).Count();

                PolSourceTypeList.Add(polSourceObsInfoEnumTextAndID.Text);
                CountOfPolSourceType.Add(count);
            }

            chart = xlCharts.Add(100, 100, 600, 200);
            chartPage = chart.Chart;

            chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;

            seriesCollection = chartPage.SeriesCollection();
            series = seriesCollection.NewSeries();

            series.XValues = PolSourceTypeList.ToArray();
            series.Values = CountOfPolSourceType.ToArray();

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

            chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = TaskRunnerServiceRes.NumberOfIssuesByWaterBasePollutionType;

            // need to save the file with a unique name under the TVItemID
            FileInfo fiImageMarine = new FileInfo(fi.DirectoryName + @"\MarineStat" + FileNameExtra + ".png");

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

            chartPage.Export(fiImageMarine.FullName, "PNG", false);

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 75);

            // ---------------------------------------------------------------------------------------------------------------
            // Recreational graphic showing the number of polsource issues of different pollution source type
            // ---------------------------------------------------------------------------------------------------------------

            PolSourceTypeList = new List<string>();
            CountOfPolSourceType = new List<int>();

            foreach (PolSourceObsInfoEnumTextAndID polSourceObsInfoEnumTextAndID in polSourceObsInfoEnumTextAndIDListLandBaseRecreational)
            {
                List<PolSourceSiteModel> polSourceSiteModelListActive = (from pss in polSourceSiteModelList
                                                                         from t in tvItemModelListPolSourceSite
                                                                         where pss.PolSourceSiteTVItemID == t.TVItemID
                                                                         && t.IsActive == true
                                                                         select pss).ToList();

                List<PolSourceObservationModel> polSourceObservationModelLastOfActive = (from pss in polSourceSiteModelListActive
                                                                                         let psoLast = (from pso in polSourceObservationModelList
                                                                                                        where pss.PolSourceSiteID == pso.PolSourceSiteID
                                                                                                        orderby pso.ObservationDate_Local descending
                                                                                                        select pso).FirstOrDefault()
                                                                                         select psoLast).ToList();

                int count = (from pso in polSourceObservationModelLastOfActive
                             from psoi in polSourceObservationIssueModelList
                             where pso.PolSourceObservationID == psoi.PolSourceObservationID
                             && psoi.ObservationInfo.Contains("," + polSourceObsInfoEnumTextAndID.ID.ToString() + ",")
                             select psoi.PolSourceObservationIssueID).Count();

                PolSourceTypeList.Add(polSourceObsInfoEnumTextAndID.Text);
                CountOfPolSourceType.Add(count);
            }

            chart = xlCharts.Add(100, 100, 600, 200);
            chartPage = chart.Chart;

            chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;

            seriesCollection = chartPage.SeriesCollection();
            series = seriesCollection.NewSeries();

            series.XValues = PolSourceTypeList.ToArray();
            series.Values = CountOfPolSourceType.ToArray();

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

            chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = TaskRunnerServiceRes.NumberOfIssuesByWaterBasePollutionType;

            // need to save the file with a unique name under the TVItemID
            FileInfo fiImageRecreational = new FileInfo(fi.DirectoryName + @"\RecreationalStat" + FileNameExtra + ".png");

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

            chartPage.Export(fiImageRecreational.FullName, "PNG", false);

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 85);

            // ---------------------------------------------------------------------------------------------------------------
            // Urban graphic showing the number of polsource issues of different pollution source type
            // ---------------------------------------------------------------------------------------------------------------

            PolSourceTypeList = new List<string>();
            CountOfPolSourceType = new List<int>();

            foreach (PolSourceObsInfoEnumTextAndID polSourceObsInfoEnumTextAndID in polSourceObsInfoEnumTextAndIDListLandBaseUrban)
            {
                List<PolSourceSiteModel> polSourceSiteModelListActive = (from pss in polSourceSiteModelList
                                                                         from t in tvItemModelListPolSourceSite
                                                                         where pss.PolSourceSiteTVItemID == t.TVItemID
                                                                         && t.IsActive == true
                                                                         select pss).ToList();

                List<PolSourceObservationModel> polSourceObservationModelLastOfActive = (from pss in polSourceSiteModelListActive
                                                                                         let psoLast = (from pso in polSourceObservationModelList
                                                                                                        where pss.PolSourceSiteID == pso.PolSourceSiteID
                                                                                                        orderby pso.ObservationDate_Local descending
                                                                                                        select pso).FirstOrDefault()
                                                                                         select psoLast).ToList();

                int count = (from pso in polSourceObservationModelLastOfActive
                             from psoi in polSourceObservationIssueModelList
                             where pso.PolSourceObservationID == psoi.PolSourceObservationID
                             && psoi.ObservationInfo.Contains("," + polSourceObsInfoEnumTextAndID.ID.ToString() + ",")
                             select psoi.PolSourceObservationIssueID).Count();

                PolSourceTypeList.Add(polSourceObsInfoEnumTextAndID.Text);
                CountOfPolSourceType.Add(count);
            }

            chart = xlCharts.Add(100, 100, 600, 200);
            chartPage = chart.Chart;

            chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;

            seriesCollection = chartPage.SeriesCollection();
            series = seriesCollection.NewSeries();

            series.XValues = PolSourceTypeList.ToArray();
            series.Values = CountOfPolSourceType.ToArray();

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

            chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = TaskRunnerServiceRes.NumberOfIssuesByWaterBasePollutionType;

            // need to save the file with a unique name under the TVItemID
            FileInfo fiImageUrban = new FileInfo(fi.DirectoryName + @"\UrbanStat" + FileNameExtra + ".png");

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

            chartPage.Export(fiImageUrban.FullName, "PNG", false);

            if (workbook != null)
            {
                workbook.Close(false);
            }
            if (xlApp != null)
            {
                xlApp.Quit();
            }


            sbHTML.AppendLine($@" <h3>{ TaskRunnerServiceRes.LandBasePollutionSourceSiteObservationAndIssues }</h3>");

            sbHTML.AppendLine($@"<div class=""textAlignCenter"">|||Image|FileName,{ fiImageLand.FullName }|width,400|height,150|||</div>");

            foreach (PolSourceObsInfoEnumTextAndID PolSourceObsInfoEnumTextAndID in polSourceObsInfoEnumTextAndIDListLandBasePolSourceType)
            {
                sbHTML.AppendLine($@" <h4>{ PolSourceObsInfoEnumTextAndID.Text }</h4>");

                switch (PolSourceObsInfoEnumTextAndID.ID)
                {
                    case 10501: // Agriculture
                        {
                            sbHTML.AppendLine($@"<div class=""textAlignCenter"">|||Image|FileName,{ fiImageAgriculture.FullName }|width,400|height,150|||</div>");
                        }
                        break;
                    case 10502: // Forested
                        {
                            sbHTML.AppendLine($@"<div class=""textAlignCenter"">|||Image|FileName,{ fiImageForested.FullName }|width,400|height,150|||</div>");
                        }
                        break;
                    case 10503: // Industry
                        {
                            sbHTML.AppendLine($@"<div class=""textAlignCenter"">|||Image|FileName,{ fiImageIndustry.FullName }|width,400|height,150|||</div>");
                        }
                        break;
                    case 10504: // Marine
                        {
                            sbHTML.AppendLine($@"<div class=""textAlignCenter"">|||Image|FileName,{ fiImageMarine.FullName }|width,400|height,150|||</div>");
                        }
                        break;
                    case 10505: // Recreational
                        {
                            sbHTML.AppendLine($@"<div class=""textAlignCenter"">|||Image|FileName,{ fiImageRecreational.FullName }|width,400|height,150|||</div>");
                        }
                        break;
                    case 10506: // Urban
                        {
                            sbHTML.AppendLine($@"<div class=""textAlignCenter"">|||Image|FileName,{ fiImageUrban.FullName }|width,400|height,150|||</div>");
                        }
                        break;
                    default:
                        break;
                }


                int countPolSourceSiteForPollutionType = 0;
                foreach (PolSourceSiteModel polSourceSiteModel in polSourceSiteModelList)
                {
                    TVItemModel tvItemModel = tvItemModelListPolSourceSite.Where(c => c.TVItemID == polSourceSiteModel.PolSourceSiteTVItemID).FirstOrDefault();
                    if (tvItemModel != null && tvItemModel.IsActive)
                    {
                        if (polSourceSiteModel != null)
                        {
                            PolSourceObservationModel polSourceObservationModel = polSourceObservationModelList.Where(c => c.PolSourceSiteID == polSourceSiteModel.PolSourceSiteID).OrderByDescending(c => c.ObservationDate_Local).FirstOrDefault();
                            if (polSourceObservationModel != null)
                            {
                                List<PolSourceObservationIssueModel> polSourceObservationIssueModelList2 = polSourceObservationIssueModelList.Where(c => c.PolSourceObservationID == polSourceObservationModel.PolSourceObservationID).OrderBy(c => c.Ordinal).ToList();
                                if (polSourceObservationIssueModelList2.Count > 0)
                                {
                                    if (polSourceObservationIssueModelList2[0].ObservationInfo.Contains(((int)PolSourceObsInfoEnum.LandBased).ToString() + ","))
                                    {
                                        if (polSourceObservationModel == null)
                                        {
                                            countPolSourceSiteForPollutionType += 1;

                                            sbHTML.AppendLine($@" <blockquote>");
                                            sbHTML.AppendLine($@" <span><b>{ TaskRunnerServiceRes.Site }:</b> { polSourceSiteModel.Site }</span>");

                                            List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(polSourceSiteModel.PolSourceSiteTVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);
                                            if (mapInfoPointModelList.Count > 0)
                                            {
                                                sbHTML.AppendLine($@" <span>{ TaskRunnerServiceRes.Coord }: { mapInfoPointModelList[0].Lat }, { mapInfoPointModelList[0].Lng }</span>");
                                            }
                                            sbHTML.AppendLine($@"           <br />");
                                            sbHTML.AppendLine($@"           <span>{ TaskRunnerServiceRes.NoObservationForThisPollutionSourceSite }</span>");
                                            sbHTML.AppendLine($@" </blockquote>");
                                        }
                                        else
                                        {
                                            if (polSourceObservationIssueModelList2.Count == 0)
                                            {
                                                countPolSourceSiteForPollutionType += 1;

                                                sbHTML.AppendLine($@" <blockquote>");
                                                sbHTML.AppendLine($@" <span><b>{ TaskRunnerServiceRes.Site }:</b> { polSourceSiteModel.Site }</span>");

                                                List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(polSourceSiteModel.PolSourceSiteTVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);
                                                if (mapInfoPointModelList.Count > 0)
                                                {
                                                    sbHTML.AppendLine($@" <span>{ TaskRunnerServiceRes.Coord }: { mapInfoPointModelList[0].Lat }, { mapInfoPointModelList[0].Lng }</span>");
                                                }
                                                sbHTML.AppendLine($@"           <br />");
                                                sbHTML.AppendLine($@"           <span>{ polSourceObservationModel.Observation_ToBeDeleted }</span>");
                                                sbHTML.AppendLine($@" </blockquote>");
                                            }
                                            else
                                            {
                                                bool IssueOfSourceTypeExist = false;
                                                foreach (PolSourceObservationIssueModel polSourceObservationIssueModel in polSourceObservationIssueModelList2)
                                                {
                                                    if (polSourceObservationIssueModel.ObservationInfo.Contains(PolSourceObsInfoEnumTextAndID.ID.ToString()))
                                                    {
                                                        IssueOfSourceTypeExist = true;
                                                    }
                                                }

                                                if (IssueOfSourceTypeExist)
                                                {
                                                    sbHTML.AppendLine($@" <blockquote>");
                                                    sbHTML.AppendLine($@" <span><b>{ TaskRunnerServiceRes.Site }:</b> { polSourceSiteModel.Site }</span>");

                                                    List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(polSourceSiteModel.PolSourceSiteTVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);
                                                    if (mapInfoPointModelList.Count > 0)
                                                    {
                                                        sbHTML.AppendLine($@" <span><b>{ TaskRunnerServiceRes.Coord }:</b> { mapInfoPointModelList[0].Lat.ToString("F5") } { mapInfoPointModelList[0].Lng.ToString("F5") }</span>");
                                                    }

                                                    int CountIssues = 0;
                                                    foreach (PolSourceObservationIssueModel polSourceObservationIssueModel in polSourceObservationIssueModelList2)
                                                    {
                                                        if (polSourceObservationIssueModel.ObservationInfo.Contains(PolSourceObsInfoEnumTextAndID.ID.ToString()))
                                                        {
                                                            countPolSourceSiteForPollutionType += 1;
                                                            CountIssues += 1;
                                                            List<int> obsInfoList = polSourceObservationIssueModel.ObservationInfo.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Where(c => !string.IsNullOrWhiteSpace(c)).Select(c => int.Parse(c)).ToList();
                                                            if (CountIssues == 1)
                                                            {
                                                                if (obsInfoList.Count > 1)
                                                                {
                                                                    sbHTML.AppendLine($@"           <span><b>{ TaskRunnerServiceRes.Dist }:</b> { polSourceObsInfoEnumTextAndIDList.Where(c => c.ID == (int)obsInfoList[1]).FirstOrDefault().Text }</span>");
                                                                }
                                                                if (obsInfoList.Count > 2)
                                                                {
                                                                    sbHTML.AppendLine($@"           <span><b>{ TaskRunnerServiceRes.Slope }:</b> { polSourceObsInfoEnumTextAndIDList.Where(c => c.ID == (int)obsInfoList[2]).FirstOrDefault().Text }</span>");
                                                                    sbHTML.AppendLine($@"           <span><b>{ TaskRunnerServiceRes.Risk }:</b> { polSourceObsInfoEnumTextAndIDList.Where(c => c.ID == (int)obsInfoList[obsInfoList.Count - 1]).FirstOrDefault().Text }</span>");
                                                                }
                                                                sbHTML.AppendLine($@"           <br />");
                                                            }

                                                            sbHTML.AppendLine($@"&nbsp;&nbsp;&nbsp;&nbsp;{ countPolSourceSiteForPollutionType }) - ");
                                                            int CountObsInfo = 0;
                                                            foreach (int obsInfo in obsInfoList)
                                                            {
                                                                CountObsInfo += 1;
                                                                if (CountObsInfo > 3 && CountObsInfo < obsInfoList.Count - 2)
                                                                {
                                                                    sbHTML.AppendLine($@"           <span>{ polSourceObsInfoEnumTextAndIDList.Where(c => c.ID == obsInfo).FirstOrDefault().Text }</span> | ");
                                                                }
                                                            }
                                                            sbHTML.AppendLine($@"           <br />");
                                                        }
                                                    }
                                                    sbHTML.AppendLine($@"           <br />");
                                                    sbHTML.AppendLine($@"           &nbsp;&nbsp;&nbsp;&nbsp;<span>Photo</span>");
                                                    sbHTML.AppendLine($@" </blockquote>");
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            sbHTML.AppendLine(@"<span>|||PageBreak|||</span>");

            sbHTML.AppendLine($@" <h3>{ TaskRunnerServiceRes.WaterBasePollutionSourceSiteObservationAndIssues }</h3>");

            sbHTML.AppendLine($@"<div class=""textAlignCenter"">|||Image|FileName,{ fiImageWater.FullName }|width,400|height,150|||</div>");

            foreach (PolSourceObsInfoEnumTextAndID PolSourceObsInfoEnumTextAndID in polSourceObsInfoEnumTextAndIDListWaterBasePolSourceType)
            {
                sbHTML.AppendLine($@" <h4>{ PolSourceObsInfoEnumTextAndID.Text }</h4>");
                foreach (PolSourceSiteModel polSourceSiteModel in polSourceSiteModelList)
                {
                    TVItemModel tvItemModel = tvItemModelListPolSourceSite.Where(c => c.TVItemID == polSourceSiteModel.PolSourceSiteTVItemID).FirstOrDefault();
                    if (tvItemModel != null && tvItemModel.IsActive)
                    {
                        if (polSourceSiteModel != null)
                        {
                            PolSourceObservationModel polSourceObservationModel = polSourceObservationModelList.Where(c => c.PolSourceSiteID == polSourceSiteModel.PolSourceSiteID).OrderByDescending(c => c.ObservationDate_Local).FirstOrDefault();
                            if (polSourceObservationModel != null)
                            {
                                List<PolSourceObservationIssueModel> polSourceObservationIssueModelList2 = polSourceObservationIssueModelList.Where(c => c.PolSourceObservationID == polSourceObservationModel.PolSourceObservationID).OrderBy(c => c.Ordinal).ToList();
                                if (polSourceObservationIssueModelList2.Count > 0)
                                {
                                    if (polSourceObservationIssueModelList2[0].ObservationInfo.Contains(((int)PolSourceObsInfoEnum.WaterBased).ToString() + ","))
                                    {
                                        if (polSourceObservationModel == null)
                                        {
                                            sbHTML.AppendLine($@" <blockquote>");
                                            sbHTML.AppendLine($@" <span>{ TaskRunnerServiceRes.ID }: { polSourceSiteModel.Site }</span>");

                                            List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(polSourceSiteModel.PolSourceSiteTVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);
                                            if (mapInfoPointModelList.Count > 0)
                                            {
                                                sbHTML.AppendLine($@" <span>{ TaskRunnerServiceRes.Coord }: { mapInfoPointModelList[0].Lat }, { mapInfoPointModelList[0].Lng }</span>");
                                            }
                                            sbHTML.AppendLine($@"           <br />");
                                            sbHTML.AppendLine($@"           <span>{ TaskRunnerServiceRes.NoObservationForThisPollutionSourceSite }</span>");
                                            sbHTML.AppendLine($@" </blockquote>");
                                        }
                                        else
                                        {
                                            if (polSourceObservationIssueModelList2.Count == 0)
                                            {
                                                sbHTML.AppendLine($@" <blockquote>");
                                                sbHTML.AppendLine($@" <span>{ TaskRunnerServiceRes.ID }: { polSourceSiteModel.Site }</span>");

                                                List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(polSourceSiteModel.PolSourceSiteTVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);
                                                if (mapInfoPointModelList.Count > 0)
                                                {
                                                    sbHTML.AppendLine($@" <span>{ TaskRunnerServiceRes.Coord }: { mapInfoPointModelList[0].Lat }, { mapInfoPointModelList[0].Lng }</span>");
                                                }
                                                sbHTML.AppendLine($@"           <br />");
                                                sbHTML.AppendLine($@"           <span>{ polSourceObservationModel.Observation_ToBeDeleted }</span>");
                                                sbHTML.AppendLine($@" </blockquote>");
                                            }
                                            else
                                            {
                                                bool IssueOfSourceTypeExist = false;
                                                foreach (PolSourceObservationIssueModel polSourceObservationIssueModel in polSourceObservationIssueModelList2)
                                                {
                                                    if (polSourceObservationIssueModel.ObservationInfo.Contains(PolSourceObsInfoEnumTextAndID.ID.ToString()))
                                                    {
                                                        IssueOfSourceTypeExist = true;
                                                    }
                                                }

                                                if (IssueOfSourceTypeExist)
                                                {
                                                    sbHTML.AppendLine($@" <blockquote>");
                                                    sbHTML.AppendLine($@" <span><b>{ TaskRunnerServiceRes.Site }:</b> { polSourceSiteModel.Site }</span>");

                                                    List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(polSourceSiteModel.PolSourceSiteTVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);
                                                    if (mapInfoPointModelList.Count > 0)
                                                    {
                                                        sbHTML.AppendLine($@" <span><b>{ TaskRunnerServiceRes.Coord }:</b> { mapInfoPointModelList[0].Lat.ToString("F5") }, { mapInfoPointModelList[0].Lng.ToString("F5") }</span>");
                                                    }

                                                    int CountIssues = 0;
                                                    foreach (PolSourceObservationIssueModel polSourceObservationIssueModel in polSourceObservationIssueModelList2)
                                                    {
                                                        if (polSourceObservationIssueModel.ObservationInfo.Contains(PolSourceObsInfoEnumTextAndID.ID.ToString()))
                                                        {
                                                            CountIssues += 1;
                                                            List<int> obsInfoList = polSourceObservationIssueModel.ObservationInfo.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Where(c => !string.IsNullOrWhiteSpace(c)).Select(c => int.Parse(c)).ToList();
                                                            if (CountIssues == 1)
                                                            {
                                                                if (obsInfoList.Count > 1)
                                                                {
                                                                    sbHTML.AppendLine($@"           <span><b>{ TaskRunnerServiceRes.Dist }:</b> { polSourceObsInfoEnumTextAndIDList.Where(c => c.ID == (int)obsInfoList[1]).FirstOrDefault().Text }</span>");
                                                                }
                                                                if (obsInfoList.Count > 2)
                                                                {
                                                                    sbHTML.AppendLine($@"           <span><b>{ TaskRunnerServiceRes.Slope }:</b> { polSourceObsInfoEnumTextAndIDList.Where(c => c.ID == (int)obsInfoList[2]).FirstOrDefault().Text }</span>");
                                                                    sbHTML.AppendLine($@"           <span><b>{ TaskRunnerServiceRes.Risk }:</b> { polSourceObsInfoEnumTextAndIDList.Where(c => c.ID == (int)obsInfoList[obsInfoList.Count - 1]).FirstOrDefault().Text }</span>");
                                                                }
                                                                sbHTML.AppendLine($@"           <br />");
                                                            }

                                                            int CountObsInfo = 0;
                                                            foreach (int obsInfo in obsInfoList)
                                                            {
                                                                CountObsInfo += 1;
                                                                if (CountObsInfo > 3 && CountObsInfo < obsInfoList.Count - 2)
                                                                {
                                                                    sbHTML.AppendLine($@"           <span>{ polSourceObsInfoEnumTextAndIDList.Where(c => c.ID == obsInfo).FirstOrDefault().Text }</span> | ");
                                                                }
                                                            }
                                                            sbHTML.AppendLine($@"           <br />");
                                                        }

                                                    }
                                                    sbHTML.AppendLine($@"           <span>Photo</span>");
                                                    sbHTML.AppendLine($@" </blockquote>");
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            sbHTML.AppendLine(@"<span>|||PageBreak|||</span>");

            sbHTML.AppendLine($@"|||FileNameExtra|Random,{ FileNameExtra }|||");

            if (!GetBottomHTML(sbHTML, fi, parameters))
            {
                return false;
            }

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 90);

            return true;
        }

        // for testing only can comment out when test is completed
        public bool PublicGenerateHTMLSubsectorPollutionSourceSites(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            bool retBool = GenerateHTMLSubsectorPollutionSourceSites(fi, sbHTML, parameters, reportTypeModel);

            StreamWriter sw = fi.CreateText();
            sw.Write(sbHTML.ToString());
            sw.Flush();
            sw.Close();

            return retBool;
        }
    }
}
