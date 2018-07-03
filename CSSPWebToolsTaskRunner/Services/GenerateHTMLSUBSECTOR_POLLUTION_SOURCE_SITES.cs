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
//using System.Web.Helpers;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        private bool GenerateHTMLSUBSECTOR_POLLUTION_SOURCE_SITES(StringBuilder sbTemp)
        {
            int Percent = 10;
            //string NotUsed = "";

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", ReportGenerateObjectsKeywordEnum.SUBSECTOR_POLLUTION_SOURCE_SITES.ToString()));

            if (!GetTopHTML())
            {
                return false;
            }

            sbTemp.AppendLine("<h2>Not implemented</h2>");

            //List<string> ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            //// TVItemID and Year already loaded

            //TVItemModel tvItemModelSubsector = _TVItemService.GetTVItemModelWithTVItemIDDB(TVItemID);
            //if (!string.IsNullOrWhiteSpace(tvItemModelSubsector.Error))
            //{
            //    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
            //    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
            //    return false;
            //}

            //string ServerPath = _TVFileService.GetServerFilePath(tvItemModelSubsector.TVItemID);

            //_TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 5);

            //List<TVItemModel> tvItemModelListPolSourceSite = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.PolSourceSite).Where(c => c.IsActive == true).ToList();
            //List<PolSourceSiteModel> polSourceSiteModelList = _PolSourceSiteService.GetPolSourceSiteModelListWithSubsectorTVItemIDDB(TVItemID).OrderBy(c => c.Site).ToList();
            //List<PolSourceObservationModel> polSourceObservationModelList = _PolSourceObservationService.GetPolSourceObservationModelListWithSubsectorTVItemIDDB(TVItemID);
            //List<PolSourceObservationIssueModel> polSourceObservationIssueModelList = _PolSourceObservationIssueService.GetPolSourceObservationIssueModelListWithSubsectorTVItemIDDB(TVItemID);
            //List<PolSourceObsInfoEnumTextAndID> polSourceObsInfoEnumTextAndIDList = new List<PolSourceObsInfoEnumTextAndID>();
            //List<PolSourceObsInfoEnumTextAndID> polSourceObsInfoEnumTextAndIDListLandBasePolSourceType = new List<PolSourceObsInfoEnumTextAndID>();
            //List<PolSourceObsInfoEnumTextAndID> polSourceObsInfoEnumTextAndIDListWaterBasePolSourceType = new List<PolSourceObsInfoEnumTextAndID>();

            //List<PolSourceObsInfoEnumTextAndID> polSourceObsInfoEnumTextAndIDListLandBaseAgriculture = new List<PolSourceObsInfoEnumTextAndID>();
            //List<PolSourceObsInfoEnumTextAndID> polSourceObsInfoEnumTextAndIDListLandBaseForested = new List<PolSourceObsInfoEnumTextAndID>();
            //List<PolSourceObsInfoEnumTextAndID> polSourceObsInfoEnumTextAndIDListLandBaseIndustry = new List<PolSourceObsInfoEnumTextAndID>();
            //List<PolSourceObsInfoEnumTextAndID> polSourceObsInfoEnumTextAndIDListLandBaseMarine = new List<PolSourceObsInfoEnumTextAndID>();
            //List<PolSourceObsInfoEnumTextAndID> polSourceObsInfoEnumTextAndIDListLandBaseRecreational = new List<PolSourceObsInfoEnumTextAndID>();
            //List<PolSourceObsInfoEnumTextAndID> polSourceObsInfoEnumTextAndIDListLandBaseUrban = new List<PolSourceObsInfoEnumTextAndID>();

            //_TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 10);

            //foreach (int id in Enum.GetValues(typeof(PolSourceObsInfoEnum)))
            //{
            //    if (id == 0)
            //        continue;

            //    polSourceObsInfoEnumTextAndIDList.Add(new PolSourceObsInfoEnumTextAndID() { Text = _BaseEnumService.GetEnumText_PolSourceObsInfoEnum((PolSourceObsInfoEnum)id), ID = id });

            //    if (id.ToString().StartsWith("105") && !id.ToString().EndsWith("00"))
            //    {
            //        polSourceObsInfoEnumTextAndIDListLandBasePolSourceType.Add(new PolSourceObsInfoEnumTextAndID() { Text = _BaseEnumService.GetEnumText_PolSourceObsInfoEnum((PolSourceObsInfoEnum)id), ID = id });
            //    }

            //    if (id.ToString().StartsWith("152") && !id.ToString().EndsWith("00"))
            //    {
            //        polSourceObsInfoEnumTextAndIDListWaterBasePolSourceType.Add(new PolSourceObsInfoEnumTextAndID() { Text = _BaseEnumService.GetEnumText_PolSourceObsInfoEnum((PolSourceObsInfoEnum)id), ID = id });
            //    }

            //    if (id.ToString().StartsWith("106") && !id.ToString().EndsWith("00"))
            //    {
            //        polSourceObsInfoEnumTextAndIDListLandBaseAgriculture.Add(new PolSourceObsInfoEnumTextAndID() { Text = _BaseEnumService.GetEnumText_PolSourceObsInfoEnum((PolSourceObsInfoEnum)id), ID = id });
            //    }

            //    if (id.ToString().StartsWith("111") && !id.ToString().EndsWith("00"))
            //    {
            //        polSourceObsInfoEnumTextAndIDListLandBaseForested.Add(new PolSourceObsInfoEnumTextAndID() { Text = _BaseEnumService.GetEnumText_PolSourceObsInfoEnum((PolSourceObsInfoEnum)id), ID = id });
            //    }

            //    if (id.ToString().StartsWith("112") && !id.ToString().EndsWith("00"))
            //    {
            //        polSourceObsInfoEnumTextAndIDListLandBaseIndustry.Add(new PolSourceObsInfoEnumTextAndID() { Text = _BaseEnumService.GetEnumText_PolSourceObsInfoEnum((PolSourceObsInfoEnum)id), ID = id });
            //    }

            //    if (id.ToString().StartsWith("114") && !id.ToString().EndsWith("00"))
            //    {
            //        polSourceObsInfoEnumTextAndIDListLandBaseMarine.Add(new PolSourceObsInfoEnumTextAndID() { Text = _BaseEnumService.GetEnumText_PolSourceObsInfoEnum((PolSourceObsInfoEnum)id), ID = id });
            //    }

            //    if (id.ToString().StartsWith("117") && !id.ToString().EndsWith("00"))
            //    {
            //        polSourceObsInfoEnumTextAndIDListLandBaseRecreational.Add(new PolSourceObsInfoEnumTextAndID() { Text = _BaseEnumService.GetEnumText_PolSourceObsInfoEnum((PolSourceObsInfoEnum)id), ID = id });
            //    }

            //    if (id.ToString().StartsWith("121") && !id.ToString().EndsWith("00"))
            //    {
            //        polSourceObsInfoEnumTextAndIDListLandBaseUrban.Add(new PolSourceObsInfoEnumTextAndID() { Text = _BaseEnumService.GetEnumText_PolSourceObsInfoEnum((PolSourceObsInfoEnum)id), ID = id });
            //    }
            //}

            //polSourceObsInfoEnumTextAndIDListLandBasePolSourceType = polSourceObsInfoEnumTextAndIDListLandBasePolSourceType.OrderBy(c => c.Text).ToList();
            //polSourceObsInfoEnumTextAndIDListWaterBasePolSourceType = polSourceObsInfoEnumTextAndIDListWaterBasePolSourceType.OrderBy(c => c.Text).ToList();
            //polSourceObsInfoEnumTextAndIDListLandBaseAgriculture = polSourceObsInfoEnumTextAndIDListLandBaseAgriculture.OrderBy(c => c.Text).ToList();
            //polSourceObsInfoEnumTextAndIDListLandBaseForested = polSourceObsInfoEnumTextAndIDListLandBaseForested.OrderBy(c => c.Text).ToList();
            //polSourceObsInfoEnumTextAndIDListLandBaseIndustry = polSourceObsInfoEnumTextAndIDListLandBaseIndustry.OrderBy(c => c.Text).ToList();
            //polSourceObsInfoEnumTextAndIDListLandBaseMarine = polSourceObsInfoEnumTextAndIDListLandBaseMarine.OrderBy(c => c.Text).ToList();
            //polSourceObsInfoEnumTextAndIDListLandBaseRecreational = polSourceObsInfoEnumTextAndIDListLandBaseRecreational.OrderBy(c => c.Text).ToList();
            //polSourceObsInfoEnumTextAndIDListLandBaseUrban = polSourceObsInfoEnumTextAndIDListLandBaseUrban.OrderBy(c => c.Text).ToList();

            //// ---------------------------------------------------------------------------------------------------------------
            //// Land Base graphic showing the number of polsource issues of different pollution source type (Agriculture, Forested, ...)
            //// ---------------------------------------------------------------------------------------------------------------

            //Percent = 15;
            //_TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);

            //List<string> PolSourceTypeList = new List<string>();
            //List<int> CountOfPolSourceType = new List<int>();

            //foreach (PolSourceObsInfoEnumTextAndID polSourceObsInfoEnumTextAndID in polSourceObsInfoEnumTextAndIDListLandBasePolSourceType)
            //{
            //    List<PolSourceSiteModel> polSourceSiteModelListActive = (from pss in polSourceSiteModelList
            //                                                             from t in tvItemModelListPolSourceSite
            //                                                             where pss.PolSourceSiteTVItemID == t.TVItemID
            //                                                             && t.IsActive == true
            //                                                             select pss).ToList();

            //    List<PolSourceObservationModel> polSourceObservationModelLastOfActive = (from pss in polSourceSiteModelListActive
            //                                                                             let psoLast = (from pso in polSourceObservationModelList
            //                                                                                            where pss.PolSourceSiteID == pso.PolSourceSiteID
            //                                                                                            orderby pso.ObservationDate_Local descending
            //                                                                                            select pso).FirstOrDefault()
            //                                                                             select psoLast).ToList();

            //    string enumTextID = "," + polSourceObsInfoEnumTextAndID.ID.ToString() + ",";

            //    int count = (from pso in polSourceObservationModelLastOfActive
            //                 from psoi in polSourceObservationIssueModelList
            //                 where pso != null
            //                 && pso.PolSourceObservationID == psoi.PolSourceObservationID
            //                 && psoi.ObservationInfo.Contains(enumTextID)
            //                 select psoi).Count();

            //    PolSourceTypeList.Add(polSourceObsInfoEnumTextAndID.Text);
            //    CountOfPolSourceType.Add(count);
            //}

            //Percent = 20;
            //_TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);

            //if (xlApp == null)
            //{
            //    xlApp = new Microsoft.Office.Interop.Excel.Application();
            //    workbook = xlApp.Workbooks.Add();
            //    worksheet = workbook.Worksheets.get_Item(1);
            //    xlCharts = (Microsoft.Office.Interop.Excel.ChartObjects)worksheet.ChartObjects();
            //}

            //Microsoft.Office.Interop.Excel.ChartObject chart = xlCharts.Add(100, 100, 600, 200);
            //Microsoft.Office.Interop.Excel.Chart chartPage = chart.Chart;

            //chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;

            //Microsoft.Office.Interop.Excel.SeriesCollection seriesCollection = chartPage.SeriesCollection();
            //Microsoft.Office.Interop.Excel.Series series = seriesCollection.NewSeries();

            //series.XValues = PolSourceTypeList.ToArray();
            //series.Values = CountOfPolSourceType.ToArray();

            //chartPage.ApplyLayout(9, Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered);
            //chartPage.ChartTitle.Select();
            //xlApp.Selection.Delete();
            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue).AxisTitle.Select();
            //xlApp.Selection.Delete();
            //chartPage.Legend.Select();
            //xlApp.Selection.Delete();
            ////chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).TickLabelSpacing = 5;
            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).MajorTickMark = Microsoft.Office.Interop.Excel.Constants.xlOutside;
            //chartPage.Parent.RoundedCorners = true;

            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = TaskRunnerServiceRes.NumberOfIssuesByLandBasePollutionType;

            //// need to save the file with a unique name under the TVItemID
            //FileInfo fiImageLand = new FileInfo(fi.DirectoryName + @"\LandBaseStat" + FileNameExtra + ".png");

            //DirectoryInfo di = new DirectoryInfo(fi.DirectoryName);

            //if (!di.Exists)
            //{
            //    try
            //    {
            //        di.Create();
            //    }
            //    catch (Exception ex)
            //    {
            //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateDirectory__, di.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
            //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateDirectory__", di.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
            //        return false;
            //    }
            //}

            //chartPage.Export(fiImageLand.FullName, "PNG", false);

            //Percent = 25;
            //_TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);

            //// ---------------------------------------------------------------------------------------------------------------
            //// Water Base graphic showing the number of polsource issues of different pollution source type (Aquaculture, Seaport ...)
            //// ---------------------------------------------------------------------------------------------------------------

            //PolSourceTypeList = new List<string>();
            //CountOfPolSourceType = new List<int>();

            //foreach (PolSourceObsInfoEnumTextAndID polSourceObsInfoEnumTextAndID in polSourceObsInfoEnumTextAndIDListWaterBasePolSourceType)
            //{
            //    List<PolSourceSiteModel> polSourceSiteModelListActive = (from pss in polSourceSiteModelList
            //                                                             from t in tvItemModelListPolSourceSite
            //                                                             where pss.PolSourceSiteTVItemID == t.TVItemID
            //                                                             && t.IsActive == true
            //                                                             select pss).ToList();

            //    List<PolSourceObservationModel> polSourceObservationModelLastOfActive = (from pss in polSourceSiteModelListActive
            //                                                                             let psoLast = (from pso in polSourceObservationModelList
            //                                                                                            where pss.PolSourceSiteID == pso.PolSourceSiteID
            //                                                                                            orderby pso.ObservationDate_Local descending
            //                                                                                            select pso).FirstOrDefault()
            //                                                                             select psoLast).ToList();

            //    int count = (from pso in polSourceObservationModelLastOfActive
            //                 from psoi in polSourceObservationIssueModelList
            //                 where pso != null
            //                 && pso.PolSourceObservationID == psoi.PolSourceObservationID
            //                 && psoi.ObservationInfo.Contains("," + polSourceObsInfoEnumTextAndID.ID.ToString() + ",")
            //                 select psoi.PolSourceObservationIssueID).Count();

            //    PolSourceTypeList.Add(polSourceObsInfoEnumTextAndID.Text);
            //    CountOfPolSourceType.Add(count);
            //}

            //chart = xlCharts.Add(100, 100, 600, 200);
            //chartPage = chart.Chart;

            //chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;

            //seriesCollection = chartPage.SeriesCollection();
            //series = seriesCollection.NewSeries();

            //series.XValues = PolSourceTypeList.ToArray();
            //series.Values = CountOfPolSourceType.ToArray();

            //chartPage.ApplyLayout(9, Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered);
            //chartPage.ChartTitle.Select();
            //xlApp.Selection.Delete();
            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue).AxisTitle.Select();
            //xlApp.Selection.Delete();
            //chartPage.Legend.Select();
            //xlApp.Selection.Delete();
            ////chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).TickLabelSpacing = 5;
            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).MajorTickMark = Microsoft.Office.Interop.Excel.Constants.xlOutside;
            //chartPage.Parent.RoundedCorners = true;

            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = TaskRunnerServiceRes.NumberOfIssuesByWaterBasePollutionType;

            //// need to save the file with a unique name under the TVItemID
            //FileInfo fiImageWater = new FileInfo(fi.DirectoryName + @"\WaterBaseStat" + FileNameExtra + ".png");

            //di = new DirectoryInfo(fi.DirectoryName);

            //if (!di.Exists)
            //{
            //    try
            //    {
            //        di.Create();
            //    }
            //    catch (Exception ex)
            //    {
            //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateDirectory__, di.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
            //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateDirectory__", di.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
            //        return false;
            //    }
            //}

            //chartPage.Export(fiImageWater.FullName, "PNG", false);

            //Percent = 35;
            //_TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);

            //// ---------------------------------------------------------------------------------------------------------------
            //// Agriculture graphic showing the number of polsource issues of different pollution source type
            //// ---------------------------------------------------------------------------------------------------------------

            //PolSourceTypeList = new List<string>();
            //CountOfPolSourceType = new List<int>();

            //foreach (PolSourceObsInfoEnumTextAndID polSourceObsInfoEnumTextAndID in polSourceObsInfoEnumTextAndIDListLandBaseAgriculture)
            //{
            //    List<PolSourceSiteModel> polSourceSiteModelListActive = (from pss in polSourceSiteModelList
            //                                                             from t in tvItemModelListPolSourceSite
            //                                                             where pss.PolSourceSiteTVItemID == t.TVItemID
            //                                                             && t.IsActive == true
            //                                                             select pss).ToList();

            //    List<PolSourceObservationModel> polSourceObservationModelLastOfActive = (from pss in polSourceSiteModelListActive
            //                                                                             let psoLast = (from pso in polSourceObservationModelList
            //                                                                                            where pss.PolSourceSiteID == pso.PolSourceSiteID
            //                                                                                            orderby pso.ObservationDate_Local descending
            //                                                                                            select pso).FirstOrDefault()
            //                                                                             select psoLast).ToList();

            //    int count = (from pso in polSourceObservationModelLastOfActive
            //                 from psoi in polSourceObservationIssueModelList
            //                 where pso != null
            //                 && pso.PolSourceObservationID == psoi.PolSourceObservationID
            //                 && psoi.ObservationInfo.Contains("," + polSourceObsInfoEnumTextAndID.ID.ToString() + ",")
            //                 select psoi.PolSourceObservationIssueID).Count();

            //    PolSourceTypeList.Add(polSourceObsInfoEnumTextAndID.Text);
            //    CountOfPolSourceType.Add(count);
            //}

            //chart = xlCharts.Add(100, 100, 600, 200);
            //chartPage = chart.Chart;

            //chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;

            //seriesCollection = chartPage.SeriesCollection();
            //series = seriesCollection.NewSeries();

            //series.XValues = PolSourceTypeList.ToArray();
            //series.Values = CountOfPolSourceType.ToArray();

            //chartPage.ApplyLayout(9, Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered);
            //chartPage.ChartTitle.Select();
            //xlApp.Selection.Delete();
            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue).AxisTitle.Select();
            //xlApp.Selection.Delete();
            //chartPage.Legend.Select();
            //xlApp.Selection.Delete();
            ////chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).TickLabelSpacing = 5;
            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).MajorTickMark = Microsoft.Office.Interop.Excel.Constants.xlOutside;
            //chartPage.Parent.RoundedCorners = true;

            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = TaskRunnerServiceRes.NumberOfIssuesByWaterBasePollutionType;

            //// need to save the file with a unique name under the TVItemID
            //FileInfo fiImageAgriculture = new FileInfo(fi.DirectoryName + @"\AgricultureStat" + FileNameExtra + ".png");

            //di = new DirectoryInfo(fi.DirectoryName);

            //if (!di.Exists)
            //{
            //    try
            //    {
            //        di.Create();
            //    }
            //    catch (Exception ex)
            //    {
            //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateDirectory__, di.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
            //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateDirectory__", di.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
            //        return false;
            //    }
            //}

            //chartPage.Export(fiImageAgriculture.FullName, "PNG", false);

            //Percent = 45;
            //_TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);

            //// ---------------------------------------------------------------------------------------------------------------
            //// Forested graphic showing the number of polsource issues of different pollution source type
            //// ---------------------------------------------------------------------------------------------------------------

            //PolSourceTypeList = new List<string>();
            //CountOfPolSourceType = new List<int>();

            //foreach (PolSourceObsInfoEnumTextAndID polSourceObsInfoEnumTextAndID in polSourceObsInfoEnumTextAndIDListLandBaseForested)
            //{
            //    List<PolSourceSiteModel> polSourceSiteModelListActive = (from pss in polSourceSiteModelList
            //                                                             from t in tvItemModelListPolSourceSite
            //                                                             where pss.PolSourceSiteTVItemID == t.TVItemID
            //                                                             && t.IsActive == true
            //                                                             select pss).ToList();

            //    List<PolSourceObservationModel> polSourceObservationModelLastOfActive = (from pss in polSourceSiteModelListActive
            //                                                                             let psoLast = (from pso in polSourceObservationModelList
            //                                                                                            where pss.PolSourceSiteID == pso.PolSourceSiteID
            //                                                                                            orderby pso.ObservationDate_Local descending
            //                                                                                            select pso).FirstOrDefault()
            //                                                                             select psoLast).ToList();

            //    int count = (from pso in polSourceObservationModelLastOfActive
            //                 from psoi in polSourceObservationIssueModelList
            //                 where pso != null
            //                 && pso.PolSourceObservationID == psoi.PolSourceObservationID
            //                 && psoi.ObservationInfo.Contains("," + polSourceObsInfoEnumTextAndID.ID.ToString() + ",")
            //                 select psoi.PolSourceObservationIssueID).Count();

            //    PolSourceTypeList.Add(polSourceObsInfoEnumTextAndID.Text);
            //    CountOfPolSourceType.Add(count);
            //}

            //chart = xlCharts.Add(100, 100, 600, 200);
            //chartPage = chart.Chart;

            //chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;

            //seriesCollection = chartPage.SeriesCollection();
            //series = seriesCollection.NewSeries();

            //series.XValues = PolSourceTypeList.ToArray();
            //series.Values = CountOfPolSourceType.ToArray();

            //chartPage.ApplyLayout(9, Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered);
            //chartPage.ChartTitle.Select();
            //xlApp.Selection.Delete();
            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue).AxisTitle.Select();
            //xlApp.Selection.Delete();
            //chartPage.Legend.Select();
            //xlApp.Selection.Delete();
            ////chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).TickLabelSpacing = 5;
            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).MajorTickMark = Microsoft.Office.Interop.Excel.Constants.xlOutside;
            //chartPage.Parent.RoundedCorners = true;

            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = TaskRunnerServiceRes.NumberOfIssuesByWaterBasePollutionType;

            //// need to save the file with a unique name under the TVItemID
            //FileInfo fiImageForested = new FileInfo(fi.DirectoryName + @"\ForestedStat" + FileNameExtra + ".png");

            //di = new DirectoryInfo(fi.DirectoryName);

            //if (!di.Exists)
            //{
            //    try
            //    {
            //        di.Create();
            //    }
            //    catch (Exception ex)
            //    {
            //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateDirectory__, di.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
            //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateDirectory__", di.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
            //        return false;
            //    }
            //}

            //chartPage.Export(fiImageForested.FullName, "PNG", false);

            //Percent = 55;
            //_TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);

            //// ---------------------------------------------------------------------------------------------------------------
            //// Industry graphic showing the number of polsource issues of different pollution source type
            //// ---------------------------------------------------------------------------------------------------------------

            //PolSourceTypeList = new List<string>();
            //CountOfPolSourceType = new List<int>();

            //foreach (PolSourceObsInfoEnumTextAndID polSourceObsInfoEnumTextAndID in polSourceObsInfoEnumTextAndIDListLandBaseIndustry)
            //{
            //    List<PolSourceSiteModel> polSourceSiteModelListActive = (from pss in polSourceSiteModelList
            //                                                             from t in tvItemModelListPolSourceSite
            //                                                             where pss.PolSourceSiteTVItemID == t.TVItemID
            //                                                             && t.IsActive == true
            //                                                             select pss).ToList();

            //    List<PolSourceObservationModel> polSourceObservationModelLastOfActive = (from pss in polSourceSiteModelListActive
            //                                                                             let psoLast = (from pso in polSourceObservationModelList
            //                                                                                            where pss.PolSourceSiteID == pso.PolSourceSiteID
            //                                                                                            orderby pso.ObservationDate_Local descending
            //                                                                                            select pso).FirstOrDefault()
            //                                                                             select psoLast).ToList();

            //    int count = (from pso in polSourceObservationModelLastOfActive
            //                 from psoi in polSourceObservationIssueModelList
            //                 where pso != null
            //                 && pso.PolSourceObservationID == psoi.PolSourceObservationID
            //                 && psoi.ObservationInfo.Contains("," + polSourceObsInfoEnumTextAndID.ID.ToString() + ",")
            //                 select psoi.PolSourceObservationIssueID).Count();

            //    PolSourceTypeList.Add(polSourceObsInfoEnumTextAndID.Text);
            //    CountOfPolSourceType.Add(count);
            //}

            //chart = xlCharts.Add(100, 100, 600, 200);
            //chartPage = chart.Chart;

            //chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;

            //seriesCollection = chartPage.SeriesCollection();
            //series = seriesCollection.NewSeries();

            //series.XValues = PolSourceTypeList.ToArray();
            //series.Values = CountOfPolSourceType.ToArray();

            //chartPage.ApplyLayout(9, Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered);
            //chartPage.ChartTitle.Select();
            //xlApp.Selection.Delete();
            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue).AxisTitle.Select();
            //xlApp.Selection.Delete();
            //chartPage.Legend.Select();
            //xlApp.Selection.Delete();
            ////chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).TickLabelSpacing = 5;
            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).MajorTickMark = Microsoft.Office.Interop.Excel.Constants.xlOutside;
            //chartPage.Parent.RoundedCorners = true;

            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = TaskRunnerServiceRes.NumberOfIssuesByWaterBasePollutionType;

            //// need to save the file with a unique name under the TVItemID
            //FileInfo fiImageIndustry = new FileInfo(fi.DirectoryName + @"\IndustryStat" + FileNameExtra + ".png");

            //di = new DirectoryInfo(fi.DirectoryName);

            //if (!di.Exists)
            //{
            //    try
            //    {
            //        di.Create();
            //    }
            //    catch (Exception ex)
            //    {
            //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateDirectory__, di.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
            //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateDirectory__", di.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
            //        return false;
            //    }
            //}

            //chartPage.Export(fiImageIndustry.FullName, "PNG", false);

            //Percent = 65;
            //_TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);

            //// ---------------------------------------------------------------------------------------------------------------
            //// Marine graphic showing the number of polsource issues of different pollution source type
            //// ---------------------------------------------------------------------------------------------------------------

            //PolSourceTypeList = new List<string>();
            //CountOfPolSourceType = new List<int>();

            //foreach (PolSourceObsInfoEnumTextAndID polSourceObsInfoEnumTextAndID in polSourceObsInfoEnumTextAndIDListLandBaseMarine)
            //{
            //    List<PolSourceSiteModel> polSourceSiteModelListActive = (from pss in polSourceSiteModelList
            //                                                             from t in tvItemModelListPolSourceSite
            //                                                             where pss.PolSourceSiteTVItemID == t.TVItemID
            //                                                             && t.IsActive == true
            //                                                             select pss).ToList();

            //    List<PolSourceObservationModel> polSourceObservationModelLastOfActive = (from pss in polSourceSiteModelListActive
            //                                                                             let psoLast = (from pso in polSourceObservationModelList
            //                                                                                            where pss.PolSourceSiteID == pso.PolSourceSiteID
            //                                                                                            orderby pso.ObservationDate_Local descending
            //                                                                                            select pso).FirstOrDefault()
            //                                                                             select psoLast).ToList();

            //    int count = (from pso in polSourceObservationModelLastOfActive
            //                 from psoi in polSourceObservationIssueModelList
            //                 where pso != null
            //                 && pso.PolSourceObservationID == psoi.PolSourceObservationID
            //                 && psoi.ObservationInfo.Contains("," + polSourceObsInfoEnumTextAndID.ID.ToString() + ",")
            //                 select psoi.PolSourceObservationIssueID).Count();

            //    PolSourceTypeList.Add(polSourceObsInfoEnumTextAndID.Text);
            //    CountOfPolSourceType.Add(count);
            //}

            //chart = xlCharts.Add(100, 100, 600, 200);
            //chartPage = chart.Chart;

            //chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;

            //seriesCollection = chartPage.SeriesCollection();
            //series = seriesCollection.NewSeries();

            //series.XValues = PolSourceTypeList.ToArray();
            //series.Values = CountOfPolSourceType.ToArray();

            //chartPage.ApplyLayout(9, Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered);
            //chartPage.ChartTitle.Select();
            //xlApp.Selection.Delete();
            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue).AxisTitle.Select();
            //xlApp.Selection.Delete();
            //chartPage.Legend.Select();
            //xlApp.Selection.Delete();
            ////chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).TickLabelSpacing = 5;
            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).MajorTickMark = Microsoft.Office.Interop.Excel.Constants.xlOutside;
            //chartPage.Parent.RoundedCorners = true;

            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = TaskRunnerServiceRes.NumberOfIssuesByWaterBasePollutionType;

            //// need to save the file with a unique name under the TVItemID
            //FileInfo fiImageMarine = new FileInfo(fi.DirectoryName + @"\MarineStat" + FileNameExtra + ".png");

            //di = new DirectoryInfo(fi.DirectoryName);

            //if (!di.Exists)
            //{
            //    try
            //    {
            //        di.Create();
            //    }
            //    catch (Exception ex)
            //    {
            //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateDirectory__, di.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
            //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateDirectory__", di.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
            //        return false;
            //    }
            //}

            //chartPage.Export(fiImageMarine.FullName, "PNG", false);

            //Percent = 75;
            //_TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);

            //// ---------------------------------------------------------------------------------------------------------------
            //// Recreational graphic showing the number of polsource issues of different pollution source type
            //// ---------------------------------------------------------------------------------------------------------------

            //PolSourceTypeList = new List<string>();
            //CountOfPolSourceType = new List<int>();

            //foreach (PolSourceObsInfoEnumTextAndID polSourceObsInfoEnumTextAndID in polSourceObsInfoEnumTextAndIDListLandBaseRecreational)
            //{
            //    List<PolSourceSiteModel> polSourceSiteModelListActive = (from pss in polSourceSiteModelList
            //                                                             from t in tvItemModelListPolSourceSite
            //                                                             where pss.PolSourceSiteTVItemID == t.TVItemID
            //                                                             && t.IsActive == true
            //                                                             select pss).ToList();

            //    List<PolSourceObservationModel> polSourceObservationModelLastOfActive = (from pss in polSourceSiteModelListActive
            //                                                                             let psoLast = (from pso in polSourceObservationModelList
            //                                                                                            where pss.PolSourceSiteID == pso.PolSourceSiteID
            //                                                                                            orderby pso.ObservationDate_Local descending
            //                                                                                            select pso).FirstOrDefault()
            //                                                                             select psoLast).ToList();

            //    int count = (from pso in polSourceObservationModelLastOfActive
            //                 from psoi in polSourceObservationIssueModelList
            //                 where pso != null
            //                 && pso.PolSourceObservationID == psoi.PolSourceObservationID
            //                 && psoi.ObservationInfo.Contains("," + polSourceObsInfoEnumTextAndID.ID.ToString() + ",")
            //                 select psoi.PolSourceObservationIssueID).Count();

            //    PolSourceTypeList.Add(polSourceObsInfoEnumTextAndID.Text);
            //    CountOfPolSourceType.Add(count);
            //}

            //chart = xlCharts.Add(100, 100, 600, 200);
            //chartPage = chart.Chart;

            //chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;

            //seriesCollection = chartPage.SeriesCollection();
            //series = seriesCollection.NewSeries();

            //series.XValues = PolSourceTypeList.ToArray();
            //series.Values = CountOfPolSourceType.ToArray();

            //chartPage.ApplyLayout(9, Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered);
            //chartPage.ChartTitle.Select();
            //xlApp.Selection.Delete();
            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue).AxisTitle.Select();
            //xlApp.Selection.Delete();
            //chartPage.Legend.Select();
            //xlApp.Selection.Delete();
            ////chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).TickLabelSpacing = 5;
            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).MajorTickMark = Microsoft.Office.Interop.Excel.Constants.xlOutside;
            //chartPage.Parent.RoundedCorners = true;

            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = TaskRunnerServiceRes.NumberOfIssuesByWaterBasePollutionType;

            //// need to save the file with a unique name under the TVItemID
            //FileInfo fiImageRecreational = new FileInfo(fi.DirectoryName + @"\RecreationalStat" + FileNameExtra + ".png");

            //di = new DirectoryInfo(fi.DirectoryName);

            //if (!di.Exists)
            //{
            //    try
            //    {
            //        di.Create();
            //    }
            //    catch (Exception ex)
            //    {
            //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateDirectory__, di.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
            //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateDirectory__", di.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
            //        return false;
            //    }
            //}

            //chartPage.Export(fiImageRecreational.FullName, "PNG", false);

            //Percent = 85;
            //_TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);

            //// ---------------------------------------------------------------------------------------------------------------
            //// Urban graphic showing the number of polsource issues of different pollution source type
            //// ---------------------------------------------------------------------------------------------------------------

            //PolSourceTypeList = new List<string>();
            //CountOfPolSourceType = new List<int>();

            //foreach (PolSourceObsInfoEnumTextAndID polSourceObsInfoEnumTextAndID in polSourceObsInfoEnumTextAndIDListLandBaseUrban)
            //{
            //    List<PolSourceSiteModel> polSourceSiteModelListActive = (from pss in polSourceSiteModelList
            //                                                             from t in tvItemModelListPolSourceSite
            //                                                             where pss.PolSourceSiteTVItemID == t.TVItemID
            //                                                             && t.IsActive == true
            //                                                             select pss).ToList();

            //    List<PolSourceObservationModel> polSourceObservationModelLastOfActive = (from pss in polSourceSiteModelListActive
            //                                                                             let psoLast = (from pso in polSourceObservationModelList
            //                                                                                            where pss.PolSourceSiteID == pso.PolSourceSiteID
            //                                                                                            orderby pso.ObservationDate_Local descending
            //                                                                                            select pso).FirstOrDefault()
            //                                                                             select psoLast).ToList();

            //    int count = (from pso in polSourceObservationModelLastOfActive
            //                 from psoi in polSourceObservationIssueModelList
            //                 where pso != null
            //                 && pso.PolSourceObservationID == psoi.PolSourceObservationID
            //                 && psoi.ObservationInfo.Contains("," + polSourceObsInfoEnumTextAndID.ID.ToString() + ",")
            //                 select psoi.PolSourceObservationIssueID).Count();

            //    PolSourceTypeList.Add(polSourceObsInfoEnumTextAndID.Text);
            //    CountOfPolSourceType.Add(count);
            //}

            //chart = xlCharts.Add(100, 100, 600, 200);
            //chartPage = chart.Chart;

            //chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;

            //seriesCollection = chartPage.SeriesCollection();
            //series = seriesCollection.NewSeries();

            //series.XValues = PolSourceTypeList.ToArray();
            //series.Values = CountOfPolSourceType.ToArray();

            //chartPage.ApplyLayout(9, Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered);
            //chartPage.ChartTitle.Select();
            //xlApp.Selection.Delete();
            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue).AxisTitle.Select();
            //xlApp.Selection.Delete();
            //chartPage.Legend.Select();
            //xlApp.Selection.Delete();
            ////chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).TickLabelSpacing = 5;
            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).MajorTickMark = Microsoft.Office.Interop.Excel.Constants.xlOutside;
            //chartPage.Parent.RoundedCorners = true;

            //chartPage.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = TaskRunnerServiceRes.NumberOfIssuesByWaterBasePollutionType;

            //// need to save the file with a unique name under the TVItemID
            //FileInfo fiImageUrban = new FileInfo(fi.DirectoryName + @"\UrbanStat" + FileNameExtra + ".png");

            //di = new DirectoryInfo(fi.DirectoryName);

            //if (!di.Exists)
            //{
            //    try
            //    {
            //        di.Create();
            //    }
            //    catch (Exception ex)
            //    {
            //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateDirectory__, di.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
            //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateDirectory__", di.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
            //        return false;
            //    }
            //}

            //chartPage.Export(fiImageUrban.FullName, "PNG", false);

            //Percent = 95;
            //_TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);

            //sbTemp.AppendLine($@" <h3>{ TaskRunnerServiceRes.LandBasePollutionSourceSiteObservationAndIssues }</h3>");

            //sbTemp.AppendLine($@"<div class=""textAlignCenter"">|||Image|FileName,{ fiImageLand.FullName }|width,400|height,150|||</div>");

            //foreach (PolSourceObsInfoEnumTextAndID PolSourceObsInfoEnumTextAndID in polSourceObsInfoEnumTextAndIDListLandBasePolSourceType)
            //{
            //    sbTemp.AppendLine($@" <h4>{ PolSourceObsInfoEnumTextAndID.Text }</h4>");

            //    switch (PolSourceObsInfoEnumTextAndID.ID)
            //    {
            //        case 10501: // Agriculture
            //            {
            //                sbTemp.AppendLine($@"<div class=""textAlignCenter"">|||Image|FileName,{ fiImageAgriculture.FullName }|width,400|height,150|||</div>");
            //            }
            //            break;
            //        case 10502: // Forested
            //            {
            //                sbTemp.AppendLine($@"<div class=""textAlignCenter"">|||Image|FileName,{ fiImageForested.FullName }|width,400|height,150|||</div>");
            //            }
            //            break;
            //        case 10503: // Industry
            //            {
            //                sbTemp.AppendLine($@"<div class=""textAlignCenter"">|||Image|FileName,{ fiImageIndustry.FullName }|width,400|height,150|||</div>");
            //            }
            //            break;
            //        case 10504: // Marine
            //            {
            //                sbTemp.AppendLine($@"<div class=""textAlignCenter"">|||Image|FileName,{ fiImageMarine.FullName }|width,400|height,150|||</div>");
            //            }
            //            break;
            //        case 10505: // Recreational
            //            {
            //                sbTemp.AppendLine($@"<div class=""textAlignCenter"">|||Image|FileName,{ fiImageRecreational.FullName }|width,400|height,150|||</div>");
            //            }
            //            break;
            //        case 10506: // Urban
            //            {
            //                sbTemp.AppendLine($@"<div class=""textAlignCenter"">|||Image|FileName,{ fiImageUrban.FullName }|width,400|height,150|||</div>");
            //            }
            //            break;
            //        default:
            //            break;
            //    }


            //    int countPolSourceSiteForPollutionType = 0;
            //    foreach (PolSourceSiteModel polSourceSiteModel in polSourceSiteModelList)
            //    {
            //        TVItemModel tvItemModel = tvItemModelListPolSourceSite.Where(c => c.TVItemID == polSourceSiteModel.PolSourceSiteTVItemID).FirstOrDefault();
            //        if (tvItemModel != null && tvItemModel.IsActive)
            //        {
            //            if (polSourceSiteModel != null)
            //            {
            //                PolSourceObservationModel polSourceObservationModel = polSourceObservationModelList.Where(c => c.PolSourceSiteID == polSourceSiteModel.PolSourceSiteID).OrderByDescending(c => c.ObservationDate_Local).FirstOrDefault();
            //                if (polSourceObservationModel != null)
            //                {
            //                    List<PolSourceObservationIssueModel> polSourceObservationIssueModelList2 = polSourceObservationIssueModelList.Where(c => c.PolSourceObservationID == polSourceObservationModel.PolSourceObservationID).OrderBy(c => c.Ordinal).ToList();
            //                    if (polSourceObservationIssueModelList2.Count > 0)
            //                    {
            //                        if (polSourceObservationIssueModelList2[0].ObservationInfo.Contains(((int)PolSourceObsInfoEnum.SourceHumanPollution).ToString() + ","))
            //                        {
            //                            if (polSourceObservationModel == null)
            //                            {
            //                                countPolSourceSiteForPollutionType += 1;

            //                                sbTemp.AppendLine($@" <blockquote>");
            //                                sbTemp.AppendLine($@" <span><b>{ TaskRunnerServiceRes.Site }:</b> { polSourceSiteModel.Site }</span>");

            //                                List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(polSourceSiteModel.PolSourceSiteTVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);
            //                                if (mapInfoPointModelList.Count > 0)
            //                                {
            //                                    sbTemp.AppendLine($@" <span>{ TaskRunnerServiceRes.Coord }: { mapInfoPointModelList[0].Lat }, { mapInfoPointModelList[0].Lng }</span>");
            //                                }
            //                                sbTemp.AppendLine($@"           <br />");
            //                                sbTemp.AppendLine($@"           <span>{ TaskRunnerServiceRes.NoObservationForThisPollutionSourceSite }</span>");
            //                                sbTemp.AppendLine($@" </blockquote>");
            //                            }
            //                            else
            //                            {
            //                                if (polSourceObservationIssueModelList2.Count == 0)
            //                                {
            //                                    countPolSourceSiteForPollutionType += 1;

            //                                    sbTemp.AppendLine($@" <blockquote>");
            //                                    sbTemp.AppendLine($@" <span><b>{ TaskRunnerServiceRes.Site }:</b> { polSourceSiteModel.Site }</span>");

            //                                    List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(polSourceSiteModel.PolSourceSiteTVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);
            //                                    if (mapInfoPointModelList.Count > 0)
            //                                    {
            //                                        sbTemp.AppendLine($@" <span>{ TaskRunnerServiceRes.Coord }: { mapInfoPointModelList[0].Lat }, { mapInfoPointModelList[0].Lng }</span>");
            //                                    }
            //                                    sbTemp.AppendLine($@"           <br />");
            //                                    sbTemp.AppendLine($@"           <span>{ polSourceObservationModel.Observation_ToBeDeleted }</span>");
            //                                    sbTemp.AppendLine($@" </blockquote>");
            //                                }
            //                                else
            //                                {
            //                                    bool IssueOfSourceTypeExist = false;
            //                                    foreach (PolSourceObservationIssueModel polSourceObservationIssueModel in polSourceObservationIssueModelList2)
            //                                    {
            //                                        if (polSourceObservationIssueModel.ObservationInfo.Contains(PolSourceObsInfoEnumTextAndID.ID.ToString()))
            //                                        {
            //                                            IssueOfSourceTypeExist = true;
            //                                        }
            //                                    }

            //                                    if (IssueOfSourceTypeExist)
            //                                    {
            //                                        sbTemp.AppendLine($@" <blockquote>");
            //                                        sbTemp.AppendLine($@" <span><b>{ TaskRunnerServiceRes.Site }:</b> { polSourceSiteModel.Site }</span>");

            //                                        List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(polSourceSiteModel.PolSourceSiteTVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);
            //                                        if (mapInfoPointModelList.Count > 0)
            //                                        {
            //                                            sbTemp.AppendLine($@" <span><b>{ TaskRunnerServiceRes.Coord }:</b> { mapInfoPointModelList[0].Lat.ToString("F5") } { mapInfoPointModelList[0].Lng.ToString("F5") }</span>");
            //                                        }

            //                                        int CountIssues = 0;
            //                                        foreach (PolSourceObservationIssueModel polSourceObservationIssueModel in polSourceObservationIssueModelList2)
            //                                        {
            //                                            if (polSourceObservationIssueModel.ObservationInfo.Contains(PolSourceObsInfoEnumTextAndID.ID.ToString()))
            //                                            {
            //                                                countPolSourceSiteForPollutionType += 1;
            //                                                CountIssues += 1;
            //                                                List<int> obsInfoList = polSourceObservationIssueModel.ObservationInfo.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Where(c => !string.IsNullOrWhiteSpace(c)).Select(c => int.Parse(c)).ToList();
            //                                                if (CountIssues == 1)
            //                                                {
            //                                                    if (obsInfoList.Count > 1)
            //                                                    {
            //                                                        PolSourceObsInfoEnumTextAndID polSourceObsInfoEnumTextAndID = polSourceObsInfoEnumTextAndIDList.Where(c => c.ID == (int)obsInfoList[1]).FirstOrDefault();
            //                                                        sbTemp.AppendLine($@"           <span><b>{ TaskRunnerServiceRes.Dist }:</b> { (polSourceObsInfoEnumTextAndID != null ? polSourceObsInfoEnumTextAndID.Text : TaskRunnerServiceRes.Empty) }</span>");
            //                                                    }
            //                                                    if (obsInfoList.Count > 2)
            //                                                    {
            //                                                        PolSourceObsInfoEnumTextAndID polSourceObsInfoEnumTextAndID = polSourceObsInfoEnumTextAndIDList.Where(c => c.ID == (int)obsInfoList[2]).FirstOrDefault();
            //                                                        sbTemp.AppendLine($@"           <span><b>{ TaskRunnerServiceRes.Slope }:</b> { (polSourceObsInfoEnumTextAndID != null ? polSourceObsInfoEnumTextAndID.Text : TaskRunnerServiceRes.Empty) }</span>");
            //                                                        polSourceObsInfoEnumTextAndID = polSourceObsInfoEnumTextAndIDList.Where(c => c.ID == (int)obsInfoList[obsInfoList.Count - 1]).FirstOrDefault();
            //                                                        sbTemp.AppendLine($@"           <span><b>{ TaskRunnerServiceRes.Risk }:</b> { (polSourceObsInfoEnumTextAndID != null ? polSourceObsInfoEnumTextAndID.Text : TaskRunnerServiceRes.Empty) }</span>");
            //                                                    }
            //                                                    sbTemp.AppendLine($@"           <br />");
            //                                                }

            //                                                sbTemp.AppendLine($@"&nbsp;&nbsp;&nbsp;&nbsp;{ countPolSourceSiteForPollutionType }) - ");
            //                                                int CountObsInfo = 0;
            //                                                foreach (int obsInfo in obsInfoList)
            //                                                {
            //                                                    CountObsInfo += 1;
            //                                                    if (CountObsInfo > 3 && CountObsInfo < obsInfoList.Count - 2)
            //                                                    {
            //                                                        PolSourceObsInfoEnumTextAndID polSourceObsInfoEnumTextAndID = polSourceObsInfoEnumTextAndIDList.Where(c => c.ID == obsInfo).FirstOrDefault();
            //                                                        string text = polSourceObsInfoEnumTextAndID != null ? polSourceObsInfoEnumTextAndID.Text : TaskRunnerServiceRes.Empty;
            //                                                        sbTemp.AppendLine($@"           <span>{ text }</span> | ");
            //                                                    }
            //                                                }
            //                                                sbTemp.AppendLine($@"           <br />");
            //                                            }
            //                                        }
            //                                        sbTemp.AppendLine($@"           <br />");
            //                                        sbTemp.AppendLine($@"           &nbsp;&nbsp;&nbsp;&nbsp;<span>Photo</span>");
            //                                        sbTemp.AppendLine($@" </blockquote>");
            //                                    }
            //                                }
            //                            }
            //                        }
            //                    }
            //                }
            //            }
            //        }
            //    }
            //}

            //sbTemp.AppendLine(@"<span>|||PageBreak|||</span>");

            //sbTemp.AppendLine($@" <h3>{ TaskRunnerServiceRes.WaterBasePollutionSourceSiteObservationAndIssues }</h3>");

            //sbTemp.AppendLine($@"<div class=""textAlignCenter"">|||Image|FileName,{ fiImageWater.FullName }|width,400|height,150|||</div>");

            //foreach (PolSourceObsInfoEnumTextAndID PolSourceObsInfoEnumTextAndID in polSourceObsInfoEnumTextAndIDListWaterBasePolSourceType)
            //{
            //    sbTemp.AppendLine($@" <h4>{ PolSourceObsInfoEnumTextAndID.Text }</h4>");
            //    foreach (PolSourceSiteModel polSourceSiteModel in polSourceSiteModelList)
            //    {
            //        TVItemModel tvItemModel = tvItemModelListPolSourceSite.Where(c => c.TVItemID == polSourceSiteModel.PolSourceSiteTVItemID).FirstOrDefault();
            //        if (tvItemModel != null && tvItemModel.IsActive)
            //        {
            //            if (polSourceSiteModel != null)
            //            {
            //                PolSourceObservationModel polSourceObservationModel = polSourceObservationModelList.Where(c => c.PolSourceSiteID == polSourceSiteModel.PolSourceSiteID).OrderByDescending(c => c.ObservationDate_Local).FirstOrDefault();
            //                if (polSourceObservationModel != null)
            //                {
            //                    List<PolSourceObservationIssueModel> polSourceObservationIssueModelList2 = polSourceObservationIssueModelList.Where(c => c.PolSourceObservationID == polSourceObservationModel.PolSourceObservationID).OrderBy(c => c.Ordinal).ToList();
            //                    if (polSourceObservationIssueModelList2.Count > 0)
            //                    {
            //                        if (polSourceObservationIssueModelList2[0].ObservationInfo.Contains(((int)PolSourceObsInfoEnum.SourceAnimalPollution).ToString() + ","))
            //                        {
            //                            if (polSourceObservationModel == null)
            //                            {
            //                                sbTemp.AppendLine($@" <blockquote>");
            //                                sbTemp.AppendLine($@" <span>{ TaskRunnerServiceRes.ID }: { polSourceSiteModel.Site }</span>");

            //                                List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(polSourceSiteModel.PolSourceSiteTVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);
            //                                if (mapInfoPointModelList.Count > 0)
            //                                {
            //                                    sbTemp.AppendLine($@" <span>{ TaskRunnerServiceRes.Coord }: { mapInfoPointModelList[0].Lat }, { mapInfoPointModelList[0].Lng }</span>");
            //                                }
            //                                sbTemp.AppendLine($@"           <br />");
            //                                sbTemp.AppendLine($@"           <span>{ TaskRunnerServiceRes.NoObservationForThisPollutionSourceSite }</span>");
            //                                sbTemp.AppendLine($@" </blockquote>");
            //                            }
            //                            else
            //                            {
            //                                if (polSourceObservationIssueModelList2.Count == 0)
            //                                {
            //                                    sbTemp.AppendLine($@" <blockquote>");
            //                                    sbTemp.AppendLine($@" <span>{ TaskRunnerServiceRes.ID }: { polSourceSiteModel.Site }</span>");

            //                                    List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(polSourceSiteModel.PolSourceSiteTVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);
            //                                    if (mapInfoPointModelList.Count > 0)
            //                                    {
            //                                        sbTemp.AppendLine($@" <span>{ TaskRunnerServiceRes.Coord }: { mapInfoPointModelList[0].Lat }, { mapInfoPointModelList[0].Lng }</span>");
            //                                    }
            //                                    sbTemp.AppendLine($@"           <br />");
            //                                    sbTemp.AppendLine($@"           <span>{ polSourceObservationModel.Observation_ToBeDeleted }</span>");
            //                                    sbTemp.AppendLine($@" </blockquote>");
            //                                }
            //                                else
            //                                {
            //                                    bool IssueOfSourceTypeExist = false;
            //                                    foreach (PolSourceObservationIssueModel polSourceObservationIssueModel in polSourceObservationIssueModelList2)
            //                                    {
            //                                        if (polSourceObservationIssueModel.ObservationInfo.Contains(PolSourceObsInfoEnumTextAndID.ID.ToString()))
            //                                        {
            //                                            IssueOfSourceTypeExist = true;
            //                                        }
            //                                    }

            //                                    if (IssueOfSourceTypeExist)
            //                                    {
            //                                        sbTemp.AppendLine($@" <blockquote>");
            //                                        sbTemp.AppendLine($@" <span><b>{ TaskRunnerServiceRes.Site }:</b> { polSourceSiteModel.Site }</span>");

            //                                        List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(polSourceSiteModel.PolSourceSiteTVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);
            //                                        if (mapInfoPointModelList.Count > 0)
            //                                        {
            //                                            sbTemp.AppendLine($@" <span><b>{ TaskRunnerServiceRes.Coord }:</b> { mapInfoPointModelList[0].Lat.ToString("F5") }, { mapInfoPointModelList[0].Lng.ToString("F5") }</span>");
            //                                        }

            //                                        int CountIssues = 0;
            //                                        foreach (PolSourceObservationIssueModel polSourceObservationIssueModel in polSourceObservationIssueModelList2)
            //                                        {
            //                                            if (polSourceObservationIssueModel.ObservationInfo.Contains(PolSourceObsInfoEnumTextAndID.ID.ToString()))
            //                                            {
            //                                                CountIssues += 1;
            //                                                List<int> obsInfoList = polSourceObservationIssueModel.ObservationInfo.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Where(c => !string.IsNullOrWhiteSpace(c)).Select(c => int.Parse(c)).ToList();
            //                                                if (CountIssues == 1)
            //                                                {
            //                                                    if (obsInfoList.Count > 1)
            //                                                    {
            //                                                        PolSourceObsInfoEnumTextAndID polSourceObsInfoEnumTextAndID = polSourceObsInfoEnumTextAndIDList.Where(c => c.ID == (int)obsInfoList[1]).FirstOrDefault();
            //                                                        string text = polSourceObsInfoEnumTextAndID == null ? TaskRunnerServiceRes.Empty : polSourceObsInfoEnumTextAndID.Text;
            //                                                        sbTemp.AppendLine($@"           <span><b>{ TaskRunnerServiceRes.Dist }:</b> { text }</span>");
            //                                                    }
            //                                                    if (obsInfoList.Count > 2)
            //                                                    {
            //                                                        PolSourceObsInfoEnumTextAndID polSourceObsInfoEnumTextAndID = polSourceObsInfoEnumTextAndIDList.Where(c => c.ID == (int)obsInfoList[2]).FirstOrDefault();
            //                                                        string text = polSourceObsInfoEnumTextAndID == null ? TaskRunnerServiceRes.Empty : polSourceObsInfoEnumTextAndID.Text;

            //                                                        sbTemp.AppendLine($@"           <span><b>{ TaskRunnerServiceRes.Slope }:</b> { text }</span>");

            //                                                        PolSourceObsInfoEnumTextAndID polSourceObsInfoEnumTextAndID2 = polSourceObsInfoEnumTextAndIDList.Where(c => c.ID == (int)obsInfoList[obsInfoList.Count - 1]).FirstOrDefault();
            //                                                        string text2 = polSourceObsInfoEnumTextAndID == null ? TaskRunnerServiceRes.Empty : polSourceObsInfoEnumTextAndID.Text;
            //                                                        sbTemp.AppendLine($@"           <span><b>{ TaskRunnerServiceRes.Risk }:</b> { text2 }</span>");
            //                                                    }
            //                                                    sbTemp.AppendLine($@"           <br />");
            //                                                }

            //                                                int CountObsInfo = 0;
            //                                                foreach (int obsInfo in obsInfoList)
            //                                                {
            //                                                    CountObsInfo += 1;
            //                                                    if (CountObsInfo > 3 && CountObsInfo < obsInfoList.Count - 2)
            //                                                    {
            //                                                        PolSourceObsInfoEnumTextAndID polSourceObsInfoEnumTextAndID = polSourceObsInfoEnumTextAndIDList.Where(c => c.ID == obsInfo).FirstOrDefault();
            //                                                        string text = polSourceObsInfoEnumTextAndID == null ? TaskRunnerServiceRes.Empty : polSourceObsInfoEnumTextAndID.Text;

            //                                                        sbTemp.AppendLine($@"           <span>{ text }</span> | ");
            //                                                    }
            //                                                }
            //                                                sbTemp.AppendLine($@"           <br />");
            //                                            }

            //                                        }
            //                                        sbTemp.AppendLine($@"           <span>Photo</span>");
            //                                        sbTemp.AppendLine($@" </blockquote>");
            //                                    }
            //                                }
            //                            }
            //                        }
            //                    }
            //                }
            //            }
            //        }
            //    }
            //}

            //sbTemp.AppendLine(@"<span>|||PageBreak|||</span>");

            if (!GetBottomHTML())
            {
                return false;
            }

            Percent = 98;
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);

            return true;
        }

        // for testing only can comment out when test is completed
        public bool PublicGenerateHTMLSUBSECTOR_POLLUTION_SOURCE_SITES(StringBuilder sbTemp)
        {
            bool retBool = GenerateHTMLSUBSECTOR_POLLUTION_SOURCE_SITES(sbTemp);

            StreamWriter sw = fi.CreateText();
            sw.Write(sbTemp.ToString());
            sw.Flush();
            sw.Close();

            return retBool;
        }
    }
}
