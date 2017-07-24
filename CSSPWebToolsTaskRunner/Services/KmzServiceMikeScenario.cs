using DHI.Generic.MikeZero.DFS.dfsu;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CSSPWebToolsTaskRunner.Services;
using CSSPWebToolsTaskRunner.Services.Resources;
using System.Windows.Forms;
using System.Drawing;
using System.Diagnostics;
using System.Transactions;
using System.Xml;
using System.Threading;
using DHI.Generic.MikeZero.DFS;
using DHI.Generic.MikeZero;
using CSSPWebToolsDBDLL.Models;
using CSSPWebToolsDBDLL;
using CSSPWebToolsDBDLL.Services;
using CSSPWebToolsDBDLL.Services.Resources;
using System.Globalization;
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;
using DHI.PFS;
using CSSPDHI;

namespace CSSPWebToolsTaskRunner.Services

{
    public class KmzServiceMikeScenario
    {
        #region Variables
        private List<Node> InterpolatedContourNodeList = new List<Node>();
        private Dictionary<String, Vector> ForwardVector = new Dictionary<String, Vector>();
        private Dictionary<String, Vector> BackwardVector = new Dictionary<String, Vector>();
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        public MapInfoService _MapInfoService { get; private set; }
        public MikeScenarioService _MikeScenarioService { get; private set; }
        public AppTaskService _AppTaskService { get; private set; }
        public TVItemService _TVItemService { get; private set; }
        public TVFileService _TVFileService { get; private set; }
        #endregion Variables

        #region Constructors
        public KmzServiceMikeScenario(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
            _MapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _MikeScenarioService = new MikeScenarioService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _AppTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _TVItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _TVFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        }
        #endregion Constructors

        #region Functions public
        public void GenerateMikeScenarioBoundaryConditions(FileInfo fi)
        {
            string NotUsed = "";

            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            if (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr)
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-CA");
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("fr-CA");
            }
            else
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-CA");
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-CA");
            }

            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);

            DirectoryInfo di = new DirectoryInfo(ServerFilePath);
            if (!di.Exists)
                di.Create();


            if (fi.Exists)
                fi.Delete();

            FileInfo fiKML = new FileInfo(fi.FullName.Replace(".kmz", ".kml"));

            if (fiKML.Exists)
                fiKML.Delete();

            StringBuilder sbKML = new StringBuilder();

            sbKML.AppendLine(@"<?xml version=""1.0"" encoding=""UTF-8""?>");
            sbKML.AppendLine(@"<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom"">");
            sbKML.AppendLine(@"<Document>");
            sbKML.AppendLine(string.Format(@"<name>{0}</name>", fi.Name));

            string[] Colors = { "ylw", "grn", "blue", "ltblu", "pink", "red" };

            foreach (string color in Colors)
            {
                sbKML.AppendLine(string.Format(@"	<Style id=""sn_{0}-pushpin"">", color));
                sbKML.AppendLine(@"		<IconStyle>");
                sbKML.AppendLine(@"			<scale>1.1</scale>");
                sbKML.AppendLine(@"			<Icon>");
                sbKML.AppendLine(string.Format(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/{0}-pushpin.png</href>", color));
                sbKML.AppendLine(@"			</Icon>");
                sbKML.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
                sbKML.AppendLine(@"		</IconStyle>");
                sbKML.AppendLine(@"		<ListStyle>");
                sbKML.AppendLine(@"		</ListStyle>");
                sbKML.AppendLine(@"	</Style>");
                sbKML.AppendLine(string.Format(@"	<StyleMap id=""msn_{0}-pushpin"">", color));
                sbKML.AppendLine(@"		<Pair>");
                sbKML.AppendLine(@"			<key>normal</key>");
                sbKML.AppendLine(string.Format(@"			<styleUrl>#sn_{0}-pushpin</styleUrl>", color));
                sbKML.AppendLine(@"		</Pair>");
                sbKML.AppendLine(@"		<Pair>");
                sbKML.AppendLine(@"			<key>highlight</key>");
                sbKML.AppendLine(string.Format(@"			<styleUrl>#sh_{0}-pushpin</styleUrl>", color));
                sbKML.AppendLine(@"		</Pair>");
                sbKML.AppendLine(@"	</StyleMap>");
                sbKML.AppendLine(string.Format(@"	<Style id=""sh_{0}-pushpin"">", color));
                sbKML.AppendLine(@"		<IconStyle>");
                sbKML.AppendLine(@"			<scale>1.3</scale>");
                sbKML.AppendLine(@"			<Icon>");
                sbKML.AppendLine(string.Format(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/{0}-pushpin.png</href>", color));
                sbKML.AppendLine(@"			</Icon>");
                sbKML.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
                sbKML.AppendLine(@"		</IconStyle>");
                sbKML.AppendLine(@"		<ListStyle>");
                sbKML.AppendLine(@"		</ListStyle>");
                sbKML.AppendLine(@"	</Style>");
            }

            sbKML.AppendLine(@"<Folder>");
            sbKML.AppendLine(@"<name>" + KmzServiceMikeScenarioRes.Nodes + " (" + KmzServiceMikeScenarioRes.Mesh + @")</name>");

            TVItemModel tvItemModelMikeScenario = _TVItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelMikeScenario.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                return;
            }

            MikeBoundaryConditionService mikeBoundaryConditionService = new MikeBoundaryConditionService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            List<MikeBoundaryConditionModel> mbcModelList = mikeBoundaryConditionService.GetMikeBoundaryConditionModelListWithMikeScenarioTVItemIDAndTVTypeDB(tvItemModelMikeScenario.TVItemID, TVTypeEnum.MikeBoundaryConditionMesh);

            int countColor = 0;
            foreach (MikeBoundaryConditionModel mbcm in mbcModelList)
            {
                sbKML.AppendLine(@"<Folder>");
                sbKML.AppendLine(@"<name>" + mbcm.MikeBoundaryConditionName + " (" + mbcm.MikeBoundaryConditionCode + ") " + mbcm.MikeBoundaryConditionLength_m.ToString("F0") + " (m)</name>");

                // drawing Boundary Nodes
                List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(mbcm.MikeBoundaryConditionTVItemID, TVTypeEnum.MikeBoundaryConditionMesh, MapInfoDrawTypeEnum.Polyline);

                sbKML.AppendLine(@"<Folder>");
                sbKML.AppendLine(@"<name>" + KmzServiceMikeScenarioRes.Nodes + @"</name>");
                sbKML.AppendLine(@"<open>1</open>");
                foreach (MapInfoPointModel mapInfoPointModel in mapInfoPointModelList)
                {
                    sbKML.AppendLine(@"<Placemark>");
                    sbKML.AppendLine(@"<name>" + mapInfoPointModel.Ordinal + "</name>");
                    sbKML.AppendLine(string.Format(@"<styleUrl>#msn_{0}-pushpin</styleUrl>", Colors[countColor]));
                    sbKML.AppendLine(@"<Point>");
                    sbKML.AppendLine(@"<coordinates>" + mapInfoPointModel.Lng.ToString().Replace(",", ".") + @"," + mapInfoPointModel.Lat.ToString().Replace(",", ".") + @",0</coordinates>");
                    sbKML.AppendLine(@"</Point>");
                    sbKML.AppendLine(@"</Placemark>");
                }
                sbKML.AppendLine(@"</Folder>");

                sbKML.AppendLine(@"</Folder>");
                countColor += 1;
            }

            sbKML.AppendLine(@"</Folder>");

            sbKML.AppendLine(@"<Folder>");
            sbKML.AppendLine(@"<name>" + KmzServiceMikeScenarioRes.Nodes + " (" + KmzServiceMikeScenarioRes.WebTide + @")</name>");


            mbcModelList = mikeBoundaryConditionService.GetMikeBoundaryConditionModelListWithMikeScenarioTVItemIDAndTVTypeDB(tvItemModelMikeScenario.TVItemID, TVTypeEnum.MikeBoundaryConditionWebTide);

            countColor = 0;
            foreach (MikeBoundaryConditionModel mbcm in mbcModelList)
            {
                sbKML.AppendLine(@"<Folder>");
                sbKML.AppendLine(@"<name>" + mbcm.MikeBoundaryConditionName + " (" + (mbcm.MikeBoundaryConditionLevelOrVelocity == MikeBoundaryConditionLevelOrVelocityEnum.Level ? KmzServiceMikeScenarioRes.WaterLevels : KmzServiceMikeScenarioRes.Currents) + ") " + mbcm.WebTideDataSet + " " + mbcm.NumberOfWebTideNodes + " " + KmzServiceMikeScenarioRes.Nodes + "</name>");

                // drawing Boundary Nodes
                List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(mbcm.MikeBoundaryConditionTVItemID, TVTypeEnum.MikeBoundaryConditionWebTide, MapInfoDrawTypeEnum.Polyline);

                sbKML.AppendLine(@"<Folder>");
                sbKML.AppendLine(@"<name>" + KmzServiceMikeScenarioRes.Nodes + @"</name>");
                sbKML.AppendLine(@"<open>1</open>");
                foreach (MapInfoPointModel mapInfoPointModel in mapInfoPointModelList)
                {
                    sbKML.AppendLine(@"<Placemark>");
                    sbKML.AppendLine(@"<name>" + mapInfoPointModel.Ordinal + "</name>");
                    sbKML.AppendLine(string.Format(@"<styleUrl>#msn_{0}-pushpin</styleUrl>", Colors[countColor]));
                    sbKML.AppendLine(@"<Point>");
                    sbKML.AppendLine(@"<coordinates>" + mapInfoPointModel.Lng.ToString().Replace(",", ".") + @"," + mapInfoPointModel.Lat.ToString().Replace(",", ".") + @",0</coordinates>");
                    sbKML.AppendLine(@"</Point>");
                    sbKML.AppendLine(@"</Placemark>");
                }
                sbKML.AppendLine(@"</Folder>");

                sbKML.AppendLine(@"</Folder>");
                countColor += 1;
            }

            sbKML.AppendLine(@"</Folder>");

            sbKML.AppendLine(@"</Document>");
            sbKML.AppendLine(@"</kml>");

            SaveInKMZFileStream(fi, fiKML, sbKML);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

        }
        public void GenerateMikeScenarioConcentrationAnimation(FileInfo fi)
        {
            string NotUsed = "";
            //string asliefj = _AppTaskService.GetAppTaskParamStr(_TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "FileGeneratorType");

            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            if (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr)
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-CA");
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("fr-CA");
            }
            else
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-CA");
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-CA");
            }

            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);

            DirectoryInfo di = new DirectoryInfo(ServerFilePath);
            if (!di.Exists)
                di.Create();

            if (fi.Exists)
                fi.Delete();

            MikeScenarioModel SamplingPlanModel = _MikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(SamplingPlanModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, ServiceRes.MikeScenario, ServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", ServiceRes.MikeScenario, ServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                return;
            }


            StringBuilder sbKMZ = new StringBuilder();

        }
        //public void GenerateMikeScenarioConcentrationLimits(FileInfo fi)
        //{
        //    FileInfo fiDfsu;
        //    DfsuFile dfsuFile;
        //    string NotUsed = "";

        //    TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

        //    if (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr)
        //    {
        //        Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-CA");
        //        Thread.CurrentThread.CurrentUICulture = new CultureInfo("fr-CA");
        //    }
        //    else
        //    {
        //        Thread.CurrentThread.CurrentCulture = new CultureInfo("en-CA");
        //        Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-CA");
        //    }

        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);

        //    DirectoryInfo di = new DirectoryInfo(ServerFilePath);
        //    if (!di.Exists)
        //        di.Create();


        //    if (fi.Exists)
        //        fi.Delete();

        //    FileInfo fiKML = new FileInfo(fi.FullName.Replace(".kmz", ".kml"));

        //    if (fiKML.Exists)
        //        fiKML.Delete();

        //    MikeScenarioModel mikeScenarioModel = _MikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //    if (!string.IsNullOrWhiteSpace(mikeScenarioModel.Error))
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, ServiceRes.MikeScenario, ServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", ServiceRes.MikeScenario, ServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        return;
        //    }

        //    TVFileModel tvFileModelm21_3fm = tvFileService.GetTVFileModelWithTVItemIDAndTVFileTypeM21FMOrM3FMDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);

        //    if (tvFileModelm21_3fm == null)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.FileType, ".m21fm, .m3fm");
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.FileType, ".m21fm, .m3fm");
        //        return;
        //    }

        //    string ShortFileName = tvFileModelm21_3fm.ServerFileName.Replace(".m21fm", "").Replace(".m3fm", "");
        //    FileInfo fiM21_3FM = new FileInfo(tvFileService.ChoseEDriveOrCDrive(tvFileModelm21_3fm.ServerFilePath) + tvFileModelm21_3fm.ServerFileName);

        //    M21_3FMService m21_3fmInput = new M21_3FMService(_TaskRunnerBaseService);
        //    m21_3fmInput.Read_M21_3FM_File(fiM21_3FM);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //    {
        //        return;
        //    }

        //    string TransFileName = m21_3fmInput.femEngineHD.transport_module.outputs.output.First().Value.file_name.Replace("'", "");

        //    fiDfsu = new FileInfo(TransFileName);

        //    try
        //    {
        //        dfsuFile = DfsuFile.Open(fiDfsu.FullName);
        //    }
        //    catch (Exception ex)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotReadFile_Error_, fiDfsu.FullName, ex.Message + (ex.InnerException != null ? ex.InnerException.Message : ""));
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotReadFile_Error_", fiDfsu.FullName, ex.Message);
        //        return;
        //    }


        //    List<Element> ElementList = new List<Element>();
        //    List<Node> NodeList = new List<Node>();

        //    // ----------------------------------------
        //    // filling the ElementList and The NodeList
        //    // ----------------------------------------

        //    FillElementListAndNodeList(dfsuFile, ElementList, NodeList);

        //    List<float> ContourValueList = new List<float>();

        //    string ContourValueString = _AppTaskService.GetAppTaskParamStr(_TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "ContourValues");

        //    try
        //    {
        //        if (ContourValueString.Trim().Length != 0)
        //        {
        //            foreach (string s in ContourValueString.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries))
        //            {
        //                ContourValueList.Add(float.Parse(s));
        //            }
        //        }
        //    }
        //    catch (Exception)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.PlsEnterValidNumSeparatedByCommasFor_, TaskRunnerServiceRes.PollutionContour);
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("PlsEnterValidNumSeparatedByCommasFor_", TaskRunnerServiceRes.PollutionContour);
        //        return;
        //    }

        //    #region SigmaZLayerDepth
        //    string SigmaLayerValuesString = "";
        //    string ZLayerValuesString = "";
        //    string DepthValuesString = "";

        //    List<int> SigmaLayerValueList = new List<int>();
        //    List<int> ZLayerValueList = new List<int>();
        //    List<float> DepthValueList = new List<float>();

        //    if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
        //    {
        //        try
        //        {
        //            if (SigmaLayerValuesString.Trim().Length != 0)
        //            {
        //                foreach (string s in SigmaLayerValuesString.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries))
        //                {
        //                    SigmaLayerValueList.Add(int.Parse(s));
        //                }
        //            }
        //        }
        //        catch (Exception)
        //        {
        //            NotUsed = TaskRunnerServiceRes.PlsEnterValidNumSeparatedByCommasForSigmaLayers;
        //            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("PlsEnterValidNumSeparatedByCommasForSigmaLayers");
        //            return;
        //        }

        //        try
        //        {
        //            if (ZLayerValuesString.Trim().Length != 0)
        //            {
        //                foreach (string s in ZLayerValuesString.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries))
        //                {
        //                    ZLayerValueList.Add(int.Parse(s));
        //                }
        //            }
        //        }
        //        catch (Exception)
        //        {
        //            NotUsed = TaskRunnerServiceRes.PlsEnterValidNumSeparatedByCommasForZLayers;
        //            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("PlsEnterValidNumSeparatedByCommasForZLayers");
        //            return;
        //        }

        //        try
        //        {
        //            if (DepthValuesString.Trim().Length != 0)
        //            {
        //                foreach (string s in DepthValuesString.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries))
        //                {
        //                    DepthValueList.Add(float.Parse(s));
        //                }
        //            }
        //        }
        //        catch (Exception)
        //        {
        //            NotUsed = TaskRunnerServiceRes.PlsEnterValidNumSeparatedByCommasForDepth;
        //            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("PlsEnterValidNumSeparatedByCommasForDepth");
        //            return;
        //        }
        //    }
        //    else if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu2D)
        //    {
        //        SigmaLayerValueList.Add(0);
        //    }
        //    else
        //    {
        //        NotUsed = TaskRunnerServiceRes.NeedToImplement;
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("NeedToImplement");
        //        return;
        //    }
        //    #endregion SigmaZLayerDepth

        //    List<NodeLayer> TopNodeLayerList = new List<NodeLayer>();
        //    List<NodeLayer> BottomNodeLayerList = new List<NodeLayer>();
        //    List<ElementLayer> ElementLayerList = new List<ElementLayer>();

        //    FillElementLayerList(dfsuFile, SigmaLayerValueList, ElementList, ElementLayerList, TopNodeLayerList, BottomNodeLayerList);

        //    StringBuilder sbKML = new StringBuilder();

        //    sbKML.AppendLine(@"<?xml version=""1.0"" encoding=""UTF-8""?>");
        //    sbKML.AppendLine(@"<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom"">");
        //    sbKML.AppendLine(@"<Document>");
        //    sbKML.AppendLine(string.Format(@"<name>{0}</name>", fi.Name));

        //    int ItemNumber = 0;

        //    // getting the ItemNumber
        //    foreach (IDfsSimpleDynamicItemInfo dfsDyInfo in dfsuFile.ItemInfo)
        //    {
        //        if (dfsDyInfo.Quantity.Item == eumItem.eumIConcentration)
        //        {
        //            ItemNumber = dfsDyInfo.ItemNumber;
        //            break;
        //        }
        //    }

        //    if (ItemNumber == 0)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind__, TaskRunnerServiceRes.ParameterType, eumItem.eumIConcentration.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind__", TaskRunnerServiceRes.ParameterType, eumItem.eumIConcentration.ToString());
        //        return;
        //    }

        //    #region Contour Style
        //    sbKML.AppendLine(@"	<StyleMap id=""msn_ylw-pushpin"">");
        //    sbKML.AppendLine(@"		<Pair>");
        //    sbKML.AppendLine(@"			<key>normal</key>");
        //    sbKML.AppendLine(@"			<styleUrl>#sn_ylw-pushpin</styleUrl>");
        //    sbKML.AppendLine(@"		</Pair>");
        //    sbKML.AppendLine(@"		<Pair>");
        //    sbKML.AppendLine(@"			<key>highlight</key>");
        //    sbKML.AppendLine(@"			<styleUrl>#sh_ylw-pushpin</styleUrl>");
        //    sbKML.AppendLine(@"		</Pair>");
        //    sbKML.AppendLine(@"	</StyleMap>");
        //    sbKML.AppendLine(@"	<Style id=""sn_ylw-pushpin"">");
        //    sbKML.AppendLine(@"		<IconStyle>");
        //    sbKML.AppendLine(@"			<scale>1.1</scale>");
        //    sbKML.AppendLine(@"			<Icon>");
        //    sbKML.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>");
        //    sbKML.AppendLine(@"			</Icon>");
        //    sbKML.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
        //    sbKML.AppendLine(@"		</IconStyle>");
        //    sbKML.AppendLine(@"      <LineStyle>");
        //    sbKML.AppendLine(@"         <color>ff000000</color>");
        //    sbKML.AppendLine(@"       </LineStyle>");
        //    sbKML.AppendLine(@"	</Style>");
        //    sbKML.AppendLine(@"	<Style id=""sh_ylw-pushpin"">");
        //    sbKML.AppendLine(@"		<IconStyle>");
        //    sbKML.AppendLine(@"			<scale>1.3</scale>");
        //    sbKML.AppendLine(@"			<Icon>");
        //    sbKML.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>");
        //    sbKML.AppendLine(@"			</Icon>");
        //    sbKML.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
        //    sbKML.AppendLine(@"		</IconStyle>");
        //    sbKML.AppendLine(@"      <LineStyle>");
        //    sbKML.AppendLine(@"         <color>ff000000</color>");
        //    sbKML.AppendLine(@"       </LineStyle>");
        //    sbKML.AppendLine(@"	</Style>");

        //    sbKML.AppendLine(@"	<StyleMap id=""msn_grn-pushpin"">");
        //    sbKML.AppendLine(@"		<Pair>");
        //    sbKML.AppendLine(@"			<key>normal</key>");
        //    sbKML.AppendLine(@"			<styleUrl>#sn_grn-pushpin</styleUrl>");
        //    sbKML.AppendLine(@"		</Pair>");
        //    sbKML.AppendLine(@"		<Pair>");
        //    sbKML.AppendLine(@"			<key>highlight</key>");
        //    sbKML.AppendLine(@"			<styleUrl>#sh_grn-pushpin</styleUrl>");
        //    sbKML.AppendLine(@"		</Pair>");
        //    sbKML.AppendLine(@"	</StyleMap>");
        //    sbKML.AppendLine(@"	<Style id=""sn_grn-pushpin"">");
        //    sbKML.AppendLine(@"		<IconStyle>");
        //    sbKML.AppendLine(@"			<scale>1.1</scale>");
        //    sbKML.AppendLine(@"			<Icon>");
        //    sbKML.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/grn-pushpin.png</href>");
        //    sbKML.AppendLine(@"			</Icon>");
        //    sbKML.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
        //    sbKML.AppendLine(@"		</IconStyle>");
        //    sbKML.AppendLine(@"      <LineStyle>");
        //    sbKML.AppendLine(@"         <color>ff000000</color>");
        //    sbKML.AppendLine(@"       </LineStyle>");
        //    sbKML.AppendLine(@"	</Style>");
        //    sbKML.AppendLine(@"	<Style id=""sh_grn-pushpin"">");
        //    sbKML.AppendLine(@"		<IconStyle>");
        //    sbKML.AppendLine(@"			<scale>1.3</scale>");
        //    sbKML.AppendLine(@"			<Icon>");
        //    sbKML.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/grn-pushpin.png</href>");
        //    sbKML.AppendLine(@"			</Icon>");
        //    sbKML.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
        //    sbKML.AppendLine(@"		</IconStyle>");
        //    sbKML.AppendLine(@"      <LineStyle>");
        //    sbKML.AppendLine(@"         <color>ff000000</color>");
        //    sbKML.AppendLine(@"       </LineStyle>");
        //    sbKML.AppendLine(@"	</Style>");

        //    sbKML.AppendLine(@"	<StyleMap id=""msn_blue-pushpin"">");
        //    sbKML.AppendLine(@"		<Pair>");
        //    sbKML.AppendLine(@"			<key>normal</key>");
        //    sbKML.AppendLine(@"			<styleUrl>#sn_blue-pushpin</styleUrl>");
        //    sbKML.AppendLine(@"		</Pair>");
        //    sbKML.AppendLine(@"		<Pair>");
        //    sbKML.AppendLine(@"			<key>highlight</key>");
        //    sbKML.AppendLine(@"			<styleUrl>#sh_blue-pushpin</styleUrl>");
        //    sbKML.AppendLine(@"		</Pair>");
        //    sbKML.AppendLine(@"	</StyleMap>");
        //    sbKML.AppendLine(@"	<Style id=""sn_blue-pushpin"">");
        //    sbKML.AppendLine(@"		<IconStyle>");
        //    sbKML.AppendLine(@"			<scale>1.1</scale>");
        //    sbKML.AppendLine(@"			<Icon>");
        //    sbKML.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/blue-pushpin.png</href>");
        //    sbKML.AppendLine(@"			</Icon>");
        //    sbKML.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
        //    sbKML.AppendLine(@"		</IconStyle>");
        //    sbKML.AppendLine(@"      <LineStyle>");
        //    sbKML.AppendLine(@"         <color>ff000000</color>");
        //    sbKML.AppendLine(@"       </LineStyle>");
        //    sbKML.AppendLine(@"	</Style>");
        //    sbKML.AppendLine(@"	<Style id=""sh_blue-pushpin"">");
        //    sbKML.AppendLine(@"		<IconStyle>");
        //    sbKML.AppendLine(@"			<scale>1.3</scale>");
        //    sbKML.AppendLine(@"			<Icon>");
        //    sbKML.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/blue-pushpin.png</href>");
        //    sbKML.AppendLine(@"			</Icon>");
        //    sbKML.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
        //    sbKML.AppendLine(@"		</IconStyle>");
        //    sbKML.AppendLine(@"      <LineStyle>");
        //    sbKML.AppendLine(@"         <color>ff000000</color>");
        //    sbKML.AppendLine(@"       </LineStyle>");
        //    sbKML.AppendLine(@"	</Style>");

        //    sbKML.AppendLine(@"<Style id=""fc_LT14"">");
        //    sbKML.AppendLine(@"<LineStyle>");
        //    sbKML.AppendLine(@"<color>ff000000</color>");
        //    sbKML.AppendLine(@"</LineStyle>");
        //    sbKML.AppendLine(@"<PolyStyle>");
        //    sbKML.AppendLine(@"<color>6600ff00</color>");
        //    sbKML.AppendLine(@"<outline>1</outline>");
        //    sbKML.AppendLine(@"</PolyStyle>");
        //    sbKML.AppendLine(@"</Style>");

        //    sbKML.AppendLine(@"<Style id=""fc_14"">");
        //    sbKML.AppendLine(@"<LineStyle>");
        //    sbKML.AppendLine(@"<color>ff000000</color>");
        //    sbKML.AppendLine(@"</LineStyle>");
        //    sbKML.AppendLine(@"<PolyStyle>");
        //    sbKML.AppendLine(@"<color>66ff0000</color>");
        //    sbKML.AppendLine(@"<outline>1</outline>");
        //    sbKML.AppendLine(@"</PolyStyle>");
        //    sbKML.AppendLine(@"</Style>");

        //    sbKML.AppendLine(@"<Style id=""fc_88"">");
        //    sbKML.AppendLine(@"<LineStyle>");
        //    sbKML.AppendLine(@"<color>ff000000</color>");
        //    sbKML.AppendLine(@"</LineStyle>");
        //    sbKML.AppendLine(@"<PolyStyle>");
        //    sbKML.AppendLine(@"<color>660000ff</color>");
        //    sbKML.AppendLine(@"<outline>1</outline>");
        //    sbKML.AppendLine(@"</PolyStyle>");
        //    sbKML.AppendLine(@"</Style>");
        //    #endregion Contour Style
        //    #region Model Input Style
        //    sbKML.AppendLine(@"<Style id=""sn_grn-pushpin"">");
        //    sbKML.AppendLine(@"<IconStyle>");
        //    sbKML.AppendLine(@"<scale>1.1</scale>");
        //    sbKML.AppendLine(@"<Icon>");
        //    sbKML.AppendLine(@"<href>http://maps.google.com/mapfiles/kml/pushpin/grn-pushpin.png</href>");
        //    sbKML.AppendLine(@"</Icon>");
        //    sbKML.AppendLine(@"<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
        //    sbKML.AppendLine(@"</IconStyle>");
        //    sbKML.AppendLine(@"<ListStyle>");
        //    sbKML.AppendLine(@"</ListStyle>");
        //    sbKML.AppendLine(@"</Style>");

        //    sbKML.AppendLine(@"<Style id=""sh_grn-pushpin"">");
        //    sbKML.AppendLine(@"<IconStyle>");
        //    sbKML.AppendLine(@"<scale>1.3</scale>");
        //    sbKML.AppendLine(@"<Icon>");
        //    sbKML.AppendLine(@"<href>http://maps.google.com/mapfiles/kml/pushpin/grn-pushpin.png</href>");
        //    sbKML.AppendLine(@"</Icon>");
        //    sbKML.AppendLine(@"<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
        //    sbKML.AppendLine(@"</IconStyle>");
        //    sbKML.AppendLine(@"<ListStyle>");
        //    sbKML.AppendLine(@"</ListStyle>");
        //    sbKML.AppendLine(@"</Style>");

        //    sbKML.AppendLine(@"<StyleMap id=""msn_grn-pushpin"">");
        //    sbKML.AppendLine(@"<Pair>");
        //    sbKML.AppendLine(@"<key>normal</key>");
        //    sbKML.AppendLine(@"<styleUrl>#sn_grn-pushpin</styleUrl>");
        //    sbKML.AppendLine(@"</Pair>");
        //    sbKML.AppendLine(@"<Pair>");
        //    sbKML.AppendLine(@"<key>highlight</key>");
        //    sbKML.AppendLine(@"<styleUrl>#sh_grn-pushpin</styleUrl>");
        //    sbKML.AppendLine(@"</Pair>");
        //    sbKML.AppendLine(@"</StyleMap>");

        //    sbKML.AppendLine(@"<Style id=""sn_red-pushpin"">");
        //    sbKML.AppendLine(@"<IconStyle>");
        //    sbKML.AppendLine(@"<scale>1.1</scale>");
        //    sbKML.AppendLine(@"<Icon>");
        //    sbKML.AppendLine(@"<href>http://maps.google.com/mapfiles/kml/pushpin/red-pushpin.png</href>");
        //    sbKML.AppendLine(@"</Icon>");
        //    sbKML.AppendLine(@"<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
        //    sbKML.AppendLine(@"</IconStyle>");
        //    sbKML.AppendLine(@"<ListStyle>");
        //    sbKML.AppendLine(@"</ListStyle>");
        //    sbKML.AppendLine(@"</Style>");

        //    sbKML.AppendLine(@"<Style id=""sh_red-pushpin"">");
        //    sbKML.AppendLine(@"<IconStyle>");
        //    sbKML.AppendLine(@"<scale>1.3</scale>");
        //    sbKML.AppendLine(@"<Icon>");
        //    sbKML.AppendLine(@"<href>http://maps.google.com/mapfiles/kml/pushpin/red-pushpin.png</href>");
        //    sbKML.AppendLine(@"</Icon>");
        //    sbKML.AppendLine(@"<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
        //    sbKML.AppendLine(@"</IconStyle>");
        //    sbKML.AppendLine(@"<ListStyle>");
        //    sbKML.AppendLine(@"</ListStyle>");
        //    sbKML.AppendLine(@"</Style>");

        //    sbKML.AppendLine(@"<StyleMap id=""msn_red-pushpin"">");
        //    sbKML.AppendLine(@"<Pair>");
        //    sbKML.AppendLine(@"<key>normal</key>");
        //    sbKML.AppendLine(@"<styleUrl>#sn_red-pushpin</styleUrl>");
        //    sbKML.AppendLine(@"</Pair>");
        //    sbKML.AppendLine(@"<Pair>");
        //    sbKML.AppendLine(@"<key>highlight</key>");
        //    sbKML.AppendLine(@"<styleUrl>#sh_red-pushpin</styleUrl>");
        //    sbKML.AppendLine(@"</Pair>");
        //    sbKML.AppendLine(@"</StyleMap>");

        //    sbKML.AppendLine(@"<Style id=""sn_blue-pushpin"">");
        //    sbKML.AppendLine(@"<IconStyle>");
        //    sbKML.AppendLine(@"<scale>1.1</scale>");
        //    sbKML.AppendLine(@"<Icon>");
        //    sbKML.AppendLine(@"<href>http://maps.google.com/mapfiles/kml/pushpin/blue-pushpin.png</href>");
        //    sbKML.AppendLine(@"</Icon>");
        //    sbKML.AppendLine(@"<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
        //    sbKML.AppendLine(@"</IconStyle>");
        //    sbKML.AppendLine(@"<ListStyle>");
        //    sbKML.AppendLine(@"</ListStyle>");
        //    sbKML.AppendLine(@"</Style>");

        //    sbKML.AppendLine(@"<Style id=""sh_blue-pushpin"">");
        //    sbKML.AppendLine(@"<IconStyle>");
        //    sbKML.AppendLine(@"<scale>1.3</scale>");
        //    sbKML.AppendLine(@"<Icon>");
        //    sbKML.AppendLine(@"<href>http://maps.google.com/mapfiles/kml/pushpin/blue-pushpin.png</href>");
        //    sbKML.AppendLine(@"</Icon>");
        //    sbKML.AppendLine(@"<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
        //    sbKML.AppendLine(@"</IconStyle>");
        //    sbKML.AppendLine(@"<ListStyle>");
        //    sbKML.AppendLine(@"</ListStyle>");
        //    sbKML.AppendLine(@"</Style>");

        //    sbKML.AppendLine(@"<StyleMap id=""msn_blue-pushpin"">");
        //    sbKML.AppendLine(@"<Pair>");
        //    sbKML.AppendLine(@"<key>normal</key>");
        //    sbKML.AppendLine(@"<styleUrl>#sn_blue-pushpin</styleUrl>");
        //    sbKML.AppendLine(@"</Pair>");
        //    sbKML.AppendLine(@"<Pair>");
        //    sbKML.AppendLine(@"<key>highlight</key>");
        //    sbKML.AppendLine(@"<styleUrl>#sh_blue-pushpin</styleUrl>");
        //    sbKML.AppendLine(@"</Pair>");
        //    sbKML.AppendLine(@"</StyleMap>");
        //    #endregion Model Input Style

        //    //int pcount = 0;
        //    sbKML.AppendLine(@"<Folder><name>" + TaskRunnerServiceRes.PollutionLimits + @"</name>");
        //    sbKML.AppendLine(@"<visibility>0</visibility>");
        //    int CountLayer = 1;
        //    int CountAt = 0;

        //    foreach (int Layer in SigmaLayerValueList)
        //    {
        //        #region Top of Layer
        //        sbKML.AppendLine(string.Format(@"<Folder><name>" + TaskRunnerServiceRes.TopOfLayer + @" [{0}]</name>", Layer));
        //        sbKML.AppendLine(@"<visibility>0</visibility>");
        //        int CountContour = 1;
        //        foreach (float ContourValue in ContourValueList)
        //        {
        //            CountAt += 1;
        //            sbKML.AppendLine(string.Format(@"<Folder><name>" + TaskRunnerServiceRes.ContourValue + @" [{0}]</name>", ContourValue));
        //            sbKML.AppendLine(@"<visibility>0</visibility>");
        //            string AppTaskStatus = ((int)((CountAt * 100) / (SigmaLayerValueList.Count * ContourValueList.Count))).ToString() + " %";

        //            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, ((int)((CountAt * 100) / (SigmaLayerValueList.Count * ContourValueList.Count))));

        //            List<Node> AllNodeList = new List<Node>();
        //            List<ContourPolygon> ContourPolygonList = new List<ContourPolygon>();

        //            for (int timeStep = 0; timeStep < dfsuFile.NumberOfTimeSteps; timeStep++)
        //            {

        //                float[] ValueList = (float[])dfsuFile.ReadItemTimeStep(ItemNumber, timeStep).Data;

        //                for (int i = 0; i < ElementLayerList.Count; i++)
        //                {
        //                    if (ElementLayerList[i].Element.Value < ValueList[i])
        //                    {
        //                        ElementLayerList[i].Element.Value = ValueList[i];
        //                    }
        //                }
        //            }
        //            //}

        //            foreach (NodeLayer nl in TopNodeLayerList)
        //            {
        //                float Total = 0;
        //                foreach (Element element in nl.Node.ElementList)
        //                {
        //                    Total += element.Value;
        //                }
        //                nl.Node.Value = Total / nl.Node.ElementList.Count;
        //            }

        //            List<NodeLayer> AboveNodeLayerList = new List<NodeLayer>();

        //            AboveNodeLayerList = (from n in TopNodeLayerList
        //                                  where (n.Node.Value >= ContourValue)
        //                                  && n.Layer == Layer
        //                                  select n).ToList<NodeLayer>();

        //            foreach (NodeLayer snl in AboveNodeLayerList)
        //            {
        //                List<NodeLayer> EndNodeLayerList = null;

        //                List<NodeLayer> NodeLayerConnectedList = (from nll in TopNodeLayerList
        //                                                          from n in snl.Node.ConnectNodeList
        //                                                          where (n.ID == nll.Node.ID)
        //                                                          select nll).ToList<NodeLayer>();

        //                EndNodeLayerList = (from nll in NodeLayerConnectedList
        //                                    where (nll.Node.ID != snl.Node.ID)
        //                                    && (nll.Node.Value < ContourValue)
        //                                    && nll.Layer == Layer
        //                                    select nll).ToList<NodeLayer>();

        //                foreach (NodeLayer en in EndNodeLayerList)
        //                {
        //                    AllNodeList.Add(en.Node);
        //                }

        //                if (snl.Node.Code != 0)
        //                {
        //                    AllNodeList.Add(snl.Node);
        //                }

        //            }

        //            //if (AllNodeList.Count == 0)
        //            //{
        //            //    continue;
        //            //}

        //            List<Element> TempUniqueElementList = new List<Element>();
        //            List<Element> UniqueElementList = new List<Element>();
        //            foreach (ElementLayer el in ElementLayerList.Where(l => l.Layer == Layer))
        //            {
        //                if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
        //                {
        //                    if (el.Element.Type == 32)
        //                    {
        //                        bool NodeBigger = false;
        //                        for (int i = 3; i < 6; i++)
        //                        {
        //                            if (el.Element.NodeList[i].Value >= ContourValue)
        //                            {
        //                                NodeBigger = true;
        //                                break;
        //                            }
        //                        }
        //                        if (NodeBigger)
        //                        {
        //                            int countTrue = 0;
        //                            for (int i = 3; i < 6; i++)
        //                            {
        //                                if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
        //                                {
        //                                    countTrue += 1;
        //                                }
        //                            }
        //                            if (countTrue != el.Element.NodeList.Count)
        //                            {
        //                                TempUniqueElementList.Add(el.Element);
        //                            }
        //                        }
        //                    }
        //                    else if (el.Element.Type == 33)
        //                    {
        //                        bool NodeBigger = false;
        //                        for (int i = 4; i < 8; i++)
        //                        {
        //                            if (el.Element.NodeList[i].Value >= ContourValue)
        //                            {
        //                                NodeBigger = true;
        //                                break;
        //                            }
        //                        }
        //                        if (NodeBigger)
        //                        {
        //                            int countTrue = 0;
        //                            for (int i = 4; i < 8; i++)
        //                            {
        //                                if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
        //                                {
        //                                    countTrue += 1;
        //                                }
        //                            }
        //                            if (countTrue != el.Element.NodeList.Count)
        //                            {
        //                                TempUniqueElementList.Add(el.Element);
        //                            }
        //                        }
        //                    }
        //                    else
        //                    {
        //                        NotUsed = string.Format(TaskRunnerServiceRes.ElementType_IsNotSupported, el.Element.Type.ToString());
        //                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ElementType_IsNotSupported", el.Element.Type.ToString());
        //                        return;
        //                    }
        //                }
        //                else if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu2D)
        //                {
        //                    bool NodeBigger = false;
        //                    for (int i = 0; i < el.Element.NodeList.Count; i++)
        //                    {
        //                        if (el.Element.NodeList[i].Value >= ContourValue)
        //                        {
        //                            NodeBigger = true;
        //                            break;
        //                        }
        //                    }
        //                    if (NodeBigger)
        //                    {
        //                        int countTrue = 0;
        //                        for (int i = 0; i < el.Element.NodeList.Count; i++)
        //                        {
        //                            if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
        //                            {
        //                                countTrue += 1;
        //                            }
        //                        }
        //                        if (countTrue != el.Element.NodeList.Count)
        //                        {
        //                            TempUniqueElementList.Add(el.Element);
        //                        }
        //                    }
        //                }
        //            }

        //            UniqueElementList = (from el in TempUniqueElementList select el).Distinct().ToList<Element>();

        //            // filling InterpolatedContourNodeList
        //            InterpolatedContourNodeList = new List<Node>();

        //            foreach (Element el in UniqueElementList)
        //            {
        //                if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
        //                {
        //                    if (el.Type == 32)
        //                    {
        //                        if (el.NodeList[3].Value >= ContourValue && el.NodeList[4].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[3], el.NodeList[4], ContourValue);
        //                        }
        //                        if (el.NodeList[3].Value >= ContourValue && el.NodeList[5].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[3], el.NodeList[5], ContourValue);
        //                        }
        //                        if (el.NodeList[4].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[4], el.NodeList[3], ContourValue);
        //                        }
        //                        if (el.NodeList[4].Value >= ContourValue && el.NodeList[5].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[4], el.NodeList[5], ContourValue);
        //                        }
        //                        if (el.NodeList[5].Value >= ContourValue && el.NodeList[4].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[5], el.NodeList[4], ContourValue);
        //                        }
        //                        if (el.NodeList[5].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[5], el.NodeList[3], ContourValue);
        //                        }
        //                    }
        //                    else if (el.Type == 33)
        //                    {
        //                        if (el.NodeList[4].Value >= ContourValue && el.NodeList[5].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[4], el.NodeList[5], ContourValue);
        //                        }
        //                        if (el.NodeList[4].Value >= ContourValue && el.NodeList[7].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[4], el.NodeList[7], ContourValue);
        //                        }
        //                        if (el.NodeList[5].Value >= ContourValue && el.NodeList[4].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[5], el.NodeList[4], ContourValue);
        //                        }
        //                        if (el.NodeList[5].Value >= ContourValue && el.NodeList[6].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[5], el.NodeList[6], ContourValue);
        //                        }
        //                        if (el.NodeList[6].Value >= ContourValue && el.NodeList[5].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[6], el.NodeList[5], ContourValue);
        //                        }
        //                        if (el.NodeList[6].Value >= ContourValue && el.NodeList[7].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[6], el.NodeList[7], ContourValue);
        //                        }
        //                        if (el.NodeList[7].Value >= ContourValue && el.NodeList[4].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[7], el.NodeList[4], ContourValue);
        //                        }
        //                        if (el.NodeList[7].Value >= ContourValue && el.NodeList[6].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[7], el.NodeList[6], ContourValue);
        //                        }
        //                    }
        //                    else
        //                    {
        //                        NotUsed = string.Format(TaskRunnerServiceRes.ElementType_IsNotSupported, el.Type.ToString());
        //                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ElementType_IsNotSupported", el.Type.ToString());
        //                        return;
        //                    }
        //                }
        //                else if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu2D)
        //                {
        //                    if (el.Type == 21)
        //                    {
        //                        if (el.NodeList[0].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[0], el.NodeList[1], ContourValue);
        //                        }
        //                        if (el.NodeList[0].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[0], el.NodeList[2], ContourValue);
        //                        }
        //                        if (el.NodeList[1].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[1], el.NodeList[0], ContourValue);
        //                        }
        //                        if (el.NodeList[1].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[1], el.NodeList[2], ContourValue);
        //                        }
        //                        if (el.NodeList[2].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[2], el.NodeList[1], ContourValue);
        //                        }
        //                        if (el.NodeList[2].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[2], el.NodeList[0], ContourValue);
        //                        }
        //                    }
        //                    else if (el.Type == 24)
        //                    {
        //                    }
        //                    else if (el.Type == 25)
        //                    {
        //                        if (el.NodeList[0].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[0], el.NodeList[1], ContourValue);
        //                        }
        //                        if (el.NodeList[0].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[0], el.NodeList[3], ContourValue);
        //                        }
        //                        if (el.NodeList[1].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[1], el.NodeList[0], ContourValue);
        //                        }
        //                        if (el.NodeList[1].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[1], el.NodeList[2], ContourValue);
        //                        }
        //                        if (el.NodeList[2].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[2], el.NodeList[1], ContourValue);
        //                        }
        //                        if (el.NodeList[2].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[2], el.NodeList[3], ContourValue);
        //                        }
        //                        if (el.NodeList[3].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[3], el.NodeList[0], ContourValue);
        //                        }
        //                        if (el.NodeList[3].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
        //                        {
        //                            InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[3], el.NodeList[2], ContourValue);
        //                        }
        //                    }
        //                    else
        //                    {
        //                        NotUsed = string.Format(TaskRunnerServiceRes.ElementType_IsNotSupported, el.Type.ToString());
        //                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ElementType_IsNotSupported", el.Type.ToString());
        //                        return;
        //                    }
        //                }
        //            }

        //            List<Node> UniqueNodeList = (from n in AllNodeList orderby n.ID select n).Distinct().ToList<Node>();

        //            ForwardVector = new Dictionary<String, Vector>();
        //            BackwardVector = new Dictionary<String, Vector>();

        //            // ------------------------- new code --------------------------
        //            //                     

        //            foreach (Element el in UniqueElementList)
        //            {
        //                if (el.Type == 21)
        //                {
        //                    FillVectors21_32(el, UniqueElementList, ContourValue, false, true);
        //                }
        //                else if (el.Type == 24)
        //                {
        //                    NotUsed = string.Format(TaskRunnerServiceRes.ElementType_IsNotSupported, el.Type.ToString());
        //                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ElementType_IsNotSupported", el.Type.ToString());
        //                    return;
        //                }
        //                else if (el.Type == 25)
        //                {
        //                    FillVectors25_33(el, UniqueElementList, ContourValue, false, true);
        //                }
        //                else if (el.Type == 32)
        //                {
        //                    FillVectors21_32(el, UniqueElementList, ContourValue, true, true);
        //                }
        //                else if (el.Type == 33)
        //                {
        //                    FillVectors25_33(el, UniqueElementList, ContourValue, true, true);
        //                }
        //                else
        //                {
        //                    NotUsed = string.Format(TaskRunnerServiceRes.ElementType_IsNotSupported, el.Type.ToString());
        //                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ElementType_IsNotSupported", el.Type.ToString());
        //                    return;
        //                }

        //            }

        //            //-------------- new code ------------------------


        //            bool MoreContourLine = true;
        //            while (MoreContourLine && ForwardVector.Count > 0)
        //            {
        //                List<Node> FinalContourNodeList = new List<Node>();
        //                Vector LastVector = new Vector();
        //                LastVector = ForwardVector.First().Value;
        //                FinalContourNodeList.Add(LastVector.StartNode);
        //                FinalContourNodeList.Add(LastVector.EndNode);
        //                ForwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
        //                BackwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
        //                bool PolygonCompleted = false;
        //                while (!PolygonCompleted)
        //                {
        //                    List<string> KeyStrList = (from k in ForwardVector.Keys
        //                                               where k.StartsWith(LastVector.EndNode.ID.ToString() + ",")
        //                                               && !k.EndsWith("," + LastVector.StartNode.ID.ToString())
        //                                               select k).ToList<string>();

        //                    if (KeyStrList.Count != 1)
        //                    {
        //                        KeyStrList = (from k in BackwardVector.Keys
        //                                      where k.StartsWith(LastVector.EndNode.ID.ToString() + ",")
        //                                      && !k.EndsWith("," + LastVector.StartNode.ID.ToString())
        //                                      select k).ToList<string>();

        //                        if (KeyStrList.Count != 1)
        //                        {
        //                            PolygonCompleted = true;
        //                            break;
        //                        }
        //                        else
        //                        {
        //                            LastVector = BackwardVector[KeyStrList[0]];
        //                            BackwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
        //                            ForwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
        //                        }
        //                    }
        //                    else
        //                    {
        //                        LastVector = ForwardVector[KeyStrList[0]];
        //                        ForwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
        //                        BackwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
        //                    }
        //                    FinalContourNodeList.Add(LastVector.EndNode);
        //                    if (FinalContourNodeList[FinalContourNodeList.Count - 1] == FinalContourNodeList[0])
        //                    {
        //                        PolygonCompleted = true;
        //                    }
        //                }

        //                if (mapInfoService.CalculateAreaOfPolygon(FinalContourNodeList) < 0)
        //                {
        //                    FinalContourNodeList.Reverse();
        //                }

        //                FinalContourNodeList.Add(FinalContourNodeList[0]);
        //                ContourPolygon contourPolygon = new ContourPolygon() { };
        //                contourPolygon.ContourNodeList = FinalContourNodeList;
        //                contourPolygon.ContourValue = ContourValue;
        //                ContourPolygonList.Add(contourPolygon);

        //                if (ForwardVector.Count == 0)
        //                {
        //                    MoreContourLine = false;
        //                }
        //            }

        //            foreach (ContourPolygon contourPolygon in ContourPolygonList)
        //            {
        //                sbKML.AppendLine(@"<Folder>");
        //                sbKML.AppendLine(@"<visibility>0</visibility>");
        //                sbKML.AppendLine(@"<Placemark>");
        //                sbKML.AppendLine(@"<visibility>0</visibility>");
        //                if (contourPolygon.ContourValue >= 14 && contourPolygon.ContourValue < 88)
        //                {
        //                    sbKML.AppendLine(@"<styleUrl>#fc_14</styleUrl>");
        //                }
        //                else if (contourPolygon.ContourValue >= 88)
        //                {
        //                    sbKML.AppendLine(@"<styleUrl>#fc_88</styleUrl>");
        //                }
        //                else
        //                {
        //                    sbKML.AppendLine(@"<styleUrl>#fc_LT14</styleUrl>");
        //                }
        //                sbKML.AppendLine(@"<Polygon>");
        //                sbKML.AppendLine(@"<outerBoundaryIs>");
        //                sbKML.AppendLine(@"<LinearRing>");
        //                sbKML.AppendLine(@"<coordinates>");
        //                foreach (Node node in contourPolygon.ContourNodeList)
        //                {
        //                    sbKML.Append(node.X.ToString().Replace(",", ".") + @"," + node.Y.ToString().Replace(",", ".") + ",0 ");
        //                }
        //                sbKML.AppendLine(@"</coordinates>");
        //                sbKML.AppendLine(@"</LinearRing>");
        //                sbKML.AppendLine(@"</outerBoundaryIs>");
        //                sbKML.AppendLine(@"</Polygon>");
        //                sbKML.AppendLine(@"</Placemark>");
        //                sbKML.AppendLine(@"</Folder>");
        //            }
        //            sbKML.AppendLine(@"</Folder>");
        //            CountContour += 1;
        //        }
        //        sbKML.AppendLine(@"</Folder>");
        //        #endregion Top of Layer

        //        #region Bottom of Layer
        //        //// 
        //        //if (Layer == dfsuFile.NumberOfSigmaLayers)
        //        //{
        //        //    sbKMLPollutionLimitsContour.AppendLine(string.Format(@"<Folder><name>Bottom of Layer [{0}]</name>", Layer));
        //        //    sbKMLPollutionLimitsContour.AppendLine(@"<visibility>0</visibility>");
        //        //    CountContour = 1;
        //        //    foreach (float ContourValue in ContourValueList)
        //        //    {
        //        //        CountAt += 1;
        //        //        sbKMLPollutionLimitsContour.AppendLine(string.Format(@"<Folder><name>Contour Value [{0}]</name>", ContourValue));
        //        //        sbKMLPollutionLimitsContour.AppendLine(@"<visibility>0</visibility>");
        //        //        string AppTaskStatus = ((int)((CountAt * 100) / (SigmaLayerValueList.Count * ContourValueList.Count))).ToString() + " %";
        //        //        UpdateTask(AppTaskID, AppTaskStatus);

        //        //        List<Node> AllNodeList = new List<Node>();
        //        //        List<ContourPolygon> ContourPolygonList = new List<ContourPolygon>();

        //        //        //foreach (Dfs.Parameter.TimeSeriesValue v in p.TimeSeriesValueList)
        //        //        //{
        //        //        for (int timeStep = 0; timeStep < dfsuFile.NumberOfTimeSteps; timeStep++)
        //        //        {

        //        //            float[] ValueList = (float[])dfsuFile.ReadItemTimeStep(ItemNumber, timeStep).Data;

        //        //            for (int i = 0; i < ElementLayerList.Count; i++)
        //        //            {
        //        //                if (ElementLayerList[i].Element.Value < ValueList[i])
        //        //                {
        //        //                    ElementLayerList[i].Element.Value = ValueList[i];
        //        //                }
        //        //            }
        //        //        }
        //        //        //}

        //        //        foreach (NodeLayer nl in BottomNodeLayerList)
        //        //        {
        //        //            float Total = 0;
        //        //            foreach (Element element in nl.Node.ElementList)
        //        //            {
        //        //                Total += element.Value;
        //        //            }
        //        //            nl.Node.Value = Total / nl.Node.ElementList.Count;
        //        //        }

        //        //        List<NodeLayer> AboveNodeLayerList = new List<NodeLayer>();

        //        //        AboveNodeLayerList = (from n in BottomNodeLayerList
        //        //                              where (n.Node.Value >= ContourValue)
        //        //                              && n.Layer == Layer
        //        //                              select n).ToList<NodeLayer>();

        //        //        foreach (NodeLayer snl in AboveNodeLayerList)
        //        //        {
        //        //            List<NodeLayer> EndNodeLayerList = null;

        //        //            List<NodeLayer> NodeLayerConnectedList = (from nll in BottomNodeLayerList
        //        //                                                      from n in snl.Node.ConnectNodeList
        //        //                                                      where (n.ID == nll.Node.ID)
        //        //                                                      select nll).ToList<NodeLayer>();

        //        //            EndNodeLayerList = (from nll in NodeLayerConnectedList
        //        //                                where (nll.Node.ID != snl.Node.ID)
        //        //                                && (nll.Node.Value < ContourValue)
        //        //                                && nll.Layer == Layer
        //        //                                select nll).ToList<NodeLayer>();

        //        //            foreach (NodeLayer en in EndNodeLayerList)
        //        //            {
        //        //                AllNodeList.Add(en.Node);
        //        //            }

        //        //            if (snl.Node.Code != 0)
        //        //            {
        //        //                AllNodeList.Add(snl.Node);
        //        //            }

        //        //        }

        //        //        if (AllNodeList.Count == 0)
        //        //        {
        //        //            continue;
        //        //        }

        //        //        List<Element> TempUniqueElementList = new List<Element>();
        //        //        List<Element> UniqueElementList = new List<Element>();
        //        //        foreach (ElementLayer el in ElementLayerList.Where(l => l.Layer == Layer))
        //        //        {
        //        //            if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
        //        //            {
        //        //                if (el.Element.Type == 32)
        //        //                {
        //        //                    bool NodeBigger = false;
        //        //                    for (int i = 0; i < 3; i++)
        //        //                    {
        //        //                        if (el.Element.NodeList[i].Value >= ContourValue)
        //        //                        {
        //        //                            NodeBigger = true;
        //        //                            break;
        //        //                        }
        //        //                    }
        //        //                    if (NodeBigger)
        //        //                    {
        //        //                        int countTrue = 0;
        //        //                        for (int i = 0; i < 3; i++)
        //        //                        {
        //        //                            if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
        //        //                            {
        //        //                                countTrue += 1;
        //        //                            }
        //        //                        }
        //        //                        if (countTrue != el.Element.NodeList.Count)
        //        //                        {
        //        //                            TempUniqueElementList.Add(el.Element);
        //        //                        }
        //        //                    }
        //        //                }
        //        //                else if (el.Element.Type == 33)
        //        //                {
        //        //                    bool NodeBigger = false;
        //        //                    for (int i = 0; i < 4; i++)
        //        //                    {
        //        //                        if (el.Element.NodeList[i].Value >= ContourValue)
        //        //                        {
        //        //                            NodeBigger = true;
        //        //                            break;
        //        //                        }
        //        //                    }
        //        //                    if (NodeBigger)
        //        //                    {
        //        //                        int countTrue = 0;
        //        //                        for (int i = 0; i < 4; i++)
        //        //                        {
        //        //                            if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
        //        //                            {
        //        //                                countTrue += 1;
        //        //                            }
        //        //                        }
        //        //                        if (countTrue != el.Element.NodeList.Count)
        //        //                        {
        //        //                            TempUniqueElementList.Add(el.Element);
        //        //                        }
        //        //                    }
        //        //                }
        //        //                else
        //        //                {
        //        //                    UpdateTask(AppTaskID, "");
        //        //                    throw new Exception("Element type is not supported: Element type = [" + el.Element.Type + "]");
        //        //                }
        //        //            }
        //        //            else if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu2D)
        //        //            {
        //        //                bool NodeBigger = false;
        //        //                for (int i = 0; i < el.Element.NodeList.Count; i++)
        //        //                {
        //        //                    if (el.Element.NodeList[i].Value >= ContourValue)
        //        //                    {
        //        //                        NodeBigger = true;
        //        //                        break;
        //        //                    }
        //        //                }
        //        //                if (NodeBigger)
        //        //                {
        //        //                    int countTrue = 0;
        //        //                    for (int i = 0; i < el.Element.NodeList.Count; i++)
        //        //                    {
        //        //                        if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
        //        //                        {
        //        //                            countTrue += 1;
        //        //                        }
        //        //                    }
        //        //                    if (countTrue != el.Element.NodeList.Count)
        //        //                    {
        //        //                        TempUniqueElementList.Add(el.Element);
        //        //                    }
        //        //                }
        //        //            }
        //        //        }

        //        //        UniqueElementList = (from el in TempUniqueElementList select el).Distinct().ToList<Element>();

        //        //        // filling InterpolatedContourNodeList
        //        //        InterpolatedContourNodeList = new List<Node>();

        //        //        foreach (Element el in UniqueElementList)
        //        //        {
        //        //            if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
        //        //            {
        //        //                if (el.Type == 32)
        //        //                {
        //        //                    if (el.NodeList[0].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
        //        //                    {
        //        //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[0], el.NodeList[1], ContourValue);
        //        //                    }
        //        //                    if (el.NodeList[0].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
        //        //                    {
        //        //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[0], el.NodeList[2], ContourValue);
        //        //                    }
        //        //                    if (el.NodeList[1].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
        //        //                    {
        //        //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[1], el.NodeList[0], ContourValue);
        //        //                    }
        //        //                    if (el.NodeList[1].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
        //        //                    {
        //        //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[1], el.NodeList[2], ContourValue);
        //        //                    }
        //        //                    if (el.NodeList[2].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
        //        //                    {
        //        //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[2], el.NodeList[1], ContourValue);
        //        //                    }
        //        //                    if (el.NodeList[2].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
        //        //                    {
        //        //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[2], el.NodeList[0], ContourValue);
        //        //                    }
        //        //                }
        //        //                else if (el.Type == 33)
        //        //                {
        //        //                    if (el.NodeList[0].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
        //        //                    {
        //        //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[0], el.NodeList[1], ContourValue);
        //        //                    }
        //        //                    if (el.NodeList[0].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
        //        //                    {
        //        //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[0], el.NodeList[3], ContourValue);
        //        //                    }
        //        //                    if (el.NodeList[1].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
        //        //                    {
        //        //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[1], el.NodeList[0], ContourValue);
        //        //                    }
        //        //                    if (el.NodeList[1].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
        //        //                    {
        //        //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[1], el.NodeList[2], ContourValue);
        //        //                    }
        //        //                    if (el.NodeList[2].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
        //        //                    {
        //        //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[2], el.NodeList[1], ContourValue);
        //        //                    }
        //        //                    if (el.NodeList[2].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
        //        //                    {
        //        //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[2], el.NodeList[3], ContourValue);
        //        //                    }
        //        //                    if (el.NodeList[3].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
        //        //                    {
        //        //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[3], el.NodeList[0], ContourValue);
        //        //                    }
        //        //                    if (el.NodeList[3].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
        //        //                    {
        //        //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[3], el.NodeList[2], ContourValue);
        //        //                    }
        //        //                }
        //        //                else
        //        //                {
        //        //                    UpdateTask(AppTaskID, "");
        //        //                    throw new Exception("Element type is not supported: Element type = [" + el.Type + "]");
        //        //                }
        //        //            }
        //        //            else
        //        //            {
        //        //                UpdateTask(AppTaskID, "");
        //        //                throw new Exception("Bottom does not exist outside the Dfsu3DSigma and Dfsu3DSigmaZ.");
        //        //            }
        //        //        }

        //        //        List<Node> UniqueNodeList = (from n in AllNodeList orderby n.ID select n).Distinct().ToList<Node>();

        //        //        ForwardVector = new Dictionary<String, Vector>();
        //        //        BackwardVector = new Dictionary<String, Vector>();

        //        //        // ------------------------- new code --------------------------
        //        //        //                     

        //        //        foreach (Element el in UniqueElementList)
        //        //        {
        //        //            if (el.Type == 21)
        //        //            {
        //        //                FillVectors21_32(el, UniqueElementList, ContourValue, AppTaskID, false, false);
        //        //            }
        //        //            else if (el.Type == 24)
        //        //            {
        //        //                UpdateTask(AppTaskID, "");
        //        //                throw new Exception("Element type is not supported: Element type = [" + el.Type + "]");
        //        //            }
        //        //            else if (el.Type == 25)
        //        //            {
        //        //                FillVectors25_33(el, UniqueElementList, ContourValue, AppTaskID, false, false);
        //        //            }
        //        //            else if (el.Type == 32)
        //        //            {
        //        //                FillVectors21_32(el, UniqueElementList, ContourValue, AppTaskID, true, false);
        //        //            }
        //        //            else if (el.Type == 33)
        //        //            {
        //        //                FillVectors25_33(el, UniqueElementList, ContourValue, AppTaskID, true, false);
        //        //            }
        //        //            else
        //        //            {
        //        //                UpdateTask(AppTaskID, "");
        //        //                throw new Exception("Element type is not supported: Element type = [" + el.Type + "]");
        //        //            }

        //        //        }

        //        //        //-------------- new code ------------------------


        //        //        bool MoreContourLine = true;
        //        //        while (MoreContourLine)
        //        //        {
        //        //            List<Node> FinalContourNodeList = new List<Node>();
        //        //            Vector LastVector = new Vector();
        //        //            LastVector = ForwardVector.First().Value;
        //        //            FinalContourNodeList.Add(LastVector.StartNode);
        //        //            FinalContourNodeList.Add(LastVector.EndNode);
        //        //            ForwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
        //        //            BackwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
        //        //            bool PolygonCompleted = false;
        //        //            while (!PolygonCompleted)
        //        //            {
        //        //                List<string> KeyStrList = (from k in ForwardVector.Keys
        //        //                                           where k.StartsWith(LastVector.EndNode.ID.ToString() + ",")
        //        //                                           && !k.EndsWith("," + LastVector.StartNode.ID.ToString())
        //        //                                           select k).ToList<string>();

        //        //                if (KeyStrList.Count != 1)
        //        //                {
        //        //                    KeyStrList = (from k in BackwardVector.Keys
        //        //                                  where k.StartsWith(LastVector.EndNode.ID.ToString() + ",")
        //        //                                  && !k.EndsWith("," + LastVector.StartNode.ID.ToString())
        //        //                                  select k).ToList<string>();

        //        //                    if (KeyStrList.Count != 1)
        //        //                    {
        //        //                        PolygonCompleted = true;
        //        //                        break;
        //        //                    }
        //        //                    else
        //        //                    {
        //        //                        LastVector = BackwardVector[KeyStrList[0]];
        //        //                        BackwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
        //        //                        ForwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
        //        //                    }
        //        //                }
        //        //                else
        //        //                {
        //        //                    LastVector = ForwardVector[KeyStrList[0]];
        //        //                    ForwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
        //        //                    BackwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
        //        //                }
        //        //                FinalContourNodeList.Add(LastVector.EndNode);
        //        //                if (FinalContourNodeList[FinalContourNodeList.Count - 1] == FinalContourNodeList[0])
        //        //                {
        //        //                    PolygonCompleted = true;
        //        //                }
        //        //            }

        //        //            if (CalculateAreaOfPolygon(FinalContourNodeList) < 0)
        //        //            {
        //        //                FinalContourNodeList.Reverse();
        //        //            }

        //        //            FinalContourNodeList.Add(FinalContourNodeList[0]);
        //        //            ContourPolygon contourPolygon = new ContourPolygon() { };
        //        //            contourPolygon.ContourNodeList = FinalContourNodeList;
        //        //            contourPolygon.ContourValue = ContourValue;
        //        //            ContourPolygonList.Add(contourPolygon);

        //        //            if (ForwardVector.Count == 0)
        //        //            {
        //        //                MoreContourLine = false;
        //        //            }
        //        //        }
        //        //        //sbKMLPollutionLimitsContour.AppendLine(@"<Folder>");
        //        //        //sbKMLPollutionLimitsContour.AppendLine(string.Format(@"<name>{0} Pollution Limits Contour</name>", ContourValue));
        //        //        //sbKMLPollutionLimitsContour.AppendLine(@"<visibility>0</visibility>");

        //        //        foreach (ContourPolygon contourPolygon in ContourPolygonList)
        //        //        {
        //        //            sbKMLPollutionLimitsContour.AppendLine(@"<Folder>");
        //        //            sbKMLPollutionLimitsContour.AppendLine(@"<visibility>0</visibility>");
        //        //            sbKMLPollutionLimitsContour.AppendLine(@"<Placemark>");
        //        //            sbKMLPollutionLimitsContour.AppendLine(@"<visibility>0</visibility>");
        //        //            if (contourPolygon.ContourValue >= 14 && contourPolygon.ContourValue < 88)
        //        //            {
        //        //                sbKMLPollutionLimitsContour.AppendLine(@"<styleUrl>#fc_14</styleUrl>");
        //        //            }
        //        //            else if (contourPolygon.ContourValue >= 88)
        //        //            {
        //        //                sbKMLPollutionLimitsContour.AppendLine(@"<styleUrl>#fc_88</styleUrl>");
        //        //            }
        //        //            else
        //        //            {
        //        //                sbKMLPollutionLimitsContour.AppendLine(@"<styleUrl>#fc_LT14</styleUrl>");
        //        //            }
        //        //            sbKMLPollutionLimitsContour.AppendLine(@"<Polygon>");
        //        //            sbKMLPollutionLimitsContour.AppendLine(@"<outerBoundaryIs>");
        //        //            sbKMLPollutionLimitsContour.AppendLine(@"<LinearRing>");
        //        //            sbKMLPollutionLimitsContour.AppendLine(@"<coordinates>");
        //        //            foreach (Node node in contourPolygon.ContourNodeList)
        //        //            {
        //        //                sbKMLPollutionLimitsContour.Append(string.Format(@"{0},{1},0 ", node.X, node.Y));
        //        //            }
        //        //            sbKMLPollutionLimitsContour.AppendLine(@"</coordinates>");
        //        //            sbKMLPollutionLimitsContour.AppendLine(@"</LinearRing>");
        //        //            sbKMLPollutionLimitsContour.AppendLine(@"</outerBoundaryIs>");
        //        //            sbKMLPollutionLimitsContour.AppendLine(@"</Polygon>");
        //        //            sbKMLPollutionLimitsContour.AppendLine(@"</Placemark>");
        //        //            sbKMLPollutionLimitsContour.AppendLine(@"</Folder>");
        //        //        }
        //        //        sbKMLPollutionLimitsContour.AppendLine(@"</Folder>");
        //        //        CountContour += 1;
        //        //    }
        //        //    sbKMLPollutionLimitsContour.AppendLine(@"</Folder>");
        //        //}
        //        #endregion Bottom of Layer
        //        CountLayer += 1;
        //    }
        //    sbKML.AppendLine(@"</Folder>");

        //    #region Source Description
        //    List<MikeSourceModel> mikeSourceModelList = _MikeScenarioService._MikeSourceService.GetMikeSourceModelListWithMikeScenarioTVItemIDDB(mikeScenarioModel.MikeScenarioTVItemID);
        //    if (mikeSourceModelList.Count == 0)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBeNullOrEmpty, TaskRunnerServiceRes.mikeSourceModelList);
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_ShouldNotBeNullOrEmpty", TaskRunnerServiceRes.mikeSourceModelList);
        //        return;
        //    }

        //    sbKML.AppendLine(@"<description><![CDATA[");
        //    sbKML.AppendLine(string.Format(@"<h2>{0}</h2>", TaskRunnerServiceRes.ModelParameters));
        //    sbKML.AppendLine(@"<ul>");
        //    sbKML.AppendLine(string.Format(@"<li><b>{0}:</b> {1:yyyy/MM/dd HH:mm:ss tt} (UTC)</li>", TaskRunnerServiceRes.CreatedDate, m21_3fmInput.topfileinfo.Created));
        //    sbKML.AppendLine(string.Format(@"<li><b>{0}:</b> {1:yyyy/MM/dd HH:mm:ss tt} (UTC)</li>", TaskRunnerServiceRes.ScenarioStartTime, m21_3fmInput.femEngineHD.time.start_time));
        //    sbKML.AppendLine(string.Format(@"<li><b>{0}:</b> {1:yyyy/MM/dd HH:mm:ss tt} (UTC)</li>", TaskRunnerServiceRes.ScenarioEndTime, m21_3fmInput.femEngineHD.time.start_time.AddSeconds(m21_3fmInput.femEngineHD.time.time_step_interval * m21_3fmInput.femEngineHD.time.number_of_time_steps)));
        //    sbKML.AppendLine(@"</ul>");
        //    sbKML.AppendLine(@"<ul>");

        //    foreach (float cv in ContourValueList)
        //    {
        //        if (cv >= 14 && cv < 88)
        //        {
        //            sbKML.AppendLine(string.Format(@"<li><span style=""background-color:Blue; color:White"">{0} = {1:F0}</span</li>", TaskRunnerServiceRes.FCMPNPollutionContour, cv));
        //        }
        //        else if (cv >= 88)
        //        {
        //            sbKML.AppendLine(string.Format(@"<li><span style=""background-color:Red; color:White"">{0} = {1:F0}</span></li>", TaskRunnerServiceRes.FCMPNPollutionContour, cv));
        //        }
        //        else
        //        {
        //            sbKML.AppendLine(string.Format(@"<li><span style=""background-color:Green; color:White"">{0} = {1:F0}</span></li>", TaskRunnerServiceRes.FCMPNPollutionContour, cv));
        //        }

        //    }
        //    sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.AverageDecayFactor + @":</b> " + (m21_3fmInput.femEngineHD.transport_module.decay.component["COMPONENT_1"].constant_value * 3600 * 24).ToString("F6").Replace(",", ".") + @" /" + TaskRunnerServiceRes.DayLowerCase + @"</li>");
        //    if (m21_3fmInput.femEngineHD.transport_module.decay.component["COMPONENT_1"].format == 0)
        //    {
        //        sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.DecayIsConstant + @"</b></li>");
        //    }
        //    else
        //    {
        //        sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.DecayIsVariable + @"</b></li>");
        //        sbKML.AppendLine(@"<ul><li><b>" + TaskRunnerServiceRes.Amplitude + @":</b> " + ((double)mikeScenarioModel.DecayFactorAmplitude).ToString("F6").Replace(",", ".") + @"</li></ul>");
        //    }
        //    if (m21_3fmInput.femEngineHD.hydrodynamic_module.wind_forcing.constant_speed > 0)
        //    {
        //        sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.Wind + @":</b> " + (m21_3fmInput.femEngineHD.hydrodynamic_module.wind_forcing.constant_speed * 3.6).ToString("F1").Replace(",", ".") + @" (km/h)   " + (m21_3fmInput.femEngineHD.hydrodynamic_module.wind_forcing.constant_speed).ToString("F1").Replace(",", ".") + @" (m/s)</li>");
        //        sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.WindDirection + @":</b> " + (m21_3fmInput.femEngineHD.hydrodynamic_module.wind_forcing.constant_direction).ToString("F1").Replace(",", ".") + @" " + TaskRunnerServiceRes.DegreeLowerCase + " (0 = " + TaskRunnerServiceRes.NorthClockwiseLowerCase + @")</li>");
        //    }
        //    else
        //    {
        //        sbKML.AppendLine("<li><b>" + TaskRunnerServiceRes.NoWind + @"</b></li>");
        //    }
        //    switch (m21_3fmInput.femEngineHD.hydrodynamic_module.density.type)
        //    {
        //        case M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.DENSITY.TYPE.Barotropic:
        //        case M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.DENSITY.TYPE.FunctionOfTemperatureAndSalinity:
        //        case M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.DENSITY.TYPE.FunctionOfTemperature:
        //        case M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.DENSITY.TYPE.FunctionOfSalinity:
        //            sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.DensityType + @":</b> {0}</li>", GetSpaceBeforeCapital(m21_3fmInput.femEngineHD.hydrodynamic_module.density.type.ToString())));
        //            sbKML.AppendLine("<ul>");
        //            sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.TemperatureReference + @":</b> " + (m21_3fmInput.femEngineHD.hydrodynamic_module.density.temperature_reference).ToString("F1").Replace(",", ".") + TaskRunnerServiceRes.Celcius + @"</li>");
        //            sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.SalinityReference + @":</b> " + (m21_3fmInput.femEngineHD.hydrodynamic_module.density.salinity_reference).ToString("F1").Replace(",", ".") + TaskRunnerServiceRes.PSU + @"</li>");
        //            sbKML.AppendLine("</ul>");
        //            break;
        //        default:
        //            sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.DensityTypeNotDefine + @"</b></li>");
        //            break;
        //    }
        //    switch (m21_3fmInput.femEngineHD.hydrodynamic_module.bed_resistance.type)
        //    {
        //        case M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.BED_RESISTANCE.TYPE.NoBedResistance:
        //            sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.BedResistanceType + @":</b> <ul><li>" + GetSpaceBeforeCapital(m21_3fmInput.femEngineHD.hydrodynamic_module.bed_resistance.type.ToString()) + ": " + (m21_3fmInput.femEngineHD.hydrodynamic_module.bed_resistance.manning_number.constant_value).ToString().Replace(",", ".") + @"</li></ul></li>");
        //            break;
        //        case M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.BED_RESISTANCE.TYPE.ChezyNumber:
        //            sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.BedResistanceType + @":</b> <ul><li>" + GetSpaceBeforeCapital(m21_3fmInput.femEngineHD.hydrodynamic_module.bed_resistance.type.ToString()) + ": " + (m21_3fmInput.femEngineHD.hydrodynamic_module.bed_resistance.chezy_number.constant_value).ToString().Replace(",", ".") + @"</li></ul></li>");
        //            break;
        //        case M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.BED_RESISTANCE.TYPE.ManningNumber:
        //            sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.BedResistanceType + @":</b> <ul><li>" + GetSpaceBeforeCapital(m21_3fmInput.femEngineHD.hydrodynamic_module.bed_resistance.type.ToString()) + ": " + (m21_3fmInput.femEngineHD.hydrodynamic_module.bed_resistance.manning_number.constant_value).ToString().Replace(",", ".") + @"</li></ul></li>");
        //            break;
        //        default:
        //            sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.BedResistanceType + @":</b> " + TaskRunnerServiceRes.NotDetermined + @"</li>");
        //            break;
        //    }
        //    sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.ResultFrequency + @":</b> {0:F0} {1}</li>", mikeScenarioModel.ResultFrequency_min, TaskRunnerServiceRes.MinutesLowerCase));
        //    sbKML.AppendLine(@"</ul>");

        //    // showing not used sources 
        //    foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.SOURCES.SOURCE> kvp in m21_3fmInput.femEngineHD.hydrodynamic_module.sources.source)
        //    {
        //        M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.SOURCES.SOURCE shd = kvp.Value;
        //        if (kvp.Value.include == 1)
        //        {
        //            if (shd.Name.Substring(1, shd.Name.Length - 2).ToLower().StartsWith("river"))
        //            {
        //                sbKML.AppendLine(string.Format(@"<h2 style='Color: Blue'>{0} ({1})</h2>", shd.Name.Substring(1, shd.Name.Length - 2), TaskRunnerServiceRes.IncludedLowerCase));
        //            }
        //            else
        //            {
        //                sbKML.AppendLine(string.Format(@"<h2 style='Color: Green'>{0} ({1})</h2>", shd.Name.Substring(1, shd.Name.Length - 2), TaskRunnerServiceRes.IncludedLowerCase));
        //            }
        //        }
        //        else
        //        {
        //            sbKML.AppendLine(string.Format(@"<h2 style='Color: Red'>{0} ({1})</h2>", shd.Name.Substring(1, shd.Name.Length - 2), TaskRunnerServiceRes.NotIncludedLowerCase));
        //        }
        //        sbKML.AppendLine(@"<h3>" + TaskRunnerServiceRes.Effluent + @"</h3>");
        //        sbKML.AppendLine(@"<ul>");
        //        sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.Coordinates + @"</b>");
        //        sbKML.AppendLine(string.Format(@"&nbsp;&nbsp;&nbsp; {0:F5} &nbsp; {1:F5}</li><br />", shd.coordinates.y, shd.coordinates.x).Replace(",", "."));

        //        MikeSourceModel mikeSourceModelLocal = (from msl in mikeSourceModelList
        //                                                where msl.SourceNumberString == kvp.Key
        //                                                select msl).FirstOrDefault<MikeSourceModel>();

        //        if (!string.IsNullOrWhiteSpace(mikeSourceModelLocal.Error))
        //        {
        //            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, kvp.Key);
        //            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", kvp.Key);
        //            return;
        //        }

        //        if ((bool)mikeSourceModelLocal.IsContinuous)
        //        {
        //            sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.IsContinuous + @"</b></li>");
        //            sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.Flow + @":</b> " + (shd.constant_value).ToString("F6").Replace(",", ".") + " (m3/s)  " + (shd.constant_value * 24 * 3600).ToString("F1").Replace(",", ".") + @" (m3/" + TaskRunnerServiceRes.DayLowerCase + @")</li>");
        //            sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.FCMPNPer100ML + @":</b> " + (m21_3fmInput.femEngineHD.transport_module.sources.source[kvp.Key].component["COMPONENT_1"].constant_value).ToString("F0").Replace(",", ".") + @"</li>");
        //            sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.Temperature + @":</b> " + (m21_3fmInput.femEngineHD.hydrodynamic_module.temperature_salinity_module.sources.source[kvp.Key].temperature.constant_value).ToString("F1").Replace(",", ".") + @" " + TaskRunnerServiceRes.Celcius + @"</li>");
        //            sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.Salinity + @":</b> " + (m21_3fmInput.femEngineHD.hydrodynamic_module.temperature_salinity_module.sources.source[kvp.Key].salinity.constant_value).ToString("F1").Replace(",", ".") + @" " + TaskRunnerServiceRes.PSU + @"</li>");
        //        }
        //        else
        //        {
        //            sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.IsNotContinuous + @"</b></li>");

        //            MikeSourceStartEndService mikeSourceStartEndService = new MikeSourceStartEndService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //            List<MikeSourceStartEndModel> mikeSourceStartEndModelListLocal = mikeSourceStartEndService.GetMikeSourceStartEndModelListWithMikeSourceIDDB(mikeSourceModelLocal.MikeSourceID);
        //            int CountMikeSourceStartEnd = 0;
        //            foreach (MikeSourceStartEndModel mssem in mikeSourceStartEndModelListLocal)
        //            {
        //                CountMikeSourceStartEnd += 1;
        //                sbKML.AppendLine(@"<br /><b>" + TaskRunnerServiceRes.Spill + @": " + CountMikeSourceStartEnd + "</b><br />");
        //                sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.SpillStartTime + @":</b> {0:yyyy/MM/dd HH:mm:ss tt} (UTC)</li>", mssem.StartDateAndTime_Local));
        //                sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.SpillEndTime + @":</b> {0:yyyy/MM/dd HH:mm:ss tt} (UTC)</li>", mssem.EndDateAndTime_Local));
        //                sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.FlowStart + @":</b> " + ((double)mssem.SourceFlowStart_m3_day / 24 / 3600).ToString("F6").Replace(",", ".") + @" (m3/s)  " + ((double)mssem.SourceFlowStart_m3_day).ToString("F0").Replace(",", ".") + @" (m3/" + TaskRunnerServiceRes.DayLowerCase + @")</li>");
        //                sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.FlowEnd + @":</b> " + ((double)mssem.SourceFlowEnd_m3_day / 24 / 3600).ToString("F6").Replace(",", ".") + @" (m3/s)  " + ((double)mssem.SourceFlowEnd_m3_day).ToString("F0").Replace(",", ".") + @" (m3/" + TaskRunnerServiceRes.DayLowerCase + @")</li>");
        //                sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.FCMPNPer100MLStart + @":</b> " + ((double)mssem.SourcePollutionStart_MPN_100ml).ToString("F0").Replace(",", ".") + @"</li>");
        //                sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.FCMPNPer100MLEnd + @":</b> " + ((double)mssem.SourcePollutionEnd_MPN_100ml).ToString("F0").Replace(",", ".") + @"</li>");
        //                sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.TemperatureStart + @":</b> " + ((double)mssem.SourceTemperatureStart_C).ToString("F0").Replace(",", ".") + @" " + TaskRunnerServiceRes.Celcius + @"</li>");
        //                sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.TempertureEnd + @":</b> " + ((double)mssem.SourceTemperatureEnd_C).ToString("F0").Replace(",", ".") + @" " + TaskRunnerServiceRes.Celcius + @"</li>");
        //                sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.SalinityStart + @":</b> " + ((double)mssem.SourceSalinityStart_PSU).ToString("F0").Replace(",", ".") + @" " + TaskRunnerServiceRes.PSU + @"</li>");
        //                sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.SalinityEnd + @":</b> " + ((double)mssem.SourceSalinityEnd_PSU).ToString("F0").Replace(",", ".") + @" " + TaskRunnerServiceRes.PSU + @"</li>");
        //            }
        //        }
        //        sbKML.AppendLine(@"</ul>");
        //    }


        //    sbKML.AppendLine(@"<iframe src=""about:"" width=""600"" height=""1"" />");
        //    sbKML.AppendLine(@"]]></description>");
        //    #endregion Source Description

        //    sbKML.Append("<Folder><name>" + TaskRunnerServiceRes.ModelInput + "</name>");
        //    sbKML.AppendLine(@"<visibility>0</visibility>");

        //    sbKML.Append("<Folder><name>" + TaskRunnerServiceRes.SourceIncluded + @"</name>");
        //    sbKML.AppendLine(@"<visibility>0</visibility>");

        //    // showing used sources 
        //    foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.SOURCES.SOURCE> kvp in m21_3fmInput.femEngineHD.hydrodynamic_module.sources.source)
        //    {
        //        M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.SOURCES.SOURCE shd = kvp.Value;
        //        if (shd.include == 1)
        //        {
        //            #region Source Placemark
        //            MikeSourceModel mikeSourceModelLocal = (from msl in mikeSourceModelList
        //                                                    where msl.SourceNumberString == kvp.Key
        //                                                    select msl).FirstOrDefault<MikeSourceModel>();

        //            if (mikeSourceModelLocal == null)
        //            {
        //                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeSource, TaskRunnerServiceRes.SourceNumberString, kvp.Key);
        //                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeSource, TaskRunnerServiceRes.SourceNumberString, kvp.Key);
        //                return;
        //            }

        //            sbKML.Append("<Placemark>");

        //            if (shd.include == 1)
        //            {
        //                if (shd.Name.Substring(1, shd.Name.Length - 2).ToLower().StartsWith("river") || shd.Name.Substring(1, shd.Name.Length - 2).ToLower().StartsWith("rivière") || shd.Name.Substring(1, shd.Name.Length - 2).ToLower().StartsWith("brook") || shd.Name.Substring(1, shd.Name.Length - 2).ToLower().StartsWith("ruisseau"))
        //                {
        //                    sbKML.AppendLine(string.Format(@"<name>{0} (" + TaskRunnerServiceRes.UsedLowerCase + @")</name>", shd.Name.Substring(1, shd.Name.Length - 2)));
        //                    sbKML.AppendLine(@"<styleUrl>#msn_blue-pushpin</styleUrl>");
        //                }
        //                else
        //                {
        //                    sbKML.AppendLine(string.Format(@"<name>{0} (" + TaskRunnerServiceRes.UsedLowerCase + @")</name>", shd.Name.Substring(1, shd.Name.Length - 2)));
        //                    sbKML.AppendLine(@"<styleUrl>#msn_grn-pushpin</styleUrl>");
        //                }
        //            }
        //            else
        //            {
        //                sbKML.AppendLine(string.Format(@"<name>{0} (" + TaskRunnerServiceRes.NotUsedLowerCase + @")</name>", shd.Name.Substring(1, shd.Name.Length - 2)));
        //                sbKML.AppendLine(@"<styleUrl>#msn_red-pushpin</styleUrl>");
        //            }

        //            sbKML.AppendLine(@"<visibility>0</visibility>");
        //            sbKML.AppendLine(@"<description><![CDATA[");
        //            sbKML.AppendLine(string.Format(@"<h2>{0}</h2>", shd.Name.Substring(1, shd.Name.Length - 2)));
        //            sbKML.AppendLine(@"<h3>" + TaskRunnerServiceRes.Effluent + @"</h3>");
        //            sbKML.AppendLine(@"<ul>");
        //            sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.Coordinates + @"</b>");
        //            sbKML.AppendLine(string.Format(@"&nbsp;&nbsp;&nbsp; {0:F5} &nbsp; {1:F5}</li>", shd.coordinates.y, shd.coordinates.x).Replace(",", "."));

        //            if ((bool)mikeSourceModelLocal.IsContinuous)
        //            {
        //                sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.IsContinuous + "</b></li>");
        //                sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.Flow + @":</b> {0:F6} (m3/s)  {1:F0} (m3/" + TaskRunnerServiceRes.DayLowerCase + @")</li>", shd.constant_value, shd.constant_value * 24 * 3600).ToString().Replace(",", "."));
        //                sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.FCMPNPer100ML + @":</b> {0:F0}</li>", m21_3fmInput.femEngineHD.transport_module.sources.source[kvp.Key].component["COMPONENT_1"].constant_value).ToString().Replace(",", "."));
        //                sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.Temperature + @":</b> {0:F0} " + TaskRunnerServiceRes.Celcius + @"</li>", m21_3fmInput.femEngineHD.hydrodynamic_module.temperature_salinity_module.sources.source[kvp.Key].temperature.constant_value).ToString().Replace(",", "."));
        //                sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.Salinity + @":</b> {0:F0} " + TaskRunnerServiceRes.PSU + @"</li>", m21_3fmInput.femEngineHD.hydrodynamic_module.temperature_salinity_module.sources.source[kvp.Key].salinity.constant_value).ToString().Replace(",", "."));
        //            }
        //            else
        //            {
        //                sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.IsNotContinuous + @"</b></li>");

        //                MikeSourceStartEndService mikeSourceStartEndService = new MikeSourceStartEndService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //                List<MikeSourceStartEndModel> mikeSourceStartEndModelList = mikeSourceStartEndService.GetMikeSourceStartEndModelListWithMikeSourceIDDB(mikeSourceModelLocal.MikeSourceID);
        //                int CountMikeSourceStartEnd = 0;
        //                foreach (MikeSourceStartEndModel mssem in mikeSourceStartEndModelList)
        //                {
        //                    CountMikeSourceStartEnd += 1;
        //                    sbKML.AppendLine(@"<ul>");
        //                    sbKML.AppendLine("<b>" + TaskRunnerServiceRes.Spill + @" " + CountMikeSourceStartEnd + "</b>");
        //                    sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.SpillStartTime + @":</b> {0:yyyy/MM/dd HH:mm:ss tt}</li>", mssem.StartDateAndTime_Local));
        //                    sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.SpillEndTime + @":</b> {0:yyyy/MM/dd HH:mm:ss tt}</li>", mssem.EndDateAndTime_Local));
        //                    sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.FlowStart + @":</b> {0:F6} (m3/s)  {1:F0} (m3/" + TaskRunnerServiceRes.DayLowerCase + @")</li>", mssem.SourceFlowStart_m3_day, mssem.SourceFlowStart_m3_day * 24 * 3600).ToString().Replace(",", "."));
        //                    sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.FlowEnd + @":</b> {0:F6} (m3/s)  {1:F0} (m3/" + TaskRunnerServiceRes.DayLowerCase + @")</li>", mssem.SourceFlowEnd_m3_day, mssem.SourceFlowEnd_m3_day * 24 * 3600).ToString().Replace(",", "."));
        //                    sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.FCMPNPer100MLStart + @":</b> {0:F0}</li>", mssem.SourcePollutionStart_MPN_100ml).ToString().Replace(",", "."));
        //                    sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.FCMPNPer100MLEnd + @":</b> {0:F0}</li>", mssem.SourcePollutionEnd_MPN_100ml).ToString().Replace(",", "."));
        //                    sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.TemperatureStart + @":</b> {0:F0} " + TaskRunnerServiceRes.Celcius + @"</li>", mssem.SourceTemperatureStart_C).ToString().Replace(",", "."));
        //                    sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.TempertureEnd + @":</b> {0:F0} " + TaskRunnerServiceRes.Celcius + @"</li>", mssem.SourceTemperatureEnd_C).ToString().Replace(",", "."));
        //                    sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.SalinityStart + @":</b> {0:F0} " + TaskRunnerServiceRes.PSU + @"</li>", mssem.SourceSalinityStart_PSU).ToString().Replace(",", "."));
        //                    sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.SalinityEnd + @":</b> {0:F0} " + TaskRunnerServiceRes.PSU + @"</li>", mssem.SourceSalinityEnd_PSU).ToString().Replace(",", "."));
        //                    sbKML.AppendLine(@"</ul>");
        //                }
        //            }
        //            sbKML.AppendLine(@"</ul>");
        //            sbKML.AppendLine(@"<iframe src=""about:"" width=""500"" height=""1"" />");
        //            sbKML.AppendLine(@"]]></description>");

        //            sbKML.AppendLine(@"<Point>");
        //            sbKML.AppendLine(@"<coordinates>");
        //            sbKML.AppendLine(shd.coordinates.x.ToString().Replace(",", ".") + @"," + shd.coordinates.y.ToString().Replace(",", ".") + ",0 ");
        //            sbKML.AppendLine(@"</coordinates>");
        //            sbKML.AppendLine(@"</Point>");
        //            sbKML.AppendLine(@"</Placemark>");
        //            #endregion Source Placemark
        //        }
        //    }

        //    sbKML.Append("</Folder>");

        //    sbKML.Append("<Folder><name>" + TaskRunnerServiceRes.SourceNotIncluded + @"</name>");
        //    sbKML.AppendLine(@"<visibility>0</visibility>");

        //    // showing not used sources 
        //    foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.SOURCES.SOURCE> kvp in m21_3fmInput.femEngineHD.hydrodynamic_module.sources.source)
        //    {
        //        M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.SOURCES.SOURCE shd = kvp.Value;
        //        if (shd.include != 1)
        //        {
        //            #region Source Placemark
        //            MikeSourceModel mikeSourceModelLocal = (from msl in mikeSourceModelList
        //                                                    where msl.SourceNumberString == kvp.Key
        //                                                    select msl).FirstOrDefault<MikeSourceModel>();

        //            if (mikeSourceModelLocal == null)
        //            {
        //                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeSource, TaskRunnerServiceRes.SourceNumberString, kvp.Key);
        //                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeSource, TaskRunnerServiceRes.SourceNumberString, kvp.Key);
        //                return;
        //            }

        //            sbKML.Append("<Placemark>");

        //            if (shd.include == 1)
        //            {
        //                if (shd.Name.Substring(1, shd.Name.Length - 2).ToLower().StartsWith("river") || shd.Name.Substring(1, shd.Name.Length - 2).ToLower().StartsWith("rivière") || shd.Name.Substring(1, shd.Name.Length - 2).ToLower().StartsWith("brook") || shd.Name.Substring(1, shd.Name.Length - 2).ToLower().StartsWith("ruisseau"))
        //                {
        //                    sbKML.AppendLine(string.Format(@"<name>{0} (" + TaskRunnerServiceRes.UsedLowerCase + @")</name>", shd.Name.Substring(1, shd.Name.Length - 2)));
        //                    sbKML.AppendLine(@"<styleUrl>#msn_blue-pushpin</styleUrl>");
        //                }
        //                else
        //                {
        //                    sbKML.AppendLine(string.Format(@"<name>{0} (" + TaskRunnerServiceRes.UsedLowerCase + @")</name>", shd.Name.Substring(1, shd.Name.Length - 2)));
        //                    sbKML.AppendLine(@"<styleUrl>#msn_grn-pushpin</styleUrl>");
        //                }
        //            }
        //            else
        //            {
        //                sbKML.AppendLine(string.Format(@"<name>{0} (" + TaskRunnerServiceRes.NotUsedLowerCase + @")</name>", shd.Name.Substring(1, shd.Name.Length - 2)));
        //                sbKML.AppendLine(@"<styleUrl>#msn_red-pushpin</styleUrl>");
        //            }

        //            sbKML.AppendLine(@"<visibility>0</visibility>");
        //            sbKML.AppendLine(@"<description><![CDATA[");
        //            sbKML.AppendLine(string.Format(@"<h2>{0}</h2>", shd.Name.Substring(1, shd.Name.Length - 2)));
        //            sbKML.AppendLine(@"<h3>" + TaskRunnerServiceRes.Effluent + @"</h3>");
        //            sbKML.AppendLine(@"<ul>");
        //            sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.Coordinates + @"</b>");
        //            sbKML.AppendLine(string.Format(@"&nbsp;&nbsp;&nbsp; {0:F5} &nbsp; {1:F5}</li>", shd.coordinates.y, shd.coordinates.x).Replace(",", "."));

        //            if ((bool)mikeSourceModelLocal.IsContinuous)
        //            {
        //                sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.IsContinuous + "</b></li>");
        //                sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.Flow + @":</b> {0:F6} (m3/s)  {1:F0} (m3/" + TaskRunnerServiceRes.DayLowerCase + @")</li>", shd.constant_value, shd.constant_value * 24 * 3600).ToString().Replace(",", "."));
        //                sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.FCMPNPer100ML + @":</b> {0:F0}</li>", m21_3fmInput.femEngineHD.transport_module.sources.source[kvp.Key].component["COMPONENT_1"].constant_value).ToString().Replace(",", "."));
        //                sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.Temperature + @":</b> {0:F0} " + TaskRunnerServiceRes.Celcius + @"</li>", m21_3fmInput.femEngineHD.hydrodynamic_module.temperature_salinity_module.sources.source[kvp.Key].temperature.constant_value).ToString().Replace(",", "."));
        //                sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.Salinity + @":</b> {0:F0} " + TaskRunnerServiceRes.PSU + @"</li>", m21_3fmInput.femEngineHD.hydrodynamic_module.temperature_salinity_module.sources.source[kvp.Key].salinity.constant_value).ToString().Replace(",", "."));
        //            }
        //            else
        //            {
        //                sbKML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.IsNotContinuous + @"</b></li>");

        //                MikeSourceStartEndService mikeSourceStartEndService = new MikeSourceStartEndService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //                List<MikeSourceStartEndModel> mikeSourceStartEndModelList = mikeSourceStartEndService.GetMikeSourceStartEndModelListWithMikeSourceIDDB(mikeSourceModelLocal.MikeSourceID);
        //                int CountMikeSourceStartEnd = 0;
        //                foreach (MikeSourceStartEndModel mssem in mikeSourceStartEndModelList)
        //                {
        //                    CountMikeSourceStartEnd += 1;
        //                    sbKML.AppendLine(@"<ul>");
        //                    sbKML.AppendLine("<b>" + TaskRunnerServiceRes.Spill + @" " + CountMikeSourceStartEnd + "</b>");
        //                    sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.SpillStartTime + @":</b> {0:yyyy/MM/dd HH:mm:ss tt}</li>", mssem.StartDateAndTime_Local));
        //                    sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.SpillEndTime + @":</b> {0:yyyy/MM/dd HH:mm:ss tt}</li>", mssem.EndDateAndTime_Local));
        //                    sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.FlowStart + @":</b> {0:F6} (m3/s)  {1:F0} (m3/" + TaskRunnerServiceRes.DayLowerCase + @")</li>", mssem.SourceFlowStart_m3_day, mssem.SourceFlowStart_m3_day * 24 * 3600).ToString().Replace(",", "."));
        //                    sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.FlowEnd + @":</b> {0:F6} (m3/s)  {1:F0} (m3/" + TaskRunnerServiceRes.DayLowerCase + @")</li>", mssem.SourceFlowEnd_m3_day, mssem.SourceFlowEnd_m3_day * 24 * 3600).ToString().Replace(",", "."));
        //                    sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.FCMPNPer100MLStart + @":</b> {0:F0}</li>", mssem.SourcePollutionStart_MPN_100ml).ToString().Replace(",", "."));
        //                    sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.FCMPNPer100MLEnd + @":</b> {0:F0}</li>", mssem.SourcePollutionEnd_MPN_100ml).ToString().Replace(",", "."));
        //                    sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.TemperatureStart + @":</b> {0:F0} " + TaskRunnerServiceRes.Celcius + @"</li>", mssem.SourceTemperatureStart_C).ToString().Replace(",", "."));
        //                    sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.TempertureEnd + @":</b> {0:F0} " + TaskRunnerServiceRes.Celcius + @"</li>", mssem.SourceTemperatureEnd_C).ToString().Replace(",", "."));
        //                    sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.SalinityStart + @":</b> {0:F0} " + TaskRunnerServiceRes.PSU + @"</li>", mssem.SourceSalinityStart_PSU).ToString().Replace(",", "."));
        //                    sbKML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.SalinityEnd + @":</b> {0:F0} " + TaskRunnerServiceRes.PSU + @"</li>", mssem.SourceSalinityEnd_PSU).ToString().Replace(",", "."));
        //                    sbKML.AppendLine(@"</ul>");
        //                }
        //            }
        //            sbKML.AppendLine(@"</ul>");
        //            sbKML.AppendLine(@"<iframe src=""about:"" width=""500"" height=""1"" />");
        //            sbKML.AppendLine(@"]]></description>");

        //            sbKML.AppendLine(@"<Point>");
        //            sbKML.AppendLine(@"<coordinates>");
        //            sbKML.AppendLine(shd.coordinates.x.ToString().Replace(",", ".") + @"," + shd.coordinates.y.ToString().Replace(",", ".") + ",0 ");
        //            sbKML.AppendLine(@"</coordinates>");
        //            sbKML.AppendLine(@"</Point>");
        //            sbKML.AppendLine(@"</Placemark>");
        //            #endregion Source Placemark
        //        }
        //    }
        //    sbKML.Append("</Folder>");

        //    sbKML.Append("</Folder>");

        //    sbKML.AppendLine(@"</Document>");
        //    sbKML.AppendLine(@"</kml>");

        //    SaveInKMZFileStream(fi, fiKML, sbKML);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //}
        public void GenerateMikeScenarioCurrentAnimation(FileInfo fi)
        {
            string NotUsed = "";
            //string asliefj = _AppTaskService.GetAppTaskParamStr(_TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "FileGeneratorType");

            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            if (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr)
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-CA");
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("fr-CA");
            }
            else
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-CA");
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-CA");
            }

            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);

            DirectoryInfo di = new DirectoryInfo(ServerFilePath);
            if (!di.Exists)
                di.Create();

            if (fi.Exists)
                fi.Delete();

            MikeScenarioModel SamplingPlanModel = _MikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(SamplingPlanModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, ServiceRes.MikeScenario, ServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", ServiceRes.MikeScenario, ServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                return;
            }

            StringBuilder sbKMZ = new StringBuilder();

        }
        public void GenerateMikeScenarioCurrentMaximum(FileInfo fi)
        {
            string NotUsed = "";
            //string asliefj = _AppTaskService.GetAppTaskParamStr(_TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "FileGeneratorType");

            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            if (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr)
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-CA");
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("fr-CA");
            }
            else
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-CA");
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-CA");
            }

            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);

            DirectoryInfo di = new DirectoryInfo(ServerFilePath);
            if (!di.Exists)
                di.Create();

            if (fi.Exists)
                fi.Delete();

            MikeScenarioModel SamplingPlanModel = _MikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(SamplingPlanModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, ServiceRes.MikeScenario, ServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", ServiceRes.MikeScenario, ServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                return;
            }

            StringBuilder sbKMZ = new StringBuilder();

        }
        public void GenerateMikeScenarioMesh(FileInfo fi)
        {
            string NotUsed = "";
            //string asliefj = _AppTaskService.GetAppTaskParamStr(_TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "FileGeneratorType");

            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            if (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr)
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-CA");
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("fr-CA");
            }
            else
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-CA");
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-CA");
            }

            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);

            DirectoryInfo di = new DirectoryInfo(ServerFilePath);
            if (!di.Exists)
                di.Create();

            if (fi.Exists)
                fi.Delete();

            MikeScenarioModel SamplingPlanModel = _MikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(SamplingPlanModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, ServiceRes.MikeScenario, ServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", ServiceRes.MikeScenario, ServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                return;
            }

            StringBuilder sbKMZ = new StringBuilder();

        }
        public void GenerateMikeScenarioStudyArea(FileInfo fi)
        {
            string NotUsed = "";
            //string asliefj = _AppTaskService.GetAppTaskParamStr(_TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "FileGeneratorType");

            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            if (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr)
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-CA");
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("fr-CA");
            }
            else
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-CA");
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-CA");
            }

            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            string ServerFilePath = tvFileService.GetServerFilePath(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);

            DirectoryInfo di = new DirectoryInfo(ServerFilePath);
            if (!di.Exists)
                di.Create();

            if (fi.Exists)
                fi.Delete();

            MikeScenarioModel SamplingPlanModel = _MikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(SamplingPlanModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, ServiceRes.MikeScenario, ServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", ServiceRes.MikeScenario, ServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                return;
            }

            StringBuilder sbKMZ = new StringBuilder();

        }
        //public void CreateKMZResultFiles(KMZActionEnum kmzAction, DfsuFile dfsuFile, int MikeScenarioID, int TVFileID, string dfsParamItem /* ex: eumIuVelocity */, string ContourValueString, string SigmaLayerValuesString, string ZLayerValuesString, string DepthValuesString, string KMLTextPathForVector, int VectorSizeInMeterForEach10cm_s)
        //{
        //    string NotUsed = "";
        //    m21_3fm m21_3fmInput = null;
        //    string FileName = "";
        //    string ShortFileName = "";
        //    string LocalFileName = "";
        //    string delimStr = ",";
        //    char[] delimiter = delimStr.ToCharArray();
        //    FileInfo fiDfsu;

        //    // is this the one that i should open or is it the transport_module one
        //    TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    TVFile tvFile = tvFileService.GetTVFileWithTVFileIDDB(TVFileID);
        //    if (tvFile == null)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.TVFileID, TVFileID.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.TVFileID, TVFileID.ToString());
        //        return;
        //    }

        //    TVFile tvFilem21_3fm = tvFileService.GetTVFileWithTVItemIDAndTVFileTypeM21FMOrM3FM(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);

        //    if (tvFilem21_3fm == null)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.FileType, ".m21fmcf, .m3fm");
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.FileType, ".m21fmcf, .m3fm");
        //        return;
        //    }

        //    ShortFileName = tvFilem21_3fm.ServerFileName.Replace(".m21fm", "").Replace(".m3fm", "");
        //    FileInfo fi = new FileInfo(tvFilem21_3fm.ServerFilePath + tvFilem21_3fm.ServerFileName);

        //    m21_3fmInput = new m21_3fm();
        //    if (!m21_3fmInput.Read_M21_3FM_File(fi.FullName))
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, fi.FullName);
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", fi.FullName);
        //        return;
        //    }

        //    if (tvFile.FileType == (int)FileTypeEnum.LOG)
        //    {
        //        // finding the MikeScenarioFiles with FileType == ".m21fm"

        //        if (!fi.Exists)
        //        {
        //            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, fi.FullName);
        //            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", fi.FullName);
        //            return;
        //        }

        //        FileName = tvFilem21_3fm.ServerFilePath + tvFilem21_3fm.ServerFileName;

        //        string FirstPart = FileName.Substring(0, FileName.LastIndexOf("\\"));
        //        string SecondPart = m21_3fmInput.system.ResultRootFolder.Replace("|", "\\");
        //        string ThirdPart = FileName.Substring(FileName.LastIndexOf("\\") + 1) + " - Result Files\\";

        //        // try hydrodynamic module result file
        //        string ForthPart = m21_3fmInput.femEngineHD.hydrodynamic_module.outputs.output.First().Value.file_name.Replace("'", "");


        //        if (kmzAction != KMZActionEnum.GenerateKMZBoundaryConditionNodes)
        //        {
        //            if (tvFile.ServerFileName == ForthPart)
        //            {
        //            }
        //            else
        //            {
        //                // try transport module result file
        //                ForthPart = m21_3fmInput.femEngineHD.transport_module.outputs.output.First().Value.file_name.Replace("'", "");

        //                if (tvFile.ServerFileName != ForthPart)
        //                {
        //                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.FileFullName, fi.FullName);
        //                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.FileFullName, fi.FullName);
        //                    return;
        //                }
        //            }
        //        }

        //        fiDfsu = new FileInfo(FirstPart + SecondPart + ThirdPart + ForthPart);

        //        try
        //        {
        //            dfsuFile = DfsuFile.Open(fiDfsu.FullName);
        //        }
        //        catch (Exception)
        //        {
        //            // try transport module result file
        //            ForthPart = m21_3fmInput.femEngineHD.transport_module.outputs.output.First().Value.file_name.Replace("'", "");

        //            fiDfsu = new FileInfo(FirstPart + SecondPart + ThirdPart + ForthPart);

        //            try
        //            {
        //                dfsuFile = DfsuFile.Open(fiDfsu.FullName);
        //            }
        //            catch (Exception ex)
        //            {
        //                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotReadFile_Error_, fiDfsu.FullName, ex.Message);
        //                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotReadFile_Error_", fiDfsu.FullName, ex.Message);
        //                return;
        //            }
        //        }
        //    }
        //    else
        //    {
        //        fiDfsu = new FileInfo(tvFile.ServerFilePath + tvFile.ServerFileName);

        //        try
        //        {
        //            dfsuFile = DfsuFile.Open(fiDfsu.FullName);
        //        }
        //        catch (Exception ex)
        //        {
        //            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotReadFile_Error_, fiDfsu.FullName, ex.Message);
        //            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotReadFile_Error_", fiDfsu.FullName, ex.Message);
        //            return;
        //        }
        //    }

        //    List<Element> ElementList = new List<Element>();
        //    List<Node> NodeList = new List<Node>();

        //    // ----------------------------------------
        //    // filling the ElementList and The NodeList
        //    // ----------------------------------------

        //    FillElementListAndNodeList(dfsuFile, ElementList, NodeList);


        //    // Reseting appliation variables sbKML, sbStyle, sbPlacemark
        //    StringBuilder sbKML = new StringBuilder();
        //    StringBuilder sbStyle = new StringBuilder();
        //    StringBuilder sbPlacemark = new StringBuilder();

        //    // private KML and KMZ file names
        //    string TempFileName = fiDfsu.FullName.Substring(0, fiDfsu.FullName.LastIndexOf(@"\"));
        //    TempFileName = TempFileName.Substring(0, TempFileName.LastIndexOf(@"\") + 1);
        //    string KMLFileNameRoot = TempFileName + @"KMZ\";
        //    string KMZFileNameRoot = TempFileName + @"KMZ\";

        //    List<float> ContourValueList = new List<float>();

        //    try
        //    {
        //        if (ContourValueString.Trim().Length != 0)
        //        {
        //            foreach (string s in ContourValueString.Split(delimiter))
        //            {
        //                ContourValueList.Add(float.Parse(s));
        //            }
        //        }
        //    }
        //    catch (Exception)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.PlsEnterValidNumSeparatedByCommasFor_, TaskRunnerServiceRes.PollutionContour);
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("PlsEnterValidNumSeparatedByCommasFor_", TaskRunnerServiceRes.PollutionContour);
        //        return;
        //    }

        //    List<int> SigmaLayerValueList = new List<int>();
        //    List<int> ZLayerValueList = new List<int>();
        //    List<float> DepthValueList = new List<float>();

        //    if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
        //    {
        //        try
        //        {
        //            if (SigmaLayerValuesString.Trim().Length != 0)
        //            {
        //                foreach (string s in SigmaLayerValuesString.Split(delimiter))
        //                {
        //                    SigmaLayerValueList.Add(int.Parse(s));
        //                }
        //            }
        //        }
        //        catch (Exception)
        //        {
        //            NotUsed = string.Format(TaskRunnerServiceRes.PlsEnterValidNumSeparatedByCommasFor_, TaskRunnerServiceRes.SigmaLayers);
        //            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("PlsEnterValidNumSeparatedByCommasForSigmaLayers");
        //            return;
        //        }

        //        try
        //        {
        //            if (ZLayerValuesString.Trim().Length != 0)
        //            {
        //                foreach (string s in ZLayerValuesString.Split(delimiter))
        //                {
        //                    ZLayerValueList.Add(int.Parse(s));
        //                }
        //            }
        //        }
        //        catch (Exception)
        //        {
        //            NotUsed = string.Format(TaskRunnerServiceRes.PlsEnterValidNumSeparatedByCommasFor_, TaskRunnerServiceRes.ZLayers);
        //            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("PlsEnterValidNumSeparatedByCommasFor_", TaskRunnerServiceRes.ZLayers);
        //            return;
        //        }

        //        try
        //        {
        //            if (DepthValuesString.Trim().Length != 0)
        //            {
        //                foreach (string s in DepthValuesString.Split(delimiter))
        //                {
        //                    DepthValueList.Add(float.Parse(s));
        //                }
        //            }
        //        }
        //        catch (Exception)
        //        {
        //            NotUsed = string.Format(TaskRunnerServiceRes.PlsEnterValidNumSeparatedByCommasFor_, TaskRunnerServiceRes.Depth);
        //            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("PlsEnterValidNumSeparatedByCommasFor_", TaskRunnerServiceRes.Depth);
        //            return;
        //        }
        //    }
        //    else if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu2D)
        //    {
        //        SigmaLayerValueList.Add(0);
        //    }
        //    else
        //    {
        //        NotUsed = TaskRunnerServiceRes.NotImplementedYet;
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("NotImplementedYet");
        //        return;
        //    }

        //    List<NodeLayer> TopNodeLayerList = new List<NodeLayer>();
        //    List<NodeLayer> BottomNodeLayerList = new List<NodeLayer>();
        //    List<ElementLayer> ElementLayerList = new List<ElementLayer>();

        //    FillElementLayerList(dfsuFile, SigmaLayerValueList, ElementList, ElementLayerList, TopNodeLayerList, BottomNodeLayerList);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    switch (kmzAction)
        //    {
        //        case KMZActionEnum.GenerateKMZContourAnimation:
        //            {
        //                #region ContourAnimation
        //                // __________________________________________________________
        //                // Writting top part of KML file
        //                // __________________________________________________________
        //                StringBuilder sbKMLTop = new StringBuilder();
        //                string SigmaLayerFileString = "";
        //                string ZLayerFileString = "";
        //                string DepthFileString = "";

        //                //SigmaLayerValueList.Clear();
        //                //ZLayerValueList.Clear();
        //                //DepthValueList.Clear();

        //                if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
        //                {
        //                    if (SigmaLayerValueList.Count > 0)
        //                    {
        //                        string TempStr = "";
        //                        foreach (int val in SigmaLayerValueList)
        //                        {
        //                            TempStr = TempStr + "," + val;
        //                        }
        //                        TempStr = TempStr.Substring(1).Replace(",", "_");
        //                        SigmaLayerFileString = "_Sigma[" + TempStr + "]";
        //                    }
        //                    if (ZLayerValueList.Count > 0)
        //                    {
        //                        string TempStr = "";
        //                        foreach (int val in ZLayerValueList)
        //                        {
        //                            TempStr = TempStr + "," + val;
        //                        }
        //                        TempStr = TempStr.Substring(1).Replace(",", "_");
        //                        ZLayerFileString = "_Z[" + TempStr + "]";
        //                    }
        //                    if (DepthValueList.Count > 0)
        //                    {
        //                        string TempStr = "";
        //                        foreach (int val in SigmaLayerValueList)
        //                        {
        //                            TempStr = TempStr + "," + val;
        //                        }
        //                        TempStr = TempStr.Substring(1).Replace(",", "_");
        //                        DepthFileString = "_Depth[" + TempStr + "]";
        //                    }
        //                }
        //                else if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu2D)
        //                {
        //                    SigmaLayerValueList.Clear();
        //                    SigmaLayerValueList.Add(1);
        //                }
        //                else
        //                {
        //                    NotUsed = string.Format(TaskRunnerServiceRes._NotImplementedYet, dfsuFile.DfsuFileType.ToString());
        //                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_NotImplementedYet", dfsuFile.DfsuFileType.ToString());
        //                    return;
        //                }

        //                string ExtendedFileName = ShortFileName + "_Pol_Anim[" + ContourValueString.Replace(",", "_") + "]" + SigmaLayerFileString;  //+ ZLayerFileString + DepthFileString;

        //                WriteKMLTop(ExtendedFileName, sbKMLTop);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                sbKML = new StringBuilder();
        //                sbStyle = new StringBuilder();
        //                sbPlacemark = new StringBuilder();

        //                // __________________________________________________________
        //                // Creating Pollution Video
        //                // __________________________________________________________
        //                StringBuilder sbKMLFeacalColiformContour = new StringBuilder();
        //                StringBuilder sbStyleFeacalColiformContour = new StringBuilder();

        //                WriteKMLFeacalColiformContourLine(dfsuFile, dfsParamItem, sbStyleFeacalColiformContour, sbKMLFeacalColiformContour, ContourValueList, SigmaLayerValueList, ZLayerValueList, DepthValueList, ElementLayerList, TopNodeLayerList, BottomNodeLayerList);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                sbStyle.Append(sbStyleFeacalColiformContour.ToString());
        //                sbPlacemark.Append(sbKMLFeacalColiformContour.ToString());

        //                // __________________________________________________________
        //                // Creating Input To Model
        //                // __________________________________________________________
        //                StringBuilder sbKMLModelInput = new StringBuilder();
        //                StringBuilder sbStyleModelInput = new StringBuilder();
        //                sbKMLModelInput.Append(GetModelAndSourceDesc(ContourValueList, MikeScenarioID, m21_3fmInput));

        //                WriteKMLModelInput(sbStyleModelInput, sbKMLModelInput, ContourValueList, MikeScenarioID, m21_3fmInput);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                sbStyle.Append(sbStyleModelInput.ToString());
        //                sbPlacemark.Append(sbKMLModelInput.ToString());


        //                // __________________________________________________________
        //                // Writting bottom part of KML file
        //                // __________________________________________________________
        //                StringBuilder sbKMLBottom = new StringBuilder();

        //                WriteKMLBottom(sbKMLBottom);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                // __________________________________________________________
        //                // Concatenating KML file
        //                // __________________________________________________________
        //                sbKML.Append(sbKMLTop.ToString());
        //                sbKML.Append(sbStyle.ToString());
        //                sbKML.Append(sbPlacemark.ToString());
        //                sbKML.Append(sbKMLBottom.ToString());

        //                string KMLFileName = KMLFileNameRoot + ExtendedFileName + "_" + _TaskRunnerBaseService._BWObj.appTaskModel.Language + ".kml";
        //                string KMZFileName = KMZFileNameRoot + ExtendedFileName + "_" + _TaskRunnerBaseService._BWObj.appTaskModel.Language + ".kmz";

        //                // make sure directory exist if not create it
        //                DirectoryInfo di = new DirectoryInfo(KMZFileName.Substring(0, KMZFileName.LastIndexOf("\\")));
        //                if (!di.Exists)
        //                    di.Create();

        //                LocalFileName = tvFilem21_3fm.ServerFilePath + ExtendedFileName + "_" + _TaskRunnerBaseService._BWObj.appTaskModel.Language + ".kmz";

        //                SaveInKMZFileStream(fi, fiKML, sbKML);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                #endregion ContourAnimation
        //            }
        //            break;
        //        case KMZActionEnum.GenerateKMZContourLimit:
        //            {
        //                #region ContourLimit
        //                //string DocName = fi.FullName.Substring(0, fi.FullName.LastIndexOf("\\")) + "_ContourLimit";

        //                // __________________________________________________________
        //                // Writting top part of KML file
        //                // __________________________________________________________
        //                StringBuilder sbKMLTop = new StringBuilder();
        //                string SigmaLayerFileString = "";
        //                string ZLayerFileString = "";
        //                string DepthFileString = "";

        //                //SigmaLayerValueList.Clear();
        //                //ZLayerValueList.Clear();
        //                //DepthValueList.Clear();

        //                if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
        //                {
        //                    if (SigmaLayerValueList.Count > 0)
        //                    {
        //                        string TempStr = "";
        //                        foreach (int val in SigmaLayerValueList)
        //                        {
        //                            TempStr = TempStr + "," + val;
        //                        }
        //                        TempStr = TempStr.Substring(1).Replace(",", "_");
        //                        SigmaLayerFileString = "_Sigma[" + TempStr + "]";
        //                    }
        //                    if (ZLayerValueList.Count > 0)
        //                    {
        //                        string TempStr = "";
        //                        foreach (int val in ZLayerValueList)
        //                        {
        //                            TempStr = TempStr + "," + val;
        //                        }
        //                        TempStr = TempStr.Substring(1).Replace(",", "_");
        //                        ZLayerFileString = "_Z[" + TempStr + "]";
        //                    }
        //                    if (DepthValueList.Count > 0)
        //                    {
        //                        string TempStr = "";
        //                        foreach (int val in SigmaLayerValueList)
        //                        {
        //                            TempStr = TempStr + "," + val;
        //                        }
        //                        TempStr = TempStr.Substring(1).Replace(",", "_");
        //                        DepthFileString = "_Depth[" + TempStr + "]";
        //                    }
        //                }
        //                else if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu2D)
        //                {
        //                    SigmaLayerValueList.Clear();
        //                    SigmaLayerValueList.Add(1);
        //                }
        //                else
        //                {
        //                    NotUsed = string.Format(TaskRunnerServiceRes._NotImplementedYet, dfsuFile.DfsuFileType.ToString());
        //                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_NotImplementedYet", dfsuFile.DfsuFileType.ToString());
        //                    return;
        //                }

        //                string ExtendedFileName = ShortFileName + "_Pol_Limits[" + ContourValueString.Replace(",", "_") + "]" + SigmaLayerFileString; //+ ZLayerFileString + DepthFileString;

        //                WriteKMLTop(ExtendedFileName, sbKMLTop);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                sbKML = new StringBuilder();
        //                sbStyle = new StringBuilder();
        //                sbPlacemark = new StringBuilder();

        //                // __________________________________________________________
        //                // Creating Pollution Limits
        //                // __________________________________________________________
        //                StringBuilder sbKMLPollutionLimitsContour = new StringBuilder();
        //                StringBuilder sbStylePollutionLimitsContour = new StringBuilder();

        //                WriteKMLPollutionLimitsContourLine(dfsuFile, dfsParamItem, sbStylePollutionLimitsContour, sbKMLPollutionLimitsContour, ContourValueList, SigmaLayerValueList, ZLayerValueList, DepthValueList, ElementLayerList, TopNodeLayerList, BottomNodeLayerList);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                sbStyle.Append(sbStylePollutionLimitsContour.ToString());
        //                sbPlacemark.Append(sbKMLPollutionLimitsContour.ToString());

        //                // __________________________________________________________
        //                // Creating Input To Model
        //                // __________________________________________________________
        //                StringBuilder sbKMLModelInput = new StringBuilder();
        //                StringBuilder sbStyleModelInput = new StringBuilder();
        //                sbKMLModelInput.Append(GetModelAndSourceDesc(ContourValueList, MikeScenarioID, m21_3fmInput));

        //                WriteKMLModelInput(sbStyleModelInput, sbKMLModelInput, ContourValueList, MikeScenarioID, m21_3fmInput);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                sbStyle.Append(sbStyleModelInput.ToString());
        //                sbPlacemark.Append(sbKMLModelInput.ToString());


        //                // __________________________________________________________
        //                // Writting bottom part of KML file
        //                // __________________________________________________________
        //                StringBuilder sbKMLBottom = new StringBuilder();

        //                WriteKMLBottom(sbKMLBottom);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                // __________________________________________________________
        //                // Concatenating KML file
        //                // __________________________________________________________
        //                sbKML.Append(sbKMLTop.ToString());
        //                sbKML.Append(sbStyle.ToString());
        //                sbKML.Append(sbPlacemark.ToString());
        //                sbKML.Append(sbKMLBottom.ToString());

        //                //string PolCont = "_ContourLimit";
        //                //foreach (float f in ContourValueList)
        //                //{
        //                //    PolCont += string.Format("[{0}]", f);
        //                //}
        //                string KMLFileName = KMLFileNameRoot + ExtendedFileName + "_" + _TaskRunnerBaseService._BWObj.appTaskModel.Language + ".kml";
        //                string KMZFileName = KMZFileNameRoot + ExtendedFileName + "_" + _TaskRunnerBaseService._BWObj.appTaskModel.Language + ".kmz";

        //                // make sure directory exist if not create it
        //                DirectoryInfo di = new DirectoryInfo(KMZFileName.Substring(0, KMZFileName.LastIndexOf("\\")));
        //                if (!di.Exists)
        //                    di.Create();

        //                LocalFileName = tvFilem21_3fm.ServerFilePath + ExtendedFileName + "_" + _TaskRunnerBaseService._BWObj.appTaskModel.Language + ".kmz";

        //                SaveInKMZFileStream(LocalFileName, MikeScenarioID, KMZFileName, KMLFileName, sbKML);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                #endregion ContourLimit
        //            }
        //            break;
        //        case KMZActionEnum.GenerateKMZCurrentAnimation:
        //            {
        //                #region CurrentAnimation
        //                // __________________________________________________________
        //                // Writting top part of KML file
        //                // __________________________________________________________
        //                StringBuilder sbKMLTop = new StringBuilder();
        //                string SigmaLayerFileString = "";
        //                string ZLayerFileString = "";
        //                string DepthFileString = "";

        //                //SigmaLayerValueList.Clear();
        //                //ZLayerValueList.Clear();
        //                //DepthValueList.Clear();

        //                if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
        //                {
        //                    if (SigmaLayerValueList.Count > 0)
        //                    {
        //                        string TempStr = "";
        //                        foreach (int val in SigmaLayerValueList)
        //                        {
        //                            TempStr = TempStr + "," + val;
        //                        }
        //                        TempStr = TempStr.Substring(1).Replace(",", "_");
        //                        SigmaLayerFileString = "_Sigma[" + TempStr + "]";
        //                    }
        //                    if (ZLayerValueList.Count > 0)
        //                    {
        //                        string TempStr = "";
        //                        foreach (int val in ZLayerValueList)
        //                        {
        //                            TempStr = TempStr + "," + val;
        //                        }
        //                        TempStr = TempStr.Substring(1).Replace(",", "_");
        //                        ZLayerFileString = "_Z[" + TempStr + "]";
        //                    }
        //                    if (DepthValueList.Count > 0)
        //                    {
        //                        string TempStr = "";
        //                        foreach (int val in SigmaLayerValueList)
        //                        {
        //                            TempStr = TempStr + "," + val;
        //                        }
        //                        TempStr = TempStr.Substring(1).Replace(",", "_");
        //                        DepthFileString = "_Depth[" + TempStr + "]";
        //                    }
        //                }
        //                else if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu2D)
        //                {
        //                    SigmaLayerValueList.Add(1);
        //                }
        //                else
        //                {
        //                    NotUsed = string.Format(TaskRunnerServiceRes._NotImplementedYet, dfsuFile.DfsuFileType.ToString());
        //                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_NotImplementedYet", dfsuFile.DfsuFileType.ToString());
        //                    return;
        //                }

        //                string ExtendedFileName = ShortFileName + "_Cur_Anim" + SigmaLayerFileString; // +ZLayerFileString + DepthFileString;

        //                WriteKMLTop(ExtendedFileName, sbKMLTop);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                sbKML = new StringBuilder();
        //                sbStyle = new StringBuilder();
        //                sbPlacemark = new StringBuilder();

        //                // __________________________________________________________
        //                // Creating Current Video
        //                // __________________________________________________________
        //                StringBuilder sbKMLCurrentAnim = new StringBuilder();
        //                StringBuilder sbStyleCurrentAnim = new StringBuilder();

        //                List<ElementLayer> SelectedElementLayerList = ParseKMLPath(KMLTextPathForVector, ElementLayerList);

        //                WriteKMLCurrentsAnimation(dfsuFile, dfsParamItem, sbStyleCurrentAnim, sbKMLCurrentAnim, ContourValueList, SigmaLayerValueList, ZLayerValueList, DepthValueList, SelectedElementLayerList, VectorSizeInMeterForEach10cm_s);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                sbStyle.Append(sbStyleCurrentAnim.ToString());
        //                sbPlacemark.Append(sbKMLCurrentAnim.ToString());

        //                //// __________________________________________________________
        //                //// Creating Input To Model
        //                //// __________________________________________________________
        //                //StringBuilder sbKMLModelInput = new StringBuilder();
        //                //StringBuilder sbStyleModelInput = new StringBuilder();
        //                //sbKMLModelInput.Append(GetModelAndSourceDesc(ContourValueList, MikeScenarioID, m21_3fmInput));
        //                //if (!WriteKMLModelInput(sbStyleModelInput, sbKMLModelInput, ContourValueList, MikeScenarioID, m21_3fmInput))
        //                //    return false;
        //                //sbStyle.Append(sbStyleModelInput.ToString());
        //                //sbPlacemark.Append(sbKMLModelInput.ToString());


        //                // __________________________________________________________
        //                // Writting bottom part of KML file
        //                // __________________________________________________________
        //                StringBuilder sbKMLBottom = new StringBuilder();

        //                WriteKMLBottom(sbKMLBottom);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                // __________________________________________________________
        //                // Concatenating KML file
        //                // __________________________________________________________
        //                sbKML.Append(sbKMLTop.ToString());
        //                sbKML.Append(sbStyle.ToString());
        //                sbKML.Append(sbPlacemark.ToString());
        //                sbKML.Append(sbKMLBottom.ToString());

        //                string KMLFileName = KMLFileNameRoot + ExtendedFileName + "_" + _TaskRunnerBaseService._BWObj.appTaskModel.Language + ".kml";
        //                string KMZFileName = KMZFileNameRoot + ExtendedFileName + "_" + _TaskRunnerBaseService._BWObj.appTaskModel.Language + ".kmz";

        //                // make sure directory exist if not create it
        //                DirectoryInfo di = new DirectoryInfo(KMZFileName.Substring(0, KMZFileName.LastIndexOf("\\")));
        //                if (!di.Exists)
        //                    di.Create();

        //                LocalFileName = tvFilem21_3fm.ServerFilePath + ExtendedFileName + "_" + _TaskRunnerBaseService._BWObj.appTaskModel.Language + ".kmz";

        //                SaveInKMZFileStream(LocalFileName, MikeScenarioID, KMZFileName, KMLFileName, sbKML);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                #endregion CurrentAnimation
        //            }
        //            break;
        //        case KMZActionEnum.GenerateKMZCurrentMaximum:
        //            {
        //                #region CurrentMaximum
        //                #endregion CurrentMaximum
        //            }
        //            break;
        //        case KMZActionEnum.GenerateKMZMesh:
        //            {
        //                #region Mesh
        //                //string DocName = fi.FullName.Substring(0, fi.FullName.LastIndexOf("\\")) + "_ModelInput";
        //                // __________________________________________________________
        //                // Creating Mesh
        //                // __________________________________________________________
        //                // __________________________________________________________
        //                // Writting top part of KML file
        //                // __________________________________________________________
        //                StringBuilder sbKMLTop = new StringBuilder();

        //                WriteKMLTop(ShortFileName + "_Mesh", sbKMLTop);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                sbKML = new StringBuilder();
        //                sbStyle = new StringBuilder();
        //                sbPlacemark = new StringBuilder();

        //                StringBuilder sbKMLMesh = new StringBuilder();
        //                StringBuilder sbStyleMesh = new StringBuilder();

        //                WriteKMLMesh(sbStyleMesh, sbKMLMesh, ElementLayerList);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                sbStyle.Append(sbStyleMesh.ToString());
        //                sbPlacemark.Append(sbKMLMesh.ToString());

        //                // __________________________________________________________
        //                // Writting bottom part of KML file
        //                // __________________________________________________________
        //                StringBuilder sbKMLBottom = new StringBuilder();

        //                WriteKMLBottom(sbKMLBottom);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                // __________________________________________________________
        //                // Concatenating KML file
        //                // __________________________________________________________
        //                sbKML.Append(sbKMLTop.ToString());
        //                sbKML.Append(sbStyle.ToString());
        //                sbKML.Append(sbPlacemark.ToString());
        //                sbKML.Append(sbKMLBottom.ToString());

        //                string KMLFileName = KMLFileNameRoot + ShortFileName + "_Mesh_" + _TaskRunnerBaseService._BWObj.appTaskModel.Language + ".kml";
        //                string KMZFileName = KMZFileNameRoot + ShortFileName + "_Mesh_" + _TaskRunnerBaseService._BWObj.appTaskModel.Language + ".kmz";

        //                // make sure directory exist if not create it
        //                DirectoryInfo di = new DirectoryInfo(KMZFileName.Substring(0, KMZFileName.LastIndexOf("\\")));
        //                if (!di.Exists)
        //                    di.Create();

        //                LocalFileName = tvFilem21_3fm.ServerFilePath + ShortFileName + "_Mesh_" + _TaskRunnerBaseService._BWObj.appTaskModel.Language + ".kmz";

        //                SaveInKMZFileStream(LocalFileName, MikeScenarioID, KMZFileName, KMLFileName, sbKML);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                #endregion Mesh
        //            }
        //            break;
        //        case KMZActionEnum.GenerateKMZStudyArea:
        //            {
        //                #region Study Area
        //                //string DocName = fi.FullName.Substring(0, fi.FullName.LastIndexOf("\\")) + "_ModelInput";
        //                // __________________________________________________________
        //                // Creating Study Area line
        //                // __________________________________________________________
        //                // __________________________________________________________
        //                // Writting top part of KML file
        //                // __________________________________________________________
        //                StringBuilder sbKMLTop = new StringBuilder();

        //                WriteKMLTop(ShortFileName + "_StudyArea", sbKMLTop);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                sbKML = new StringBuilder();
        //                sbStyle = new StringBuilder();
        //                sbPlacemark = new StringBuilder();

        //                StringBuilder sbKMLStudyAreaLine = new StringBuilder();
        //                StringBuilder sbStyleStudyAreaLine = new StringBuilder();

        //                WriteKMLStudyAreaLine(sbStyleStudyAreaLine, sbKMLStudyAreaLine, ElementList, NodeList);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                sbStyle.Append(sbStyleStudyAreaLine.ToString());
        //                sbPlacemark.Append(sbKMLStudyAreaLine.ToString());

        //                // __________________________________________________________
        //                // Writting bottom part of KML file
        //                // __________________________________________________________
        //                StringBuilder sbKMLBottom = new StringBuilder();

        //                WriteKMLBottom(sbKMLBottom);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                // __________________________________________________________
        //                // Concatenating KML file
        //                // __________________________________________________________
        //                sbKML.Append(sbKMLTop.ToString());
        //                sbKML.Append(sbStyle.ToString());
        //                sbKML.Append(sbPlacemark.ToString());
        //                sbKML.Append(sbKMLBottom.ToString());

        //                string KMLFileName = KMLFileNameRoot + ShortFileName + "_StudyArea_" + _TaskRunnerBaseService._BWObj.appTaskModel.Language + ".kml";
        //                string KMZFileName = KMZFileNameRoot + ShortFileName + "_StudyArea_" + _TaskRunnerBaseService._BWObj.appTaskModel.Language + ".kmz";

        //                // make sure directory exist if not create it
        //                DirectoryInfo di = new DirectoryInfo(KMZFileName.Substring(0, KMZFileName.LastIndexOf("\\")));
        //                if (!di.Exists)
        //                    di.Create();

        //                LocalFileName = tvFilem21_3fm.ServerFilePath + ShortFileName + "_StudyArea_" + _TaskRunnerBaseService._BWObj.appTaskModel.Language + ".kmz";

        //                SaveInKMZFileStream(LocalFileName, MikeScenarioID, KMZFileName, KMLFileName, sbKML);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                #endregion Study Area
        //            }
        //            break;
        //        case KMZActionEnum.GenerateKMZBoundaryConditionNodes:
        //            {
        //                #region Boundary Condition Nodes
        //                // __________________________________________________________
        //                // Creating Boundary Condition Nodes
        //                // __________________________________________________________
        //                // __________________________________________________________
        //                // Writting top part of KML file
        //                // __________________________________________________________
        //                StringBuilder sbKMLTop = new StringBuilder();

        //                WriteKMLTop(ShortFileName + "_Boundary_Condition_Nodes", sbKMLTop);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                sbKML = new StringBuilder();
        //                sbStyle = new StringBuilder();
        //                sbPlacemark = new StringBuilder();

        //                StringBuilder sbKMLStudyAreaLine = new StringBuilder();
        //                StringBuilder sbStyleStudyAreaLine = new StringBuilder();

        //                WriteKMLBoundaryConditionNode(MikeScenarioID, sbStyleStudyAreaLine, sbKMLStudyAreaLine);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                sbStyle.Append(sbStyleStudyAreaLine.ToString());
        //                sbPlacemark.Append(sbKMLStudyAreaLine.ToString());

        //                // __________________________________________________________
        //                // Writting bottom part of KML file
        //                // __________________________________________________________
        //                StringBuilder sbKMLBottom = new StringBuilder();

        //                WriteKMLBottom(sbKMLBottom);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                // __________________________________________________________
        //                // Concatenating KML file
        //                // __________________________________________________________
        //                sbKML.Append(sbKMLTop.ToString());
        //                sbKML.Append(sbStyle.ToString());
        //                sbKML.Append(sbPlacemark.ToString());
        //                sbKML.Append(sbKMLBottom.ToString());

        //                string KMLFileName = KMLFileNameRoot + ShortFileName + "_BoundaryCondition_" + _TaskRunnerBaseService._BWObj.appTaskModel.Language + ".kml";
        //                string KMZFileName = KMZFileNameRoot + ShortFileName + "_BoundaryCondition_" + _TaskRunnerBaseService._BWObj.appTaskModel.Language + ".kmz";

        //                // make sure directory exist if not create it
        //                DirectoryInfo di = new DirectoryInfo(KMZFileName.Substring(0, KMZFileName.LastIndexOf("\\")));
        //                if (!di.Exists)
        //                    di.Create();
        //                // Saving sbKML StringBuilder to the KML file [KMZFileName]

        //                LocalFileName = tvFilem21_3fm.ServerFilePath + ShortFileName + "_BoundaryCondition_" + _TaskRunnerBaseService._BWObj.appTaskModel.Language + ".kmz";

        //                // should verify that this is at the right place
        //                // and that MikeScenarioFiles, CSSPFiles etc... are updated with this file info.
        //                SaveInKMZFileStream(LocalFileName, MikeScenarioID, KMZFileName, KMLFileName, sbKML);
        //                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                    return;

        //                #endregion Boundary Condition Nodes
        //            }
        //            break;
        //        default:
        //            break;
        //    }
        //}

        #endregion Functions public

        #region Functions private
        private void DrawKMLContourPolygon(List<ContourPolygon> ContourPolygonList, DfsuFile dfsuFile, int ParamCount, StringBuilder sbStyleFeacalColiformContour, StringBuilder sbPlacemarkFeacalColiformContour)
        {
            int Count = 0;
            float MaxXCoord = -180;
            float MaxYCoord = -90;
            float MinXCoord = 180;
            float MinYCoord = 90;
            sbPlacemarkFeacalColiformContour.AppendLine(@"<Folder>");
            sbPlacemarkFeacalColiformContour.AppendLine(@"<visibility>0</visibility>");
            sbPlacemarkFeacalColiformContour.AppendLine(string.Format(@"<name>{0:yyyy-MM-dd} {0:HH:mm:ss tt}</name>", dfsuFile.StartDateTime.AddSeconds(ParamCount * dfsuFile.TimeStepInSeconds)));
            sbPlacemarkFeacalColiformContour.AppendLine(@"<TimeSpan>");
            sbPlacemarkFeacalColiformContour.AppendLine(string.Format(@"<begin>{0:yyyy-MM-dd}T{0:HH:mm:ss}</begin>", dfsuFile.StartDateTime.AddSeconds(ParamCount * dfsuFile.TimeStepInSeconds)));
            sbPlacemarkFeacalColiformContour.AppendLine(string.Format(@"<end>{0:yyyy-MM-dd}T{0:HH:mm:ss}</end>", dfsuFile.StartDateTime.AddSeconds((ParamCount + 1) * dfsuFile.TimeStepInSeconds)));
            sbPlacemarkFeacalColiformContour.AppendLine(@"</TimeSpan>");
            foreach (ContourPolygon contourPolygon in ContourPolygonList)
            {
                Count += 1;
                // draw the polygons
                sbPlacemarkFeacalColiformContour.AppendLine(@"<Placemark>");
                sbPlacemarkFeacalColiformContour.AppendLine(@"<visibility>0</visibility>");
                sbPlacemarkFeacalColiformContour.AppendLine(string.Format(@"<name>{0} {1}</name>", contourPolygon.ContourValue, TaskRunnerServiceRes.PollutionContour));
                if (contourPolygon.ContourValue >= 14 && contourPolygon.ContourValue < 88)
                {
                    sbPlacemarkFeacalColiformContour.AppendLine(@"<styleUrl>#fc_14</styleUrl>");
                }
                else if (contourPolygon.ContourValue >= 88)
                {
                    sbPlacemarkFeacalColiformContour.AppendLine(@"<styleUrl>#fc_88</styleUrl>");
                }
                else
                {
                    sbPlacemarkFeacalColiformContour.AppendLine(@"<styleUrl>#fc_LT14</styleUrl>");
                }

                sbPlacemarkFeacalColiformContour.AppendLine(@"<Polygon>");
                //sbPlacemarkFeacalColiformContour.AppendLine(@"<extrude>1</extrude>");
                //sbPlacemarkFeacalColiformContour.AppendLine(@"<tessellate>1</tessellate>");
                //sbPlacemarkFeacalColiformContour.AppendLine(@"<altitudeMode>absolute</altitudeMode>");
                //sbPlacemarkFeacalColiformContour.AppendLine(@"<gx:drawOrder>" + contourPolygon.Layer +"</gx:drawOrder>"); 
                sbPlacemarkFeacalColiformContour.AppendLine(@"<outerBoundaryIs>");
                sbPlacemarkFeacalColiformContour.AppendLine(@"<LinearRing>");
                sbPlacemarkFeacalColiformContour.AppendLine(@"<coordinates>");
                foreach (Node node in contourPolygon.ContourNodeList)
                {
                    if (MaxXCoord < node.X) MaxXCoord = node.X;
                    if (MaxYCoord < node.Y) MaxYCoord = node.Y;
                    if (MinXCoord > node.X) MinXCoord = node.X;
                    if (MinYCoord > node.Y) MinYCoord = node.Y;
                    sbPlacemarkFeacalColiformContour.Append(node.X.ToString().Replace(",", ".") + @"," + node.Y.ToString().Replace(",", ".") + "," + node.Z.ToString().Replace(",", ".") + " ");
                }
                sbPlacemarkFeacalColiformContour.AppendLine(@"</coordinates>");
                sbPlacemarkFeacalColiformContour.AppendLine(@"</LinearRing>");
                sbPlacemarkFeacalColiformContour.AppendLine(@"</outerBoundaryIs>");
                sbPlacemarkFeacalColiformContour.AppendLine(@"</Polygon>");
                sbPlacemarkFeacalColiformContour.AppendLine(@"</Placemark>");
            }
            sbPlacemarkFeacalColiformContour.AppendLine(@"</Folder>");
        }
        private void DrawKMLContourStyle(StringBuilder sbStyleFeacalColiformContour, StringBuilder sbPlacemarkFeacalColiformContour)
        {
            sbStyleFeacalColiformContour.AppendLine(@"	<StyleMap id=""msn_ylw-pushpin"">");
            sbStyleFeacalColiformContour.AppendLine(@"		<Pair>");
            sbStyleFeacalColiformContour.AppendLine(@"			<key>normal</key>");
            sbStyleFeacalColiformContour.AppendLine(@"			<styleUrl>#sn_ylw-pushpin</styleUrl>");
            sbStyleFeacalColiformContour.AppendLine(@"		</Pair>");
            sbStyleFeacalColiformContour.AppendLine(@"		<Pair>");
            sbStyleFeacalColiformContour.AppendLine(@"			<key>highlight</key>");
            sbStyleFeacalColiformContour.AppendLine(@"			<styleUrl>#sh_ylw-pushpin</styleUrl>");
            sbStyleFeacalColiformContour.AppendLine(@"		</Pair>");
            sbStyleFeacalColiformContour.AppendLine(@"	</StyleMap>");
            sbStyleFeacalColiformContour.AppendLine(@"	<Style id=""sn_ylw-pushpin"">");
            sbStyleFeacalColiformContour.AppendLine(@"		<IconStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"			<scale>1.1</scale>");
            sbStyleFeacalColiformContour.AppendLine(@"			<Icon>");
            sbStyleFeacalColiformContour.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>");
            sbStyleFeacalColiformContour.AppendLine(@"			</Icon>");
            sbStyleFeacalColiformContour.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbStyleFeacalColiformContour.AppendLine(@"		</IconStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"      <LineStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"         <color>ff000000</color>");
            sbStyleFeacalColiformContour.AppendLine(@"       </LineStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"	</Style>");
            sbStyleFeacalColiformContour.AppendLine(@"	<Style id=""sh_ylw-pushpin"">");
            sbStyleFeacalColiformContour.AppendLine(@"		<IconStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"			<scale>1.3</scale>");
            sbStyleFeacalColiformContour.AppendLine(@"			<Icon>");
            sbStyleFeacalColiformContour.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>");
            sbStyleFeacalColiformContour.AppendLine(@"			</Icon>");
            sbStyleFeacalColiformContour.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbStyleFeacalColiformContour.AppendLine(@"		</IconStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"      <LineStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"         <color>ff000000</color>");
            sbStyleFeacalColiformContour.AppendLine(@"       </LineStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"	</Style>");

            sbStyleFeacalColiformContour.AppendLine(@"	<StyleMap id=""msn_grn-pushpin"">");
            sbStyleFeacalColiformContour.AppendLine(@"		<Pair>");
            sbStyleFeacalColiformContour.AppendLine(@"			<key>normal</key>");
            sbStyleFeacalColiformContour.AppendLine(@"			<styleUrl>#sn_grn-pushpin</styleUrl>");
            sbStyleFeacalColiformContour.AppendLine(@"		</Pair>");
            sbStyleFeacalColiformContour.AppendLine(@"		<Pair>");
            sbStyleFeacalColiformContour.AppendLine(@"			<key>highlight</key>");
            sbStyleFeacalColiformContour.AppendLine(@"			<styleUrl>#sh_grn-pushpin</styleUrl>");
            sbStyleFeacalColiformContour.AppendLine(@"		</Pair>");
            sbStyleFeacalColiformContour.AppendLine(@"	</StyleMap>");
            sbStyleFeacalColiformContour.AppendLine(@"	<Style id=""sn_grn-pushpin"">");
            sbStyleFeacalColiformContour.AppendLine(@"		<IconStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"			<scale>1.1</scale>");
            sbStyleFeacalColiformContour.AppendLine(@"			<Icon>");
            sbStyleFeacalColiformContour.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/grn-pushpin.png</href>");
            sbStyleFeacalColiformContour.AppendLine(@"			</Icon>");
            sbStyleFeacalColiformContour.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbStyleFeacalColiformContour.AppendLine(@"		</IconStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"      <LineStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"         <color>ff000000</color>");
            sbStyleFeacalColiformContour.AppendLine(@"       </LineStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"	</Style>");
            sbStyleFeacalColiformContour.AppendLine(@"	<Style id=""sh_grn-pushpin"">");
            sbStyleFeacalColiformContour.AppendLine(@"		<IconStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"			<scale>1.3</scale>");
            sbStyleFeacalColiformContour.AppendLine(@"			<Icon>");
            sbStyleFeacalColiformContour.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/grn-pushpin.png</href>");
            sbStyleFeacalColiformContour.AppendLine(@"			</Icon>");
            sbStyleFeacalColiformContour.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbStyleFeacalColiformContour.AppendLine(@"		</IconStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"      <LineStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"         <color>ff000000</color>");
            sbStyleFeacalColiformContour.AppendLine(@"       </LineStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"	</Style>");

            sbStyleFeacalColiformContour.AppendLine(@"	<StyleMap id=""msn_blue-pushpin"">");
            sbStyleFeacalColiformContour.AppendLine(@"		<Pair>");
            sbStyleFeacalColiformContour.AppendLine(@"			<key>normal</key>");
            sbStyleFeacalColiformContour.AppendLine(@"			<styleUrl>#sn_blue-pushpin</styleUrl>");
            sbStyleFeacalColiformContour.AppendLine(@"		</Pair>");
            sbStyleFeacalColiformContour.AppendLine(@"		<Pair>");
            sbStyleFeacalColiformContour.AppendLine(@"			<key>highlight</key>");
            sbStyleFeacalColiformContour.AppendLine(@"			<styleUrl>#sh_blue-pushpin</styleUrl>");
            sbStyleFeacalColiformContour.AppendLine(@"		</Pair>");
            sbStyleFeacalColiformContour.AppendLine(@"	</StyleMap>");
            sbStyleFeacalColiformContour.AppendLine(@"	<Style id=""sn_blue-pushpin"">");
            sbStyleFeacalColiformContour.AppendLine(@"		<IconStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"			<scale>1.1</scale>");
            sbStyleFeacalColiformContour.AppendLine(@"			<Icon>");
            sbStyleFeacalColiformContour.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/blue-pushpin.png</href>");
            sbStyleFeacalColiformContour.AppendLine(@"			</Icon>");
            sbStyleFeacalColiformContour.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbStyleFeacalColiformContour.AppendLine(@"		</IconStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"      <LineStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"         <color>ff000000</color>");
            sbStyleFeacalColiformContour.AppendLine(@"       </LineStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"	</Style>");
            sbStyleFeacalColiformContour.AppendLine(@"	<Style id=""sh_blue-pushpin"">");
            sbStyleFeacalColiformContour.AppendLine(@"		<IconStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"			<scale>1.3</scale>");
            sbStyleFeacalColiformContour.AppendLine(@"			<Icon>");
            sbStyleFeacalColiformContour.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/blue-pushpin.png</href>");
            sbStyleFeacalColiformContour.AppendLine(@"			</Icon>");
            sbStyleFeacalColiformContour.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbStyleFeacalColiformContour.AppendLine(@"		</IconStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"      <LineStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"         <color>ff000000</color>");
            sbStyleFeacalColiformContour.AppendLine(@"       </LineStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"	</Style>");

            sbStyleFeacalColiformContour.AppendLine(@"<Style id=""fc_LT14"">");
            sbStyleFeacalColiformContour.AppendLine(@"<LineStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"<color>ff000000</color>");
            sbStyleFeacalColiformContour.AppendLine(@"</LineStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"<PolyStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"<color>6600ff00</color>");
            sbStyleFeacalColiformContour.AppendLine(@"<outline>1</outline>");
            sbStyleFeacalColiformContour.AppendLine(@"</PolyStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"</Style>");

            sbStyleFeacalColiformContour.AppendLine(@"<Style id=""fc_14"">");
            sbStyleFeacalColiformContour.AppendLine(@"<LineStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"<color>ff000000</color>");
            sbStyleFeacalColiformContour.AppendLine(@"</LineStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"<PolyStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"<color>66ff0000</color>");
            sbStyleFeacalColiformContour.AppendLine(@"<outline>1</outline>");
            sbStyleFeacalColiformContour.AppendLine(@"</PolyStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"</Style>");

            sbStyleFeacalColiformContour.AppendLine(@"<Style id=""fc_88"">");
            sbStyleFeacalColiformContour.AppendLine(@"<LineStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"<color>ff000000</color>");
            sbStyleFeacalColiformContour.AppendLine(@"</LineStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"<PolyStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"<color>660000ff</color>");
            sbStyleFeacalColiformContour.AppendLine(@"<outline>1</outline>");
            sbStyleFeacalColiformContour.AppendLine(@"</PolyStyle>");
            sbStyleFeacalColiformContour.AppendLine(@"</Style>");
        }
        private void DrawKMLCurrentsStyle(StringBuilder sbStyleCurrentAnim)
        {
            sbStyleCurrentAnim.AppendLine(@"<Style id=""pink"">");
            sbStyleCurrentAnim.AppendLine(@"<LineStyle>");
            sbStyleCurrentAnim.AppendLine(@"<color>ffff00ff</color>");
            //sbStyleCurrentAnim.AppendLine(@"<gx:physicalWidth>12</gx:physicalWidth>");
            sbStyleCurrentAnim.AppendLine(@"</LineStyle>");
            sbStyleCurrentAnim.AppendLine(@"</Style>");

            sbStyleCurrentAnim.AppendLine(@"<Style id=""green"">");
            sbStyleCurrentAnim.AppendLine(@"<LineStyle>");
            sbStyleCurrentAnim.AppendLine(@"<color>ff00ff00</color>");
            //sbStyleCurrentAnim.AppendLine(@"<gx:physicalWidth>12</gx:physicalWidth>");
            sbStyleCurrentAnim.AppendLine(@"</LineStyle>");
            sbStyleCurrentAnim.AppendLine(@"</Style>");

            sbStyleCurrentAnim.AppendLine(@"<Style id=""yellow"">");
            sbStyleCurrentAnim.AppendLine(@"<LineStyle>");
            sbStyleCurrentAnim.AppendLine(@"<color>ff00ffff</color>");
            //sbStyleCurrentAnim.AppendLine(@"<gx:physicalWidth>12</gx:physicalWidth>");
            sbStyleCurrentAnim.AppendLine(@"</LineStyle>");
            sbStyleCurrentAnim.AppendLine(@"</Style>");
        }
        public void FillElementLayerList(DfsuFile dfsuFile, List<int> SigmaLayerValueList, List<Element> ElementList, List<ElementLayer> ElementLayerList, List<NodeLayer> TopNodeLayerList, List<NodeLayer> BottomNodeLayerList)
        {
            string NotUsed = "";
            if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
            {
                foreach (int layer in SigmaLayerValueList)
                {
                    if (layer > dfsuFile.NumberOfSigmaLayers)
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.MaximumNumberOfSigmaLayers_, dfsuFile.NumberOfSigmaLayers.ToString());
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("MaximumNumberOfSigmaLayers_", dfsuFile.NumberOfSigmaLayers.ToString());
                        return;
                    }
                }
                float Depth = 0.0f;
                List<Element> TempElementList;

                // doing type 32
                TempElementList = (from el in ElementList
                                   where el.Type == 32
                                   && (dfsuFile.Z[el.NodeList[3].ID - 1] == Depth
                                   && dfsuFile.Z[el.NodeList[4].ID - 1] == Depth
                                   && dfsuFile.Z[el.NodeList[5].ID - 1] == Depth)
                                   select el).ToList<Element>();

                foreach (Element el in TempElementList)
                {
                    int Layer = 1;
                    List<Element> ColumnElementList = (from el1 in ElementList
                                                      where el1.Type == 32
                                                      && el1.NodeList[0].X == el.NodeList[0].X
                                                      && el1.NodeList[1].X == el.NodeList[1].X
                                                      && el1.NodeList[2].X == el.NodeList[2].X
                                                      orderby dfsuFile.Z[el1.NodeList[0].ID - 1] descending
                                                      select el1).ToList<Element>();

                    for (int j = 0; j < ColumnElementList.Count; j++)
                    {
                        ElementLayer elementLayer = new ElementLayer();
                        elementLayer.Layer = Layer;
                        elementLayer.ZMin = (from nz in ColumnElementList[j].NodeList select dfsuFile.Z[nz.ID - 1]).Min();
                        elementLayer.ZMax = (from nz in ColumnElementList[j].NodeList select dfsuFile.Z[nz.ID - 1]).Max();
                        elementLayer.Element = ColumnElementList[j];
                        ElementLayerList.Add(elementLayer);

                        NodeLayer nl3 = new NodeLayer();
                        nl3.Layer = Layer;
                        nl3.Z = dfsuFile.Z[ColumnElementList[j].NodeList[3].ID - 1];
                        nl3.Node = ColumnElementList[j].NodeList[3];

                        NodeLayer nl4 = new NodeLayer();
                        nl4.Layer = Layer;
                        nl4.Z = dfsuFile.Z[ColumnElementList[j].NodeList[4].ID - 1];
                        nl4.Node = ColumnElementList[j].NodeList[4];

                        NodeLayer nl5 = new NodeLayer();
                        nl5.Layer = Layer;
                        nl5.Z = dfsuFile.Z[ColumnElementList[j].NodeList[5].ID - 1];
                        nl5.Node = ColumnElementList[j].NodeList[5];

                        TopNodeLayerList.Add(nl3);
                        TopNodeLayerList.Add(nl4);
                        TopNodeLayerList.Add(nl5);

                        NodeLayer nl0 = new NodeLayer();
                        nl0.Layer = Layer;
                        nl0.Z = dfsuFile.Z[ColumnElementList[j].NodeList[0].ID - 1];
                        nl0.Node = ColumnElementList[j].NodeList[0];

                        NodeLayer nl1 = new NodeLayer();
                        nl1.Layer = Layer;
                        nl1.Z = dfsuFile.Z[ColumnElementList[j].NodeList[1].ID - 1];
                        nl1.Node = ColumnElementList[j].NodeList[1];

                        NodeLayer nl2 = new NodeLayer();
                        nl2.Layer = Layer;
                        nl2.Z = dfsuFile.Z[ColumnElementList[j].NodeList[2].ID - 1];
                        nl2.Node = ColumnElementList[j].NodeList[2];

                        BottomNodeLayerList.Add(nl0);
                        BottomNodeLayerList.Add(nl1);
                        BottomNodeLayerList.Add(nl2);

                        Layer += 1;
                    }
                }

                // doing type 33
                TempElementList = (from el in ElementList
                                   where el.Type == 33
                                   && (dfsuFile.Z[el.NodeList[4].ID - 1] == Depth
                                   && dfsuFile.Z[el.NodeList[5].ID - 1] == Depth
                                   && dfsuFile.Z[el.NodeList[6].ID - 1] == Depth
                                   && dfsuFile.Z[el.NodeList[7].ID - 1] == Depth)
                                   select el).ToList<Element>();

                foreach (Element el in TempElementList)
                {
                    int Layer = 1;
                    List<Element> ColumElementList = (from el1 in ElementList
                                                      where el1.Type == 33
                                                      && el1.NodeList[0].X == el.NodeList[0].X
                                                      && el1.NodeList[1].X == el.NodeList[1].X
                                                      && el1.NodeList[2].X == el.NodeList[2].X
                                                      && el1.NodeList[3].X == el.NodeList[3].X
                                                      orderby dfsuFile.Z[el1.NodeList[0].ID - 1] descending
                                                      select el1).ToList<Element>();

                    for (int j = 0; j < ColumElementList.Count; j++)
                    {
                        ElementLayer elementLayer = new ElementLayer();
                        elementLayer.Layer = Layer;
                        elementLayer.ZMin = (from nz in ColumElementList[j].NodeList select dfsuFile.Z[nz.ID - 1]).Min();
                        elementLayer.ZMax = (from nz in ColumElementList[j].NodeList select dfsuFile.Z[nz.ID - 1]).Max();
                        elementLayer.Element = ColumElementList[j];
                        ElementLayerList.Add(elementLayer);

                        NodeLayer nl4 = new NodeLayer();
                        nl4.Layer = Layer;
                        nl4.Z = 0;
                        nl4.Node = ColumElementList[j].NodeList[4];

                        NodeLayer nl5 = new NodeLayer();
                        nl5.Layer = Layer;
                        nl5.Z = 0;
                        nl5.Node = ColumElementList[j].NodeList[5];

                        NodeLayer nl6 = new NodeLayer();
                        nl6.Layer = Layer;
                        nl6.Z = 0;
                        nl6.Node = ColumElementList[j].NodeList[6];

                        NodeLayer nl7 = new NodeLayer();
                        nl7.Layer = Layer;
                        nl7.Z = 0;
                        nl7.Node = ColumElementList[j].NodeList[7];


                        TopNodeLayerList.Add(nl4);
                        TopNodeLayerList.Add(nl5);
                        TopNodeLayerList.Add(nl6);
                        TopNodeLayerList.Add(nl7);

                        NodeLayer nl0 = new NodeLayer();
                        nl0.Layer = Layer;
                        nl0.Z = dfsuFile.Z[ColumElementList[j].NodeList[0].ID - 1];
                        nl0.Node = ColumElementList[j].NodeList[0];

                        NodeLayer nl1 = new NodeLayer();
                        nl1.Layer = Layer;
                        nl1.Z = dfsuFile.Z[ColumElementList[j].NodeList[1].ID - 1];
                        nl1.Node = ColumElementList[j].NodeList[1];

                        NodeLayer nl2 = new NodeLayer();
                        nl2.Layer = Layer;
                        nl2.Z = dfsuFile.Z[ColumElementList[j].NodeList[2].ID - 1];
                        nl2.Node = ColumElementList[j].NodeList[2];

                        NodeLayer nl3 = new NodeLayer();
                        nl3.Layer = Layer;
                        nl3.Z = dfsuFile.Z[ColumElementList[j].NodeList[3].ID - 1];
                        nl3.Node = ColumElementList[j].NodeList[3];

                        BottomNodeLayerList.Add(nl0);
                        BottomNodeLayerList.Add(nl1);
                        BottomNodeLayerList.Add(nl2);
                        BottomNodeLayerList.Add(nl3);

                        Layer += 1;
                    }
                }

                List<ElementLayer> TempElementLayerList = (from el in ElementLayerList
                                                           orderby el.Element.ID
                                                           select el).Distinct().ToList();

                //ElementLayerList = new List<ElementLayer>();
                int OldElemID = 0;
                foreach (ElementLayer el in TempElementLayerList)
                {
                    if (OldElemID == el.Element.ID)
                    {
                        ElementLayerList.Remove(el);
                    }
                    OldElemID = el.Element.ID;
                }

                List<NodeLayer> TempNodeLayerList = (from nl in TopNodeLayerList
                                                     orderby nl.Node.ID
                                                     select nl).Distinct().ToList();

                TopNodeLayerList = new List<NodeLayer>();
                int OldID = 0;
                foreach (NodeLayer nl in TempNodeLayerList)
                {
                    if (OldID != nl.Node.ID)
                    {
                        TopNodeLayerList.Add(nl);
                        OldID = nl.Node.ID;
                    }
                }

                if (BottomNodeLayerList.Count() > 0)
                {
                    TempNodeLayerList = (from nl in BottomNodeLayerList
                                         orderby nl.Node.ID
                                         select nl).Distinct().ToList();

                    BottomNodeLayerList = new List<NodeLayer>();
                    OldID = 0;
                    foreach (NodeLayer nl in TempNodeLayerList)
                    {
                        if (OldID != nl.Node.ID)
                        {
                            BottomNodeLayerList.Add(nl);
                            OldID = nl.Node.ID;
                        }
                    }
                }

            }
            else if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu2D)
            {
                int Layer = 1;
                for (int j = 0; j < ElementList.Count; j++)
                {
                    ElementLayer elementLayer = new ElementLayer();
                    elementLayer.Layer = Layer;
                    elementLayer.ZMin = 0;
                    elementLayer.ZMax = 0;
                    elementLayer.Element = ElementList[j];
                    ElementLayerList.Add(elementLayer);

                    // doing Nodes
                    if (ElementList[j].Type == 21)
                    {
                        NodeLayer nl0 = new NodeLayer();
                        nl0.Layer = Layer;
                        nl0.Z = 0;
                        nl0.Node = ElementList[j].NodeList[0];

                        NodeLayer nl1 = new NodeLayer();
                        nl1.Layer = Layer;
                        nl1.Z = 0;
                        nl1.Node = ElementList[j].NodeList[1];

                        NodeLayer nl2 = new NodeLayer();
                        nl2.Layer = Layer;
                        nl2.Z = 0;
                        nl2.Node = ElementList[j].NodeList[2];

                        TopNodeLayerList.Add(nl0);
                        TopNodeLayerList.Add(nl1);
                        TopNodeLayerList.Add(nl2);
                    }
                    else if (ElementList[j].Type == 24)
                    {
                        NotUsed = TaskRunnerServiceRes.NotImplementedYet;
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("NotImplementedYet");
                        return;
                    }
                    else if (ElementList[j].Type == 25)
                    {
                        NodeLayer nl0 = new NodeLayer();
                        nl0.Layer = Layer;
                        nl0.Z = 0;
                        nl0.Node = ElementList[j].NodeList[0];

                        NodeLayer nl1 = new NodeLayer();
                        nl1.Layer = Layer;
                        nl1.Z = 0;
                        nl1.Node = ElementList[j].NodeList[1];

                        NodeLayer nl2 = new NodeLayer();
                        nl2.Layer = Layer;
                        nl2.Z = 0;
                        nl2.Node = ElementList[j].NodeList[2];

                        NodeLayer nl3 = new NodeLayer();
                        nl3.Layer = Layer;
                        nl3.Z = 0;
                        nl3.Node = ElementList[j].NodeList[3];


                        TopNodeLayerList.Add(nl0);
                        TopNodeLayerList.Add(nl1);
                        TopNodeLayerList.Add(nl2);
                        TopNodeLayerList.Add(nl3);
                    }
                }

                List<ElementLayer> TempElementLayerList = (from el in ElementLayerList
                                                           orderby el.Element.ID
                                                           select el).Distinct().ToList();

                ElementLayerList = new List<ElementLayer>();
                foreach (ElementLayer el in TempElementLayerList)
                {
                    ElementLayerList.Add(el);
                }

                List<NodeLayer> TempNodeLayerList = (from nl in TopNodeLayerList
                                                     orderby nl.Node.ID
                                                     select nl).Distinct().ToList();

                TopNodeLayerList = new List<NodeLayer>();
                int OldID = 0;
                foreach (NodeLayer nl in TempNodeLayerList)
                {
                    if (OldID != nl.Node.ID)
                    {
                        TopNodeLayerList.Add(nl);
                        OldID = nl.Node.ID;
                    }
                }
            }
            else
            {
                NotUsed = TaskRunnerServiceRes.NotImplementedYet;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("NotImplementedYet");
                return;
            }
        }
        public void FillElementListAndNodeList(DfsuFile dfsuFile, List<Element> ElementList, List<Node> NodeList)
        {
            for (int i = 0; i < dfsuFile.NumberOfNodes; i++)
            {
                Node n = new Node()
                {
                    Code = dfsuFile.Code[i],
                    ID = dfsuFile.NodeIds[i],
                    X = (float)dfsuFile.X[i],
                    Y = (float)dfsuFile.Y[i],
                    Z = dfsuFile.Z[i],
                    Value = 0,
                    ConnectNodeList = new List<Node>(),
                    ElementList = new List<Element>()
                };
                NodeList.Add(n);
            }

            for (int i = 0; i < dfsuFile.NumberOfElements; i++)
            {
                Element el = new Element()
                {
                    ID = dfsuFile.ElementIds[i],
                    Type = dfsuFile.ElementType[i],
                    Value = 0,
                    NodeList = new List<Node>(),
                    NumbOfNodes = 0
                };
                ElementList.Add(el);
            }

            for (int i = 0; i < dfsuFile.NumberOfElements; i++)
            {
                int CountNode = 0;
                for (int j = 0; j < dfsuFile.ElementTable[i].Count(); j++)
                {
                    CountNode += 1;
                    ElementList[i].NodeList.Add(NodeList[dfsuFile.ElementTable[i][j] - 1]);
                    if (!NodeList[dfsuFile.ElementTable[i][j] - 1].ElementList.Contains(ElementList[i]))
                    {
                        NodeList[dfsuFile.ElementTable[i][j] - 1].ElementList.Add(ElementList[i]);
                    }
                    for (int k = 0; k < dfsuFile.ElementTable[i].Count(); k++)
                    {
                        if (k != j)
                        {
                            if (!NodeList[dfsuFile.ElementTable[i][j] - 1].ConnectNodeList.Contains(NodeList[dfsuFile.ElementTable[i][k] - 1]))
                            {
                                NodeList[dfsuFile.ElementTable[i][j] - 1].ConnectNodeList.Add(NodeList[dfsuFile.ElementTable[i][k] - 1]);
                            }
                        }
                    }
                }
                ElementList[i].NumbOfNodes = CountNode;
            }
        }
        private void FillVectors21_32(Element el, List<Element> UniqueElementList, float ContourValue, bool Is3D, bool IsTop)
        {
            string NotUsed = "";

            Node Node0 = new Node();
            Node Node1 = new Node();
            Node Node2 = new Node();
            if (Is3D && IsTop)
            {
                Node0 = el.NodeList[3];
                Node1 = el.NodeList[4];
                Node2 = el.NodeList[5];
            }
            else
            {
                Node0 = el.NodeList[0];
                Node1 = el.NodeList[1];
                Node2 = el.NodeList[2];
            }

            int ElemCount01 = (from el1 in UniqueElementList
                               from el2 in Node0.ElementList
                               from el3 in Node1.ElementList
                               where el1 == el2 && el1 == el3
                               select el1).Count();

            int ElemCount02 = (from el1 in UniqueElementList
                               from el2 in Node0.ElementList
                               from el3 in Node2.ElementList
                               where el1 == el2 && el1 == el3
                               select el1).Count();

            int ElemCount12 = (from el1 in UniqueElementList
                               from el2 in Node1.ElementList
                               from el3 in Node2.ElementList
                               where el1 == el2 && el1 == el3
                               select el1).Count();

            if (Node0.Value >= ContourValue && Node1.Value >= ContourValue && Node2.Value >= ContourValue)
            {
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    ForwardVector.Add(Node0.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node1 });
                    BackwardVector.Add(Node1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node0 });
                }
                if (Node0.Code != 0 && Node2.Code != 0 && ElemCount02 == 1)
                {
                    ForwardVector.Add(Node0.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node2 });
                    BackwardVector.Add(Node2.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node2, EndNode = Node0 });
                }
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    ForwardVector.Add(Node1.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node2 });
                    BackwardVector.Add(Node2.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node2, EndNode = Node1 });
                }
            }
            else if (Node0.Value >= ContourValue && Node1.Value >= ContourValue && Node2.Value < ContourValue)
            {
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    ForwardVector.Add(Node0.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node1 });
                    BackwardVector.Add(Node1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node0 });
                }
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node0.ID * 100000 + Node2.ID).First();
                if (Node0.Code != 0 && Node2.Code != 0 && ElemCount02 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node0.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node0 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node1.ID * 100000 + Node2.ID).First();
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node1, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node1 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }
            }
            else if (Node0.Value >= ContourValue && Node1.Value < ContourValue && Node2.Value >= ContourValue)
            {
                if (Node0.Code != 0 && Node2.Code != 0 && ElemCount02 == 1)
                {
                    ForwardVector.Add(Node0.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node2 });
                    BackwardVector.Add(Node2.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node2, EndNode = Node0 });
                }
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node0.ID * 100000 + Node1.ID).First();
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node0.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node0 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node2.ID * 100000 + Node1.ID).First();
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node2.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node2, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node2 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }
            }
            else if (Node0.Value >= ContourValue && Node1.Value < ContourValue && Node2.Value < ContourValue)
            {
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node0.ID * 100000 + Node1.ID).First();
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node0.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node0 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node0.ID * 100000 + Node2.ID).First();
                if (Node0.Code != 0 && Node2.Code != 0 && ElemCount02 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node0.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node0, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node0 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }
            }
            else if (Node0.Value < ContourValue && Node1.Value >= ContourValue && Node2.Value >= ContourValue)
            {
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    ForwardVector.Add(Node1.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node2 });
                    BackwardVector.Add(Node2.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node2, EndNode = Node1 });
                }
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node1.ID * 100000 + Node0.ID).First();
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node1.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node1, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node1 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node2.ID * 100000 + Node0.ID).First();
                if (Node0.Code != 0 && Node2.Code != 0 && ElemCount02 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node2.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node2, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node2 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }
            }
            else if (Node0.Value < ContourValue && Node1.Value >= ContourValue && Node2.Value < ContourValue)
            {
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node1.ID * 100000 + Node0.ID).First();
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node1.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node1, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node1 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node1.ID * 100000 + Node2.ID).First();
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node1, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node1 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }
            }
            else if (Node0.Value < ContourValue && Node1.Value < ContourValue && Node2.Value >= ContourValue)
            {
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node2.ID * 100000 + Node0.ID).First();
                if (Node0.Code != 0 && Node2.Code != 0 && ElemCount02 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node2, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node2 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node2.ID * 100000 + Node1.ID).First();
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node2.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node2, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node2 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }
            }
            else if (Node0.Value < ContourValue && Node1.Value < ContourValue && Node2.Value < ContourValue)
            {
                // no vector to create
            }
            else
            {
                NotUsed = TaskRunnerServiceRes.AllNodesAreSmallerThanContourValue;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("AllNodesAreSmallerThanContourValue");
                return;
            }
        }
        private void FillVectors25_33(Element el, List<Element> UniqueElementList, float ContourValue, bool Is3D, bool IsTop)
        {
            string NotUsed = "";

            Node Node0 = new Node();
            Node Node1 = new Node();
            Node Node2 = new Node();
            Node Node3 = new Node();

            if (Is3D && IsTop)
            {
                Node0 = el.NodeList[4];
                Node1 = el.NodeList[5];
                Node2 = el.NodeList[6];
                Node3 = el.NodeList[7];
            }
            else
            {
                Node0 = el.NodeList[0];
                Node1 = el.NodeList[1];
                Node2 = el.NodeList[2];
                Node3 = el.NodeList[3];
            }

            int ElemCount01 = (from el1 in UniqueElementList
                               from el2 in Node0.ElementList
                               from el3 in Node1.ElementList
                               where el1 == el2 && el1 == el3
                               select el1).Count();

            int ElemCount03 = (from el1 in UniqueElementList
                               from el2 in Node0.ElementList
                               from el3 in Node3.ElementList
                               where el1 == el2 && el1 == el3
                               select el1).Count();

            int ElemCount12 = (from el1 in UniqueElementList
                               from el2 in Node1.ElementList
                               from el3 in Node2.ElementList
                               where el1 == el2 && el1 == el3
                               select el1).Count();

            int ElemCount23 = (from el1 in UniqueElementList
                               from el2 in Node2.ElementList
                               from el3 in Node3.ElementList
                               where el1 == el2 && el1 == el3
                               select el1).Count();

            if (Node0.Value >= ContourValue && Node1.Value >= ContourValue && Node2.Value >= ContourValue && Node3.Value >= ContourValue)
            {
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    ForwardVector.Add(Node0.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node1 });
                    BackwardVector.Add(Node1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node0 });
                }
                if (Node0.Code != 0 && Node2.Code != 0 && ElemCount03 == 1)
                {
                    ForwardVector.Add(Node0.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node3 });
                    BackwardVector.Add(Node3.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node3, EndNode = Node0 });
                }
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    ForwardVector.Add(Node1.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node2 });
                    BackwardVector.Add(Node2.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node2, EndNode = Node1 });
                }
                if (Node2.Code != 0 && Node3.Code != 0 && ElemCount23 == 1)
                {
                    ForwardVector.Add(Node2.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = Node2, EndNode = Node3 });
                    BackwardVector.Add(Node3.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = Node3, EndNode = Node2 });
                }
            }
            else if (Node0.Value >= ContourValue && Node1.Value >= ContourValue && Node2.Value >= ContourValue && Node3.Value < ContourValue)
            {
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    ForwardVector.Add(Node0.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node1 });
                    BackwardVector.Add(Node1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node0 });
                }
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    ForwardVector.Add(Node1.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node2 });
                    BackwardVector.Add(Node2.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node2, EndNode = Node1 });
                }
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node0.ID * 100000 + Node3.ID).First();
                if (Node0.Code != 0 && Node3.Code != 0 && ElemCount03 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node0.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node0 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node2.ID * 100000 + Node3.ID).First();
                if (Node2.Code != 0 && Node3.Code != 0 && ElemCount23 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node2.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node2, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node2 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }
            }
            else if (Node0.Value >= ContourValue && Node1.Value >= ContourValue && Node2.Value < ContourValue && Node3.Value >= ContourValue)
            {
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    ForwardVector.Add(Node0.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node1 });
                    BackwardVector.Add(Node1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node0 });
                }
                if (Node0.Code != 0 && Node3.Code != 0 && ElemCount03 == 1)
                {
                    ForwardVector.Add(Node0.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node3 });
                    BackwardVector.Add(Node3.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node3, EndNode = Node0 });
                }
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node1.ID * 100000 + Node2.ID).First();
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node1.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node1, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node1 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node3.ID * 100000 + Node2.ID).First();
                if (Node3.Code != 0 && Node2.Code != 0 && ElemCount23 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node3.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node3, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node3 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }
            }
            else if (Node0.Value >= ContourValue && Node1.Value < ContourValue && Node2.Value >= ContourValue && Node3.Value >= ContourValue)
            {
                if (Node0.Code != 0 && Node3.Code != 0 && ElemCount03 == 1)
                {
                    ForwardVector.Add(Node0.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node3 });
                    BackwardVector.Add(Node3.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node3, EndNode = Node0 });
                }
                if (Node2.Code != 0 && Node3.Code != 0 && ElemCount23 == 1)
                {
                    ForwardVector.Add(Node2.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = Node2, EndNode = Node3 });
                    BackwardVector.Add(Node3.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = Node3, EndNode = Node2 });
                }
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node0.ID * 100000 + Node1.ID).First();
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node0.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node0 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node2.ID * 100000 + Node1.ID).First();
                if (Node2.Code != 0 && Node1.Code != 0 && ElemCount12 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node2.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node2, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node2 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }
            }
            else if (Node0.Value < ContourValue && Node1.Value >= ContourValue && Node2.Value >= ContourValue && Node3.Value >= ContourValue)
            {
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    ForwardVector.Add(Node1.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node2 });
                    BackwardVector.Add(Node2.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node2, EndNode = Node1 });
                }
                if (Node2.Code != 0 && Node3.Code != 0 && ElemCount23 == 1)
                {
                    ForwardVector.Add(Node2.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = Node2, EndNode = Node3 });
                    BackwardVector.Add(Node3.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = Node3, EndNode = Node2 });
                }
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node1.ID * 100000 + Node0.ID).First();
                if (Node1.Code != 0 && Node0.Code != 0 && ElemCount01 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node1.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node1, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node1 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node3.ID * 100000 + Node0.ID).First();
                if (Node3.Code != 0 && Node0.Code != 0 && ElemCount03 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node3.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node3, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node3 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }
            }
            else if (Node0.Value >= ContourValue && Node1.Value >= ContourValue && Node2.Value < ContourValue && Node3.Value < ContourValue)
            {
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    ForwardVector.Add(Node0.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node1 });
                    BackwardVector.Add(Node1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node0 });
                }
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node0.ID * 100000 + Node3.ID).First();
                if (Node0.Code != 0 && Node3.Code != 0 && ElemCount03 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node0.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node0 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node1.ID * 100000 + Node2.ID).First();
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node1, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node1 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }
            }
            else if (Node0.Value >= ContourValue && Node1.Value < ContourValue && Node2.Value >= ContourValue && Node3.Value < ContourValue)
            {
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node0.ID * 100000 + Node3.ID).First();
                if (Node0.Code != 0 && Node3.Code != 0 && ElemCount03 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node0.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node0 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node0.ID * 100000 + Node1.ID).First();
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node0.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node0, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node0 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }

                Node TempInt3 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node2.ID * 100000 + Node1.ID).First();
                if (Node2.Code != 0 && Node1.Code != 0 && ElemCount12 == 1)
                {
                    if (TempInt3 != null)
                    {
                        ForwardVector.Add(Node2.ID.ToString() + "," + TempInt3.ID.ToString(), new Vector() { StartNode = Node2, EndNode = TempInt3 });
                        BackwardVector.Add(TempInt3.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = TempInt3, EndNode = Node2 });
                    }


                }
                Node TempInt4 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node2.ID * 100000 + Node3.ID).First();
                if (Node2.Code != 0 && Node3.Code != 0 && ElemCount23 == 1)
                {
                    if (TempInt4 != null)
                    {
                        ForwardVector.Add(Node2.ID.ToString() + "," + TempInt4.ID.ToString(), new Vector() { StartNode = Node2, EndNode = TempInt4 });
                        BackwardVector.Add(TempInt4.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = TempInt4, EndNode = Node2 });
                    }
                }

                if (TempInt3 != null && TempInt4 != null)
                {
                    ForwardVector.Add(TempInt3.ID.ToString() + "," + TempInt4.ID.ToString(), new Vector() { StartNode = TempInt3, EndNode = TempInt4 });
                    BackwardVector.Add(TempInt4.ID.ToString() + "," + TempInt3.ID.ToString(), new Vector() { StartNode = TempInt4, EndNode = TempInt3 });
                }

            }
            else if (Node0.Value < ContourValue && Node1.Value >= ContourValue && Node2.Value >= ContourValue && Node3.Value < ContourValue)
            {
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node1.ID * 100000 + Node0.ID).First();
                if (Node1.Code != 0 && Node0.Code != 0 && ElemCount01 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node1.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node1, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node1 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node2.ID * 100000 + Node3.ID).First();
                if (Node2.Code != 0 && Node3.Code != 0 && ElemCount23 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node2.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node2, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node2 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }

            }
            else if (Node0.Value >= ContourValue && Node1.Value < ContourValue && Node2.Value < ContourValue && Node3.Value >= ContourValue)
            {
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node0.ID * 100000 + Node1.ID).First();
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node0.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node0 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node3.ID * 100000 + Node2.ID).First();
                if (Node3.Code != 0 && Node2.Code != 0 && ElemCount23 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node3.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node3, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node3 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }

            }
            else if (Node0.Value < ContourValue && Node1.Value >= ContourValue && Node2.Value < ContourValue && Node3.Value >= ContourValue)
            {
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node3.ID * 100000 + Node0.ID).First();
                if (Node3.Code != 0 && Node0.Code != 0 && ElemCount03 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node3.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node3, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node3 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node3.ID * 100000 + Node2.ID).First();
                if (Node3.Code != 0 && Node2.Code != 0 && ElemCount23 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node3.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node3, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node3 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }

                Node TempInt3 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node1.ID * 100000 + Node0.ID).First();
                if (Node1.Code != 0 && Node0.Code != 0 && ElemCount01 == 1)
                {
                    if (TempInt3 != null)
                    {
                        ForwardVector.Add(Node1.ID.ToString() + "," + TempInt3.ID.ToString(), new Vector() { StartNode = Node1, EndNode = TempInt3 });
                        BackwardVector.Add(TempInt3.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = TempInt3, EndNode = Node1 });
                    }


                }
                Node TempInt4 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node1.ID * 100000 + Node2.ID).First();
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    if (TempInt4 != null)
                    {
                        ForwardVector.Add(Node1.ID.ToString() + "," + TempInt4.ID.ToString(), new Vector() { StartNode = Node1, EndNode = TempInt4 });
                        BackwardVector.Add(TempInt4.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = TempInt4, EndNode = Node1 });
                    }
                }

                if (TempInt3 != null && TempInt4 != null)
                {
                    ForwardVector.Add(TempInt3.ID.ToString() + "," + TempInt4.ID.ToString(), new Vector() { StartNode = TempInt3, EndNode = TempInt4 });
                    BackwardVector.Add(TempInt4.ID.ToString() + "," + TempInt3.ID.ToString(), new Vector() { StartNode = TempInt4, EndNode = TempInt3 });
                }

            }
            else if (Node0.Value < ContourValue && Node1.Value < ContourValue && Node2.Value >= ContourValue && Node3.Value >= ContourValue)
            {
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node3.ID * 100000 + Node0.ID).First();
                if (Node3.Code != 0 && Node0.Code != 0 && ElemCount03 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node3.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node3, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node3 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node2.ID * 100000 + Node1.ID).First();
                if (Node2.Code != 0 && Node1.Code != 0 && ElemCount12 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node2.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node2, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node2 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }

            }
            else if (Node0.Value >= ContourValue && Node1.Value < ContourValue && Node2.Value < ContourValue && Node3.Value < ContourValue)
            {
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node0.ID * 100000 + Node1.ID).First();
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node0.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node0 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node0.ID * 100000 + Node3.ID).First();
                if (Node0.Code != 0 && Node3.Code != 0 && ElemCount03 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node0.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node0, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node0 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }

            }
            else if (Node0.Value < ContourValue && Node1.Value >= ContourValue && Node2.Value < ContourValue && Node3.Value < ContourValue)
            {
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node1.ID * 100000 + Node0.ID).First();
                if (Node1.Code != 0 && Node0.Code != 0 && ElemCount01 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node1.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node1, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node1 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node1.ID * 100000 + Node2.ID).First();
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node1, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node1 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }

            }
            else if (Node0.Value < ContourValue && Node1.Value < ContourValue && Node2.Value >= ContourValue && Node3.Value < ContourValue)
            {
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node2.ID * 100000 + Node1.ID).First();
                if (Node2.Code != 0 && Node1.Code != 0 && ElemCount12 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node2, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node2 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node2.ID * 100000 + Node3.ID).First();
                if (Node2.Code != 0 && Node3.Code != 0 && ElemCount23 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node2.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node2, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node2 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }

            }
            else if (Node0.Value < ContourValue && Node1.Value < ContourValue && Node2.Value < ContourValue && Node3.Value >= ContourValue)
            {
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node3.ID * 100000 + Node0.ID).First();
                if (Node3.Code != 0 && Node0.Code != 0 && ElemCount03 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node3.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node3, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node3 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node3.ID * 100000 + Node2.ID).First();
                if (Node3.Code != 0 && Node2.Code != 0 && ElemCount23 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node3.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node3, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node3 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }

            }
            else if (Node0.Value < ContourValue && Node1.Value < ContourValue && Node2.Value < ContourValue && Node3.Value < ContourValue)
            {
                // no vector to create
            }
            else
            {
                NotUsed = TaskRunnerServiceRes.AllNodesAreSmallerThanContourValue;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("AllNodesAreSmallerThanContourValue");
                return;
            }
        }
        public List<ElementLayer> GetElementSurrondingEachPoint(List<ElementLayer> ElementLayerList, List<Node> Nodes)
        {
            List<ElementLayer> AllElementList = new List<ElementLayer>();

            foreach (ElementLayer el in ElementLayerList)
            {
                float XMin = (from a in el.Element.NodeList
                              select a.X).Min();
                float YMin = (from a in el.Element.NodeList
                              select a.Y).Min();
                float XMax = (from a in el.Element.NodeList
                              select a.X).Max();
                float YMax = (from a in el.Element.NodeList
                              select a.Y).Max();

                foreach (Node n in Nodes)
                {
                    if ((n.X > XMin && n.X < XMax) && (n.Y > YMin && n.Y < YMax))
                    {
                        Point p = new Point((int)(n.X * 10000000), (int)(n.Y * 10000000));
                        if (el.Element.Type == 21 || el.Element.Type == 32)
                        {
                            Point[] poly = 
                            { 
                                new Point() { X = (int)(el.Element.NodeList[0].X*10000000), Y = (int)(el.Element.NodeList[0].Y*10000000) }, 
                                new Point() { X = (int)(el.Element.NodeList[1].X*10000000), Y = (int)(el.Element.NodeList[1].Y*10000000) }, 
                                new Point() { X = (int)(el.Element.NodeList[2].X*10000000), Y = (int)(el.Element.NodeList[2].Y*10000000) }, 
                                new Point() { X = (int)(el.Element.NodeList[0].X*10000000), Y = (int)(el.Element.NodeList[0].Y*10000000) }, 
                                           };

                            if (PointInPolygon(p, poly))
                            {
                                AllElementList.Add(el);
                            }
                        }
                        else if (el.Element.Type == 25 || el.Element.Type == 33)
                        {
                            Point[] poly = 
                            { 
                                new Point() { X = (int)(el.Element.NodeList[0].X*10000000), Y = (int)(el.Element.NodeList[0].Y*10000000) }, 
                                new Point() { X = (int)(el.Element.NodeList[1].X*10000000), Y = (int)(el.Element.NodeList[1].Y*10000000) }, 
                                new Point() { X = (int)(el.Element.NodeList[2].X*10000000), Y = (int)(el.Element.NodeList[2].Y*10000000) }, 
                                new Point() { X = (int)(el.Element.NodeList[3].X*10000000), Y = (int)(el.Element.NodeList[3].Y*10000000) }, 
                                new Point() { X = (int)(el.Element.NodeList[0].X*10000000), Y = (int)(el.Element.NodeList[0].Y*10000000) }, 
                                           };
                            if (PointInPolygon(p, poly))
                            {
                                AllElementList.Add(el);
                            }
                        }
                        else
                        {
                            return null;
                        }
                    }
                }
            }

            return AllElementList;
        }
        //private StringBuilder GetModelAndSourceDesc(List<float> ContourValueList, int mikeScenarioID, PFSFile pfsFile)
        //{
        //    StringBuilder sbToReturn = new StringBuilder();

        //    MikeScenarioService mikeScenarioService = new MikeScenarioService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    MikeScenarioModel mikeScenarioModel = mikeScenarioService.GetMikeScenarioModelWithMikeScenarioIDDB(mikeScenarioID);
        //    if (mikeScenarioModel == null)
        //    {
        //        return new StringBuilder(string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, "MIKEScenario", "MikeScenarioID", mikeScenarioModel.MikeScenarioID));
        //    }

        //    MikeSourceService mikeSourceService = new MikeSourceService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    List<MikeSourceModel> mikeSourceModelList = mikeSourceService.GetMikeSourceModelListWithMikeScenarioTVItemIDDB(mikeScenarioModel.MikeScenarioTVItemID);
        //    if (mikeSourceModelList.Count == 0)
        //    {
        //        return new StringBuilder(string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, "MikeSource", "MikeScenarioID", mikeScenarioModel.MikeScenarioID));
        //    }

        //    sbToReturn.AppendLine(@"<description><![CDATA[");
        //    sbToReturn.AppendLine(string.Format(@"<h2>{0}</h2>", TaskRunnerServiceRes.ModelParameters));
        //    sbToReturn.AppendLine(@"<ul>");
        //    sbToReturn.AppendLine(string.Format(@"<li><b>{0}:</b> {1:yyyy/MM/dd HH:mm:ss tt} (UTC)</li>", TaskRunnerServiceRes.CreatedDate, m21_3fmInput.topfileinfo.Created));
        //    sbToReturn.AppendLine(string.Format(@"<li><b>{0}:</b> {1:yyyy/MM/dd HH:mm:ss tt} (UTC)</li>", TaskRunnerServiceRes.ScenarioStartTime, m21_3fmInput.femEngineHD.time.start_time));
        //    sbToReturn.AppendLine(string.Format(@"<li><b>{0}:</b> {1:yyyy/MM/dd HH:mm:ss tt} (UTC)</li>", TaskRunnerServiceRes.ScenarioEndTime, m21_3fmInput.femEngineHD.time.start_time.AddSeconds(m21_3fmInput.femEngineHD.time.time_step_interval * m21_3fmInput.femEngineHD.time.number_of_time_steps)));
        //    sbToReturn.AppendLine(@"</ul>");
        //    sbToReturn.AppendLine(@"<ul>");

        //    foreach (float cv in ContourValueList)
        //    {
        //        if (cv >= 14 && cv < 88)
        //        {
        //            sbToReturn.AppendLine(string.Format(@"<li><span style=""background-color:Blue; color:White"">{0} = {1:F0}</span</li>", TaskRunnerServiceRes.FCMPNPollutionContour, cv));
        //        }
        //        else if (cv >= 88)
        //        {
        //            sbToReturn.AppendLine(string.Format(@"<li><span style=""background-color:Red; color:White"">{0} = {1:F0}</span></li>", TaskRunnerServiceRes.FCMPNPollutionContour, cv));
        //        }
        //        else
        //        {
        //            sbToReturn.AppendLine(string.Format(@"<li><span style=""background-color:Green; color:White"">{0} = {1:F0}</span></li>", TaskRunnerServiceRes.FCMPNPollutionContour, cv));
        //        }

        //    }
        //    sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.AverageDecayFactor + @":</b> " + (m21_3fmInput.femEngineHD.transport_module.decay.component["COMPONENT_1"].constant_value * 3600 * 24).ToString("F6").Replace(",", ".") + @" /" + TaskRunnerServiceRes.DayLowerCase + @"</li>");
        //    if (m21_3fmInput.femEngineHD.transport_module.decay.component["COMPONENT_1"].format == 0)
        //    {
        //        sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.DecayIsConstant + @"</b></li>");
        //    }
        //    else
        //    {
        //        sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.DecayIsVariable + @"</b></li>");
        //        sbToReturn.AppendLine(@"<ul><li><b>" + TaskRunnerServiceRes.Amplitude + @":</b> " + ((double)mikeScenarioModel.DecayFactorAmplitude).ToString("F6").Replace(",", ".") + @"</li></ul>");
        //    }
        //    if (m21_3fmInput.femEngineHD.hydrodynamic_module.wind_forcing.constant_speed > 0)
        //    {
        //        sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.Wind + @":</b> " + (m21_3fmInput.femEngineHD.hydrodynamic_module.wind_forcing.constant_speed * 3.6).ToString("F1").Replace(",", ".") + @" (km/h)   " + (m21_3fmInput.femEngineHD.hydrodynamic_module.wind_forcing.constant_speed).ToString("F1").Replace(",", ".") + @" (m/s)</li>");
        //        sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.WindDirection + @":</b> " + (m21_3fmInput.femEngineHD.hydrodynamic_module.wind_forcing.constant_direction).ToString("F1").Replace(",", ".") + @" " + TaskRunnerServiceRes.DegreeLowerCase + " (0 = " + TaskRunnerServiceRes.NorthClockwiseLowerCase + @")</li>");
        //    }
        //    else
        //    {
        //        sbToReturn.AppendLine("<li><b>" + TaskRunnerServiceRes.NoWind + @"</b></li>");
        //    }
        //    switch (m21_3fmInput.femEngineHD.hydrodynamic_module.density.type)
        //    {
        //        case M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.DENSITY.TYPE.Barotropic:
        //        case M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.DENSITY.TYPE.FunctionOfTemperatureAndSalinity:
        //        case M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.DENSITY.TYPE.FunctionOfTemperature:
        //        case M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.DENSITY.TYPE.FunctionOfSalinity:
        //            sbToReturn.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.DensityType + @":</b> {0}</li>", GetSpaceBeforeCapital(m21_3fmInput.femEngineHD.hydrodynamic_module.density.type.ToString())));
        //            sbToReturn.AppendLine("<ul>");
        //            sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.TemperatureReference + @":</b> " + (m21_3fmInput.femEngineHD.hydrodynamic_module.density.temperature_reference).ToString("F1").Replace(",", ".") + TaskRunnerServiceRes.Celcius + @"</li>");
        //            sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.SalinityReference + @":</b> " + (m21_3fmInput.femEngineHD.hydrodynamic_module.density.salinity_reference).ToString("F1").Replace(",", ".") + TaskRunnerServiceRes.PSU + @"</li>");
        //            sbToReturn.AppendLine("</ul>");
        //            break;
        //        default:
        //            sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.DensityTypeNotDefine + @"</b></li>");
        //            break;
        //    }
        //    switch (m21_3fmInput.femEngineHD.hydrodynamic_module.bed_resistance.type)
        //    {
        //        case M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.BED_RESISTANCE.TYPE.NoBedResistance:
        //            sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.BedResistanceType + @":</b> <ul><li>" + GetSpaceBeforeCapital(m21_3fmInput.femEngineHD.hydrodynamic_module.bed_resistance.type.ToString()) + ": " + (m21_3fmInput.femEngineHD.hydrodynamic_module.bed_resistance.manning_number.constant_value).ToString().Replace(",", ".") + @"</li></ul></li>");
        //            break;
        //        case M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.BED_RESISTANCE.TYPE.ChezyNumber:
        //            sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.BedResistanceType + @":</b> <ul><li>" + GetSpaceBeforeCapital(m21_3fmInput.femEngineHD.hydrodynamic_module.bed_resistance.type.ToString()) + ": " + (m21_3fmInput.femEngineHD.hydrodynamic_module.bed_resistance.chezy_number.constant_value).ToString().Replace(",", ".") + @"</li></ul></li>");
        //            break;
        //        case M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.BED_RESISTANCE.TYPE.ManningNumber:
        //            sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.BedResistanceType + @":</b> <ul><li>" + GetSpaceBeforeCapital(m21_3fmInput.femEngineHD.hydrodynamic_module.bed_resistance.type.ToString()) + ": " + (m21_3fmInput.femEngineHD.hydrodynamic_module.bed_resistance.manning_number.constant_value).ToString().Replace(",", ".") + @"</li></ul></li>");
        //            break;
        //        default:
        //            sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.BedResistanceType + @":</b> " + TaskRunnerServiceRes.NotDetermined + @"</li>");
        //            break;
        //    }
        //    sbToReturn.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.ResultFrequency + @":</b> {0:F0} {1}</li>", mikeScenarioModel.ResultFrequency_min, TaskRunnerServiceRes.MinutesLowerCase));
        //    sbToReturn.AppendLine(@"</ul>");

        //    // showing not used sources 
        //    foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.SOURCES.SOURCE> kvp in m21_3fmInput.femEngineHD.hydrodynamic_module.sources.source)
        //    {
        //        M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.SOURCES.SOURCE shd = kvp.Value;
        //        if (kvp.Value.include == 1)
        //        {
        //            if (shd.Name.Substring(1, shd.Name.Length - 2).ToLower().StartsWith("river"))
        //            {
        //                sbToReturn.AppendLine(string.Format(@"<h2 style='Color: Blue'>{0} ({1})</h2>", shd.Name.Substring(1, shd.Name.Length - 2), TaskRunnerServiceRes.IncludedLowerCase));
        //            }
        //            else
        //            {
        //                sbToReturn.AppendLine(string.Format(@"<h2 style='Color: Green'>{0} ({1})</h2>", shd.Name.Substring(1, shd.Name.Length - 2), TaskRunnerServiceRes.IncludedLowerCase));
        //            }
        //        }
        //        else
        //        {
        //            sbToReturn.AppendLine(string.Format(@"<h2 style='Color: Red'>{0} ({1})</h2>", shd.Name.Substring(1, shd.Name.Length - 2), TaskRunnerServiceRes.NotIncludedLowerCase));
        //        }
        //        sbToReturn.AppendLine(@"<h3>" + TaskRunnerServiceRes.Effluent + @"</h3>");
        //        sbToReturn.AppendLine(@"<ul>");
        //        sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.Coordinates + @"</b>");
        //        sbToReturn.AppendLine(string.Format(@"&nbsp;&nbsp;&nbsp; {0:F5} &nbsp; {1:F5}</li><br />", shd.coordinates.y, shd.coordinates.x).Replace(",", "."));

        //        MikeSourceModel mikeSourceModelLocal = (from msl in mikeSourceModelList
        //                                                where msl.SourceNumberString == kvp.Key
        //                                                select msl).FirstOrDefault<MikeSourceModel>();

        //        if (mikeSourceModelLocal == null)
        //        {
        //            return new StringBuilder(string.Format(TaskRunnerServiceRes.CouldNotFind_, kvp.Key));
        //        }

        //        if ((bool)mikeSourceModelLocal.IsContinuous)
        //        {
        //            sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.IsContinuous + @"</b></li>");
        //            sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.Flow + @":</b> " + (shd.constant_value).ToString("F6").Replace(",", ".") + " (m3/s)  " + (shd.constant_value * 24 * 3600).ToString("F1").Replace(",", ".") + @" (m3/" + TaskRunnerServiceRes.DayLowerCase + @")</li>");
        //            sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.FCMPNPer100ML + @":</b> " + (m21_3fmInput.femEngineHD.transport_module.sources.source[kvp.Key].component["COMPONENT_1"].constant_value).ToString("F0").Replace(",", ".") + @"</li>");
        //            sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.Temperature + @":</b> " + (m21_3fmInput.femEngineHD.hydrodynamic_module.temperature_salinity_module.sources.source[kvp.Key].temperature.constant_value).ToString("F1").Replace(",", ".") + @" " + TaskRunnerServiceRes.Celcius + @"</li>");
        //            sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.Salinity + @":</b> " + (m21_3fmInput.femEngineHD.hydrodynamic_module.temperature_salinity_module.sources.source[kvp.Key].salinity.constant_value).ToString("F1").Replace(",", ".") + @" " + TaskRunnerServiceRes.PSU + @"</li>");
        //        }
        //        else
        //        {
        //            sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.IsNotContinuous + @"</b></li>");

        //            MikeSourceStartEndService mikeSourceStartEndService = new MikeSourceStartEndService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //            List<MikeSourceStartEndModel> mikeSourceStartEndModelListLocal = mikeSourceStartEndService.GetMikeSourceStartEndModelListWithMikeSourceIDDB(mikeSourceModelLocal.MikeSourceID);
        //            int CountMikeSourceStartEnd = 0;
        //            foreach (MikeSourceStartEndModel mssem in mikeSourceStartEndModelListLocal)
        //            {
        //                CountMikeSourceStartEnd += 1;
        //                sbToReturn.AppendLine(@"<br /><b>" + TaskRunnerServiceRes.Spill + @": " + CountMikeSourceStartEnd + "</b><br />");
        //                sbToReturn.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.SpillStartTime + @":</b> {0:yyyy/MM/dd HH:mm:ss tt} (UTC)</li>", mssem.StartDateAndTime_Local));
        //                sbToReturn.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.SpillEndTime + @":</b> {0:yyyy/MM/dd HH:mm:ss tt} (UTC)</li>", mssem.EndDateAndTime_Local));
        //                sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.FlowStart + @":</b> " + ((double)mssem.SourceFlowStart_m3_day / 24 / 3600).ToString("F6").Replace(",", ".") + @" (m3/s)  " + ((double)mssem.SourceFlowStart_m3_day).ToString("F0").Replace(",", ".") + @" (m3/" + TaskRunnerServiceRes.DayLowerCase + @")</li>");
        //                sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.FlowEnd + @":</b> " + ((double)mssem.SourceFlowEnd_m3_day / 24 / 3600).ToString("F6").Replace(",", ".") + @" (m3/s)  " + ((double)mssem.SourceFlowEnd_m3_day).ToString("F0").Replace(",", ".") + @" (m3/" + TaskRunnerServiceRes.DayLowerCase + @")</li>");
        //                sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.FCMPNPer100MLStart + @":</b> " + ((double)mssem.SourcePollutionStart_MPN_100ml).ToString("F0").Replace(",", ".") + @"</li>");
        //                sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.FCMPNPer100MLEnd + @":</b> " + ((double)mssem.SourcePollutionEnd_MPN_100ml).ToString("F0").Replace(",", ".") + @"</li>");
        //                sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.TemperatureStart + @":</b> " + ((double)mssem.SourceTemperatureStart_C).ToString("F0").Replace(",", ".") + @" " + TaskRunnerServiceRes.Celcius + @"</li>");
        //                sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.TempertureEnd + @":</b> " + ((double)mssem.SourceTemperatureEnd_C).ToString("F0").Replace(",", ".") + @" " + TaskRunnerServiceRes.Celcius + @"</li>");
        //                sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.SalinityStart + @":</b> " + ((double)mssem.SourceSalinityStart_PSU).ToString("F0").Replace(",", ".") + @" " + TaskRunnerServiceRes.PSU + @"</li>");
        //                sbToReturn.AppendLine(@"<li><b>" + TaskRunnerServiceRes.SalinityEnd + @":</b> " + ((double)mssem.SourceSalinityEnd_PSU).ToString("F0").Replace(",", ".") + @" " + TaskRunnerServiceRes.PSU + @"</li>");
        //            }
        //        }
        //        sbToReturn.AppendLine(@"</ul>");
        //    }


        //    sbToReturn.AppendLine(@"<iframe src=""about:"" width=""600"" height=""1"" />");
        //    sbToReturn.AppendLine(@"]]></description>");

        //    return sbToReturn;
        //}
        private string GetNodeSurroundingPolygon(Node node)
        {
            StringBuilder coordStr = new StringBuilder();
            List<Node> ContourNodeList = new List<Node>();
            Element LastElement = new Element();
            Node LastNode = new Node();
            Element FirstElement = new Element();

            if (node.Code != 0)
            {
                ContourNodeList.Add(node);

                // get first element
                bool ElementFound = false;
                while (!ElementFound)
                {
                    List<Element> PotentialElemList = (from el in node.ElementList
                                                       where ((el.NodeList[0] != node && el.NodeList[0].Code != 0)
                                                       || (el.NodeList[1] != node && el.NodeList[1].Code != 0)
                                                       || (el.NodeList[2] != node && el.NodeList[2].Code != 0))
                                                       select el).ToList<Element>();

                    foreach (Element elem in PotentialElemList)
                    {
                        LastElement = elem;

                        foreach (Node n in LastElement.NodeList.Where(nn => nn != node))
                        {
                            if (n.Code != 0)
                            {
                                // check to see if the node and n are part of a single element
                                List<Element> AllElem = (from el in node.ElementList
                                                         where ((el.NodeList[0] == n)
                                                         || (el.NodeList[1] == n)
                                                         || (el.NodeList[2] == n))
                                                         select el).ToList<Element>();

                                if (AllElem.Count == 1)
                                {
                                    LastNode = n;
                                    break;
                                }
                            }
                        }
                        if (LastNode.ID != 0)
                        {
                            break;
                        }
                    }

                    FirstElement = LastElement;

                    if (LastElement.ID != 0 && LastNode.ID != 0)
                    {
                        ElementFound = true;
                    }
                }

                if (LastElement == null || LastNode == null)
                {
                    // this should not happen
                    return "";
                }

                Node NewNode = new Node();
                NewNode.ID = LastNode.ID + 1000000;
                NewNode.X = (node.X + LastNode.X) / 2;
                NewNode.Y = (node.Y + LastNode.Y) / 2;
                ContourNodeList.Add(NewNode);

                foreach (Node n in LastElement.NodeList.Where(nn => nn != node && nn != LastNode))
                {
                    LastNode = n;
                    break;
                }

                Node NewNode2 = new Node();
                NewNode2.ID = LastNode.ID + 2000000;
                NewNode2.X = (LastElement.NodeList[0].X + LastElement.NodeList[1].X + LastElement.NodeList[2].X) / 3;
                NewNode2.Y = (LastElement.NodeList[0].Y + LastElement.NodeList[1].Y + LastElement.NodeList[2].Y) / 3;
                ContourNodeList.Add(NewNode2);

            }
            else
            {
                // Get first Element
                LastElement = (from el in node.ElementList select el).FirstOrDefault<Element>();

                if (LastElement == null)
                {
                    // this should not happen
                    return "";
                }

                FirstElement = LastElement;

                foreach (Node n in LastElement.NodeList.Where(nn => nn != node))
                {
                    LastNode = n;
                    break;
                }

                Node NewNode = new Node();
                NewNode.ID = LastNode.ID + 2000000;
                NewNode.X = (LastElement.NodeList[0].X + LastElement.NodeList[1].X + LastElement.NodeList[2].X) / 3;
                NewNode.Y = (LastElement.NodeList[0].Y + LastElement.NodeList[1].Y + LastElement.NodeList[2].Y) / 3;
                ContourNodeList.Add(NewNode);
            }


            bool MoreNodes = true;
            while (MoreNodes)
            {
                Element NextElement = (from el in node.ElementList
                                       where el != LastElement
                                       && ((el.NodeList[0] == LastNode)
                                       || (el.NodeList[1] == LastNode)
                                       || (el.NodeList[2] == LastNode))
                                       select el).FirstOrDefault<Element>();

                if (NextElement == null)
                {
                    if (LastNode.Code != 0 && node.Code != 0)
                    {
                        Node NewNode2 = new Node();
                        NewNode2.ID = LastNode.ID + 1000000;
                        NewNode2.X = (node.X + LastNode.X) / 2;
                        NewNode2.Y = (node.Y + LastNode.Y) / 2;
                        ContourNodeList.Add(NewNode2);

                    }
                    MoreNodes = false;
                    break;
                }

                LastElement = NextElement;

                if (FirstElement == LastElement)
                {
                    if (LastNode.Code != 0 && node.Code != 0)
                    {
                        Node NewNode2 = new Node();
                        NewNode2.ID = LastNode.ID + 1000000;
                        NewNode2.X = (node.X + LastNode.X) / 2;
                        NewNode2.Y = (node.Y + LastNode.Y) / 2;
                        ContourNodeList.Add(NewNode2);

                    }
                    MoreNodes = false;
                    break;
                }

                foreach (Node n in LastElement.NodeList.Where(nn => nn != node && nn != LastNode))
                {
                    LastNode = n;
                    break;
                }

                Node NewNode = new Node();
                NewNode.ID = LastNode.ID + 2000000;
                NewNode.X = (LastElement.NodeList[0].X + LastElement.NodeList[1].X + LastElement.NodeList[2].X) / 3;
                NewNode.Y = (LastElement.NodeList[0].Y + LastElement.NodeList[1].Y + LastElement.NodeList[2].Y) / 3;
                ContourNodeList.Add(NewNode);

            }

            ContourNodeList.Add(ContourNodeList[0]);

            MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            if (mapInfoService.CalculateAreaOfPolygon(ContourNodeList) < 0)
            {
                ContourNodeList.Reverse();
            }

            foreach (Node n in ContourNodeList)
            {
                coordStr.Append(n.X.ToString().Replace(",", ".") + @"," + n.Y.ToString().Replace(",", ".") + ",0 ");
            }

            return coordStr.ToString();

        }
        private string GetSpaceBeforeCapital(string TextToParse)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(TextToParse[0]);
            for (int i = 1; i < TextToParse.Count(); i++)
            {
                if (TextToParse[i] >= "A".ToCharArray()[0] && TextToParse[i] <= "Z".ToCharArray()[0])
                {
                    sb.Append(" " + TextToParse[i]);
                }
                else
                {
                    sb.Append(TextToParse[i]);
                }
            }
            return sb.ToString();
        }
        private void InsertNewNodeInInterpolatedBathymetryContourNodeList(Node NodeLarge, Node NodeSmall, float ContourValue)
        {
            PolyPoint point = new PolyPoint();
            point.XCoord = NodeSmall.X + (NodeLarge.X - NodeSmall.X) * (ContourValue - Math.Abs(NodeSmall.Z)) / (Math.Abs(NodeLarge.Z) - Math.Abs(NodeSmall.Z));
            point.YCoord = NodeSmall.Y + (NodeLarge.Y - NodeSmall.Y) * (ContourValue - Math.Abs(NodeSmall.Z)) / (Math.Abs(NodeLarge.Z) - Math.Abs(NodeSmall.Z));

            Node NewNode = new Node();
            NewNode.ID = 100000 * NodeLarge.ID + NodeSmall.ID;
            NewNode.X = point.XCoord;
            NewNode.Y = point.YCoord;
            NewNode.Value = ContourValue;

            if (InterpolatedContourNodeList.Where(nn => nn.ID == NewNode.ID).Count() == 0)
            {
                InterpolatedContourNodeList.Add(NewNode);
            }
        }
        private void InsertNewNodeInInterpolatedContourNodeList(DfsuFile dfsuFile, Node NodeLarge, Node NodeSmall, float ContourValue)
        {
            PolyPoint point = new PolyPoint();
            point.XCoord = NodeSmall.X + (NodeLarge.X - NodeSmall.X) * (ContourValue - NodeSmall.Value) / (NodeLarge.Value - NodeSmall.Value);
            point.YCoord = NodeSmall.Y + (NodeLarge.Y - NodeSmall.Y) * (ContourValue - NodeSmall.Value) / (NodeLarge.Value - NodeSmall.Value);
            point.Z = dfsuFile.Z[NodeSmall.ID - 1] + (dfsuFile.Z[NodeLarge.ID - 1] - dfsuFile.Z[NodeSmall.ID - 1]) * (ContourValue - NodeSmall.Value) / (NodeLarge.Value - NodeSmall.Value);

            Node NewNode = new Node();
            NewNode.ID = 100000 * NodeLarge.ID + NodeSmall.ID;
            NewNode.X = point.XCoord;
            NewNode.Y = point.YCoord;
            NewNode.Z = point.Z;
            NewNode.Value = ContourValue;

            if (InterpolatedContourNodeList.Where(nn => nn.ID == NewNode.ID).Count() == 0)
            {
                InterpolatedContourNodeList.Add(NewNode);
            }
        }
        private List<ElementLayer> ParseKMLPath(string KMLTextPathForVector, List<ElementLayer> ElementLayerList)
        {
            string NotUsed = "";

            List<ElementLayer> AllElementList = new List<ElementLayer>();
            List<Node> PathNodeList = new List<Node>();

            if (KMLTextPathForVector.Trim() == "")
            {
                foreach (ElementLayer el in ElementLayerList)
                {
                    AllElementList.Add(el);
                }
            }
            else
            {
                try
                {
                    XmlReader reader = XmlReader.Create(new StringReader(KMLTextPathForVector));
                    while (reader.Read())
                    {
                        if (reader.Name == "coordinates")
                        {
                            string AllCoordinates = reader.ReadElementContentAsString().Trim();

                            string[] xyzArray = AllCoordinates.Split(" ".ToCharArray()[0]);
                            foreach (string xyz in xyzArray)
                            {
                                string[] xyzStr = xyz.Split(",".ToCharArray()[0]);
                                if (xyzStr.Count() != 3)
                                {
                                    return null;
                                }
                                Node n = new Node();
                                if (Thread.CurrentThread.CurrentCulture.Name == "fr-CA")
                                {
                                    n.X = float.Parse(xyzStr[0].Replace(".", ","));
                                    n.Y = float.Parse(xyzStr[1].Replace(".", ","));
                                }
                                else
                                {
                                    n.X = float.Parse(xyzStr[0]);
                                    n.Y = float.Parse(xyzStr[1]);
                                }

                                PathNodeList.Add(n);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, "KMLPath" + ex.Message);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", "KMLPath" + ex.Message);
                    return new List<ElementLayer>();
                }

                AllElementList = GetElementSurrondingEachPoint(ElementLayerList, PathNodeList);
            }

            return AllElementList.Distinct().ToList();
        }
        private bool PointInPolygon(Point p, Point[] poly)
        {
            Point p1, p2;
            bool inside = false;
            if (poly.Length < 3)
            {
                return inside;
            }

            Point oldPoint = new Point(poly[poly.Length - 1].X, poly[poly.Length - 1].Y);

            for (int i = 0; i < poly.Length; i++)
            {
                Point newPoint = new Point(poly[i].X, poly[i].Y);

                if (newPoint.X > oldPoint.X)
                {
                    p1 = oldPoint;
                    p2 = newPoint;
                }
                else
                {
                    p1 = newPoint;
                    p2 = oldPoint;
                }

                if ((newPoint.X < p.X) == (p.X <= oldPoint.X) && ((long)p.Y - (long)p1.Y) * (long)(p2.X - p1.X) < ((long)p2.Y - (long)p1.Y) * (long)(p.X - p1.X))
                {
                    inside = !inside;
                }

                oldPoint = newPoint;

            }
            return inside;
        }
        private void SaveInKMZFileStream(FileInfo fi, FileInfo fiKML, StringBuilder sbKML)
        {
            string NotUsed = "";
            StreamWriter sr = fiKML.CreateText();
            sr.Write(sbKML);
            sr.Flush();
            sr.Close();

            ProcessStartInfo pZip = new ProcessStartInfo();
            pZip.Arguments = "a -tzip \"" + fi.FullName + "\" \"" + fiKML.FullName + "\"";
            pZip.RedirectStandardInput = true;
            pZip.UseShellExecute = false;
            pZip.CreateNoWindow = true;
            pZip.WindowStyle = ProcessWindowStyle.Hidden;

            Process processZip = new Process();
            processZip.StartInfo = pZip;
            try
            {
                pZip.FileName = @"C:\Program Files\7-Zip\7z.exe";
                processZip.Start();
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CompressKMLDidNotWorkWith7zError_, ex.Message);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CompressKMLDidNotWorkWith7zError_", ex.Message);
                return;
            }

            while (!processZip.HasExited)
            {
                // waiting for the processZip to finish then continue
            }

            fiKML.Delete();

            fi = new FileInfo(fi.FullName);

            string DirText = (fi.Directory + @"\");
            TVFileModel tvFileModel = _TVFileService.GetTVFileModelWithServerFilePathAndServerFileNameDB(DirText.Replace(@"C:\", @"E:\"), fi.Name);
            if (!string.IsNullOrWhiteSpace(tvFileModel.Error))
            {
                TVItemModel tvItemModelMikeScenario = _TVItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
                if (!string.IsNullOrWhiteSpace(tvItemModelMikeScenario.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                    return;
                }

                TVItemModel tvItemFileModel = _TVItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelMikeScenario.TVItemID, fi.Name, TVTypeEnum.File);
                if (!string.IsNullOrWhiteSpace(tvItemFileModel.Error))
                {
                    tvItemFileModel = _TVItemService.PostAddChildTVItemDB(tvItemModelMikeScenario.TVItemID, fi.Name, TVTypeEnum.File);
                    if (!string.IsNullOrEmpty(tvItemFileModel.Error))
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_For_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.MIKEScenarioDocumentation, TaskRunnerServiceRes.TVText, fi.Name);
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat4List("CouldNotCreate_For_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.MIKEScenarioDocumentation, TaskRunnerServiceRes.TVText, fi.Name);
                        return;
                    }
                }

                TVFileModel tvFileModelNew = new TVFileModel();
                tvFileModelNew.TVFileTVItemID = tvItemFileModel.TVItemID;
                tvFileModelNew.FilePurpose = FilePurposeEnum.MikeResultKMZ;
                tvFileModelNew.Language = _TaskRunnerBaseService._BWObj.appTaskModel.Language;
                tvFileModelNew.FileDescription = null;
                tvFileModelNew.FileType = _TVFileService.GetFileType(fi.Extension);
                tvFileModelNew.FileSize_kb = (int)(fi.Length / 1024);
                tvFileModelNew.FileInfo = "Mike Scenario Documentation";
                tvFileModelNew.FileCreatedDate_UTC = DateTime.UtcNow;
                tvFileModelNew.ServerFileName = fi.Name;
                tvFileModelNew.ServerFilePath = (fi.Directory + @"\").Replace(@"C:\", @"E:\");

                TVFileModel tvFileModelRet = _TVFileService.PostAddTVFileDB(tvFileModelNew);
                if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_For_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.MIKEScenarioDocumentation, TaskRunnerServiceRes.TVText, fi.Name);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat4List("CouldNotCreate_For_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.MIKEScenarioDocumentation, TaskRunnerServiceRes.TVText, fi.Name);
                    return;
                }
            }
            else
            {
                tvFileModel.FileSize_kb = (int)(fi.Length / 1024);
                tvFileModel.FileCreatedDate_UTC = DateTime.UtcNow;
                tvFileModel.LastUpdateDate_UTC = DateTime.UtcNow;
                tvFileModel.LastUpdateContactTVItemID = _TaskRunnerBaseService._BWObj.appTaskModel.LastUpdateContactTVItemID;

                TVFileModel tvFileModelRet = _TVFileService.PostUpdateTVFileDB(tvFileModel);
                if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_For_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.MIKEScenarioDocumentation, TaskRunnerServiceRes.TVText, fi.Name);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat4List("CouldNotCreate_For_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.MIKEScenarioDocumentation, TaskRunnerServiceRes.TVText, fi.Name);
                    return;
                }
            }

        }
        private void WriteKMLBottom(StringBuilder sbKML)
        {
            sbKML.AppendLine(@"</Document>");
            sbKML.AppendLine(@"</kml>");
        }
        private void WriteKMLBoundaryConditionNode(int MikeScenarioID, StringBuilder sbStyleBoundaryCondition, StringBuilder sbKMLBoundaryCondition)
        {
            string NotUsed = "";
            string[] Colors = { "ylw", "grn", "blue", "ltblu", "pink", "red" };

            foreach (string color in Colors)
            {
                sbKMLBoundaryCondition.AppendLine(string.Format(@"	<Style id=""sn_{0}-pushpin"">", color));
                sbKMLBoundaryCondition.AppendLine(@"		<IconStyle>");
                sbKMLBoundaryCondition.AppendLine(@"			<scale>1.1</scale>");
                sbKMLBoundaryCondition.AppendLine(@"			<Icon>");
                sbKMLBoundaryCondition.AppendLine(string.Format(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/{0}-pushpin.png</href>", color));
                sbKMLBoundaryCondition.AppendLine(@"			</Icon>");
                sbKMLBoundaryCondition.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
                sbKMLBoundaryCondition.AppendLine(@"		</IconStyle>");
                sbKMLBoundaryCondition.AppendLine(@"		<ListStyle>");
                sbKMLBoundaryCondition.AppendLine(@"		</ListStyle>");
                sbKMLBoundaryCondition.AppendLine(@"	</Style>");
                sbKMLBoundaryCondition.AppendLine(string.Format(@"	<StyleMap id=""msn_{0}-pushpin"">", color));
                sbKMLBoundaryCondition.AppendLine(@"		<Pair>");
                sbKMLBoundaryCondition.AppendLine(@"			<key>normal</key>");
                sbKMLBoundaryCondition.AppendLine(string.Format(@"			<styleUrl>#sn_{0}-pushpin</styleUrl>", color));
                sbKMLBoundaryCondition.AppendLine(@"		</Pair>");
                sbKMLBoundaryCondition.AppendLine(@"		<Pair>");
                sbKMLBoundaryCondition.AppendLine(@"			<key>highlight</key>");
                sbKMLBoundaryCondition.AppendLine(string.Format(@"			<styleUrl>#sh_{0}-pushpin</styleUrl>", color));
                sbKMLBoundaryCondition.AppendLine(@"		</Pair>");
                sbKMLBoundaryCondition.AppendLine(@"	</StyleMap>");
                sbKMLBoundaryCondition.AppendLine(string.Format(@"	<Style id=""sh_{0}-pushpin"">", color));
                sbKMLBoundaryCondition.AppendLine(@"		<IconStyle>");
                sbKMLBoundaryCondition.AppendLine(@"			<scale>1.3</scale>");
                sbKMLBoundaryCondition.AppendLine(@"			<Icon>");
                sbKMLBoundaryCondition.AppendLine(string.Format(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/{0}-pushpin.png</href>", color));
                sbKMLBoundaryCondition.AppendLine(@"			</Icon>");
                sbKMLBoundaryCondition.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
                sbKMLBoundaryCondition.AppendLine(@"		</IconStyle>");
                sbKMLBoundaryCondition.AppendLine(@"		<ListStyle>");
                sbKMLBoundaryCondition.AppendLine(@"		</ListStyle>");
                sbKMLBoundaryCondition.AppendLine(@"	</Style>");
            }

            //UpdateTask(AppTaskID, "30 %");

            sbKMLBoundaryCondition.AppendLine(@"<Folder>");
            sbKMLBoundaryCondition.AppendLine(@"<name>" + TaskRunnerServiceRes.MikeBoundaryConditionName + @"</name>");
            sbKMLBoundaryCondition.AppendLine(@"<visibility>1</visibility>");

            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemModel tvItemModelMikeScenario = tvItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelMikeScenario.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                return;
            }

            MikeBoundaryConditionService mikeBoundaryConditionService = new MikeBoundaryConditionService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            List<MikeBoundaryConditionModel> mbcModelList = mikeBoundaryConditionService.GetMikeBoundaryConditionModelListWithMikeScenarioTVItemIDAndTVTypeDB(tvItemModelMikeScenario.TVItemID, TVTypeEnum.MikeBoundaryConditionMesh);

            int countColor = 0;
            foreach (MikeBoundaryConditionModel mbcm in mbcModelList)
            {
                sbKMLBoundaryCondition.AppendLine(@"<Folder>");
                sbKMLBoundaryCondition.AppendLine(@"<name>" + mbcm.MikeBoundaryConditionName + "</name>");
                sbKMLBoundaryCondition.AppendLine(@"<visibility>1</visibility>");
                sbKMLBoundaryCondition.AppendLine(@"<description><![CDATA[");
                sbKMLBoundaryCondition.AppendLine(@"<ul>");
                sbKMLBoundaryCondition.AppendLine(@"<li><b>" + TaskRunnerServiceRes.MikeBoundaryConditionCode + @"</b>: " + mbcm.MikeBoundaryConditionCode + "</li>");
                sbKMLBoundaryCondition.AppendLine(@"<li><b>" + TaskRunnerServiceRes.MikeBoundaryConditionName + @"</b>: " + mbcm.MikeBoundaryConditionName + "</li>");
                sbKMLBoundaryCondition.AppendLine(@"<li><b>" + TaskRunnerServiceRes.MikeBoundaryConditionLength + @"</b>: " + mbcm.MikeBoundaryConditionLength_m + "</li>");
                sbKMLBoundaryCondition.AppendLine(@"<li><b>" + TaskRunnerServiceRes.MikeBoundaryConditionFormat + @"</b>: " + mbcm.MikeBoundaryConditionFormat + "</li>");
                sbKMLBoundaryCondition.AppendLine(@"<li><b>" + TaskRunnerServiceRes.MikeBoundaryConditionLevelOrVelocity + @"</b>: " + mbcm.MikeBoundaryConditionLevelOrVelocity + "</li>");
                sbKMLBoundaryCondition.AppendLine(@"<li><b>" + TaskRunnerServiceRes.WebTideDataSet + @"</b>: " + mbcm.WebTideDataSet.ToString() + "</li>");
                sbKMLBoundaryCondition.AppendLine(@"<li><b>" + TaskRunnerServiceRes.NumberOfWebTideNodes + @"</b>: " + mbcm.NumberOfWebTideNodes + "</li>");
                sbKMLBoundaryCondition.AppendLine(@"</ul>");
                sbKMLBoundaryCondition.AppendLine(@"]]></description>");

                // drawing Boundary Nodes
                MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                List<MapInfoPointModel> mapInfoPointModelList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(mbcm.MikeBoundaryConditionTVItemID, TVTypeEnum.MikeBoundaryConditionMesh, MapInfoDrawTypeEnum.Polyline);

                sbKMLBoundaryCondition.AppendLine(@"<Folder>");
                sbKMLBoundaryCondition.AppendLine(@"<name>" + TaskRunnerServiceRes.MikeBoundaryElementNodes + @"</name>");
                sbKMLBoundaryCondition.AppendLine(@"<open>1</open>");
                foreach (MapInfoPointModel mapInfoPointModel in mapInfoPointModelList)
                {
                    sbKMLBoundaryCondition.AppendLine(@"<Placemark>");
                    sbKMLBoundaryCondition.AppendLine(@"<name>Node " + mapInfoPointModel.Ordinal + "</name>");
                    sbKMLBoundaryCondition.AppendLine(string.Format(@"<styleUrl>#msn_{0}-pushpin</styleUrl>", Colors[countColor]));
                    sbKMLBoundaryCondition.AppendLine(@"<Point>");
                    sbKMLBoundaryCondition.AppendLine(@"<coordinates>" + mapInfoPointModel.Lng.ToString().Replace(",", ".") + @"," + mapInfoPointModel.Lat.ToString().Replace(",", ".") + @",0</coordinates>");
                    sbKMLBoundaryCondition.AppendLine(@"</Point>");
                    sbKMLBoundaryCondition.AppendLine(@"</Placemark>");
                }
                sbKMLBoundaryCondition.AppendLine(@"</Folder>");

                sbKMLBoundaryCondition.AppendLine(@"</Folder>");
                countColor += 1;
            }

            mbcModelList = mikeBoundaryConditionService.GetMikeBoundaryConditionModelListWithMikeScenarioTVItemIDAndTVTypeDB(tvItemModelMikeScenario.TVItemID, TVTypeEnum.MikeBoundaryConditionWebTide);

            countColor = 0;
            foreach (MikeBoundaryConditionModel mbcm in mbcModelList)
            {
                sbKMLBoundaryCondition.AppendLine(@"<Folder>");
                sbKMLBoundaryCondition.AppendLine(@"<name>" + mbcm.MikeBoundaryConditionName + "</name>");
                sbKMLBoundaryCondition.AppendLine(@"<visibility>1</visibility>");
                sbKMLBoundaryCondition.AppendLine(@"<description><![CDATA[");
                sbKMLBoundaryCondition.AppendLine(@"<ul>");
                sbKMLBoundaryCondition.AppendLine(@"<li><b>" + TaskRunnerServiceRes.MikeBoundaryConditionCode + @"</b>: " + mbcm.MikeBoundaryConditionCode + "</li>");
                sbKMLBoundaryCondition.AppendLine(@"<li><b>" + TaskRunnerServiceRes.MikeBoundaryConditionName + @"</b>: " + mbcm.MikeBoundaryConditionName + "</li>");
                sbKMLBoundaryCondition.AppendLine(@"<li><b>" + TaskRunnerServiceRes.MikeBoundaryConditionLength + @"</b>: " + mbcm.MikeBoundaryConditionLength_m + "</li>");
                sbKMLBoundaryCondition.AppendLine(@"<li><b>" + TaskRunnerServiceRes.MikeBoundaryConditionFormat + @"</b>: " + mbcm.MikeBoundaryConditionFormat + "</li>");
                sbKMLBoundaryCondition.AppendLine(@"<li><b>" + TaskRunnerServiceRes.MikeBoundaryConditionLevelOrVelocity + @"</b>: " + mbcm.MikeBoundaryConditionLevelOrVelocity + "</li>");
                sbKMLBoundaryCondition.AppendLine(@"<li><b>" + TaskRunnerServiceRes.WebTideDataSet + @"</b>: " + mbcm.WebTideDataSet + "</li>");
                sbKMLBoundaryCondition.AppendLine(@"<li><b>" + TaskRunnerServiceRes.NumberOfWebTideNodes + @"</b>: " + mbcm.NumberOfWebTideNodes + "</li>");
                sbKMLBoundaryCondition.AppendLine(@"</ul>");
                sbKMLBoundaryCondition.AppendLine(@"]]></description>");

                // drawing Boundary Nodes
                MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                List<MapInfoPointModel> mapInfoPointModelList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(mbcm.MikeBoundaryConditionTVItemID, TVTypeEnum.MikeBoundaryConditionWebTide, MapInfoDrawTypeEnum.Polyline);

                sbKMLBoundaryCondition.AppendLine(@"<Folder>");
                sbKMLBoundaryCondition.AppendLine(@"<name>" + TaskRunnerServiceRes.MikeBoundaryElementNodes + @"</name>");
                sbKMLBoundaryCondition.AppendLine(@"<open>1</open>");
                foreach (MapInfoPointModel mapInfoPointModel in mapInfoPointModelList)
                {
                    sbKMLBoundaryCondition.AppendLine(@"<Placemark>");
                    sbKMLBoundaryCondition.AppendLine(@"<name>Node " + mapInfoPointModel.Ordinal + "</name>");
                    sbKMLBoundaryCondition.AppendLine(string.Format(@"<styleUrl>#msn_{0}-pushpin</styleUrl>", Colors[countColor]));
                    sbKMLBoundaryCondition.AppendLine(@"<Point>");
                    sbKMLBoundaryCondition.AppendLine(@"<coordinates>" + mapInfoPointModel.Lng.ToString().Replace(",", ".") + @"," + mapInfoPointModel.Lat.ToString().Replace(",", ".") + @",0</coordinates>");
                    sbKMLBoundaryCondition.AppendLine(@"</Point>");
                    sbKMLBoundaryCondition.AppendLine(@"</Placemark>");
                }
                sbKMLBoundaryCondition.AppendLine(@"</Folder>");

                sbKMLBoundaryCondition.AppendLine(@"</Folder>");
                countColor += 1;
            }


            sbKMLBoundaryCondition.AppendLine(@"</Folder>");
        }
        private void WriteKMLCurrentsAnimation(DfsuFile dfsuFile, string dfsParamItem, StringBuilder sbStyleCurrentAnim, StringBuilder sbKMLCurrentAnim, List<float> ContourValueList, List<int> SigmaLayerValueList, List<int> ZLayerValueList, List<float> DepthValueList, List<ElementLayer> SelectedElementLayerList, double VectorSizeInMeterForEach10cm_s)
        {
            string NotUsed = "";

            int ItemUVelocity = 0;
            int ItemVVelocity = 0;

            // getting the ItemNumber
            foreach (IDfsSimpleDynamicItemInfo dfsDyInfo in dfsuFile.ItemInfo)
            {
                if (dfsDyInfo.Quantity.Item == eumItem.eumIuVelocity)
                {
                    ItemUVelocity = dfsDyInfo.ItemNumber;
                }
                if (dfsDyInfo.Quantity.Item == eumItem.eumIvVelocity)
                {
                    ItemVVelocity = dfsDyInfo.ItemNumber;
                }
            }

            if (ItemUVelocity == 0 || ItemVVelocity == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, TaskRunnerServiceRes.Parameters);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", TaskRunnerServiceRes.Parameters);
                return;
            }

            DrawKMLCurrentsStyle(sbStyleCurrentAnim);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            //int pcount = 0;
            sbKMLCurrentAnim.AppendLine(@"<Folder><name>" + TaskRunnerServiceRes.CurrentsAnim + @"</name>");
            sbKMLCurrentAnim.AppendLine(@"<visibility>0</visibility>");

            foreach (int Layer in SigmaLayerValueList)
            {
                sbKMLCurrentAnim.AppendLine(string.Format(@"<Folder><name>" + TaskRunnerServiceRes.Layer + " {0}</name>", Layer));
                sbKMLCurrentAnim.AppendLine(@"<visibility>0</visibility>");

                Dictionary<int, Node> ElemCenter = new Dictionary<int, Node>();

                foreach (ElementLayer el in SelectedElementLayerList.Where(c => c.Layer == Layer))
                {
                    float XCenter = 0.0f;
                    float YCenter = 0.0f;

                    foreach (Node n in el.Element.NodeList)
                    {
                        XCenter += n.X;
                        YCenter += n.Y;
                    }
                    XCenter = XCenter / el.Element.NodeList.Count();
                    YCenter = YCenter / el.Element.NodeList.Count();

                    ElemCenter.Add(el.Element.ID, new Node() { X = XCenter, Y = YCenter });
                }

                int vCount = 0;
                for (int timeStep = 0; timeStep < dfsuFile.NumberOfTimeSteps; timeStep++)
                {
                    sbKMLCurrentAnim.AppendLine(@"<Folder>");
                    sbKMLCurrentAnim.AppendLine(string.Format(@"<name>{0:yyyy-MM-dd} {0:HH:mm:ss tt}</name>", dfsuFile.StartDateTime.AddSeconds(vCount * dfsuFile.TimeStepInSeconds)));
                    sbKMLCurrentAnim.AppendLine(@"<visibility>0</visibility>");
                    sbKMLCurrentAnim.AppendLine(@"<TimeSpan>");
                    sbKMLCurrentAnim.AppendLine(string.Format(@"<begin>{0:yyyy-MM-dd}T{0:HH:mm:ss}</begin>", dfsuFile.StartDateTime.AddSeconds(vCount * dfsuFile.TimeStepInSeconds)));
                    sbKMLCurrentAnim.AppendLine(string.Format(@"<end>{0:yyyy-MM-dd}T{0:HH:mm:ss}</end>", dfsuFile.StartDateTime.AddSeconds((vCount + 1) * dfsuFile.TimeStepInSeconds)));
                    sbKMLCurrentAnim.AppendLine(@"</TimeSpan>");

                    float[] UvelocityList = (float[])dfsuFile.ReadItemTimeStep(ItemUVelocity, timeStep).Data;
                    float[] VvelocityList = (float[])dfsuFile.ReadItemTimeStep(ItemVVelocity, timeStep).Data;

                    MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                    foreach (ElementLayer el in SelectedElementLayerList.Where(c => c.Layer == Layer))
                    {
                        float UV = UvelocityList[el.Element.ID - 1];
                        float VV = VvelocityList[el.Element.ID - 1];

                        double VectorVal = Math.Sqrt((UV * UV) + (VV * VV));
                        double VectorDir = 0.0D;
                        double VectorDirCartesian = Math.Acos(Math.Abs(UV / VectorVal)) * mapInfoService.r2d;

                        if (VectorDirCartesian <= 360 && VectorDirCartesian >= 0)
                        {
                            // everything is ok
                        }
                        else
                        {
                            VectorDirCartesian = 0.0D;
                        }

                        if (UV >= 0 && VV >= 0)
                        {
                            VectorDir = 90 - VectorDirCartesian;
                        }
                        else if (UV < 0 && VV >= 0)
                        {
                            VectorDir = 270 + VectorDirCartesian;
                            VectorDirCartesian = 180 - VectorDirCartesian;
                        }
                        else if (UV >= 0 && VV < 0)
                        {
                            VectorDir = 90 + VectorDirCartesian;
                            VectorDirCartesian = 360 - VectorDirCartesian;
                        }
                        else if (UV < 0 && VV < 0)
                        {
                            VectorDir = 270 - VectorDirCartesian;
                            VectorDirCartesian = 180 + VectorDirCartesian;
                        }

                        if (VectorVal > 0)
                        {
                            sbKMLCurrentAnim.AppendLine(@"<Placemark>");
                            sbKMLCurrentAnim.AppendLine(@"<visibility>0</visibility>");
                            sbKMLCurrentAnim.AppendLine(string.Format(@"<name>{0:F4} m/s " + TaskRunnerServiceRes.AtLowerCase + " {1:F0}°</name>", VectorVal, VectorDirCartesian).ToString().Replace(",", "."));

                            if (Layer == 1)
                            {
                                sbKMLCurrentAnim.AppendLine(@"<styleUrl>#pink</styleUrl>");
                            }
                            else if (Layer == 2)
                            {
                                sbKMLCurrentAnim.AppendLine(@"<styleUrl>#yellow</styleUrl>");
                            }
                            else if (Layer == 3)
                            {
                                sbKMLCurrentAnim.AppendLine(@"<styleUrl>#green</styleUrl>");
                            }
                            else
                            {
                                // nothing ... It will be white by default
                            }

                            sbKMLCurrentAnim.AppendLine(@"<LineString>");
                            sbKMLCurrentAnim.AppendLine(@"<tessellate>1</tessellate>");
                            sbKMLCurrentAnim.AppendLine(@"<coordinates>");

                            sbKMLCurrentAnim.Append(((Node)ElemCenter[el.Element.ID]).X.ToString().Replace(",", ".") + @"," + ((Node)ElemCenter[el.Element.ID]).Y.ToString().Replace(",", ".") + ",0 ");

                            Node node = new Node();
                            double Fact = 0.00012;

                            double HypothDist = (VectorVal * VectorSizeInMeterForEach10cm_s * Fact);
                            node.X = (float)(ElemCenter[el.Element.ID].X + (HypothDist * Math.Cos(VectorDirCartesian * mapInfoService.d2r)));
                            node.Y = (float)(ElemCenter[el.Element.ID].Y + (HypothDist * Math.Sin(VectorDirCartesian * mapInfoService.d2r)));

                            sbKMLCurrentAnim.Append(node.X.ToString().Replace(",", ".") + @"," + node.Y.ToString().Replace(",", ".") + ",0 ");

                            Node node2 = new Node();

                            node2.X = (float)(node.X + (HypothDist * 0.1 * Math.Cos((VectorDirCartesian + 180 - 25) * mapInfoService.d2r)));
                            node2.Y = (float)(node.Y + (HypothDist * 0.1 * Math.Sin((VectorDirCartesian + 180 - 25) * mapInfoService.d2r)));

                            sbKMLCurrentAnim.Append(node2.X.ToString().Replace(",", ".") + @"," + node2.Y.ToString().Replace(",", ".") + ",0 ");
                            sbKMLCurrentAnim.Append(node.X.ToString().Replace(",", ".") + @"," + node.Y.ToString().Replace(",", ".") + ",0 ");

                            node2.X = (float)(node.X + (HypothDist * 0.1 * Math.Cos((VectorDirCartesian + 180 + 25) * mapInfoService.d2r)));
                            node2.Y = (float)(node.Y + (HypothDist * 0.1 * Math.Sin((VectorDirCartesian + 180 + 25) * mapInfoService.d2r)));
                            sbKMLCurrentAnim.Append(node2.X.ToString().Replace(",", ".") + @"," + node2.Y.ToString().Replace(",", ".") + ",0 ");

                            sbKMLCurrentAnim.AppendLine(@"</coordinates>");
                            sbKMLCurrentAnim.AppendLine(@"</LineString>");
                            sbKMLCurrentAnim.AppendLine(@"</Placemark>");
                        }

                    }
                    sbKMLCurrentAnim.AppendLine(@"</Folder>");
                    vCount += 1;
                }
                sbKMLCurrentAnim.AppendLine(@"</Folder>");

            }
            sbKMLCurrentAnim.AppendLine(@"</Folder>");
        }
        private void WriteKMLFeacalColiformContourLine(DfsuFile dfsuFile, string dfsParamItem, StringBuilder sbStyleFeacalColiformContour, StringBuilder sbPlacemarkFeacalColiformContour, List<float> ContourValueList, List<int> SigmaLayerValueList, List<int> ZLayerValueList, List<float> DepthValueList, List<ElementLayer> ElementLayerList, List<NodeLayer> TopNodeLayerList, List<NodeLayer> BottomNodeLayerList)
        {
            string NotUsed = "";

            int ItemNumber = 0;
            double RefreshEveryXSeconds = double.Parse("5");
            DateTime RefreshDateTime = DateTime.Now.AddSeconds(RefreshEveryXSeconds);

            // getting the ItemNumber
            foreach (IDfsSimpleDynamicItemInfo dfsDyInfo in dfsuFile.ItemInfo)
            {
                if (dfsDyInfo.Quantity.Item == eumItem.eumIConcentration)
                {
                    ItemNumber = dfsDyInfo.ItemNumber;
                    break;
                }
            }

            if (ItemNumber == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind__, TaskRunnerServiceRes.ParameterType, dfsParamItem);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind__", TaskRunnerServiceRes.ParameterType, dfsParamItem);
                return;
            }

            DrawKMLContourStyle(sbStyleFeacalColiformContour, sbPlacemarkFeacalColiformContour);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            //            int pcount = 0;
            sbPlacemarkFeacalColiformContour.AppendLine(@"<Folder><name>" + TaskRunnerServiceRes.PollutionVideo + "</name>");
            sbPlacemarkFeacalColiformContour.AppendLine(@"<visibility>0</visibility>");

            int CountAt = 0;
            int TotalStep = SigmaLayerValueList.Count * ContourValueList.Count * dfsuFile.NumberOfTimeSteps;
            int CurrentLayer = 0;
            int CurrentContourValue = 0;
            int CurrentTimeSteps = 0;
            foreach (int Layer in SigmaLayerValueList)
            {
                CurrentLayer += 1;
                CurrentContourValue = 1;
                CurrentTimeSteps = 1;

                #region Top of Layer
                sbPlacemarkFeacalColiformContour.AppendLine(string.Format(@"<Folder><name>" + TaskRunnerServiceRes.TopOfLayer + @" [{0}]</name>", Layer));
                sbPlacemarkFeacalColiformContour.AppendLine(@"<visibility>0</visibility>");
                int CountContourValue = 1;
                foreach (float ContourValue in ContourValueList)
                {
                    sbPlacemarkFeacalColiformContour.AppendLine(string.Format(@"<Folder><name>" + TaskRunnerServiceRes.ContourValue + @" [{0}]</name>", ContourValue));
                    sbPlacemarkFeacalColiformContour.AppendLine(@"<visibility>0</visibility>");

                    int vcount = 0;
                    //for (int timeStep = 30; timeStep < 35 /*dfsuFile.NumberOfTimeSteps */; timeStep++)
                    for (int timeStep = 0; timeStep < dfsuFile.NumberOfTimeSteps; timeStep++)
                    {
                        CountAt += 1;
                        if (DateTime.Now > RefreshDateTime)
                        {
                            float Perc = (((float)CurrentLayer * (float)CurrentContourValue * (float)CurrentTimeSteps) * (float)100.0f) / (float)TotalStep;

                            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)Perc);
                            return;
                        }

                        float[] ValueList = (float[])dfsuFile.ReadItemTimeStep(ItemNumber, timeStep).Data;

                        List<ContourPolygon> ContourPolygonList = new List<ContourPolygon>();

                        for (int i = 0; i < ElementLayerList.Count; i++)
                        {
                            ElementLayerList[i].Element.Value = ValueList[i];
                        }

                        foreach (NodeLayer nl in TopNodeLayerList)
                        {
                            float Total = 0;
                            foreach (Element element in nl.Node.ElementList)
                            {
                                Total += element.Value;
                            }
                            nl.Node.Value = Total / nl.Node.ElementList.Count;
                        }


                        List<Node> AllNodeList = new List<Node>();

                        List<NodeLayer> AboveNodeLayerList = new List<NodeLayer>();

                        AboveNodeLayerList = (from n in TopNodeLayerList
                                              where (n.Node.Value >= ContourValue)
                                              && n.Layer == Layer
                                              select n).ToList<NodeLayer>();

                        foreach (NodeLayer snl in AboveNodeLayerList)
                        {
                            List<NodeLayer> EndNodeLayerList = null;

                            List<NodeLayer> NodeLayerConnectedList = (from nll in TopNodeLayerList
                                                                      from n in snl.Node.ConnectNodeList
                                                                      where (n.ID == nll.Node.ID)
                                                                      select nll).ToList<NodeLayer>();

                            EndNodeLayerList = (from nll in NodeLayerConnectedList
                                                where (nll.Node.ID != snl.Node.ID)
                                                && (nll.Node.Value < ContourValue)
                                                && nll.Layer == Layer
                                                select nll).ToList<NodeLayer>();

                            foreach (NodeLayer en in EndNodeLayerList)
                            {
                                AllNodeList.Add(en.Node);
                            }

                            if (snl.Node.Code != 0)
                            {
                                AllNodeList.Add(snl.Node);
                            }

                        }

                        //if (AllNodeList.Count == 0)
                        //{
                        //    //vcount += 1;
                        //    continue;
                        //}

                        List<Element> TempUniqueElementList = new List<Element>();
                        List<Element> UniqueElementList = new List<Element>();
                        foreach (ElementLayer el in ElementLayerList.Where(l => l.Layer == Layer))
                        {
                            if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
                            {
                                if (el.Element.Type == 32)
                                {
                                    bool NodeBigger = false;
                                    for (int i = 3; i < 6; i++)
                                    {
                                        if (el.Element.NodeList[i].Value >= ContourValue)
                                        {
                                            NodeBigger = true;
                                            break;
                                        }
                                    }
                                    if (NodeBigger)
                                    {
                                        int countTrue = 0;
                                        for (int i = 3; i < 6; i++)
                                        {
                                            if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
                                            {
                                                countTrue += 1;
                                            }
                                        }
                                        if (countTrue != el.Element.NodeList.Count)
                                        {
                                            TempUniqueElementList.Add(el.Element);
                                        }
                                    }
                                }
                                else if (el.Element.Type == 33)
                                {
                                    bool NodeBigger = false;
                                    for (int i = 4; i < 8; i++)
                                    {
                                        if (el.Element.NodeList[i].Value >= ContourValue)
                                        {
                                            NodeBigger = true;
                                            break;
                                        }
                                    }
                                    if (NodeBigger)
                                    {
                                        int countTrue = 0;
                                        for (int i = 4; i < 8; i++)
                                        {
                                            if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
                                            {
                                                countTrue += 1;
                                            }
                                        }
                                        if (countTrue != el.Element.NodeList.Count)
                                        {
                                            TempUniqueElementList.Add(el.Element);
                                        }
                                    }
                                }
                                else
                                {
                                    NotUsed = string.Format(TaskRunnerServiceRes.ElementType_IsNotSupported, el.Element.Type.ToString());
                                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ElementType_IsNotSupported", el.Element.Type.ToString());
                                    return;
                                }
                            }
                            else if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu2D)
                            {
                                bool NodeBigger = false;
                                for (int i = 0; i < el.Element.NodeList.Count; i++)
                                {
                                    if (el.Element.NodeList[i].Value >= ContourValue)
                                    {
                                        NodeBigger = true;
                                        break;
                                    }
                                }
                                if (NodeBigger)
                                {
                                    int countTrue = 0;
                                    for (int i = 0; i < el.Element.NodeList.Count; i++)
                                    {
                                        if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
                                        {
                                            countTrue += 1;
                                        }
                                    }
                                    if (countTrue != el.Element.NodeList.Count)
                                    {
                                        TempUniqueElementList.Add(el.Element);
                                    }
                                }
                            }
                        }

                        UniqueElementList = (from el in TempUniqueElementList select el).Distinct().ToList<Element>();

                        // filling InterpolatedContourNodeList
                        InterpolatedContourNodeList = new List<Node>();

                        foreach (Element el in UniqueElementList)
                        {
                            if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
                            {
                                if (el.Type == 32)
                                {
                                    if (el.NodeList[3].Value >= ContourValue && el.NodeList[4].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[3], el.NodeList[4], ContourValue);
                                    }
                                    if (el.NodeList[3].Value >= ContourValue && el.NodeList[5].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[3], el.NodeList[5], ContourValue);
                                    }
                                    if (el.NodeList[4].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[4], el.NodeList[3], ContourValue);
                                    }
                                    if (el.NodeList[4].Value >= ContourValue && el.NodeList[5].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[4], el.NodeList[5], ContourValue);
                                    }
                                    if (el.NodeList[5].Value >= ContourValue && el.NodeList[4].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[5], el.NodeList[4], ContourValue);
                                    }
                                    if (el.NodeList[5].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[5], el.NodeList[3], ContourValue);
                                    }
                                }
                                else if (el.Type == 33)
                                {
                                    if (el.NodeList[4].Value >= ContourValue && el.NodeList[5].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[4], el.NodeList[5], ContourValue);
                                    }
                                    if (el.NodeList[4].Value >= ContourValue && el.NodeList[7].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[4], el.NodeList[7], ContourValue);
                                    }
                                    if (el.NodeList[5].Value >= ContourValue && el.NodeList[4].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[5], el.NodeList[4], ContourValue);
                                    }
                                    if (el.NodeList[5].Value >= ContourValue && el.NodeList[6].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[5], el.NodeList[6], ContourValue);
                                    }
                                    if (el.NodeList[6].Value >= ContourValue && el.NodeList[5].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[6], el.NodeList[5], ContourValue);
                                    }
                                    if (el.NodeList[6].Value >= ContourValue && el.NodeList[7].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[6], el.NodeList[7], ContourValue);
                                    }
                                    if (el.NodeList[7].Value >= ContourValue && el.NodeList[4].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[7], el.NodeList[4], ContourValue);
                                    }
                                    if (el.NodeList[7].Value >= ContourValue && el.NodeList[6].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[7], el.NodeList[6], ContourValue);
                                    }
                                }
                                else
                                {
                                    NotUsed = string.Format(TaskRunnerServiceRes.ElementType_IsNotSupported, el.Type.ToString());
                                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ElementType_IsNotSupported", el.Type.ToString());
                                    return;
                                }
                            }
                            else if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu2D)
                            {
                                if (el.Type == 21)
                                {
                                    if (el.NodeList[0].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[0], el.NodeList[1], ContourValue);
                                    }
                                    if (el.NodeList[0].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[0], el.NodeList[2], ContourValue);
                                    }
                                    if (el.NodeList[1].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[1], el.NodeList[0], ContourValue);
                                    }
                                    if (el.NodeList[1].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[1], el.NodeList[2], ContourValue);
                                    }
                                    if (el.NodeList[2].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[2], el.NodeList[1], ContourValue);
                                    }
                                    if (el.NodeList[2].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[2], el.NodeList[0], ContourValue);
                                    }
                                }
                                else if (el.Type == 24)
                                {
                                }
                                else if (el.Type == 25)
                                {
                                    if (el.NodeList[0].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[0], el.NodeList[1], ContourValue);
                                    }
                                    if (el.NodeList[0].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[0], el.NodeList[3], ContourValue);
                                    }
                                    if (el.NodeList[1].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[1], el.NodeList[0], ContourValue);
                                    }
                                    if (el.NodeList[1].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[1], el.NodeList[2], ContourValue);
                                    }
                                    if (el.NodeList[2].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[2], el.NodeList[1], ContourValue);
                                    }
                                    if (el.NodeList[2].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[2], el.NodeList[3], ContourValue);
                                    }
                                    if (el.NodeList[3].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[3], el.NodeList[0], ContourValue);
                                    }
                                    if (el.NodeList[3].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[3], el.NodeList[2], ContourValue);
                                    }
                                }
                                else
                                {
                                    NotUsed = string.Format(TaskRunnerServiceRes.ElementType_IsNotSupported, el.Type.ToString());
                                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ElementType_IsNotSupported", el.Type.ToString());
                                    return;
                                }
                            }
                        }

                        List<Node> UniqueNodeList = (from n in AllNodeList orderby n.ID select n).Distinct().ToList<Node>();

                        // ------------------------- new code --------------------------
                        //                     

                        ForwardVector = new Dictionary<string, Vector>();
                        BackwardVector = new Dictionary<string, Vector>();

                        foreach (Element el in UniqueElementList)
                        {
                            if (el.Type == 21)
                            {
                                FillVectors21_32(el, UniqueElementList, ContourValue, false, true);
                            }
                            else if (el.Type == 24)
                            {
                                NotUsed = TaskRunnerServiceRes.AllNodesAreSmallerThanContourValue;
                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("AllNodesAreSmallerThanContourValue");
                                return;
                            }
                            else if (el.Type == 25)
                            {
                                FillVectors25_33(el, UniqueElementList, ContourValue, false, true);
                            }
                            else if (el.Type == 32)
                            {
                                FillVectors21_32(el, UniqueElementList, ContourValue, true, true);
                            }
                            else if (el.Type == 33)
                            {
                                FillVectors25_33(el, UniqueElementList, ContourValue, true, true);
                            }
                            else
                            {
                                NotUsed = string.Format(TaskRunnerServiceRes.ElementType_IsNotSupported, el.Type.ToString());
                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ElementType_IsNotSupported", el.Type.ToString());
                                return;
                            }

                        }

                        //-------------- new code ------------------------



                        bool MoreContourLine = true;
                        MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        while (MoreContourLine && ForwardVector.Count > 0)
                        {
                            List<Node> FinalContourNodeList = new List<Node>();
                            Vector LastVector = new Vector();
                            LastVector = ForwardVector.First().Value;
                            FinalContourNodeList.Add(LastVector.StartNode);
                            FinalContourNodeList.Add(LastVector.EndNode);
                            ForwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
                            BackwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
                            bool PolygonCompleted = false;
                            while (!PolygonCompleted)
                            {
                                List<string> KeyStrList = (from k in ForwardVector.Keys
                                                           where k.StartsWith(LastVector.EndNode.ID.ToString() + ",")
                                                           && !k.EndsWith("," + LastVector.StartNode.ID.ToString())
                                                           select k).ToList<string>();

                                if (KeyStrList.Count != 1)
                                {
                                    KeyStrList = (from k in BackwardVector.Keys
                                                  where k.StartsWith(LastVector.EndNode.ID.ToString() + ",")
                                                  && !k.EndsWith("," + LastVector.StartNode.ID.ToString())
                                                  select k).ToList<string>();

                                    if (KeyStrList.Count != 1)
                                    {
                                        PolygonCompleted = true;
                                        break;
                                    }
                                    else
                                    {
                                        LastVector = BackwardVector[KeyStrList[0]];
                                        BackwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
                                        ForwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
                                    }
                                }
                                else
                                {
                                    LastVector = ForwardVector[KeyStrList[0]];
                                    ForwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
                                    BackwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
                                }
                                FinalContourNodeList.Add(LastVector.EndNode);
                                if (FinalContourNodeList[FinalContourNodeList.Count - 1] == FinalContourNodeList[0])
                                {
                                    PolygonCompleted = true;
                                }
                            }

                            if (mapInfoService.CalculateAreaOfPolygon(FinalContourNodeList) < 0)
                            {
                                FinalContourNodeList.Reverse();
                            }

                            FinalContourNodeList.Add(FinalContourNodeList[0]);
                            ContourPolygon contourPolygon = new ContourPolygon() { };
                            contourPolygon.ContourNodeList = FinalContourNodeList;
                            contourPolygon.ContourValue = ContourValue;
                            contourPolygon.Layer = Layer;
                            ContourPolygonList.Add(contourPolygon);

                            if (ForwardVector.Count == 0)
                            {
                                MoreContourLine = false;
                            }

                        }
                        DrawKMLContourPolygon(ContourPolygonList, dfsuFile, vcount, sbStyleFeacalColiformContour, sbPlacemarkFeacalColiformContour);
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                            return;

                        vcount += 1;
                    }
                    sbPlacemarkFeacalColiformContour.AppendLine(@"</Folder>");
                    CountContourValue += 1;
                }
                sbPlacemarkFeacalColiformContour.AppendLine(@"</Folder>");
                #endregion Top of Layer

                #region Bottom of Layer
                //// doing the bottom layer if the current layer is == NumberOfSigmaLayers
                //if (Layer == dfsuFile.NumberOfSigmaLayers)
                //{
                //    sbPlacemarkFeacalColiformContour.AppendLine(string.Format(@"<Folder><name>Bottom of Layer [{0}]</name>", Layer));
                //    sbPlacemarkFeacalColiformContour.AppendLine(@"<visibility>0</visibility>");
                //    CountContourValue = 1;
                //    foreach (float ContourValue in ContourValueList)
                //    {
                //        sbPlacemarkFeacalColiformContour.AppendLine(string.Format(@"<Folder><name>Contour Value [{0}]</name>", ContourValue));
                //        sbPlacemarkFeacalColiformContour.AppendLine(@"<visibility>0</visibility>");

                //        int vcount = 0;
                //        //for (int timeStep = 30; timeStep < 35 /*dfsuFile.NumberOfTimeSteps */; timeStep++)
                //        for (int timeStep = 0; timeStep < dfsuFile.NumberOfTimeSteps; timeStep++)
                //        {
                //            CountRefresh += 1;
                //            CountAt += 1;
                //            if (CountRefresh > UpdateAfter)
                //            {
                //                string AppTaskStatus = "";
                //                if (SigmaLayerValueList.Contains(dfsuFile.NumberOfSigmaLayers))
                //                {
                //                    AppTaskStatus = ((int)((CountAt * 100) / (dfsuFile.NumberOfTimeSteps * (SigmaLayerValueList.Count + 1) * ContourValueList.Count))).ToString() + " %";
                //                }
                //                else
                //                {
                //                    AppTaskStatus = ((int)((CountAt * 100) / (dfsuFile.NumberOfTimeSteps * SigmaLayerValueList.Count * ContourValueList.Count))).ToString() + " %";
                //                }
                //                UpdateTask(AppTaskID, AppTaskStatus);
                //                CountRefresh = 0;
                //            }

                //            float[] ValueList = (float[])dfsuFile.ReadItemTimeStep(ItemNumber, timeStep).Data;

                //            List<ContourPolygon> ContourPolygonList = new List<ContourPolygon>();

                //            for (int i = 0; i < ElementLayerList.Count; i++)
                //            {
                //                ElementLayerList[i].Element.Value = ValueList[i];
                //            }

                //            foreach (NodeLayer nl in BottomNodeLayerList)
                //            {
                //                float Total = 0;
                //                foreach (Element element in nl.Node.ElementList)
                //                {
                //                    Total += element.Value;
                //                }
                //                nl.Node.Value = Total / nl.Node.ElementList.Count;
                //            }


                //            List<Node> AllNodeList = new List<Node>();

                //            List<NodeLayer> AboveNodeLayerList = new List<NodeLayer>();

                //            AboveNodeLayerList = (from n in BottomNodeLayerList
                //                                  where (n.Node.Value >= ContourValue)
                //                                  && n.Layer == Layer
                //                                  select n).ToList<NodeLayer>();

                //            foreach (NodeLayer snl in AboveNodeLayerList)
                //            {
                //                List<NodeLayer> EndNodeLayerList = null;

                //                List<NodeLayer> NodeLayerConnectedList = (from nll in BottomNodeLayerList
                //                                                          from n in snl.Node.ConnectNodeList
                //                                                          where (n.ID == nll.Node.ID)
                //                                                          select nll).ToList<NodeLayer>();

                //                EndNodeLayerList = (from nll in NodeLayerConnectedList
                //                                    where (nll.Node.ID != snl.Node.ID)
                //                                    && (nll.Node.Value < ContourValue)
                //                                    && nll.Layer == Layer
                //                                    select nll).ToList<NodeLayer>();

                //                foreach (NodeLayer en in EndNodeLayerList)
                //                {
                //                    AllNodeList.Add(en.Node);
                //                }

                //                if (snl.Node.Code != 0)
                //                {
                //                    AllNodeList.Add(snl.Node);
                //                }

                //            }

                //            if (AllNodeList.Count == 0)
                //            {
                //                //vcount += 1;
                //                continue;
                //            }

                //            List<Element> TempUniqueElementList = new List<Element>();
                //            List<Element> UniqueElementList = new List<Element>();
                //            foreach (ElementLayer el in ElementLayerList.Where(l => l.Layer == Layer))
                //            {
                //                if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
                //                {
                //                    if (el.Element.Type == 32)
                //                    {
                //                        bool NodeBigger = false;
                //                        for (int i = 3; i < 6; i++)
                //                        {
                //                            if (el.Element.NodeList[i].Value >= ContourValue)
                //                            {
                //                                NodeBigger = true;
                //                                break;
                //                            }
                //                        }
                //                        if (NodeBigger)
                //                        {
                //                            int countTrue = 0;
                //                            for (int i = 3; i < 6; i++)
                //                            {
                //                                if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
                //                                {
                //                                    countTrue += 1;
                //                                }
                //                            }
                //                            if (countTrue != el.Element.NodeList.Count)
                //                            {
                //                                TempUniqueElementList.Add(el.Element);
                //                            }
                //                        }
                //                    }
                //                    else if (el.Element.Type == 33)
                //                    {
                //                        bool NodeBigger = false;
                //                        for (int i = 4; i < 8; i++)
                //                        {
                //                            if (el.Element.NodeList[i].Value >= ContourValue)
                //                            {
                //                                NodeBigger = true;
                //                                break;
                //                            }
                //                        }
                //                        if (NodeBigger)
                //                        {
                //                            int countTrue = 0;
                //                            for (int i = 4; i < 8; i++)
                //                            {
                //                                if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
                //                                {
                //                                    countTrue += 1;
                //                                }
                //                            }
                //                            if (countTrue != el.Element.NodeList.Count)
                //                            {
                //                                TempUniqueElementList.Add(el.Element);
                //                            }
                //                        }
                //                    }
                //                    else
                //                    {
                //                        UpdateTask(AppTaskID, "");
                //                        throw new Exception("Element type is not supported: Element type = [" + el.Element.Type + "]");
                //                    }
                //                }
                //                else
                //                {
                //                    UpdateTask(AppTaskID, "");
                //                    throw new Exception("Bottom only exist for Dfsu3DSigma and Dfsu3DSigmaZ.");
                //                }
                //            }

                //            UniqueElementList = (from el in TempUniqueElementList select el).Distinct().ToList<Element>();

                //            // filling InterpolatedContourNodeList
                //            InterpolatedContourNodeList = new List<Node>();

                //            foreach (Element el in UniqueElementList)
                //            {
                //                if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
                //                {
                //                    if (el.Type == 32)
                //                    {
                //                        if (el.NodeList[0].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[0], el.NodeList[1], ContourValue);
                //                        }
                //                        if (el.NodeList[0].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[0], el.NodeList[2], ContourValue);
                //                        }
                //                        if (el.NodeList[1].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[1], el.NodeList[0], ContourValue);
                //                        }
                //                        if (el.NodeList[1].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[1], el.NodeList[2], ContourValue);
                //                        }
                //                        if (el.NodeList[2].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[2], el.NodeList[1], ContourValue);
                //                        }
                //                        if (el.NodeList[2].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[2], el.NodeList[0], ContourValue);
                //                        }
                //                    }
                //                    else if (el.Type == 33)
                //                    {
                //                        if (el.NodeList[0].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[0], el.NodeList[1], ContourValue);
                //                        }
                //                        if (el.NodeList[0].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[0], el.NodeList[3], ContourValue);
                //                        }
                //                        if (el.NodeList[1].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[1], el.NodeList[0], ContourValue);
                //                        }
                //                        if (el.NodeList[1].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[1], el.NodeList[2], ContourValue);
                //                        }
                //                        if (el.NodeList[2].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[2], el.NodeList[1], ContourValue);
                //                        }
                //                        if (el.NodeList[2].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[2], el.NodeList[3], ContourValue);
                //                        }
                //                        if (el.NodeList[3].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[3], el.NodeList[0], ContourValue);
                //                        }
                //                        if (el.NodeList[3].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[3], el.NodeList[2], ContourValue);
                //                        }
                //                    }
                //                    else
                //                    {
                //                        UpdateTask(AppTaskID, "");
                //                        throw new Exception("Element type is not supported: Element type = [" + el.Type + "]");
                //                    }
                //                }
                //                else
                //                {
                //                    UpdateTask(AppTaskID, "");
                //                    throw new Exception("Bottom only exist for Dfsu3DSigma and Dfsu3DSigmaZ.");
                //                }
                //            }

                //            List<Node> UniqueNodeList = (from n in AllNodeList orderby n.ID select n).Distinct().ToList<Node>();

                //            // ------------------------- new code --------------------------
                //            //                     

                //            ForwardVector = new Dictionary<string, Vector>();
                //            BackwardVector = new Dictionary<string, Vector>();

                //            foreach (Element el in UniqueElementList)
                //            {
                //                if (el.Type == 21)
                //                {
                //                    FillVectors21_32(el, UniqueElementList, ContourValue, AppTaskID, false, false);
                //                }
                //                else if (el.Type == 24)
                //                {
                //                    UpdateTask(AppTaskID, "");
                //                    throw new Exception("This should never happen. Node0, Node1 nd Node2 all < ContourValue");
                //                }
                //                else if (el.Type == 25)
                //                {
                //                    FillVectors25_33(el, UniqueElementList, ContourValue, AppTaskID, false, false);
                //                }
                //                else if (el.Type == 32)
                //                {
                //                    FillVectors21_32(el, UniqueElementList, ContourValue, AppTaskID, true, false);
                //                }
                //                else if (el.Type == 33)
                //                {
                //                    FillVectors25_33(el, UniqueElementList, ContourValue, AppTaskID, true, false);
                //                }
                //                else
                //                {
                //                    UpdateTask(AppTaskID, "");
                //                    throw new Exception("Element type is not supported: Element type = [" + el.Type + "]");
                //                }

                //            }

                //            //-------------- new code ------------------------



                //            bool MoreContourLine = true;
                //            while (MoreContourLine)
                //            {
                //                List<Node> FinalContourNodeList = new List<Node>();
                //                Vector LastVector = new Vector();
                //                LastVector = ForwardVector.First().Value;
                //                FinalContourNodeList.Add(LastVector.StartNode);
                //                FinalContourNodeList.Add(LastVector.EndNode);
                //                ForwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
                //                BackwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
                //                bool PolygonCompleted = false;
                //                while (!PolygonCompleted)
                //                {
                //                    List<string> KeyStrList = (from k in ForwardVector.Keys
                //                                               where k.StartsWith(LastVector.EndNode.ID.ToString() + ",")
                //                                               && !k.EndsWith("," + LastVector.StartNode.ID.ToString())
                //                                               select k).ToList<string>();

                //                    if (KeyStrList.Count != 1)
                //                    {
                //                        KeyStrList = (from k in BackwardVector.Keys
                //                                      where k.StartsWith(LastVector.EndNode.ID.ToString() + ",")
                //                                      && !k.EndsWith("," + LastVector.StartNode.ID.ToString())
                //                                      select k).ToList<string>();

                //                        if (KeyStrList.Count != 1)
                //                        {
                //                            PolygonCompleted = true;
                //                            break;
                //                        }
                //                        else
                //                        {
                //                            LastVector = BackwardVector[KeyStrList[0]];
                //                            BackwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
                //                            ForwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
                //                        }
                //                    }
                //                    else
                //                    {
                //                        LastVector = ForwardVector[KeyStrList[0]];
                //                        ForwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
                //                        BackwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
                //                    }
                //                    FinalContourNodeList.Add(LastVector.EndNode);
                //                    if (FinalContourNodeList[FinalContourNodeList.Count - 1] == FinalContourNodeList[0])
                //                    {
                //                        PolygonCompleted = true;
                //                    }
                //                }

                //                if (CalculateAreaOfPolygon(FinalContourNodeList) < 0)
                //                {
                //                    FinalContourNodeList.Reverse();
                //                }

                //                FinalContourNodeList.Add(FinalContourNodeList[0]);
                //                ContourPolygon contourPolygon = new ContourPolygon() { };
                //                contourPolygon.ContourNodeList = FinalContourNodeList;
                //                contourPolygon.ContourValue = ContourValue;
                //                contourPolygon.Layer = Layer;
                //                ContourPolygonList.Add(contourPolygon);

                //                if (ForwardVector.Count == 0)
                //                {
                //                    MoreContourLine = false;
                //                }

                //            }
                //            DrawKMLContourPolygon(ContourPolygonList, dfsuFile, vcount, sbStyleFeacalColiformContour, sbPlacemarkFeacalColiformContour);
                //            vcount += 1;
                //        }
                //        sbPlacemarkFeacalColiformContour.AppendLine(@"</Folder>");
                //        CountContourValue += 1;
                //    }
                //    sbPlacemarkFeacalColiformContour.AppendLine(@"</Folder>");
                //}
                #endregion Bottom of Layer
            }
            sbPlacemarkFeacalColiformContour.AppendLine(@"</Folder>");

            return;
        }
        private void WriteKMLMesh(StringBuilder sbStyleMesh, StringBuilder sbKMLMesh, List<ElementLayer> ElementLayerList)
        {
            List<Node> nodeList = new List<Node>();

            sbStyleMesh.AppendLine(@"<Style id=""_line"">");
            sbStyleMesh.AppendLine("<LineStyle>");
            sbStyleMesh.AppendLine(@"<color>ff555555</color>");
            sbStyleMesh.AppendLine(@"<width>1</width>");
            sbStyleMesh.AppendLine("</LineStyle>");
            sbStyleMesh.AppendLine(@"</Style>");


            sbKMLMesh.AppendLine(@"<Folder>");
            sbKMLMesh.AppendLine(@"<visibility>0</visibility>");
            sbKMLMesh.AppendLine(@"<name>" + TaskRunnerServiceRes.Mesh + "</name>");

            int CountRefresh = 0;
            int CountAt = 0;
            int UpdateAfter = (int)(ElementLayerList.Count() / 10);
            foreach (ElementLayer ElemLayer in ElementLayerList.OrderBy(c => c.Element.ID))
            {
                CountRefresh += 1;
                CountAt += 1;
                if (CountRefresh > UpdateAfter)
                {
                    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)((CountAt * 10) / ElementLayerList.Count()));

                    CountRefresh = 0;
                }

                StringBuilder sbCoord = new StringBuilder();
                float total = 0;
                string LastPart = "";
                foreach (Node node in ElemLayer.Element.NodeList)
                {

                    nodeList.Add(node);

                    if (LastPart == "")
                        LastPart = node.X.ToString().Replace(",", ".") + @"," + node.Y.ToString().Replace(",", ".") + ",0 ";

                    total += node.Z;
                    sbCoord.Append(node.X.ToString().Replace(",", ".") + @"," + node.Y.ToString().Replace(",", ".") + ",0 ");

                }
                sbCoord.Append(LastPart);

                string PolyName = ElemLayer.Element.ID.ToString();

                // Inserting the Placemark
                sbKMLMesh.AppendLine(@"<Placemark>");
                sbKMLMesh.AppendLine(@"<visibility>0</visibility>");
                sbKMLMesh.AppendLine(string.Format(@"<name>{0}</name>", PolyName));
                sbKMLMesh.AppendLine(@"<styleUrl>#_line</styleUrl>");
                sbKMLMesh.AppendLine(@"<LineString>");
                sbKMLMesh.AppendLine(@"<coordinates>");
                sbKMLMesh.AppendLine(sbCoord.ToString());
                sbKMLMesh.AppendLine(@"</coordinates>");
                sbKMLMesh.AppendLine(@"</LineString>");
                sbKMLMesh.AppendLine(@"</Placemark>");
            }

            sbKMLMesh.AppendLine(@"</Folder>");

            //List<Node> uniqueNode = (from n in nodeList
            //                         select n).Distinct().ToList();

            //sbKMLMesh.AppendLine(@"<Folder>");

            //foreach (Node node in uniqueNode)
            //{
            //    sbKMLMesh.AppendLine("<Placemark>");
            //    sbKMLMesh.AppendLine(@"<visibility>0</visibility>");
            //    sbKMLMesh.AppendLine(string.Format(@"<name>{0}</name>", node.ID));
            //    sbKMLMesh.AppendLine(@"<Point>");
            //    sbKMLMesh.AppendLine(string.Format(@"<coordinates>{0},{1},0</coordinates>", node.X, node.Y));
            //    sbKMLMesh.AppendLine("@</Point>");
            //    sbKMLMesh.AppendLine("</Placemark>");
            //}
            //sbKMLMesh.AppendLine(@"</Folder>");

        }
        //private void WriteKMLModelInput(StringBuilder sbStyleModelInput, StringBuilder sbKMLModelInput, List<float> ContourValueList, int mikeScenarioID, M21_3FMService m21_3fmInput)
        //{
        //    string NotUsed = "";

        //    MikeScenarioService mikeScenarioService = new MikeScenarioService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    MikeScenarioModel mikeScenarioModel = mikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //    if (!string.IsNullOrWhiteSpace(mikeScenarioModel.Error))
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        return;
        //    }

        //    TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    TVItemModel tvItemModelMikeScenario = tvItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
        //    if (!string.IsNullOrWhiteSpace(tvItemModelMikeScenario.Error))
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //        return;
        //    }

        //    MikeSourceService mikeSourceService = new MikeSourceService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //    List<MikeSourceModel> mikeSourceModelList = mikeSourceService.GetMikeSourceModelListWithMikeScenarioTVItemIDDB(mikeScenarioModel.MikeScenarioTVItemID);
        //    if (mikeSourceModelList.Count == 0)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeSources, TaskRunnerServiceRes.TVItemID, mikeScenarioModel.MikeScenarioTVItemID.ToString());
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeSources, TaskRunnerServiceRes.TVItemID, mikeScenarioModel.MikeScenarioTVItemID.ToString());
        //        return;
        //    }

        //    WriteKMLStyleModelInput(sbStyleModelInput);
        //    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //        return;

        //    sbKMLModelInput.Append("<Folder><name>" + TaskRunnerServiceRes.ModelInput + "</name>");
        //    sbKMLModelInput.AppendLine(@"<visibility>0</visibility>");

        //    sbKMLModelInput.Append("<Folder><name>" + TaskRunnerServiceRes.SourceIncluded + @"</name>");
        //    sbKMLModelInput.AppendLine(@"<visibility>0</visibility>");

        //    // showing used sources 
        //    foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.SOURCES.SOURCE> kvp in m21_3fmInput.femEngineHD.hydrodynamic_module.sources.source)
        //    {
        //        M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.SOURCES.SOURCE shd = kvp.Value;
        //        if (shd.include == 1)
        //        {
        //            WriteSourcePlacemark(sbKMLModelInput, shd, kvp, mikeSourceModelList, m21_3fmInput);
        //            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                return;
        //        }
        //    }

        //    sbKMLModelInput.Append("</Folder>");

        //    sbKMLModelInput.Append("<Folder><name>" + TaskRunnerServiceRes.SourceNotIncluded + @"</name>");
        //    sbKMLModelInput.AppendLine(@"<visibility>0</visibility>");

        //    // showing not used sources 
        //    foreach (KeyValuePair<string, M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.SOURCES.SOURCE> kvp in m21_3fmInput.femEngineHD.hydrodynamic_module.sources.source)
        //    {
        //        M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.SOURCES.SOURCE shd = kvp.Value;
        //        if (shd.include != 1)
        //        {
        //            WriteSourcePlacemark(sbKMLModelInput, shd, kvp, mikeSourceModelList, m21_3fmInput);
        //            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
        //                return;
        //        }
        //    }
        //    sbKMLModelInput.Append("</Folder>");

        //    sbKMLModelInput.Append("</Folder>");

        //    return;
        //}
        private void WriteKMLPollutionLimitsContourLine(DfsuFile dfsuFile, string dfsParamItem, StringBuilder sbStylePollutionLimitsContour, StringBuilder sbKMLPollutionLimitsContour, List<float> ContourValueList, List<int> SigmaLayerValueList, List<int> ZLayerValueList, List<float> DepthValueList, List<ElementLayer> ElementLayerList, List<NodeLayer> TopNodeLayerList, List<NodeLayer> BottomNodeLayerList)
        {
            string NotUsed = "";

            int ItemNumber = 0;

            // getting the ItemNumber
            foreach (IDfsSimpleDynamicItemInfo dfsDyInfo in dfsuFile.ItemInfo)
            {
                if (dfsDyInfo.Quantity.Item == eumItem.eumIConcentration)
                {
                    ItemNumber = dfsDyInfo.ItemNumber;
                    break;
                }
            }

            if (ItemNumber == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind__, TaskRunnerServiceRes.ParameterType, dfsParamItem);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind__", TaskRunnerServiceRes.ParameterType, dfsParamItem);
                return;
            }

            DrawKMLContourStyle(sbStylePollutionLimitsContour, sbKMLPollutionLimitsContour);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            //int pcount = 0;
            sbKMLPollutionLimitsContour.AppendLine(@"<Folder><name>" + TaskRunnerServiceRes.PollutionLimits + @"</name>");
            sbKMLPollutionLimitsContour.AppendLine(@"<visibility>0</visibility>");
            int CountLayer = 1;
            int CountAt = 0;

            foreach (int Layer in SigmaLayerValueList)
            {
                #region Top of Layer
                sbKMLPollutionLimitsContour.AppendLine(string.Format(@"<Folder><name>" + TaskRunnerServiceRes.TopOfLayer + @" [{0}]</name>", Layer));
                sbKMLPollutionLimitsContour.AppendLine(@"<visibility>0</visibility>");
                int CountContour = 1;
                foreach (float ContourValue in ContourValueList)
                {
                    CountAt += 1;
                    sbKMLPollutionLimitsContour.AppendLine(string.Format(@"<Folder><name>" + TaskRunnerServiceRes.ContourValue + @" [{0}]</name>", ContourValue));
                    sbKMLPollutionLimitsContour.AppendLine(@"<visibility>0</visibility>");
                    string AppTaskStatus = ((int)((CountAt * 100) / (SigmaLayerValueList.Count * ContourValueList.Count))).ToString() + " %";

                    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, ((int)((CountAt * 100) / (SigmaLayerValueList.Count * ContourValueList.Count))));

                    List<Node> AllNodeList = new List<Node>();
                    List<ContourPolygon> ContourPolygonList = new List<ContourPolygon>();

                    for (int timeStep = 0; timeStep < dfsuFile.NumberOfTimeSteps; timeStep++)
                    {

                        float[] ValueList = (float[])dfsuFile.ReadItemTimeStep(ItemNumber, timeStep).Data;

                        for (int i = 0; i < ElementLayerList.Count; i++)
                        {
                            if (ElementLayerList[i].Element.Value < ValueList[i])
                            {
                                ElementLayerList[i].Element.Value = ValueList[i];
                            }
                        }
                    }
                    //}

                    foreach (NodeLayer nl in TopNodeLayerList)
                    {
                        float Total = 0;
                        foreach (Element element in nl.Node.ElementList)
                        {
                            Total += element.Value;
                        }
                        nl.Node.Value = Total / nl.Node.ElementList.Count;
                    }

                    List<NodeLayer> AboveNodeLayerList = new List<NodeLayer>();

                    AboveNodeLayerList = (from n in TopNodeLayerList
                                          where (n.Node.Value >= ContourValue)
                                          && n.Layer == Layer
                                          select n).ToList<NodeLayer>();

                    foreach (NodeLayer snl in AboveNodeLayerList)
                    {
                        List<NodeLayer> EndNodeLayerList = null;

                        List<NodeLayer> NodeLayerConnectedList = (from nll in TopNodeLayerList
                                                                  from n in snl.Node.ConnectNodeList
                                                                  where (n.ID == nll.Node.ID)
                                                                  select nll).ToList<NodeLayer>();

                        EndNodeLayerList = (from nll in NodeLayerConnectedList
                                            where (nll.Node.ID != snl.Node.ID)
                                            && (nll.Node.Value < ContourValue)
                                            && nll.Layer == Layer
                                            select nll).ToList<NodeLayer>();

                        foreach (NodeLayer en in EndNodeLayerList)
                        {
                            AllNodeList.Add(en.Node);
                        }

                        if (snl.Node.Code != 0)
                        {
                            AllNodeList.Add(snl.Node);
                        }

                    }

                    //if (AllNodeList.Count == 0)
                    //{
                    //    continue;
                    //}

                    List<Element> TempUniqueElementList = new List<Element>();
                    List<Element> UniqueElementList = new List<Element>();
                    foreach (ElementLayer el in ElementLayerList.Where(l => l.Layer == Layer))
                    {
                        if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
                        {
                            if (el.Element.Type == 32)
                            {
                                bool NodeBigger = false;
                                for (int i = 3; i < 6; i++)
                                {
                                    if (el.Element.NodeList[i].Value >= ContourValue)
                                    {
                                        NodeBigger = true;
                                        break;
                                    }
                                }
                                if (NodeBigger)
                                {
                                    int countTrue = 0;
                                    for (int i = 3; i < 6; i++)
                                    {
                                        if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
                                        {
                                            countTrue += 1;
                                        }
                                    }
                                    if (countTrue != el.Element.NodeList.Count)
                                    {
                                        TempUniqueElementList.Add(el.Element);
                                    }
                                }
                            }
                            else if (el.Element.Type == 33)
                            {
                                bool NodeBigger = false;
                                for (int i = 4; i < 8; i++)
                                {
                                    if (el.Element.NodeList[i].Value >= ContourValue)
                                    {
                                        NodeBigger = true;
                                        break;
                                    }
                                }
                                if (NodeBigger)
                                {
                                    int countTrue = 0;
                                    for (int i = 4; i < 8; i++)
                                    {
                                        if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
                                        {
                                            countTrue += 1;
                                        }
                                    }
                                    if (countTrue != el.Element.NodeList.Count)
                                    {
                                        TempUniqueElementList.Add(el.Element);
                                    }
                                }
                            }
                            else
                            {
                                NotUsed = string.Format(TaskRunnerServiceRes.ElementType_IsNotSupported, el.Element.Type.ToString());
                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ElementType_IsNotSupported", el.Element.Type.ToString());
                                return;
                            }
                        }
                        else if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu2D)
                        {
                            bool NodeBigger = false;
                            for (int i = 0; i < el.Element.NodeList.Count; i++)
                            {
                                if (el.Element.NodeList[i].Value >= ContourValue)
                                {
                                    NodeBigger = true;
                                    break;
                                }
                            }
                            if (NodeBigger)
                            {
                                int countTrue = 0;
                                for (int i = 0; i < el.Element.NodeList.Count; i++)
                                {
                                    if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
                                    {
                                        countTrue += 1;
                                    }
                                }
                                if (countTrue != el.Element.NodeList.Count)
                                {
                                    TempUniqueElementList.Add(el.Element);
                                }
                            }
                        }
                    }

                    UniqueElementList = (from el in TempUniqueElementList select el).Distinct().ToList<Element>();

                    // filling InterpolatedContourNodeList
                    InterpolatedContourNodeList = new List<Node>();

                    foreach (Element el in UniqueElementList)
                    {
                        if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
                        {
                            if (el.Type == 32)
                            {
                                if (el.NodeList[3].Value >= ContourValue && el.NodeList[4].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[3], el.NodeList[4], ContourValue);
                                }
                                if (el.NodeList[3].Value >= ContourValue && el.NodeList[5].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[3], el.NodeList[5], ContourValue);
                                }
                                if (el.NodeList[4].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[4], el.NodeList[3], ContourValue);
                                }
                                if (el.NodeList[4].Value >= ContourValue && el.NodeList[5].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[4], el.NodeList[5], ContourValue);
                                }
                                if (el.NodeList[5].Value >= ContourValue && el.NodeList[4].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[5], el.NodeList[4], ContourValue);
                                }
                                if (el.NodeList[5].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[5], el.NodeList[3], ContourValue);
                                }
                            }
                            else if (el.Type == 33)
                            {
                                if (el.NodeList[4].Value >= ContourValue && el.NodeList[5].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[4], el.NodeList[5], ContourValue);
                                }
                                if (el.NodeList[4].Value >= ContourValue && el.NodeList[7].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[4], el.NodeList[7], ContourValue);
                                }
                                if (el.NodeList[5].Value >= ContourValue && el.NodeList[4].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[5], el.NodeList[4], ContourValue);
                                }
                                if (el.NodeList[5].Value >= ContourValue && el.NodeList[6].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[5], el.NodeList[6], ContourValue);
                                }
                                if (el.NodeList[6].Value >= ContourValue && el.NodeList[5].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[6], el.NodeList[5], ContourValue);
                                }
                                if (el.NodeList[6].Value >= ContourValue && el.NodeList[7].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[6], el.NodeList[7], ContourValue);
                                }
                                if (el.NodeList[7].Value >= ContourValue && el.NodeList[4].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[7], el.NodeList[4], ContourValue);
                                }
                                if (el.NodeList[7].Value >= ContourValue && el.NodeList[6].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[7], el.NodeList[6], ContourValue);
                                }
                            }
                            else
                            {
                                NotUsed = string.Format(TaskRunnerServiceRes.ElementType_IsNotSupported, el.Type.ToString());
                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ElementType_IsNotSupported", el.Type.ToString());
                                return;
                            }
                        }
                        else if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu2D)
                        {
                            if (el.Type == 21)
                            {
                                if (el.NodeList[0].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[0], el.NodeList[1], ContourValue);
                                }
                                if (el.NodeList[0].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[0], el.NodeList[2], ContourValue);
                                }
                                if (el.NodeList[1].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[1], el.NodeList[0], ContourValue);
                                }
                                if (el.NodeList[1].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[1], el.NodeList[2], ContourValue);
                                }
                                if (el.NodeList[2].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[2], el.NodeList[1], ContourValue);
                                }
                                if (el.NodeList[2].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[2], el.NodeList[0], ContourValue);
                                }
                            }
                            else if (el.Type == 24)
                            {
                            }
                            else if (el.Type == 25)
                            {
                                if (el.NodeList[0].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[0], el.NodeList[1], ContourValue);
                                }
                                if (el.NodeList[0].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[0], el.NodeList[3], ContourValue);
                                }
                                if (el.NodeList[1].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[1], el.NodeList[0], ContourValue);
                                }
                                if (el.NodeList[1].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[1], el.NodeList[2], ContourValue);
                                }
                                if (el.NodeList[2].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[2], el.NodeList[1], ContourValue);
                                }
                                if (el.NodeList[2].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[2], el.NodeList[3], ContourValue);
                                }
                                if (el.NodeList[3].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[3], el.NodeList[0], ContourValue);
                                }
                                if (el.NodeList[3].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[3], el.NodeList[2], ContourValue);
                                }
                            }
                            else
                            {
                                NotUsed = string.Format(TaskRunnerServiceRes.ElementType_IsNotSupported, el.Type.ToString());
                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ElementType_IsNotSupported", el.Type.ToString());
                                return;
                            }
                        }
                    }

                    List<Node> UniqueNodeList = (from n in AllNodeList orderby n.ID select n).Distinct().ToList<Node>();

                    ForwardVector = new Dictionary<String, Vector>();
                    BackwardVector = new Dictionary<String, Vector>();

                    // ------------------------- new code --------------------------
                    //                     

                    foreach (Element el in UniqueElementList)
                    {
                        if (el.Type == 21)
                        {
                            FillVectors21_32(el, UniqueElementList, ContourValue, false, true);
                        }
                        else if (el.Type == 24)
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.ElementType_IsNotSupported, el.Type.ToString());
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ElementType_IsNotSupported", el.Type.ToString());
                            return;
                        }
                        else if (el.Type == 25)
                        {
                            FillVectors25_33(el, UniqueElementList, ContourValue, false, true);
                        }
                        else if (el.Type == 32)
                        {
                            FillVectors21_32(el, UniqueElementList, ContourValue, true, true);
                        }
                        else if (el.Type == 33)
                        {
                            FillVectors25_33(el, UniqueElementList, ContourValue, true, true);
                        }
                        else
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.ElementType_IsNotSupported, el.Type.ToString());
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ElementType_IsNotSupported", el.Type.ToString());
                            return;
                        }

                    }

                    //-------------- new code ------------------------


                    bool MoreContourLine = true;
                    MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                    while (MoreContourLine && ForwardVector.Count > 0)
                    {
                        List<Node> FinalContourNodeList = new List<Node>();
                        Vector LastVector = new Vector();
                        LastVector = ForwardVector.First().Value;
                        FinalContourNodeList.Add(LastVector.StartNode);
                        FinalContourNodeList.Add(LastVector.EndNode);
                        ForwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
                        BackwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
                        bool PolygonCompleted = false;
                        while (!PolygonCompleted)
                        {
                            List<string> KeyStrList = (from k in ForwardVector.Keys
                                                       where k.StartsWith(LastVector.EndNode.ID.ToString() + ",")
                                                       && !k.EndsWith("," + LastVector.StartNode.ID.ToString())
                                                       select k).ToList<string>();

                            if (KeyStrList.Count != 1)
                            {
                                KeyStrList = (from k in BackwardVector.Keys
                                              where k.StartsWith(LastVector.EndNode.ID.ToString() + ",")
                                              && !k.EndsWith("," + LastVector.StartNode.ID.ToString())
                                              select k).ToList<string>();

                                if (KeyStrList.Count != 1)
                                {
                                    PolygonCompleted = true;
                                    break;
                                }
                                else
                                {
                                    LastVector = BackwardVector[KeyStrList[0]];
                                    BackwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
                                    ForwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
                                }
                            }
                            else
                            {
                                LastVector = ForwardVector[KeyStrList[0]];
                                ForwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
                                BackwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
                            }
                            FinalContourNodeList.Add(LastVector.EndNode);
                            if (FinalContourNodeList[FinalContourNodeList.Count - 1] == FinalContourNodeList[0])
                            {
                                PolygonCompleted = true;
                            }
                        }

                        if (mapInfoService.CalculateAreaOfPolygon(FinalContourNodeList) < 0)
                        {
                            FinalContourNodeList.Reverse();
                        }

                        FinalContourNodeList.Add(FinalContourNodeList[0]);
                        ContourPolygon contourPolygon = new ContourPolygon() { };
                        contourPolygon.ContourNodeList = FinalContourNodeList;
                        contourPolygon.ContourValue = ContourValue;
                        ContourPolygonList.Add(contourPolygon);

                        if (ForwardVector.Count == 0)
                        {
                            MoreContourLine = false;
                        }
                    }

                    foreach (ContourPolygon contourPolygon in ContourPolygonList)
                    {
                        sbKMLPollutionLimitsContour.AppendLine(@"<Folder>");
                        sbKMLPollutionLimitsContour.AppendLine(@"<visibility>0</visibility>");
                        sbKMLPollutionLimitsContour.AppendLine(@"<Placemark>");
                        sbKMLPollutionLimitsContour.AppendLine(@"<visibility>0</visibility>");
                        if (contourPolygon.ContourValue >= 14 && contourPolygon.ContourValue < 88)
                        {
                            sbKMLPollutionLimitsContour.AppendLine(@"<styleUrl>#fc_14</styleUrl>");
                        }
                        else if (contourPolygon.ContourValue >= 88)
                        {
                            sbKMLPollutionLimitsContour.AppendLine(@"<styleUrl>#fc_88</styleUrl>");
                        }
                        else
                        {
                            sbKMLPollutionLimitsContour.AppendLine(@"<styleUrl>#fc_LT14</styleUrl>");
                        }
                        sbKMLPollutionLimitsContour.AppendLine(@"<Polygon>");
                        sbKMLPollutionLimitsContour.AppendLine(@"<outerBoundaryIs>");
                        sbKMLPollutionLimitsContour.AppendLine(@"<LinearRing>");
                        sbKMLPollutionLimitsContour.AppendLine(@"<coordinates>");
                        foreach (Node node in contourPolygon.ContourNodeList)
                        {
                            sbKMLPollutionLimitsContour.Append(node.X.ToString().Replace(",", ".") + @"," + node.Y.ToString().Replace(",", ".") + ",0 ");
                        }
                        sbKMLPollutionLimitsContour.AppendLine(@"</coordinates>");
                        sbKMLPollutionLimitsContour.AppendLine(@"</LinearRing>");
                        sbKMLPollutionLimitsContour.AppendLine(@"</outerBoundaryIs>");
                        sbKMLPollutionLimitsContour.AppendLine(@"</Polygon>");
                        sbKMLPollutionLimitsContour.AppendLine(@"</Placemark>");
                        sbKMLPollutionLimitsContour.AppendLine(@"</Folder>");
                    }
                    sbKMLPollutionLimitsContour.AppendLine(@"</Folder>");
                    CountContour += 1;
                }
                sbKMLPollutionLimitsContour.AppendLine(@"</Folder>");
                #endregion Top of Layer

                #region Bottom of Layer
                //// 
                //if (Layer == dfsuFile.NumberOfSigmaLayers)
                //{
                //    sbKMLPollutionLimitsContour.AppendLine(string.Format(@"<Folder><name>Bottom of Layer [{0}]</name>", Layer));
                //    sbKMLPollutionLimitsContour.AppendLine(@"<visibility>0</visibility>");
                //    CountContour = 1;
                //    foreach (float ContourValue in ContourValueList)
                //    {
                //        CountAt += 1;
                //        sbKMLPollutionLimitsContour.AppendLine(string.Format(@"<Folder><name>Contour Value [{0}]</name>", ContourValue));
                //        sbKMLPollutionLimitsContour.AppendLine(@"<visibility>0</visibility>");
                //        string AppTaskStatus = ((int)((CountAt * 100) / (SigmaLayerValueList.Count * ContourValueList.Count))).ToString() + " %";
                //        UpdateTask(AppTaskID, AppTaskStatus);

                //        List<Node> AllNodeList = new List<Node>();
                //        List<ContourPolygon> ContourPolygonList = new List<ContourPolygon>();

                //        //foreach (Dfs.Parameter.TimeSeriesValue v in p.TimeSeriesValueList)
                //        //{
                //        for (int timeStep = 0; timeStep < dfsuFile.NumberOfTimeSteps; timeStep++)
                //        {

                //            float[] ValueList = (float[])dfsuFile.ReadItemTimeStep(ItemNumber, timeStep).Data;

                //            for (int i = 0; i < ElementLayerList.Count; i++)
                //            {
                //                if (ElementLayerList[i].Element.Value < ValueList[i])
                //                {
                //                    ElementLayerList[i].Element.Value = ValueList[i];
                //                }
                //            }
                //        }
                //        //}

                //        foreach (NodeLayer nl in BottomNodeLayerList)
                //        {
                //            float Total = 0;
                //            foreach (Element element in nl.Node.ElementList)
                //            {
                //                Total += element.Value;
                //            }
                //            nl.Node.Value = Total / nl.Node.ElementList.Count;
                //        }

                //        List<NodeLayer> AboveNodeLayerList = new List<NodeLayer>();

                //        AboveNodeLayerList = (from n in BottomNodeLayerList
                //                              where (n.Node.Value >= ContourValue)
                //                              && n.Layer == Layer
                //                              select n).ToList<NodeLayer>();

                //        foreach (NodeLayer snl in AboveNodeLayerList)
                //        {
                //            List<NodeLayer> EndNodeLayerList = null;

                //            List<NodeLayer> NodeLayerConnectedList = (from nll in BottomNodeLayerList
                //                                                      from n in snl.Node.ConnectNodeList
                //                                                      where (n.ID == nll.Node.ID)
                //                                                      select nll).ToList<NodeLayer>();

                //            EndNodeLayerList = (from nll in NodeLayerConnectedList
                //                                where (nll.Node.ID != snl.Node.ID)
                //                                && (nll.Node.Value < ContourValue)
                //                                && nll.Layer == Layer
                //                                select nll).ToList<NodeLayer>();

                //            foreach (NodeLayer en in EndNodeLayerList)
                //            {
                //                AllNodeList.Add(en.Node);
                //            }

                //            if (snl.Node.Code != 0)
                //            {
                //                AllNodeList.Add(snl.Node);
                //            }

                //        }

                //        if (AllNodeList.Count == 0)
                //        {
                //            continue;
                //        }

                //        List<Element> TempUniqueElementList = new List<Element>();
                //        List<Element> UniqueElementList = new List<Element>();
                //        foreach (ElementLayer el in ElementLayerList.Where(l => l.Layer == Layer))
                //        {
                //            if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
                //            {
                //                if (el.Element.Type == 32)
                //                {
                //                    bool NodeBigger = false;
                //                    for (int i = 0; i < 3; i++)
                //                    {
                //                        if (el.Element.NodeList[i].Value >= ContourValue)
                //                        {
                //                            NodeBigger = true;
                //                            break;
                //                        }
                //                    }
                //                    if (NodeBigger)
                //                    {
                //                        int countTrue = 0;
                //                        for (int i = 0; i < 3; i++)
                //                        {
                //                            if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
                //                            {
                //                                countTrue += 1;
                //                            }
                //                        }
                //                        if (countTrue != el.Element.NodeList.Count)
                //                        {
                //                            TempUniqueElementList.Add(el.Element);
                //                        }
                //                    }
                //                }
                //                else if (el.Element.Type == 33)
                //                {
                //                    bool NodeBigger = false;
                //                    for (int i = 0; i < 4; i++)
                //                    {
                //                        if (el.Element.NodeList[i].Value >= ContourValue)
                //                        {
                //                            NodeBigger = true;
                //                            break;
                //                        }
                //                    }
                //                    if (NodeBigger)
                //                    {
                //                        int countTrue = 0;
                //                        for (int i = 0; i < 4; i++)
                //                        {
                //                            if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
                //                            {
                //                                countTrue += 1;
                //                            }
                //                        }
                //                        if (countTrue != el.Element.NodeList.Count)
                //                        {
                //                            TempUniqueElementList.Add(el.Element);
                //                        }
                //                    }
                //                }
                //                else
                //                {
                //                    UpdateTask(AppTaskID, "");
                //                    throw new Exception("Element type is not supported: Element type = [" + el.Element.Type + "]");
                //                }
                //            }
                //            else if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu2D)
                //            {
                //                bool NodeBigger = false;
                //                for (int i = 0; i < el.Element.NodeList.Count; i++)
                //                {
                //                    if (el.Element.NodeList[i].Value >= ContourValue)
                //                    {
                //                        NodeBigger = true;
                //                        break;
                //                    }
                //                }
                //                if (NodeBigger)
                //                {
                //                    int countTrue = 0;
                //                    for (int i = 0; i < el.Element.NodeList.Count; i++)
                //                    {
                //                        if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
                //                        {
                //                            countTrue += 1;
                //                        }
                //                    }
                //                    if (countTrue != el.Element.NodeList.Count)
                //                    {
                //                        TempUniqueElementList.Add(el.Element);
                //                    }
                //                }
                //            }
                //        }

                //        UniqueElementList = (from el in TempUniqueElementList select el).Distinct().ToList<Element>();

                //        // filling InterpolatedContourNodeList
                //        InterpolatedContourNodeList = new List<Node>();

                //        foreach (Element el in UniqueElementList)
                //        {
                //            if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
                //            {
                //                if (el.Type == 32)
                //                {
                //                    if (el.NodeList[0].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[0], el.NodeList[1], ContourValue);
                //                    }
                //                    if (el.NodeList[0].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[0], el.NodeList[2], ContourValue);
                //                    }
                //                    if (el.NodeList[1].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[1], el.NodeList[0], ContourValue);
                //                    }
                //                    if (el.NodeList[1].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[1], el.NodeList[2], ContourValue);
                //                    }
                //                    if (el.NodeList[2].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[2], el.NodeList[1], ContourValue);
                //                    }
                //                    if (el.NodeList[2].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[2], el.NodeList[0], ContourValue);
                //                    }
                //                }
                //                else if (el.Type == 33)
                //                {
                //                    if (el.NodeList[0].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[0], el.NodeList[1], ContourValue);
                //                    }
                //                    if (el.NodeList[0].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[0], el.NodeList[3], ContourValue);
                //                    }
                //                    if (el.NodeList[1].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[1], el.NodeList[0], ContourValue);
                //                    }
                //                    if (el.NodeList[1].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[1], el.NodeList[2], ContourValue);
                //                    }
                //                    if (el.NodeList[2].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[2], el.NodeList[1], ContourValue);
                //                    }
                //                    if (el.NodeList[2].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[2], el.NodeList[3], ContourValue);
                //                    }
                //                    if (el.NodeList[3].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[3], el.NodeList[0], ContourValue);
                //                    }
                //                    if (el.NodeList[3].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[3], el.NodeList[2], ContourValue);
                //                    }
                //                }
                //                else
                //                {
                //                    UpdateTask(AppTaskID, "");
                //                    throw new Exception("Element type is not supported: Element type = [" + el.Type + "]");
                //                }
                //            }
                //            else
                //            {
                //                UpdateTask(AppTaskID, "");
                //                throw new Exception("Bottom does not exist outside the Dfsu3DSigma and Dfsu3DSigmaZ.");
                //            }
                //        }

                //        List<Node> UniqueNodeList = (from n in AllNodeList orderby n.ID select n).Distinct().ToList<Node>();

                //        ForwardVector = new Dictionary<String, Vector>();
                //        BackwardVector = new Dictionary<String, Vector>();

                //        // ------------------------- new code --------------------------
                //        //                     

                //        foreach (Element el in UniqueElementList)
                //        {
                //            if (el.Type == 21)
                //            {
                //                FillVectors21_32(el, UniqueElementList, ContourValue, AppTaskID, false, false);
                //            }
                //            else if (el.Type == 24)
                //            {
                //                UpdateTask(AppTaskID, "");
                //                throw new Exception("Element type is not supported: Element type = [" + el.Type + "]");
                //            }
                //            else if (el.Type == 25)
                //            {
                //                FillVectors25_33(el, UniqueElementList, ContourValue, AppTaskID, false, false);
                //            }
                //            else if (el.Type == 32)
                //            {
                //                FillVectors21_32(el, UniqueElementList, ContourValue, AppTaskID, true, false);
                //            }
                //            else if (el.Type == 33)
                //            {
                //                FillVectors25_33(el, UniqueElementList, ContourValue, AppTaskID, true, false);
                //            }
                //            else
                //            {
                //                UpdateTask(AppTaskID, "");
                //                throw new Exception("Element type is not supported: Element type = [" + el.Type + "]");
                //            }

                //        }

                //        //-------------- new code ------------------------


                //        bool MoreContourLine = true;
                //        while (MoreContourLine)
                //        {
                //            List<Node> FinalContourNodeList = new List<Node>();
                //            Vector LastVector = new Vector();
                //            LastVector = ForwardVector.First().Value;
                //            FinalContourNodeList.Add(LastVector.StartNode);
                //            FinalContourNodeList.Add(LastVector.EndNode);
                //            ForwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
                //            BackwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
                //            bool PolygonCompleted = false;
                //            while (!PolygonCompleted)
                //            {
                //                List<string> KeyStrList = (from k in ForwardVector.Keys
                //                                           where k.StartsWith(LastVector.EndNode.ID.ToString() + ",")
                //                                           && !k.EndsWith("," + LastVector.StartNode.ID.ToString())
                //                                           select k).ToList<string>();

                //                if (KeyStrList.Count != 1)
                //                {
                //                    KeyStrList = (from k in BackwardVector.Keys
                //                                  where k.StartsWith(LastVector.EndNode.ID.ToString() + ",")
                //                                  && !k.EndsWith("," + LastVector.StartNode.ID.ToString())
                //                                  select k).ToList<string>();

                //                    if (KeyStrList.Count != 1)
                //                    {
                //                        PolygonCompleted = true;
                //                        break;
                //                    }
                //                    else
                //                    {
                //                        LastVector = BackwardVector[KeyStrList[0]];
                //                        BackwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
                //                        ForwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
                //                    }
                //                }
                //                else
                //                {
                //                    LastVector = ForwardVector[KeyStrList[0]];
                //                    ForwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
                //                    BackwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
                //                }
                //                FinalContourNodeList.Add(LastVector.EndNode);
                //                if (FinalContourNodeList[FinalContourNodeList.Count - 1] == FinalContourNodeList[0])
                //                {
                //                    PolygonCompleted = true;
                //                }
                //            }

                //            if (CalculateAreaOfPolygon(FinalContourNodeList) < 0)
                //            {
                //                FinalContourNodeList.Reverse();
                //            }

                //            FinalContourNodeList.Add(FinalContourNodeList[0]);
                //            ContourPolygon contourPolygon = new ContourPolygon() { };
                //            contourPolygon.ContourNodeList = FinalContourNodeList;
                //            contourPolygon.ContourValue = ContourValue;
                //            ContourPolygonList.Add(contourPolygon);

                //            if (ForwardVector.Count == 0)
                //            {
                //                MoreContourLine = false;
                //            }
                //        }
                //        //sbKMLPollutionLimitsContour.AppendLine(@"<Folder>");
                //        //sbKMLPollutionLimitsContour.AppendLine(string.Format(@"<name>{0} Pollution Limits Contour</name>", ContourValue));
                //        //sbKMLPollutionLimitsContour.AppendLine(@"<visibility>0</visibility>");

                //        foreach (ContourPolygon contourPolygon in ContourPolygonList)
                //        {
                //            sbKMLPollutionLimitsContour.AppendLine(@"<Folder>");
                //            sbKMLPollutionLimitsContour.AppendLine(@"<visibility>0</visibility>");
                //            sbKMLPollutionLimitsContour.AppendLine(@"<Placemark>");
                //            sbKMLPollutionLimitsContour.AppendLine(@"<visibility>0</visibility>");
                //            if (contourPolygon.ContourValue >= 14 && contourPolygon.ContourValue < 88)
                //            {
                //                sbKMLPollutionLimitsContour.AppendLine(@"<styleUrl>#fc_14</styleUrl>");
                //            }
                //            else if (contourPolygon.ContourValue >= 88)
                //            {
                //                sbKMLPollutionLimitsContour.AppendLine(@"<styleUrl>#fc_88</styleUrl>");
                //            }
                //            else
                //            {
                //                sbKMLPollutionLimitsContour.AppendLine(@"<styleUrl>#fc_LT14</styleUrl>");
                //            }
                //            sbKMLPollutionLimitsContour.AppendLine(@"<Polygon>");
                //            sbKMLPollutionLimitsContour.AppendLine(@"<outerBoundaryIs>");
                //            sbKMLPollutionLimitsContour.AppendLine(@"<LinearRing>");
                //            sbKMLPollutionLimitsContour.AppendLine(@"<coordinates>");
                //            foreach (Node node in contourPolygon.ContourNodeList)
                //            {
                //                sbKMLPollutionLimitsContour.Append(string.Format(@"{0},{1},0 ", node.X, node.Y));
                //            }
                //            sbKMLPollutionLimitsContour.AppendLine(@"</coordinates>");
                //            sbKMLPollutionLimitsContour.AppendLine(@"</LinearRing>");
                //            sbKMLPollutionLimitsContour.AppendLine(@"</outerBoundaryIs>");
                //            sbKMLPollutionLimitsContour.AppendLine(@"</Polygon>");
                //            sbKMLPollutionLimitsContour.AppendLine(@"</Placemark>");
                //            sbKMLPollutionLimitsContour.AppendLine(@"</Folder>");
                //        }
                //        sbKMLPollutionLimitsContour.AppendLine(@"</Folder>");
                //        CountContour += 1;
                //    }
                //    sbKMLPollutionLimitsContour.AppendLine(@"</Folder>");
                //}
                #endregion Bottom of Layer
                CountLayer += 1;
            }
            sbKMLPollutionLimitsContour.AppendLine(@"</Folder>");
        }
        private void WriteKMLStudyAreaLine(StringBuilder sbStyleStudyAreaLine, StringBuilder sbKMLStudyAreaLine, List<Element> ElementList, List<Node> NodeList)
        {
            sbStyleStudyAreaLine.AppendLine(@"<Style id=""StudyArea"">");
            sbStyleStudyAreaLine.AppendLine(@"<LineStyle>");
            sbStyleStudyAreaLine.AppendLine(@"<color>ffffff00</color>");
            sbStyleStudyAreaLine.AppendLine(@"<width>2</width>");
            sbStyleStudyAreaLine.AppendLine(@"</LineStyle>");
            sbStyleStudyAreaLine.AppendLine(@"</Style>");

            List<ContourPolygon> contourPolygonList = new List<ContourPolygon>();

            //using (DFSU dfsu = new DFSU() )


            sbKMLStudyAreaLine.AppendLine(@"<Folder>");
            sbKMLStudyAreaLine.AppendLine(@"<name>" + TaskRunnerServiceRes.StudyArea + @"</name>");
            sbKMLStudyAreaLine.AppendLine(@"<visibility>0</visibility>");
            foreach (ContourPolygon contourPolygon in contourPolygonList)
            {
                sbKMLStudyAreaLine.AppendLine(@"<Folder>");
                sbKMLStudyAreaLine.AppendLine(@"<visibility>0</visibility>");
                sbKMLStudyAreaLine.AppendLine(@"<Placemark>");
                sbKMLStudyAreaLine.AppendLine(@"<visibility>0</visibility>");
                sbKMLStudyAreaLine.AppendLine(@"<styleUrl>#StudyArea</styleUrl>");
                sbKMLStudyAreaLine.AppendLine(@"<LineString>");
                sbKMLStudyAreaLine.AppendLine(@"<coordinates>");
                foreach (Node node in contourPolygon.ContourNodeList)
                {
                    sbKMLStudyAreaLine.Append(node.X.ToString().Replace(",", ".") + @"," + node.Y.ToString().Replace(",", ".") + ",0 ");
                }
                sbKMLStudyAreaLine.AppendLine(@"</coordinates>");
                sbKMLStudyAreaLine.AppendLine(@"</LineString>");
                sbKMLStudyAreaLine.AppendLine(@"</Placemark>");
                sbKMLStudyAreaLine.AppendLine(@"</Folder>");
            }
            sbKMLStudyAreaLine.AppendLine(@"</Folder>");
        }
        private void WriteKMLStyleModelInput(StringBuilder sbStyleModelInput)
        {
            sbStyleModelInput.AppendLine(@"<Style id=""sn_grn-pushpin"">");
            sbStyleModelInput.AppendLine(@"<IconStyle>");
            sbStyleModelInput.AppendLine(@"<scale>1.1</scale>");
            sbStyleModelInput.AppendLine(@"<Icon>");
            sbStyleModelInput.AppendLine(@"<href>http://maps.google.com/mapfiles/kml/pushpin/grn-pushpin.png</href>");
            sbStyleModelInput.AppendLine(@"</Icon>");
            sbStyleModelInput.AppendLine(@"<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbStyleModelInput.AppendLine(@"</IconStyle>");
            sbStyleModelInput.AppendLine(@"<ListStyle>");
            sbStyleModelInput.AppendLine(@"</ListStyle>");
            sbStyleModelInput.AppendLine(@"</Style>");

            sbStyleModelInput.AppendLine(@"<Style id=""sh_grn-pushpin"">");
            sbStyleModelInput.AppendLine(@"<IconStyle>");
            sbStyleModelInput.AppendLine(@"<scale>1.3</scale>");
            sbStyleModelInput.AppendLine(@"<Icon>");
            sbStyleModelInput.AppendLine(@"<href>http://maps.google.com/mapfiles/kml/pushpin/grn-pushpin.png</href>");
            sbStyleModelInput.AppendLine(@"</Icon>");
            sbStyleModelInput.AppendLine(@"<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbStyleModelInput.AppendLine(@"</IconStyle>");
            sbStyleModelInput.AppendLine(@"<ListStyle>");
            sbStyleModelInput.AppendLine(@"</ListStyle>");
            sbStyleModelInput.AppendLine(@"</Style>");

            sbStyleModelInput.AppendLine(@"<StyleMap id=""msn_grn-pushpin"">");
            sbStyleModelInput.AppendLine(@"<Pair>");
            sbStyleModelInput.AppendLine(@"<key>normal</key>");
            sbStyleModelInput.AppendLine(@"<styleUrl>#sn_grn-pushpin</styleUrl>");
            sbStyleModelInput.AppendLine(@"</Pair>");
            sbStyleModelInput.AppendLine(@"<Pair>");
            sbStyleModelInput.AppendLine(@"<key>highlight</key>");
            sbStyleModelInput.AppendLine(@"<styleUrl>#sh_grn-pushpin</styleUrl>");
            sbStyleModelInput.AppendLine(@"</Pair>");
            sbStyleModelInput.AppendLine(@"</StyleMap>");

            sbStyleModelInput.AppendLine(@"<Style id=""sn_red-pushpin"">");
            sbStyleModelInput.AppendLine(@"<IconStyle>");
            sbStyleModelInput.AppendLine(@"<scale>1.1</scale>");
            sbStyleModelInput.AppendLine(@"<Icon>");
            sbStyleModelInput.AppendLine(@"<href>http://maps.google.com/mapfiles/kml/pushpin/red-pushpin.png</href>");
            sbStyleModelInput.AppendLine(@"</Icon>");
            sbStyleModelInput.AppendLine(@"<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbStyleModelInput.AppendLine(@"</IconStyle>");
            sbStyleModelInput.AppendLine(@"<ListStyle>");
            sbStyleModelInput.AppendLine(@"</ListStyle>");
            sbStyleModelInput.AppendLine(@"</Style>");

            sbStyleModelInput.AppendLine(@"<Style id=""sh_red-pushpin"">");
            sbStyleModelInput.AppendLine(@"<IconStyle>");
            sbStyleModelInput.AppendLine(@"<scale>1.3</scale>");
            sbStyleModelInput.AppendLine(@"<Icon>");
            sbStyleModelInput.AppendLine(@"<href>http://maps.google.com/mapfiles/kml/pushpin/red-pushpin.png</href>");
            sbStyleModelInput.AppendLine(@"</Icon>");
            sbStyleModelInput.AppendLine(@"<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbStyleModelInput.AppendLine(@"</IconStyle>");
            sbStyleModelInput.AppendLine(@"<ListStyle>");
            sbStyleModelInput.AppendLine(@"</ListStyle>");
            sbStyleModelInput.AppendLine(@"</Style>");

            sbStyleModelInput.AppendLine(@"<StyleMap id=""msn_red-pushpin"">");
            sbStyleModelInput.AppendLine(@"<Pair>");
            sbStyleModelInput.AppendLine(@"<key>normal</key>");
            sbStyleModelInput.AppendLine(@"<styleUrl>#sn_red-pushpin</styleUrl>");
            sbStyleModelInput.AppendLine(@"</Pair>");
            sbStyleModelInput.AppendLine(@"<Pair>");
            sbStyleModelInput.AppendLine(@"<key>highlight</key>");
            sbStyleModelInput.AppendLine(@"<styleUrl>#sh_red-pushpin</styleUrl>");
            sbStyleModelInput.AppendLine(@"</Pair>");
            sbStyleModelInput.AppendLine(@"</StyleMap>");

            sbStyleModelInput.AppendLine(@"<Style id=""sn_blue-pushpin"">");
            sbStyleModelInput.AppendLine(@"<IconStyle>");
            sbStyleModelInput.AppendLine(@"<scale>1.1</scale>");
            sbStyleModelInput.AppendLine(@"<Icon>");
            sbStyleModelInput.AppendLine(@"<href>http://maps.google.com/mapfiles/kml/pushpin/blue-pushpin.png</href>");
            sbStyleModelInput.AppendLine(@"</Icon>");
            sbStyleModelInput.AppendLine(@"<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbStyleModelInput.AppendLine(@"</IconStyle>");
            sbStyleModelInput.AppendLine(@"<ListStyle>");
            sbStyleModelInput.AppendLine(@"</ListStyle>");
            sbStyleModelInput.AppendLine(@"</Style>");

            sbStyleModelInput.AppendLine(@"<Style id=""sh_blue-pushpin"">");
            sbStyleModelInput.AppendLine(@"<IconStyle>");
            sbStyleModelInput.AppendLine(@"<scale>1.3</scale>");
            sbStyleModelInput.AppendLine(@"<Icon>");
            sbStyleModelInput.AppendLine(@"<href>http://maps.google.com/mapfiles/kml/pushpin/blue-pushpin.png</href>");
            sbStyleModelInput.AppendLine(@"</Icon>");
            sbStyleModelInput.AppendLine(@"<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbStyleModelInput.AppendLine(@"</IconStyle>");
            sbStyleModelInput.AppendLine(@"<ListStyle>");
            sbStyleModelInput.AppendLine(@"</ListStyle>");
            sbStyleModelInput.AppendLine(@"</Style>");

            sbStyleModelInput.AppendLine(@"<StyleMap id=""msn_blue-pushpin"">");
            sbStyleModelInput.AppendLine(@"<Pair>");
            sbStyleModelInput.AppendLine(@"<key>normal</key>");
            sbStyleModelInput.AppendLine(@"<styleUrl>#sn_blue-pushpin</styleUrl>");
            sbStyleModelInput.AppendLine(@"</Pair>");
            sbStyleModelInput.AppendLine(@"<Pair>");
            sbStyleModelInput.AppendLine(@"<key>highlight</key>");
            sbStyleModelInput.AppendLine(@"<styleUrl>#sh_blue-pushpin</styleUrl>");
            sbStyleModelInput.AppendLine(@"</Pair>");
            sbStyleModelInput.AppendLine(@"</StyleMap>");
        }
        private void WriteKMLTop(string DocName, StringBuilder sbKML)
        {
            sbKML.AppendLine(@"<?xml version=""1.0"" encoding=""UTF-8""?>");
            sbKML.AppendLine(@"<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom"">");
            sbKML.AppendLine(@"<Document>");
            sbKML.AppendLine(string.Format(@"<name>{0}</name>", DocName));

            return;
        }
        //private void WriteSourcePlacemark(StringBuilder sbKMLModelInput, M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.SOURCES.SOURCE shd, KeyValuePair<string, M21_3FMService.FemEngineHD.HYDRODYNAMIC_MODULE.SOURCES.SOURCE> kvp, List<MikeSourceModel> mikeSourceModelList, M21_3FMService m21_3fmInput)
        //{
        //    string NotUsed = "";

        //    MikeSourceModel mikeSourceModelLocal = (from msl in mikeSourceModelList
        //                                            where msl.SourceNumberString == kvp.Key
        //                                            select msl).FirstOrDefault<MikeSourceModel>();

        //    if (mikeSourceModelLocal == null)
        //    {
        //        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeSource, TaskRunnerServiceRes.SourceNumberString, kvp.Key);
        //        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeSource, TaskRunnerServiceRes.SourceNumberString, kvp.Key);
        //        return;
        //    }

        //    sbKMLModelInput.Append("<Placemark>");

        //    if (shd.include == 1)
        //    {
        //        if (shd.Name.Substring(1, shd.Name.Length - 2).ToLower().StartsWith("river") || shd.Name.Substring(1, shd.Name.Length - 2).ToLower().StartsWith("rivière") || shd.Name.Substring(1, shd.Name.Length - 2).ToLower().StartsWith("brook") || shd.Name.Substring(1, shd.Name.Length - 2).ToLower().StartsWith("ruisseau"))
        //        {
        //            sbKMLModelInput.AppendLine(string.Format(@"<name>{0} (" + TaskRunnerServiceRes.UsedLowerCase + @")</name>", shd.Name.Substring(1, shd.Name.Length - 2)));
        //            sbKMLModelInput.AppendLine(@"<styleUrl>#msn_blue-pushpin</styleUrl>");
        //        }
        //        else
        //        {
        //            sbKMLModelInput.AppendLine(string.Format(@"<name>{0} (" + TaskRunnerServiceRes.UsedLowerCase + @")</name>", shd.Name.Substring(1, shd.Name.Length - 2)));
        //            sbKMLModelInput.AppendLine(@"<styleUrl>#msn_grn-pushpin</styleUrl>");
        //        }
        //    }
        //    else
        //    {
        //        sbKMLModelInput.AppendLine(string.Format(@"<name>{0} (" + TaskRunnerServiceRes.NotUsedLowerCase + @")</name>", shd.Name.Substring(1, shd.Name.Length - 2)));
        //        sbKMLModelInput.AppendLine(@"<styleUrl>#msn_red-pushpin</styleUrl>");
        //    }
        //    sbKMLModelInput.AppendLine(@"<visibility>0</visibility>");
        //    sbKMLModelInput.AppendLine(@"<description><![CDATA[");
        //    sbKMLModelInput.AppendLine(string.Format(@"<h2>{0}</h2>", shd.Name.Substring(1, shd.Name.Length - 2)));
        //    sbKMLModelInput.AppendLine(@"<h3>" + TaskRunnerServiceRes.Effluent + @"</h3>");
        //    sbKMLModelInput.AppendLine(@"<ul>");
        //    sbKMLModelInput.AppendLine(@"<li><b>" + TaskRunnerServiceRes.Coordinates + @"</b>");
        //    sbKMLModelInput.AppendLine(string.Format(@"&nbsp;&nbsp;&nbsp; {0:F5} &nbsp; {1:F5}</li>", shd.coordinates.y, shd.coordinates.x).Replace(",", "."));
        //    if ((bool)mikeSourceModelLocal.IsContinuous)
        //    {
        //        sbKMLModelInput.AppendLine(@"<li><b>" + TaskRunnerServiceRes.IsContinuous + "</b></li>");
        //        sbKMLModelInput.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.Flow + @":</b> {0:F6} (m3/s)  {1:F0} (m3/" + TaskRunnerServiceRes.DayLowerCase + @")</li>", shd.constant_value, shd.constant_value * 24 * 3600).ToString().Replace(",", "."));
        //        sbKMLModelInput.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.FCMPNPer100ML + @":</b> {0:F0}</li>", m21_3fmInput.femEngineHD.transport_module.sources.source[kvp.Key].component["COMPONENT_1"].constant_value).ToString().Replace(",", "."));
        //        sbKMLModelInput.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.Temperature + @":</b> {0:F0} " + TaskRunnerServiceRes.Celcius + @"</li>", m21_3fmInput.femEngineHD.hydrodynamic_module.temperature_salinity_module.sources.source[kvp.Key].temperature.constant_value).ToString().Replace(",", "."));
        //        sbKMLModelInput.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.Salinity + @":</b> {0:F0} " + TaskRunnerServiceRes.PSU + @"</li>", m21_3fmInput.femEngineHD.hydrodynamic_module.temperature_salinity_module.sources.source[kvp.Key].salinity.constant_value).ToString().Replace(",", "."));
        //    }
        //    else
        //    {
        //        sbKMLModelInput.AppendLine(@"<li><b>" + TaskRunnerServiceRes.IsNotContinuous + @"</b></li>");

        //        MikeSourceStartEndService mikeSourceStartEndService = new MikeSourceStartEndService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        //        List<MikeSourceStartEndModel> mikeSourceStartEndModelList = mikeSourceStartEndService.GetMikeSourceStartEndModelListWithMikeSourceIDDB(mikeSourceModelLocal.MikeSourceID);
        //        int CountMikeSourceStartEnd = 0;
        //        foreach (MikeSourceStartEndModel mssem in mikeSourceStartEndModelList)
        //        {
        //            CountMikeSourceStartEnd += 1;
        //            sbKMLModelInput.AppendLine(@"<ul>");
        //            sbKMLModelInput.AppendLine("<b>" + TaskRunnerServiceRes.Spill + @" " + CountMikeSourceStartEnd + "</b>");
        //            sbKMLModelInput.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.SpillStartTime + @":</b> {0:yyyy/MM/dd HH:mm:ss tt}</li>", mssem.StartDateAndTime_Local));
        //            sbKMLModelInput.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.SpillEndTime + @":</b> {0:yyyy/MM/dd HH:mm:ss tt}</li>", mssem.EndDateAndTime_Local));
        //            sbKMLModelInput.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.FlowStart + @":</b> {0:F6} (m3/s)  {1:F0} (m3/" + TaskRunnerServiceRes.DayLowerCase + @")</li>", mssem.SourceFlowStart_m3_day, mssem.SourceFlowStart_m3_day * 24 * 3600).ToString().Replace(",", "."));
        //            sbKMLModelInput.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.FlowEnd + @":</b> {0:F6} (m3/s)  {1:F0} (m3/" + TaskRunnerServiceRes.DayLowerCase + @")</li>", mssem.SourceFlowEnd_m3_day, mssem.SourceFlowEnd_m3_day * 24 * 3600).ToString().Replace(",", "."));
        //            sbKMLModelInput.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.FCMPNPer100MLStart + @":</b> {0:F0}</li>", mssem.SourcePollutionStart_MPN_100ml).ToString().Replace(",", "."));
        //            sbKMLModelInput.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.FCMPNPer100MLEnd + @":</b> {0:F0}</li>", mssem.SourcePollutionEnd_MPN_100ml).ToString().Replace(",", "."));
        //            sbKMLModelInput.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.TemperatureStart + @":</b> {0:F0} " + TaskRunnerServiceRes.Celcius + @"</li>", mssem.SourceTemperatureStart_C).ToString().Replace(",", "."));
        //            sbKMLModelInput.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.TempertureEnd + @":</b> {0:F0} " + TaskRunnerServiceRes.Celcius + @"</li>", mssem.SourceTemperatureEnd_C).ToString().Replace(",", "."));
        //            sbKMLModelInput.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.SalinityStart + @":</b> {0:F0} " + TaskRunnerServiceRes.PSU + @"</li>", mssem.SourceSalinityStart_PSU).ToString().Replace(",", "."));
        //            sbKMLModelInput.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.SalinityEnd + @":</b> {0:F0} " + TaskRunnerServiceRes.PSU + @"</li>", mssem.SourceSalinityEnd_PSU).ToString().Replace(",", "."));
        //            sbKMLModelInput.AppendLine(@"</ul>");
        //        }
        //    }
        //    sbKMLModelInput.AppendLine(@"</ul>");
        //    sbKMLModelInput.AppendLine(@"<iframe src=""about:"" width=""500"" height=""1"" />");
        //    sbKMLModelInput.AppendLine(@"]]></description>");

        //    sbKMLModelInput.AppendLine(@"<Point>");
        //    sbKMLModelInput.AppendLine(@"<coordinates>");
        //    sbKMLModelInput.AppendLine(shd.coordinates.x.ToString().Replace(",", ".") + @"," + shd.coordinates.y.ToString().Replace(",", ".") + ",0 ");
        //    sbKMLModelInput.AppendLine(@"</coordinates>");
        //    sbKMLModelInput.AppendLine(@"</Point>");
        //    sbKMLModelInput.AppendLine(@"</Placemark>");
        //}

        #endregion Functions private
    }
}
