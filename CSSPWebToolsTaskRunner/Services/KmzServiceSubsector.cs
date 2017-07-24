using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using CSSPWebToolsTaskRunner.Services;
using CSSPWebToolsTaskRunner.Services.Resources;
using System.Transactions;
using System.Text;
using CSSPWebToolsDBDLL.Models;
using CSSPWebToolsDBDLL.Services;
using System.Threading;
using System.Globalization;
using System.Text.RegularExpressions;
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;
using CSSPEnumsDLL.Services;

namespace CSSPWebToolsTaskRunner.Services
{
    public class KmzServiceSubsector
    {
        #region Variables
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        #endregion Variables

        #region Constructors
        public KmzServiceSubsector(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
        }
        #endregion Constructors

        #region Functions public
        public void Generate(FileInfo fi)
        {
            BaseEnumService baseEnumService = new BaseEnumService(_TaskRunnerBaseService._BWObj.appTaskModel.Language);
            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            PolSourceSiteService polSourceSiteService = new PolSourceSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            InfrastructureService infrastructureService = new InfrastructureService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            if (_TaskRunnerBaseService._BWObj.appTaskModel.Language == "fr")
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

            StringBuilder sbKML = new StringBuilder();

            sbKML.AppendLine(@"<?xml version=""1.0"" encoding=""UTF-8""?>");
            sbKML.AppendLine(@"<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom"">");
            sbKML.AppendLine(@"<Document>");
            sbKML.AppendLine(@"	<name>" + _TaskRunnerBaseService.generateDocParams.FileName + "</name>");
      
            sbKML.AppendLine(@"    <StyleMap id=""msn_ylw-pushpin"">");
            sbKML.AppendLine(@"		<Pair>");
            sbKML.AppendLine(@"			<key>normal</key>");
            sbKML.AppendLine(@"			<styleUrl>#sn_ylw-pushpin</styleUrl>");
            sbKML.AppendLine(@"		</Pair>");
            sbKML.AppendLine(@"		<Pair>");
            sbKML.AppendLine(@"			<key>highlight</key>");
            sbKML.AppendLine(@"			<styleUrl>#sh_ylw-pushpin</styleUrl>");
            sbKML.AppendLine(@"		</Pair>");
            sbKML.AppendLine(@"	</StyleMap>");
            sbKML.AppendLine(@"	<Style id=""sh_ylw-pushpin"">");
            sbKML.AppendLine(@"		<IconStyle>");
            sbKML.AppendLine(@"			<scale>1.2</scale>");
            sbKML.AppendLine(@"		</IconStyle>");
            sbKML.AppendLine(@"		<LineStyle>");
            sbKML.AppendLine(@"			<color>ff00ff00</color>");
            sbKML.AppendLine(@"			<width>1.5</width>");
            sbKML.AppendLine(@"		</LineStyle>");
            sbKML.AppendLine(@"		<PolyStyle>");
            sbKML.AppendLine(@"			<color>0000ff00</color>");
            sbKML.AppendLine(@"		</PolyStyle>");
            sbKML.AppendLine(@"	</Style>");
            sbKML.AppendLine(@"	<Style id=""sn_ylw-pushpin"">");
            sbKML.AppendLine(@"		<LineStyle>");
            sbKML.AppendLine(@"			<color>ff00ff00</color>");
            sbKML.AppendLine(@"			<width>1.5</width>");
            sbKML.AppendLine(@"		</LineStyle>");
            sbKML.AppendLine(@"		<PolyStyle>");
            sbKML.AppendLine(@"			<color>0000ff00</color>");
            sbKML.AppendLine(@"		</PolyStyle>");
            sbKML.AppendLine(@"	</Style>");

            sbKML.AppendLine(@"    <StyleMap id=""msn_grn-pushpin"">");
            sbKML.AppendLine(@"		<Pair>");
            sbKML.AppendLine(@"			<key>normal</key>");
            sbKML.AppendLine(@"			<styleUrl>#sn_grn-pushpin</styleUrl>");
            sbKML.AppendLine(@"		</Pair>");
            sbKML.AppendLine(@"		<Pair>");
            sbKML.AppendLine(@"			<key>highlight</key>");
            sbKML.AppendLine(@"			<styleUrl>#sh_grn-pushpin</styleUrl>");
            sbKML.AppendLine(@"		</Pair>");
            sbKML.AppendLine(@"	</StyleMap>");
            sbKML.AppendLine(@"	<Style id=""sh_grn-pushpin"">");
            sbKML.AppendLine(@"		<IconStyle>");
            sbKML.AppendLine(@"			<scale>1.2</scale>");
            sbKML.AppendLine(@"		</IconStyle>");
            sbKML.AppendLine(@"		<LineStyle>");
            sbKML.AppendLine(@"			<color>ff0000ff</color>");
            sbKML.AppendLine(@"			<width>1.5</width>");
            sbKML.AppendLine(@"		</LineStyle>");
            sbKML.AppendLine(@"		<PolyStyle>");
            sbKML.AppendLine(@"			<color>000000ff</color>");
            sbKML.AppendLine(@"		</PolyStyle>");
            sbKML.AppendLine(@"	</Style>");
            sbKML.AppendLine(@"	<Style id=""sn_grn-pushpin"">");
            sbKML.AppendLine(@"		<LineStyle>");
            sbKML.AppendLine(@"			<color>ff0000ff</color>");
            sbKML.AppendLine(@"			<width>1.5</width>");
            sbKML.AppendLine(@"		</LineStyle>");
            sbKML.AppendLine(@"		<PolyStyle>");
            sbKML.AppendLine(@"			<color>000000ff</color>");
            sbKML.AppendLine(@"		</PolyStyle>");
            sbKML.AppendLine(@"	</Style>");

            TVItemModel tvItemModelSubsector = tvItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            List<MapInfoPointModel> mapInfoPointModelList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.Subsector, MapInfoDrawTypeEnum.Point);

            sbKML.AppendLine(@"	<Folder>");
            sbKML.AppendLine(@"	<name>Subsector</name>");
            sbKML.AppendLine(@"	<visibility>0</visibility>");

            // Doing Point
            sbKML.AppendLine(@"	<Placemark>");
            sbKML.AppendLine(@"	<name>" + tvItemModelSubsector.TVText + "</name>");
            sbKML.AppendLine(@"	<visibility>0</visibility>");
            sbKML.AppendLine(@"<styleUrl>#msn_ylw-pushpin</styleUrl>");
            //sbKMZ.AppendLine(@"	<description>");
            //sbKMZ.AppendLine(@"<![CDATA[");
            //sbKMZ.AppendLine(@"<a href=""" + _TaskRunnerBaseService.GetUrlFromTVItem(tvItemModelRoot) + @""">" + tvItemModelRoot.TVText + "</a>");
            //sbKMZ.AppendLine(@"]]>");
            //sbKMZ.AppendLine(@"	</description>");
            sbKML.AppendLine(@"	<Point>");
            sbKML.AppendLine(@"		<coordinates>" + mapInfoPointModelList[0].Lng + "," + mapInfoPointModelList[0].Lat + ",0</coordinates>");
            sbKML.AppendLine(@"	</Point>");
            sbKML.AppendLine(@"	</Placemark>");

            // Doing Polygon
            sbKML.AppendLine(@"	<Placemark>");
            sbKML.AppendLine(@"		<name>" + tvItemModelSubsector.TVText + " (poly)</name>");
            sbKML.AppendLine(@"	    <visibility>0</visibility>");
            sbKML.AppendLine(@"<styleUrl>#msn_ylw-pushpin</styleUrl>");
            //sbKMZ.AppendLine(@"	<description>");
            //sbKMZ.AppendLine(@"<![CDATA[");
            //sbKMZ.AppendLine(@"<a href=""" + _TaskRunnerBaseService.GetUrlFromTVItem(tvItemModelCountry) + @""">" + tvItemModelCountry.TVText + "</a>");
            //sbKMZ.AppendLine(@"]]>");
            //sbKMZ.AppendLine(@"	</description>");
            sbKML.AppendLine(@"		<Polygon>");
            sbKML.AppendLine(@"			<outerBoundaryIs>");
            sbKML.AppendLine(@"				<LinearRing>");
            sbKML.AppendLine(@"					<coordinates>");

            mapInfoPointModelList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.Subsector, MapInfoDrawTypeEnum.Polygon);
            foreach (MapInfoPointModel mapInfoPointModel in mapInfoPointModelList)
            {
                sbKML.AppendLine(mapInfoPointModel.Lng + "," + mapInfoPointModel.Lat + ",0 ");

            }

            sbKML.AppendLine(@"					</coordinates>");
            sbKML.AppendLine(@"				</LinearRing>");
            sbKML.AppendLine(@"			</outerBoundaryIs>");
            sbKML.AppendLine(@"		</Polygon>");
            sbKML.AppendLine(@"	</Placemark>");

            // Doing Municipalities
            sbKML.AppendLine(@" <Folder>");
            sbKML.AppendLine(@"	<name>Municipalities</name>");
            sbKML.AppendLine(@"	<visibility>0</visibility>");
            List<TVItemModel> tvItemModelMunicipalityList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.Municipality);

            foreach (TVItemModel tvItemModelMunicipality in tvItemModelMunicipalityList)
            {
                sbKML.AppendLine(@" <Folder>");
                sbKML.AppendLine(@"	<name>" + tvItemModelMunicipality.TVText + "</name>");
                sbKML.AppendLine(@"	<visibility>0</visibility>");
                // Doing point
                sbKML.AppendLine(@"	<Placemark>");
                sbKML.AppendLine(@"	<name>" + tvItemModelMunicipality.TVText + " ( Point)</name>");
                sbKML.AppendLine(@"	<visibility>0</visibility>");
                sbKML.AppendLine(@"<styleUrl>#msn_ylw-pushpin</styleUrl>");

                mapInfoPointModelList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelMunicipality.TVItemID, TVTypeEnum.Municipality, MapInfoDrawTypeEnum.Point);

                sbKML.AppendLine(@"		<Point>");
                sbKML.AppendLine(@"			<coordinates>" + mapInfoPointModelList[0].Lng + "," + mapInfoPointModelList[0].Lat + ",0</coordinates>");
                sbKML.AppendLine(@"		</Point>");
                sbKML.AppendLine(@"	</Placemark>");

                List<TVItemModel> tvItemModelInfrastructureList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelMunicipality.TVItemID, TVTypeEnum.Infrastructure);

                List<InfrastructureModel> infrastructureModelList = new List<InfrastructureModel>();

                foreach (TVItemModel tvItemModelInfrastructure in tvItemModelInfrastructureList)
                {
                    infrastructureModelList.Add(infrastructureService.GetInfrastructureModelWithInfrastructureTVItemIDDB(tvItemModelInfrastructure.TVItemID));
                }

                // Doing WWTP
                foreach (InfrastructureModel infrastructureModel in infrastructureModelList.Where(c => c.InfrastructureType == InfrastructureTypeEnum.WWTP).ToList())
                {
                    sbKML.AppendLine(@"	<Placemark>");
                    sbKML.AppendLine(@"	<name>" + tvItemModelInfrastructureList.Where(c => c.TVItemID == infrastructureModel.InfrastructureTVItemID).Select(c => c.TVText).FirstOrDefault() + "</name>");
                    sbKML.AppendLine(@"	<visibility>0</visibility>");
                    sbKML.AppendLine(@"<styleUrl>#msn_ylw-pushpin</styleUrl>");
                    
                    sbKML.AppendLine(@"	<description>");
                    sbKML.AppendLine(@"<![CDATA[");
                    sbKML.AppendLine(@"<pre>");
                    sbKML.AppendLine(@"Alarm System Type: " + baseEnumService.GetEnumText_AlarmSystemTypeEnum(infrastructureModel.AlarmSystemType) + "\r\n");
                    sbKML.AppendLine(@"Can overflow: " + infrastructureModel.CanOverflow.ToString() + "\r\n");
                    sbKML.AppendLine(@"Category: " + infrastructureModel.InfrastructureCategory + "\r\n");
                    sbKML.AppendLine(@"Collection System Type: " + baseEnumService.GetEnumText_CollectionSystemTypeEnum(infrastructureModel.CollectionSystemType) + "\r\n");
                    sbKML.AppendLine(@"Comments: " + infrastructureModel.Comment + "\r\n");
                    sbKML.AppendLine(@"DesignFlow (m3/day): " + infrastructureModel.DesignFlow_m3_day + "\r\n");
                    sbKML.AppendLine(@"Disinfection Type: " + baseEnumService.GetEnumText_DisinfectionTypeEnum(infrastructureModel.DisinfectionType) + "\r\n");
                    sbKML.AppendLine(@"Infrastructure Type: " + baseEnumService.GetEnumText_InfrastructureTypeEnum(infrastructureModel.InfrastructureType) + "\r\n");
                    sbKML.AppendLine(@"Average Flow (m3/day): " + infrastructureModel.AverageFlow_m3_day + "\r\n");
                    sbKML.AppendLine(@"AverageFlow_m3_day: " + infrastructureModel.PeakFlow_m3_day + "\r\n");
                    sbKML.AppendLine(@"Percent Flow Of Total (%): " + infrastructureModel.PercFlowOfTotal + "\r\n");
                    sbKML.AppendLine(@"Population Served: " + infrastructureModel.PopServed + "\r\n");
                    sbKML.AppendLine(@"Time Zone: " + infrastructureModel.TimeOffset_hour + "\r\n");
                    sbKML.AppendLine(@"Treatment Type: " + baseEnumService.GetEnumText_TreatmentTypeEnum(infrastructureModel.TreatmentType) + "\r\n");

                    List<MapInfoPointModel> mapInfoPointModelInfrastructureList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(infrastructureModel.InfrastructureTVItemID, TVTypeEnum.WasteWaterTreatmentPlant, MapInfoDrawTypeEnum.Point);

                    if (mapInfoPointModelInfrastructureList.Count > 0)
                    {
                        sbKML.AppendLine(@"Latitude Longitude: " + mapInfoPointModelInfrastructureList[0].Lat + " " + mapInfoPointModelInfrastructureList[0].Lng + "\r\n");
                    }
                    else
                    {
                        sbKML.AppendLine("Latitude Longitude: \r\n");
                    }

                    List<MapInfoPointModel> mapInfoPointModelInfrastructureOutfallList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(infrastructureModel.InfrastructureTVItemID, TVTypeEnum.Outfall, MapInfoDrawTypeEnum.Point);

                    if (mapInfoPointModelInfrastructureOutfallList.Count > 0)
                    {
                        sbKML.AppendLine(@"Outfall: Latitude Longitude: " + mapInfoPointModelInfrastructureOutfallList[0].Lat + " " + mapInfoPointModelInfrastructureOutfallList[0].Lng + "\r\n");
                    }
                    else
                    {
                        sbKML.AppendLine("Outfall: Latitude Longitude: \r\n");
                    }
                    sbKML.AppendLine("\r\n\r\n");
                    sbKML.AppendLine("Outfall Information\r\n\r\n");
                    sbKML.AppendLine(@"Average Depth (m): " + infrastructureModel.AverageDepth_m + "\r\n");
                    sbKML.AppendLine(@"Decay Rate (/day): " + infrastructureModel.DecayRate_per_day + "\r\n");
                    sbKML.AppendLine(@"Distance From Shore (m): " + infrastructureModel.DistanceFromShore_m + "\r\n");
                    sbKML.AppendLine(@"Far Field Velocity (m/s): " + infrastructureModel.FarFieldVelocity_m_s + "\r\n");
                    sbKML.AppendLine(@"Horizontal Angle (deg): " + infrastructureModel.HorizontalAngle_deg + "\r\n");
                    sbKML.AppendLine(@"Near Field Velocity (m/s): " + infrastructureModel.NearFieldVelocity_m_s + "\r\n");
                    sbKML.AppendLine(@"Number Of Ports: " + infrastructureModel.NumberOfPorts + "\r\n");
                    sbKML.AppendLine(@"Port Diameter (m): " + infrastructureModel.PortDiameter_m + "\r\n");
                    sbKML.AppendLine(@"Port Elevation (m): " + infrastructureModel.PortElevation_m + "\r\n");
                    sbKML.AppendLine(@"Port Spacing (m): " + infrastructureModel.PortSpacing_m + "\r\n");
                    sbKML.AppendLine(@"Receiving Water Concentration (FC /100 ml): " + infrastructureModel.ReceivingWater_MPN_per_100ml + "\r\n");
                    sbKML.AppendLine(@"Receiving Water Salinity (PSU): " + infrastructureModel.ReceivingWaterSalinity_PSU + "\r\n");
                    sbKML.AppendLine(@"Receiving Water Temperature (ºC): " + infrastructureModel.ReceivingWaterTemperature_C + "\r\n");
                    sbKML.AppendLine(@"Vertical Angle (deg): " + infrastructureModel.VerticalAngle_deg + "\r\n");
                    sbKML.AppendLine(@"</pre>");
                    sbKML.AppendLine(@"]]>");
                    sbKML.AppendLine(@"	</description>");

                    sbKML.AppendLine(@"		<Point>");
                    sbKML.AppendLine(@"			<coordinates>" + mapInfoPointModelInfrastructureList[0].Lng + "," + mapInfoPointModelInfrastructureList[0].Lat + ",0</coordinates>");
                    sbKML.AppendLine(@"		</Point>");
                    sbKML.AppendLine(@"	</Placemark>");

                    sbKML.AppendLine(@"	<Placemark>");
                    sbKML.AppendLine(@"	<name>Outfall " + tvItemModelInfrastructureList.Where(c => c.TVItemID == infrastructureModel.InfrastructureTVItemID).Select(c => c.TVText).FirstOrDefault() + "</name>");
                    sbKML.AppendLine(@"	<visibility>0</visibility>");
                    sbKML.AppendLine(@"<styleUrl>#msn_grn-pushpin</styleUrl>");

                    sbKML.AppendLine(@"		<Point>");
                    sbKML.AppendLine(@"			<coordinates>" + mapInfoPointModelInfrastructureOutfallList[0].Lng + "," + mapInfoPointModelInfrastructureOutfallList[0].Lat + ",0</coordinates>");
                    sbKML.AppendLine(@"		</Point>");
                    sbKML.AppendLine(@"	</Placemark>");

                }

                // Doing LS
                foreach (InfrastructureModel infrastructureModel in infrastructureModelList.Where(c => c.InfrastructureType == InfrastructureTypeEnum.LiftStation).ToList())
                {
                    sbKML.AppendLine(@"	<Placemark>");
                    sbKML.AppendLine(@"	<name>" + tvItemModelInfrastructureList.Where(c => c.TVItemID == infrastructureModel.InfrastructureTVItemID).Select(c => c.TVText).FirstOrDefault() + "</name>");
                    sbKML.AppendLine(@"	<visibility>0</visibility>");
                    sbKML.AppendLine(@"<styleUrl>#msn_ylw-pushpin</styleUrl>");

                    sbKML.AppendLine(@"	<description>");
                    sbKML.AppendLine(@"<![CDATA[");
                    sbKML.AppendLine(@"<pre>");
                    sbKML.AppendLine(@"Alarm System Type: " + baseEnumService.GetEnumText_AlarmSystemTypeEnum(infrastructureModel.AlarmSystemType) + "\r\n");
                    sbKML.AppendLine(@"Can overflow: " + infrastructureModel.CanOverflow.ToString() + "\r\n");
                    sbKML.AppendLine(@"Category: " + infrastructureModel.InfrastructureCategory + "\r\n");
                    sbKML.AppendLine(@"Collection System Type: " + baseEnumService.GetEnumText_CollectionSystemTypeEnum(infrastructureModel.CollectionSystemType) + "\r\n");
                    sbKML.AppendLine(@"Comments: " + infrastructureModel.Comment + "\r\n");
                    sbKML.AppendLine(@"Infrastructure Type: " + baseEnumService.GetEnumText_InfrastructureTypeEnum(infrastructureModel.InfrastructureType) + "\r\n");
                    sbKML.AppendLine(@"Percent Flow Of Total (%): " + infrastructureModel.PercFlowOfTotal + "\r\n");

                    List<MapInfoPointModel> mapInfoPointModelInfrastructureList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(infrastructureModel.InfrastructureTVItemID, TVTypeEnum.LiftStation, MapInfoDrawTypeEnum.Point);

                    if (mapInfoPointModelInfrastructureList.Count > 0)
                    {
                        sbKML.AppendLine(@"Latitude Longitude: " + mapInfoPointModelInfrastructureList[0].Lat + " " + mapInfoPointModelInfrastructureList[0].Lng + "\r\n");
                    }
                    else
                    {
                        sbKML.AppendLine("Latitude Longitude: \r\n");
                    }

                    List<MapInfoPointModel> mapInfoPointModelInfrastructureOutfallList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(infrastructureModel.InfrastructureTVItemID, TVTypeEnum.Outfall, MapInfoDrawTypeEnum.Point);

                    if (mapInfoPointModelInfrastructureOutfallList.Count > 0)
                    {
                        sbKML.AppendLine(@"Outfall: Latitude Longitude: " + mapInfoPointModelInfrastructureOutfallList[0].Lat + " " + mapInfoPointModelInfrastructureOutfallList[0].Lng + "\r\n");
                    }
                    else
                    {
                        sbKML.AppendLine("Outfall: Latitude Longitude: \r\n");
                    }
                    sbKML.AppendLine("\r\n\r\n");
                    sbKML.AppendLine("Outfall Information\r\n\r\n");
                    sbKML.AppendLine(@"Average Depth (m): " + infrastructureModel.AverageDepth_m + "\r\n");
                    sbKML.AppendLine(@"Decay Rate (/day): " + infrastructureModel.DecayRate_per_day + "\r\n");
                    sbKML.AppendLine(@"Distance From Shore (m): " + infrastructureModel.DistanceFromShore_m + "\r\n");
                    sbKML.AppendLine(@"Far Field Velocity (m/s): " + infrastructureModel.FarFieldVelocity_m_s + "\r\n");
                    sbKML.AppendLine(@"Horizontal Angle (deg): " + infrastructureModel.HorizontalAngle_deg + "\r\n");
                    sbKML.AppendLine(@"Near Field Velocity (m/s): " + infrastructureModel.NearFieldVelocity_m_s + "\r\n");
                    sbKML.AppendLine(@"Number Of Ports: " + infrastructureModel.NumberOfPorts + "\r\n");
                    sbKML.AppendLine(@"Port Diameter (m): " + infrastructureModel.PortDiameter_m + "\r\n");
                    sbKML.AppendLine(@"Port Elevation (m): " + infrastructureModel.PortElevation_m + "\r\n");
                    sbKML.AppendLine(@"Port Spacing (m): " + infrastructureModel.PortSpacing_m + "\r\n");
                    sbKML.AppendLine(@"Receiving Water Concentration (FC /100 ml): " + infrastructureModel.ReceivingWater_MPN_per_100ml + "\r\n");
                    sbKML.AppendLine(@"Receiving Water Salinity (PSU): " + infrastructureModel.ReceivingWaterSalinity_PSU + "\r\n");
                    sbKML.AppendLine(@"Receiving Water Temperature (ºC): " + infrastructureModel.ReceivingWaterTemperature_C + "\r\n");
                    sbKML.AppendLine(@"Vertical Angle (deg): " + infrastructureModel.VerticalAngle_deg + "\r\n");
                    sbKML.AppendLine(@"</pre>");
                    sbKML.AppendLine(@"]]>");
                    sbKML.AppendLine(@"	</description>");

                    sbKML.AppendLine(@"		<Point>");
                    sbKML.AppendLine(@"			<coordinates>" + mapInfoPointModelInfrastructureList[0].Lng + "," + mapInfoPointModelInfrastructureList[0].Lat + ",0</coordinates>");
                    sbKML.AppendLine(@"		</Point>");
                    sbKML.AppendLine(@"	</Placemark>");

                    sbKML.AppendLine(@"	<Placemark>");
                    sbKML.AppendLine(@"	<name>Outfall " + tvItemModelInfrastructureList.Where(c => c.TVItemID == infrastructureModel.InfrastructureTVItemID).Select(c => c.TVText).FirstOrDefault() + "</name>");
                    sbKML.AppendLine(@"	<visibility>0</visibility>");
                    sbKML.AppendLine(@"<styleUrl>#msn_grn-pushpin</styleUrl>");

                    sbKML.AppendLine(@"		<Point>");
                    sbKML.AppendLine(@"			<coordinates>" + mapInfoPointModelInfrastructureOutfallList[0].Lng + "," + mapInfoPointModelInfrastructureOutfallList[0].Lat + ",0</coordinates>");
                    sbKML.AppendLine(@"		</Point>");
                    sbKML.AppendLine(@"	</Placemark>");
                }

                // Doing Line Overflow
                foreach (InfrastructureModel infrastructureModel in infrastructureModelList.Where(c => c.InfrastructureType == InfrastructureTypeEnum.LineOverflow).ToList())
                {
                    sbKML.AppendLine(@"	<Placemark>");
                    sbKML.AppendLine(@"	<name>" + tvItemModelInfrastructureList.Where(c => c.TVItemID == infrastructureModel.InfrastructureTVItemID).Select(c => c.TVText).FirstOrDefault() + "</name>");
                    sbKML.AppendLine(@"	<visibility>0</visibility>");
                    sbKML.AppendLine(@"<styleUrl>#msn_ylw-pushpin</styleUrl>");

                    sbKML.AppendLine(@"	<description>");
                    sbKML.AppendLine(@"<![CDATA[");
                    sbKML.AppendLine(@"<pre>");
                    sbKML.AppendLine(@"Alarm System Type: " + baseEnumService.GetEnumText_AlarmSystemTypeEnum(infrastructureModel.AlarmSystemType) + "\r\n");
                    sbKML.AppendLine(@"Can overflow: " + infrastructureModel.CanOverflow.ToString() + "\r\n");
                    sbKML.AppendLine(@"Category: " + infrastructureModel.InfrastructureCategory + "\r\n");
                    sbKML.AppendLine(@"Collection System Type: " + baseEnumService.GetEnumText_CollectionSystemTypeEnum(infrastructureModel.CollectionSystemType) + "\r\n");
                    sbKML.AppendLine(@"Comments: " + infrastructureModel.Comment + "\r\n");
                    sbKML.AppendLine(@"Infrastructure Type: " + baseEnumService.GetEnumText_InfrastructureTypeEnum(infrastructureModel.InfrastructureType) + "\r\n");
                    sbKML.AppendLine(@"Percent Flow Of Total (%): " + infrastructureModel.PercFlowOfTotal + "\r\n");

                    List<MapInfoPointModel> mapInfoPointModelInfrastructureList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(infrastructureModel.InfrastructureTVItemID, TVTypeEnum.LineOverflow, MapInfoDrawTypeEnum.Point);

                    if (mapInfoPointModelInfrastructureList.Count > 0)
                    {
                        sbKML.AppendLine(@"Latitude Longitude: " + mapInfoPointModelInfrastructureList[0].Lat + " " + mapInfoPointModelInfrastructureList[0].Lng + "\r\n");
                    }
                    else
                    {
                        sbKML.AppendLine("Latitude Longitude: \r\n");
                    }

                    List<MapInfoPointModel> mapInfoPointModelInfrastructureOutfallList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(infrastructureModel.InfrastructureTVItemID, TVTypeEnum.Outfall, MapInfoDrawTypeEnum.Point);

                    if (mapInfoPointModelInfrastructureOutfallList.Count > 0)
                    {
                        sbKML.AppendLine(@"Outfall: Latitude Longitude: " + mapInfoPointModelInfrastructureOutfallList[0].Lat + " " + mapInfoPointModelInfrastructureOutfallList[0].Lng + "\r\n");
                    }
                    else
                    {
                        sbKML.AppendLine("Outfall: Latitude Longitude: \r\n");
                    }
                    sbKML.AppendLine("\r\n\r\n");
                    sbKML.AppendLine("Outfall Information\r\n\r\n");
                    sbKML.AppendLine(@"Average Depth (m): " + infrastructureModel.AverageDepth_m + "\r\n");
                    sbKML.AppendLine(@"Decay Rate (/day): " + infrastructureModel.DecayRate_per_day + "\r\n");
                    sbKML.AppendLine(@"Distance From Shore (m): " + infrastructureModel.DistanceFromShore_m + "\r\n");
                    sbKML.AppendLine(@"Far Field Velocity (m/s): " + infrastructureModel.FarFieldVelocity_m_s + "\r\n");
                    sbKML.AppendLine(@"Horizontal Angle (deg): " + infrastructureModel.HorizontalAngle_deg + "\r\n");
                    sbKML.AppendLine(@"Near Field Velocity (m/s): " + infrastructureModel.NearFieldVelocity_m_s + "\r\n");
                    sbKML.AppendLine(@"Number Of Ports: " + infrastructureModel.NumberOfPorts + "\r\n");
                    sbKML.AppendLine(@"Port Diameter (m): " + infrastructureModel.PortDiameter_m + "\r\n");
                    sbKML.AppendLine(@"Port Elevation (m): " + infrastructureModel.PortElevation_m + "\r\n");
                    sbKML.AppendLine(@"Port Spacing (m): " + infrastructureModel.PortSpacing_m + "\r\n");
                    sbKML.AppendLine(@"Receiving Water Concentration (FC /100 ml): " + infrastructureModel.ReceivingWater_MPN_per_100ml + "\r\n");
                    sbKML.AppendLine(@"Receiving Water Salinity (PSU): " + infrastructureModel.ReceivingWaterSalinity_PSU + "\r\n");
                    sbKML.AppendLine(@"Receiving Water Temperature (ºC): " + infrastructureModel.ReceivingWaterTemperature_C + "\r\n");
                    sbKML.AppendLine(@"Vertical Angle (deg): " + infrastructureModel.VerticalAngle_deg + "\r\n");
                    sbKML.AppendLine(@"</pre>");
                    sbKML.AppendLine(@"]]>");
                    sbKML.AppendLine(@"	</description>");

                    sbKML.AppendLine(@"		<Point>");
                    sbKML.AppendLine(@"			<coordinates>" + mapInfoPointModelInfrastructureList[0].Lng + "," + mapInfoPointModelInfrastructureList[0].Lat + ",0</coordinates>");
                    sbKML.AppendLine(@"		</Point>");
                    sbKML.AppendLine(@"	</Placemark>");

                    sbKML.AppendLine(@"	<Placemark>");
                    sbKML.AppendLine(@"	<name>Outfall " + tvItemModelInfrastructureList.Where(c => c.TVItemID == infrastructureModel.InfrastructureTVItemID).Select(c => c.TVText).FirstOrDefault() + "</name>");
                    sbKML.AppendLine(@"	<visibility>0</visibility>");
                    sbKML.AppendLine(@"<styleUrl>#msn_grn-pushpin</styleUrl>");

                    sbKML.AppendLine(@"		<Point>");
                    sbKML.AppendLine(@"			<coordinates>" + mapInfoPointModelInfrastructureOutfallList[0].Lng + "," + mapInfoPointModelInfrastructureOutfallList[0].Lat + ",0</coordinates>");
                    sbKML.AppendLine(@"		</Point>");
                    sbKML.AppendLine(@"	</Placemark>");
                }
                sbKML.AppendLine(@" </Folder>");
            }
            sbKML.AppendLine(@" </Folder>");

            // Doing Short Pollution Source Site
            sbKML.AppendLine(@" <Folder>");
            sbKML.AppendLine(@"	<name>Short Pollution Source Sites</name>");
            sbKML.AppendLine(@"	<visibility>0</visibility>");
            List<PolSourceSiteModel> polSourceSiteModelList = polSourceSiteService.GetPolSourceSiteModelListWithSubsectorTVItemIDDB(tvItemModelSubsector.TVItemID);

            foreach (PolSourceSiteModel polSourceSiteModel in polSourceSiteModelList.OrderBy(c => c.Site).ToList())
            {
                // Doing point
                sbKML.AppendLine(@"	<Placemark>");
                sbKML.AppendLine(@"	<name>" + polSourceSiteModel.Site.ToString() + "</name>");
                sbKML.AppendLine(@"	<visibility>0</visibility>");
                sbKML.AppendLine(@"<styleUrl>#msn_ylw-pushpin</styleUrl>");
                sbKML.AppendLine(@"	<description>");
                sbKML.AppendLine(@"<![CDATA[");
                sbKML.AppendLine(@"<pre>");

                PolSourceObservationModel polSourceObservationModel = polSourceSiteService._PolSourceObservationService.GetPolSourceObservationModelLatestWithPolSourceSiteTVItemIDDB(polSourceSiteModel.PolSourceSiteTVItemID);
                if (polSourceObservationModel != null)
                {
                    List<PolSourceObservationIssueModel> polSourceObservationIssueModelList = polSourceSiteService._PolSourceObservationService._PolSourceObservationIssueService.GetPolSourceObservationIssueModelListWithPolSourceObservationIDDB(polSourceObservationModel.PolSourceObservationID);

                    string SelectedObservation = "Selected: \r\n";
                    foreach (PolSourceObservationIssueModel polSourceObservationIssueModel in polSourceObservationIssueModelList)
                    {
                        foreach (PolSourceObsInfoEnum polSourceObsInfo in polSourceObservationIssueModel.PolSourceObsInfoList)
                        {
                            SelectedObservation += baseEnumService.GetEnumText_PolSourceObsInfoReportEnum(polSourceObsInfo);
                        }
                        SelectedObservation += "\r\n\r\n";
                    }

                    sbKML.AppendLine("Written: \r\n" + (string.IsNullOrWhiteSpace(polSourceObservationModel.Observation_ToBeDeleted) ? "" : polSourceObservationModel.Observation_ToBeDeleted.ToString()) + "\r\n\r\n" + SelectedObservation);
                }
                else
                {
                    string SelectedObservation = "Selected: \r\n";
                    sbKML.AppendLine("Written: \r\n\r\n" + SelectedObservation);
                }

                sbKML.AppendLine(@"</pre>");
                sbKML.AppendLine(@"]]>");
                sbKML.AppendLine(@"	</description>");

                mapInfoPointModelList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(polSourceSiteModel.PolSourceSiteTVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);

                sbKML.AppendLine(@"		<Point>");
                sbKML.AppendLine(@"			<coordinates>" + mapInfoPointModelList[0].Lng + "," + mapInfoPointModelList[0].Lat + ",0</coordinates>");
                sbKML.AppendLine(@"		</Point>");
                sbKML.AppendLine(@"	</Placemark>");

            }
            sbKML.AppendLine(@" </Folder>");

            // Doing Long Pollution Source Site
            sbKML.AppendLine(@" <Folder>");
            sbKML.AppendLine(@"	<name>Long Pollution Source Sites</name>");
            sbKML.AppendLine(@"	<visibility>0</visibility>");
            polSourceSiteModelList = polSourceSiteService.GetPolSourceSiteModelListWithSubsectorTVItemIDDB(tvItemModelSubsector.TVItemID);

            foreach (PolSourceSiteModel polSourceSiteModel in polSourceSiteModelList.OrderBy(c => c.Site).ToList())
            {
                TVItemModel tvItemModelPolSourceSite = tvItemService.GetTVItemModelWithTVItemIDDB(polSourceSiteModel.PolSourceSiteTVItemID);

                // Doing point
                sbKML.AppendLine(@"	<Placemark>");
                sbKML.AppendLine(@"	<name>" + tvItemModelPolSourceSite.TVText + "</name>");
                sbKML.AppendLine(@"	<visibility>0</visibility>");
                sbKML.AppendLine(@"<styleUrl>#msn_ylw-pushpin</styleUrl>");
                sbKML.AppendLine(@"	<description>");
                sbKML.AppendLine(@"<![CDATA[");
                sbKML.AppendLine(@"<pre>");

                PolSourceObservationModel polSourceObservationModel = polSourceSiteService._PolSourceObservationService.GetPolSourceObservationModelLatestWithPolSourceSiteTVItemIDDB(polSourceSiteModel.PolSourceSiteTVItemID);
                if (polSourceObservationModel != null)
                {
                    List<PolSourceObservationIssueModel> polSourceObservationIssueModelList = polSourceSiteService._PolSourceObservationService._PolSourceObservationIssueService.GetPolSourceObservationIssueModelListWithPolSourceObservationIDDB(polSourceObservationModel.PolSourceObservationID);

                    string SelectedObservation = "Selected: \r\n";
                    foreach (PolSourceObservationIssueModel polSourceObservationIssueModel in polSourceObservationIssueModelList)
                    {
                        foreach (PolSourceObsInfoEnum polSourceObsInfo in polSourceObservationIssueModel.PolSourceObsInfoList)
                        {
                            SelectedObservation += baseEnumService.GetEnumText_PolSourceObsInfoReportEnum(polSourceObsInfo);
                        }
                        SelectedObservation += "\r\n\r\n";
                    }

                    sbKML.AppendLine("Written: \r\n" + (string.IsNullOrWhiteSpace(polSourceObservationModel.Observation_ToBeDeleted) ? "" : polSourceObservationModel.Observation_ToBeDeleted.ToString()) + "\r\n\r\n" + SelectedObservation);
                }
                else
                {
                    string SelectedObservation = "Selected: \r\n";
                    sbKML.AppendLine("Written: \r\n\r\n" + SelectedObservation);
                }

                sbKML.AppendLine(@"</pre>");
                sbKML.AppendLine(@"]]>");
                sbKML.AppendLine(@"	</description>");

                mapInfoPointModelList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelPolSourceSite.TVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);

                sbKML.AppendLine(@"		<Point>");
                sbKML.AppendLine(@"			<coordinates>" + mapInfoPointModelList[0].Lng + "," + mapInfoPointModelList[0].Lat + ",0</coordinates>");
                sbKML.AppendLine(@"		</Point>");
                sbKML.AppendLine(@"	</Placemark>");

            }
            sbKML.AppendLine(@" </Folder>");


            // Doing MWQM Site
            sbKML.AppendLine(@" <Folder>");
            sbKML.AppendLine(@"	<name>MWQM Sites</name>");
            sbKML.AppendLine(@"	<visibility>0</visibility>");
            List<TVItemModel> tvItemModelMWQMSiteList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.MWQMSite);

            foreach (TVItemModel tvItemModelMWQMSite in tvItemModelMWQMSiteList)
            {
                // Doing point
                sbKML.AppendLine(@"	<Placemark>");
                sbKML.AppendLine(@"	<name>" + tvItemModelMWQMSite.TVText + "</name>");
                sbKML.AppendLine(@"	<visibility>0</visibility>");
                sbKML.AppendLine(@"<styleUrl>#msn_ylw-pushpin</styleUrl>");

                mapInfoPointModelList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelMWQMSite.TVItemID, TVTypeEnum.MWQMSite, MapInfoDrawTypeEnum.Point);

                sbKML.AppendLine(@"		<Point>");
                sbKML.AppendLine(@"			<coordinates>" + mapInfoPointModelList[0].Lng + "," + mapInfoPointModelList[0].Lat + ",0</coordinates>");
                sbKML.AppendLine(@"		</Point>");
                sbKML.AppendLine(@"	</Placemark>");

            }
            sbKML.AppendLine(@" </Folder>");

            sbKML.AppendLine(@"	</Folder>");

            sbKML.AppendLine(@"</Document>");
            sbKML.AppendLine(@"</kml>");

            StreamWriter sw = fi.CreateText();
            sw.Write(sbKML.ToString());
            sw.Close();
        }
        #endregion Functions public
    }
}
