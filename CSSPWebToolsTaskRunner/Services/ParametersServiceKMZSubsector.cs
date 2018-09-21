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
using DHI.Generic.MikeZero.DFS.dfsu;
using CSSPEnumsDLL.Services;
using System.Threading;
using System.Globalization;
using CSSPWebToolsDBDLL;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        private bool GenerateKMZSubsector()
        {
            switch (reportTypeModel.UniqueCode)
            {
                case "SubsectorInformationKMZ":
                    {
                        if (!GenerateSubsectorInformationKMZ())
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
                    break;
            }
            return true;
        }
        private bool GenerateSubsectorInformationKMZ()
        {
            string NotUsed = "";
            bool ErrorInDoc = false;

            if (!WriteKMLTop(this.sb))
            {
                ErrorInDoc = true;
            }

            int TVItemID = _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID;

            if (TVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.Parameter_NotFound, "TVItemID");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Parameter_NotFound", "TVItemID");
                return false;
            }

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


            sb.AppendLine(@"	<name>" + fi.Name + "</name>");

            sb.AppendLine(@"    <StyleMap id=""msn_ylw-pushpin"">");
            sb.AppendLine(@"		<Pair>");
            sb.AppendLine(@"			<key>normal</key>");
            sb.AppendLine(@"			<styleUrl>#sn_ylw-pushpin</styleUrl>");
            sb.AppendLine(@"		</Pair>");
            sb.AppendLine(@"		<Pair>");
            sb.AppendLine(@"			<key>highlight</key>");
            sb.AppendLine(@"			<styleUrl>#sh_ylw-pushpin</styleUrl>");
            sb.AppendLine(@"		</Pair>");
            sb.AppendLine(@"	</StyleMap>");
            sb.AppendLine(@"	<Style id=""sh_ylw-pushpin"">");
            sb.AppendLine(@"		<IconStyle>");
            sb.AppendLine(@"			<scale>1.2</scale>");
            sb.AppendLine(@"		</IconStyle>");
            sb.AppendLine(@"		<LineStyle>");
            sb.AppendLine(@"			<color>ff00ff00</color>");
            sb.AppendLine(@"			<width>1.5</width>");
            sb.AppendLine(@"		</LineStyle>");
            sb.AppendLine(@"		<PolyStyle>");
            sb.AppendLine(@"			<color>0000ff00</color>");
            sb.AppendLine(@"		</PolyStyle>");
            sb.AppendLine(@"	</Style>");
            sb.AppendLine(@"	<Style id=""sn_ylw-pushpin"">");
            sb.AppendLine(@"		<LineStyle>");
            sb.AppendLine(@"			<color>ff00ff00</color>");
            sb.AppendLine(@"			<width>1.5</width>");
            sb.AppendLine(@"		</LineStyle>");
            sb.AppendLine(@"		<PolyStyle>");
            sb.AppendLine(@"			<color>0000ff00</color>");
            sb.AppendLine(@"		</PolyStyle>");
            sb.AppendLine(@"	</Style>");

            sb.AppendLine(@"    <StyleMap id=""msn_grn-pushpin"">");
            sb.AppendLine(@"		<Pair>");
            sb.AppendLine(@"			<key>normal</key>");
            sb.AppendLine(@"			<styleUrl>#sn_grn-pushpin</styleUrl>");
            sb.AppendLine(@"		</Pair>");
            sb.AppendLine(@"		<Pair>");
            sb.AppendLine(@"			<key>highlight</key>");
            sb.AppendLine(@"			<styleUrl>#sh_grn-pushpin</styleUrl>");
            sb.AppendLine(@"		</Pair>");
            sb.AppendLine(@"	</StyleMap>");
            sb.AppendLine(@"	<Style id=""sh_grn-pushpin"">");
            sb.AppendLine(@"		<IconStyle>");
            sb.AppendLine(@"			<scale>1.2</scale>");
            sb.AppendLine(@"		</IconStyle>");
            sb.AppendLine(@"		<LineStyle>");
            sb.AppendLine(@"			<color>ff0000ff</color>");
            sb.AppendLine(@"			<width>1.5</width>");
            sb.AppendLine(@"		</LineStyle>");
            sb.AppendLine(@"		<PolyStyle>");
            sb.AppendLine(@"			<color>000000ff</color>");
            sb.AppendLine(@"		</PolyStyle>");
            sb.AppendLine(@"	</Style>");
            sb.AppendLine(@"	<Style id=""sn_grn-pushpin"">");
            sb.AppendLine(@"		<LineStyle>");
            sb.AppendLine(@"			<color>ff0000ff</color>");
            sb.AppendLine(@"			<width>1.5</width>");
            sb.AppendLine(@"		</LineStyle>");
            sb.AppendLine(@"		<PolyStyle>");
            sb.AppendLine(@"			<color>000000ff</color>");
            sb.AppendLine(@"		</PolyStyle>");
            sb.AppendLine(@"	</Style>");

            sb.AppendLine(@"    <StyleMap id=""msn_blue-pushpin"">");
            sb.AppendLine(@"		<Pair>");
            sb.AppendLine(@"			<key>normal</key>");
            sb.AppendLine(@"			<styleUrl>#sn_blue-pushpin</styleUrl>");
            sb.AppendLine(@"		</Pair>");
            sb.AppendLine(@"		<Pair>");
            sb.AppendLine(@"			<key>highlight</key>");
            sb.AppendLine(@"			<styleUrl>#sh_blue-pushpin</styleUrl>");
            sb.AppendLine(@"		</Pair>");
            sb.AppendLine(@"	</StyleMap>");
            sb.AppendLine(@"	<Style id=""sh_blue-pushpin"">");
            sb.AppendLine(@"		<IconStyle>");
            sb.AppendLine(@"			<scale>1.2</scale>");
            sb.AppendLine(@"		</IconStyle>");
            sb.AppendLine(@"		<LineStyle>");
            sb.AppendLine(@"			<color>ff0000ff</color>");
            sb.AppendLine(@"			<width>1.5</width>");
            sb.AppendLine(@"		</LineStyle>");
            sb.AppendLine(@"		<PolyStyle>");
            sb.AppendLine(@"			<color>000000ff</color>");
            sb.AppendLine(@"		</PolyStyle>");
            sb.AppendLine(@"	</Style>");
            sb.AppendLine(@"	<Style id=""sn_blue-pushpin"">");
            sb.AppendLine(@"		<LineStyle>");
            sb.AppendLine(@"			<color>ff0000ff</color>");
            sb.AppendLine(@"			<width>1.5</width>");
            sb.AppendLine(@"		</LineStyle>");
            sb.AppendLine(@"		<PolyStyle>");
            sb.AppendLine(@"			<color>000000ff</color>");
            sb.AppendLine(@"		</PolyStyle>");
            sb.AppendLine(@"	</Style>");

            sb.AppendLine(@"    <StyleMap id=""msn_ltblu-pushpin"">");
            sb.AppendLine(@"		<Pair>");
            sb.AppendLine(@"			<key>normal</key>");
            sb.AppendLine(@"			<styleUrl>#sn_ltblu-pushpin</styleUrl>");
            sb.AppendLine(@"		</Pair>");
            sb.AppendLine(@"		<Pair>");
            sb.AppendLine(@"			<key>highlight</key>");
            sb.AppendLine(@"			<styleUrl>#sh_ltblu-pushpin</styleUrl>");
            sb.AppendLine(@"		</Pair>");
            sb.AppendLine(@"	</StyleMap>");
            sb.AppendLine(@"	<Style id=""sh_ltblu-pushpin"">");
            sb.AppendLine(@"		<IconStyle>");
            sb.AppendLine(@"			<scale>1.2</scale>");
            sb.AppendLine(@"		</IconStyle>");
            sb.AppendLine(@"		<LineStyle>");
            sb.AppendLine(@"			<color>ff0000ff</color>");
            sb.AppendLine(@"			<width>1.5</width>");
            sb.AppendLine(@"		</LineStyle>");
            sb.AppendLine(@"		<PolyStyle>");
            sb.AppendLine(@"			<color>000000ff</color>");
            sb.AppendLine(@"		</PolyStyle>");
            sb.AppendLine(@"	</Style>");
            sb.AppendLine(@"	<Style id=""sn_ltblu-pushpin"">");
            sb.AppendLine(@"		<LineStyle>");
            sb.AppendLine(@"			<color>ff0000ff</color>");
            sb.AppendLine(@"			<width>1.5</width>");
            sb.AppendLine(@"		</LineStyle>");
            sb.AppendLine(@"		<PolyStyle>");
            sb.AppendLine(@"			<color>000000ff</color>");
            sb.AppendLine(@"		</PolyStyle>");
            sb.AppendLine(@"	</Style>");

            sb.AppendLine(@"    <StyleMap id=""msn_wht-blank"">");
            sb.AppendLine(@"		<Pair>");
            sb.AppendLine(@"			<key>normal</key>");
            sb.AppendLine(@"			<styleUrl>#sn_wht-blank</styleUrl>");
            sb.AppendLine(@"		</Pair>");
            sb.AppendLine(@"		<Pair>");
            sb.AppendLine(@"			<key>highlight</key>");
            sb.AppendLine(@"			<styleUrl>#sh_wht-blank</styleUrl>");
            sb.AppendLine(@"		</Pair>");
            sb.AppendLine(@"	</StyleMap>");
            sb.AppendLine(@"	<Style id=""sh_wht-blank"">");
            sb.AppendLine(@"		<IconStyle>");
            sb.AppendLine(@"			<scale>1.2</scale>");
            sb.AppendLine(@"		</IconStyle>");
            sb.AppendLine(@"		<LineStyle>");
            sb.AppendLine(@"			<color>ff0000ff</color>");
            sb.AppendLine(@"			<width>1.5</width>");
            sb.AppendLine(@"		</LineStyle>");
            sb.AppendLine(@"		<PolyStyle>");
            sb.AppendLine(@"			<color>000000ff</color>");
            sb.AppendLine(@"		</PolyStyle>");
            sb.AppendLine(@"	</Style>");
            sb.AppendLine(@"	<Style id=""sn_wht-blank"">");
            sb.AppendLine(@"		<LineStyle>");
            sb.AppendLine(@"			<color>ff0000ff</color>");
            sb.AppendLine(@"			<width>1.5</width>");
            sb.AppendLine(@"		</LineStyle>");
            sb.AppendLine(@"		<PolyStyle>");
            sb.AppendLine(@"			<color>000000ff</color>");
            sb.AppendLine(@"		</PolyStyle>");
            sb.AppendLine(@"	</Style>");

            TVItemModel tvItemModelSubsector = _TVItemService.GetTVItemModelWithTVItemIDDB(TVItemID);
            List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.Subsector, MapInfoDrawTypeEnum.Point);

            sb.AppendLine(@"	<Folder>");
            sb.AppendLine(@"	<name>" + tvItemModelSubsector.TVText + "</name>");

            int pos = tvItemModelSubsector.TVText.IndexOf(" ");
            string TVText = tvItemModelSubsector.TVText;
            if (pos > 0)
            {
                TVText = tvItemModelSubsector.TVText.Substring(0, pos);
            }

            // Doing Point
            sb.AppendLine(@"	<Placemark>");
            sb.AppendLine(@"	    <name>" + TaskRunnerServiceRes.GeneralStatAndInfo + "</name>");
            sb.AppendLine(@"        <styleUrl>#msn_blue-pushpin</styleUrl>");
            sb.AppendLine(@"	    <description>");
            sb.AppendLine(@"            <![CDATA[");
            sb.AppendLine(@"                <h2><a href=""" + _TaskRunnerBaseService.GetUrlFromTVItem(tvItemModelSubsector) + @""">" + tvItemModelSubsector.TVText + "</a></h2>");

            List<TVItemModel> tvItemModelMunicipalityList = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.Municipality);

            sb.AppendLine(@"                <h3>" + TaskRunnerServiceRes.Municipalities + " (" + tvItemModelMunicipalityList.Count + ")</h3>");
            sb.AppendLine(@"                <ul>");


            foreach (TVItemModel tvItemModelMunicipality in tvItemModelMunicipalityList)
            {
                sb.AppendLine(@"                    <li>");
                sb.AppendLine(@"                        <p><a href=""" + _TaskRunnerBaseService.GetUrlFromTVItem(tvItemModelMunicipality) + @""">" + tvItemModelMunicipality.TVText + "</a></p>");

                List<TVItemModel> tvItemModelInfrastructureList = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelMunicipality.TVItemID, TVTypeEnum.Infrastructure);

                List<InfrastructureModel> infrastructureModelList = new List<InfrastructureModel>();
                foreach (TVItemModel tvItemModelInfrastructure in tvItemModelInfrastructureList)
                {
                    infrastructureModelList.Add(_InfrastructureService.GetInfrastructureModelWithInfrastructureTVItemIDDB(tvItemModelInfrastructure.TVItemID));
                }

                sb.AppendLine(@"                <h4>" + TaskRunnerServiceRes.Infrastructures + " (" + infrastructureModelList.Count + ")</h4>");
                sb.AppendLine(@"                        <ul>");
                foreach (TVItemModel tvItemModelInfrastructure in tvItemModelInfrastructureList)
                {
                    sb.AppendLine(@"                            <li>");
                    sb.AppendLine(@"                                <p><a href=""" + _TaskRunnerBaseService.GetUrlFromTVItem(tvItemModelInfrastructure) + @""">" + tvItemModelInfrastructure.TVText + "</a></p>");
                    sb.AppendLine(@"                            </li>");
                }
                sb.AppendLine(@"                        </ul>");
                
                sb.AppendLine(@"                    </li>");
            }
            sb.AppendLine(@"                </ul>");

            List<TVItemModel> tvItemModelMWQMSiteList = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.MWQMSite);

            sb.AppendLine(@"                <h3>" + TaskRunnerServiceRes.MWQMSites + " (" + tvItemModelMWQMSiteList.Count + ")</h3>");
            sb.AppendLine(@"                <ul>");


            foreach (TVItemModel tvItemModelMWQMSite in tvItemModelMWQMSiteList)
            {
                sb.AppendLine(@"                    <li>");
                sb.AppendLine(@"                        <p><a href=""" + _TaskRunnerBaseService.GetUrlFromTVItem(tvItemModelMWQMSite) + @""">" + tvItemModelMWQMSite.TVText + "</a></p>");
                sb.AppendLine(@"                    </li>");
            }

            sb.AppendLine(@"                </ul>");

            List<TVItemModel> tvItemModelPolSourceSiteList = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.PolSourceSite);

            sb.AppendLine(@"                <h3>" + TaskRunnerServiceRes.PollutionSourceSites + " (" + tvItemModelPolSourceSiteList.Count + ")</h3>");
            sb.AppendLine(@"                <ul>");


            foreach (TVItemModel tvItemModelPollutionSourceSite in tvItemModelPolSourceSiteList)
            {
                sb.AppendLine(@"                    <li>");
                sb.AppendLine(@"                        <p><a href=""" + _TaskRunnerBaseService.GetUrlFromTVItem(tvItemModelPollutionSourceSite) + @""">" + tvItemModelPollutionSourceSite.TVText + "</a></p>");
                sb.AppendLine(@"                    </li>");
            }

            sb.AppendLine(@"                </ul>");

            sb.AppendLine(@"            ]]> ");
            sb.AppendLine(@"	    </description>");
            sb.AppendLine(@"	    <Point>");
            sb.AppendLine(@"	    	<coordinates>" + mapInfoPointModelList[0].Lng + "," + mapInfoPointModelList[0].Lat + ",0</coordinates>");
            sb.AppendLine(@"	    </Point>");
            sb.AppendLine(@"	</Placemark>");

            // Doing Polygon
            sb.AppendLine(@"	<Placemark>");
            sb.AppendLine(@"		<name>" + TVText + " (poly)</name>");
            sb.AppendLine(@"        <styleUrl>#msn_ylw-pushpin</styleUrl>");
            sb.AppendLine(@"		<Polygon>");
            sb.AppendLine(@"			<outerBoundaryIs>");
            sb.AppendLine(@"				<LinearRing>");
            sb.AppendLine(@"					<coordinates>");

            mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.Subsector, MapInfoDrawTypeEnum.Polygon);
            foreach (MapInfoPointModel mapInfoPointModel in mapInfoPointModelList)
            {
                sb.AppendLine(mapInfoPointModel.Lng + "," + mapInfoPointModel.Lat + ",0 ");

            }

            sb.AppendLine(@"					</coordinates>");
            sb.AppendLine(@"				</LinearRing>");
            sb.AppendLine(@"			</outerBoundaryIs>");
            sb.AppendLine(@"		</Polygon>");
            sb.AppendLine(@"	</Placemark>");

            // Doing Municipalities
            sb.AppendLine(@" <Folder>");
            sb.AppendLine(@"	<name>Municipalities</name>");

            foreach (TVItemModel tvItemModelMunicipality in tvItemModelMunicipalityList)
            {
                sb.AppendLine(@" <Folder>");
                sb.AppendLine(@"	<name>" + tvItemModelMunicipality.TVText + "</name>");
                // Doing point
                sb.AppendLine(@"	<Placemark>");
                sb.AppendLine(@"	<name>" + tvItemModelMunicipality.TVText + " ( Point)</name>");
                sb.AppendLine(@"<styleUrl>#msn_blue-pushpin</styleUrl>");

                mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelMunicipality.TVItemID, TVTypeEnum.Municipality, MapInfoDrawTypeEnum.Point);

                sb.AppendLine(@"		<Point>");
                sb.AppendLine(@"			<coordinates>" + mapInfoPointModelList[0].Lng + "," + mapInfoPointModelList[0].Lat + ",0</coordinates>");
                sb.AppendLine(@"		</Point>");
                sb.AppendLine(@"	</Placemark>");

                List<TVItemModel> tvItemModelInfrastructureList = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelMunicipality.TVItemID, TVTypeEnum.Infrastructure);

                List<InfrastructureModel> infrastructureModelList = new List<InfrastructureModel>();

                foreach (TVItemModel tvItemModelInfrastructure in tvItemModelInfrastructureList)
                {
                    infrastructureModelList.Add(_InfrastructureService.GetInfrastructureModelWithInfrastructureTVItemIDDB(tvItemModelInfrastructure.TVItemID));
                }

                // Doing WWTP
                foreach (InfrastructureModel infrastructureModel in infrastructureModelList.Where(c => c.InfrastructureType == InfrastructureTypeEnum.WWTP).ToList())
                {
                    sb.AppendLine(@"	<Placemark>");
                    sb.AppendLine(@"	<name>" + tvItemModelInfrastructureList.Where(c => c.TVItemID == infrastructureModel.InfrastructureTVItemID).Select(c => c.TVText).FirstOrDefault() + "</name>");
                    sb.AppendLine(@"<styleUrl>#msn_blue-pushpin</styleUrl>");

                    sb.AppendLine(@"	<description>");
                    sb.AppendLine(@"<![CDATA[");
                    sb.AppendLine(@"<pre>");
                    sb.AppendLine(@"Alarm System Type: " + (infrastructureModel.AlarmSystemType != null && infrastructureModel.AlarmSystemType != 0 ? _BaseEnumService.GetEnumText_AlarmSystemTypeEnum(infrastructureModel.AlarmSystemType) : "") + "\r\n");
                    sb.AppendLine(@"Can overflow: " + infrastructureModel.CanOverflow.ToString() + "\r\n");
                    sb.AppendLine(@"Category: " + infrastructureModel.InfrastructureCategory + "\r\n");
                    sb.AppendLine(@"Collection System Type: " + (infrastructureModel.CollectionSystemType != null && infrastructureModel.CollectionSystemType != 0 ? _BaseEnumService.GetEnumText_CollectionSystemTypeEnum(infrastructureModel.CollectionSystemType) : "") + "\r\n");
                    sb.AppendLine(@"Comments: " + infrastructureModel.Comment + "\r\n");
                    sb.AppendLine(@"DesignFlow (m3/day): " + infrastructureModel.DesignFlow_m3_day + "\r\n");
                    sb.AppendLine(@"Disinfection Type: " + (infrastructureModel.DisinfectionType != null && infrastructureModel.DisinfectionType != 0 ? _BaseEnumService.GetEnumText_DisinfectionTypeEnum(infrastructureModel.DisinfectionType) : "") + "\r\n");
                    sb.AppendLine(@"Infrastructure Type: " + (infrastructureModel.InfrastructureType != null && infrastructureModel.InfrastructureType != 0 ? _BaseEnumService.GetEnumText_InfrastructureTypeEnum(infrastructureModel.InfrastructureType) : "") + "\r\n");
                    sb.AppendLine(@"Average Flow (m3/day): " + infrastructureModel.AverageFlow_m3_day + "\r\n");
                    sb.AppendLine(@"AverageFlow_m3_day: " + infrastructureModel.PeakFlow_m3_day + "\r\n");
                    sb.AppendLine(@"Percent Flow Of Total (%): " + infrastructureModel.PercFlowOfTotal + "\r\n");
                    sb.AppendLine(@"Population Served: " + infrastructureModel.PopServed + "\r\n");
                    sb.AppendLine(@"Time Zone: " + infrastructureModel.TimeOffset_hour + "\r\n");
                    sb.AppendLine(@"Treatment Type: " + (infrastructureModel.TreatmentType != null && infrastructureModel.TreatmentType != 0 ? _BaseEnumService.GetEnumText_TreatmentTypeEnum(infrastructureModel.TreatmentType) : "") + "\r\n");

                    List<MapInfoPointModel> mapInfoPointModelInfrastructureList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(infrastructureModel.InfrastructureTVItemID, TVTypeEnum.WasteWaterTreatmentPlant, MapInfoDrawTypeEnum.Point);

                    if (mapInfoPointModelInfrastructureList.Count > 0)
                    {
                        sb.AppendLine(@"Latitude Longitude: " + mapInfoPointModelInfrastructureList[0].Lat + " " + mapInfoPointModelInfrastructureList[0].Lng + "\r\n");
                    }
                    else
                    {
                        sb.AppendLine("Latitude Longitude: \r\n");
                    }

                    List<MapInfoPointModel> mapInfoPointModelInfrastructureOutfallList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(infrastructureModel.InfrastructureTVItemID, TVTypeEnum.Outfall, MapInfoDrawTypeEnum.Point);

                    if (mapInfoPointModelInfrastructureOutfallList.Count > 0)
                    {
                        sb.AppendLine(@"Outfall: Latitude Longitude: " + mapInfoPointModelInfrastructureOutfallList[0].Lat + " " + mapInfoPointModelInfrastructureOutfallList[0].Lng + "\r\n");
                    }
                    else
                    {
                        sb.AppendLine("Outfall: Latitude Longitude: \r\n");
                    }
                    sb.AppendLine("\r\n\r\n");
                    sb.AppendLine("Outfall Information\r\n\r\n");
                    sb.AppendLine(@"Average Depth (m): " + infrastructureModel.AverageDepth_m + "\r\n");
                    sb.AppendLine(@"Decay Rate (/day): " + infrastructureModel.DecayRate_per_day + "\r\n");
                    sb.AppendLine(@"Distance From Shore (m): " + infrastructureModel.DistanceFromShore_m + "\r\n");
                    sb.AppendLine(@"Far Field Velocity (m/s): " + infrastructureModel.FarFieldVelocity_m_s + "\r\n");
                    sb.AppendLine(@"Horizontal Angle (deg): " + infrastructureModel.HorizontalAngle_deg + "\r\n");
                    sb.AppendLine(@"Near Field Velocity (m/s): " + infrastructureModel.NearFieldVelocity_m_s + "\r\n");
                    sb.AppendLine(@"Number Of Ports: " + infrastructureModel.NumberOfPorts + "\r\n");
                    sb.AppendLine(@"Port Diameter (m): " + infrastructureModel.PortDiameter_m + "\r\n");
                    sb.AppendLine(@"Port Elevation (m): " + infrastructureModel.PortElevation_m + "\r\n");
                    sb.AppendLine(@"Port Spacing (m): " + infrastructureModel.PortSpacing_m + "\r\n");
                    sb.AppendLine(@"Receiving Water Concentration (FC /100 ml): " + infrastructureModel.ReceivingWater_MPN_per_100ml + "\r\n");
                    sb.AppendLine(@"Receiving Water Salinity (PSU): " + infrastructureModel.ReceivingWaterSalinity_PSU + "\r\n");
                    sb.AppendLine(@"Receiving Water Temperature (ºC): " + infrastructureModel.ReceivingWaterTemperature_C + "\r\n");
                    sb.AppendLine(@"Vertical Angle (deg): " + infrastructureModel.VerticalAngle_deg + "\r\n");
                    sb.AppendLine(@"</pre>");
                    sb.AppendLine(@"]]>");
                    sb.AppendLine(@"	</description>");

                    sb.AppendLine(@"		<Point>");
                    sb.AppendLine(@"			<coordinates>" + mapInfoPointModelInfrastructureList[0].Lng + "," + mapInfoPointModelInfrastructureList[0].Lat + ",0</coordinates>");
                    sb.AppendLine(@"		</Point>");
                    sb.AppendLine(@"	</Placemark>");

                    sb.AppendLine(@"	<Placemark>");
                    sb.AppendLine(@"	<name>Outfall " + tvItemModelInfrastructureList.Where(c => c.TVItemID == infrastructureModel.InfrastructureTVItemID).Select(c => c.TVText).FirstOrDefault() + "</name>");
                    sb.AppendLine(@"<styleUrl>#msn_grn-pushpin</styleUrl>");

                    sb.AppendLine(@"		<Point>");
                    sb.AppendLine(@"			<coordinates>" + mapInfoPointModelInfrastructureOutfallList[0].Lng + "," + mapInfoPointModelInfrastructureOutfallList[0].Lat + ",0</coordinates>");
                    sb.AppendLine(@"		</Point>");
                    sb.AppendLine(@"	</Placemark>");

                }

                // Doing LS
                foreach (InfrastructureModel infrastructureModel in infrastructureModelList.Where(c => c.InfrastructureType == InfrastructureTypeEnum.LiftStation).ToList())
                {
                    sb.AppendLine(@"	<Placemark>");
                    sb.AppendLine(@"	<name>" + tvItemModelInfrastructureList.Where(c => c.TVItemID == infrastructureModel.InfrastructureTVItemID).Select(c => c.TVText).FirstOrDefault() + "</name>");
                    sb.AppendLine(@"<styleUrl>#msn_ylw-pushpin</styleUrl>");

                    sb.AppendLine(@"	<description>");
                    sb.AppendLine(@"<![CDATA[");
                    sb.AppendLine(@"<pre>");
                    sb.AppendLine(@"Alarm System Type: " + (infrastructureModel.AlarmSystemType != null && infrastructureModel.AlarmSystemType != 0 ? _BaseEnumService.GetEnumText_AlarmSystemTypeEnum(infrastructureModel.AlarmSystemType) : "") + "\r\n");
                    sb.AppendLine(@"Can overflow: " + infrastructureModel.CanOverflow.ToString() + "\r\n");
                    sb.AppendLine(@"Category: " + infrastructureModel.InfrastructureCategory + "\r\n");
                    sb.AppendLine(@"Collection System Type: " + (infrastructureModel.CollectionSystemType != null && infrastructureModel.CollectionSystemType != 0 ? _BaseEnumService.GetEnumText_CollectionSystemTypeEnum(infrastructureModel.CollectionSystemType) : "") + "\r\n");
                    sb.AppendLine(@"Comments: " + infrastructureModel.Comment + "\r\n");
                    sb.AppendLine(@"Infrastructure Type: " + (infrastructureModel.InfrastructureType != null && infrastructureModel.InfrastructureType != 0 ? _BaseEnumService.GetEnumText_InfrastructureTypeEnum(infrastructureModel.InfrastructureType) : "") + "\r\n");
                    sb.AppendLine(@"Percent Flow Of Total (%): " + infrastructureModel.PercFlowOfTotal + "\r\n");

                    List<MapInfoPointModel> mapInfoPointModelInfrastructureList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(infrastructureModel.InfrastructureTVItemID, TVTypeEnum.LiftStation, MapInfoDrawTypeEnum.Point);

                    if (mapInfoPointModelInfrastructureList.Count > 0)
                    {
                        sb.AppendLine(@"Latitude Longitude: " + mapInfoPointModelInfrastructureList[0].Lat + " " + mapInfoPointModelInfrastructureList[0].Lng + "\r\n");
                    }
                    else
                    {
                        sb.AppendLine("Latitude Longitude: \r\n");
                    }

                    List<MapInfoPointModel> mapInfoPointModelInfrastructureOutfallList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(infrastructureModel.InfrastructureTVItemID, TVTypeEnum.Outfall, MapInfoDrawTypeEnum.Point);

                    if (mapInfoPointModelInfrastructureOutfallList.Count > 0)
                    {
                        sb.AppendLine(@"Outfall: Latitude Longitude: " + mapInfoPointModelInfrastructureOutfallList[0].Lat + " " + mapInfoPointModelInfrastructureOutfallList[0].Lng + "\r\n");
                    }
                    else
                    {
                        sb.AppendLine("Outfall: Latitude Longitude: \r\n");
                    }
                    sb.AppendLine("\r\n\r\n");
                    sb.AppendLine("Outfall Information\r\n\r\n");
                    sb.AppendLine(@"Average Depth (m): " + infrastructureModel.AverageDepth_m + "\r\n");
                    sb.AppendLine(@"Decay Rate (/day): " + infrastructureModel.DecayRate_per_day + "\r\n");
                    sb.AppendLine(@"Distance From Shore (m): " + infrastructureModel.DistanceFromShore_m + "\r\n");
                    sb.AppendLine(@"Far Field Velocity (m/s): " + infrastructureModel.FarFieldVelocity_m_s + "\r\n");
                    sb.AppendLine(@"Horizontal Angle (deg): " + infrastructureModel.HorizontalAngle_deg + "\r\n");
                    sb.AppendLine(@"Near Field Velocity (m/s): " + infrastructureModel.NearFieldVelocity_m_s + "\r\n");
                    sb.AppendLine(@"Number Of Ports: " + infrastructureModel.NumberOfPorts + "\r\n");
                    sb.AppendLine(@"Port Diameter (m): " + infrastructureModel.PortDiameter_m + "\r\n");
                    sb.AppendLine(@"Port Elevation (m): " + infrastructureModel.PortElevation_m + "\r\n");
                    sb.AppendLine(@"Port Spacing (m): " + infrastructureModel.PortSpacing_m + "\r\n");
                    sb.AppendLine(@"Receiving Water Concentration (FC /100 ml): " + infrastructureModel.ReceivingWater_MPN_per_100ml + "\r\n");
                    sb.AppendLine(@"Receiving Water Salinity (PSU): " + infrastructureModel.ReceivingWaterSalinity_PSU + "\r\n");
                    sb.AppendLine(@"Receiving Water Temperature (ºC): " + infrastructureModel.ReceivingWaterTemperature_C + "\r\n");
                    sb.AppendLine(@"Vertical Angle (deg): " + infrastructureModel.VerticalAngle_deg + "\r\n");
                    sb.AppendLine(@"</pre>");
                    sb.AppendLine(@"]]>");
                    sb.AppendLine(@"	</description>");

                    sb.AppendLine(@"		<Point>");
                    sb.AppendLine(@"			<coordinates>" + mapInfoPointModelInfrastructureList[0].Lng + "," + mapInfoPointModelInfrastructureList[0].Lat + ",0</coordinates>");
                    sb.AppendLine(@"		</Point>");
                    sb.AppendLine(@"	</Placemark>");

                    sb.AppendLine(@"	<Placemark>");
                    sb.AppendLine(@"	<name>Outfall " + tvItemModelInfrastructureList.Where(c => c.TVItemID == infrastructureModel.InfrastructureTVItemID).Select(c => c.TVText).FirstOrDefault() + "</name>");
                    sb.AppendLine(@"<styleUrl>#msn_grn-pushpin</styleUrl>");

                    sb.AppendLine(@"		<Point>");
                    sb.AppendLine(@"			<coordinates>" + mapInfoPointModelInfrastructureOutfallList[0].Lng + "," + mapInfoPointModelInfrastructureOutfallList[0].Lat + ",0</coordinates>");
                    sb.AppendLine(@"		</Point>");
                    sb.AppendLine(@"	</Placemark>");
                }

                // Doing Line Overflow
                foreach (InfrastructureModel infrastructureModel in infrastructureModelList.Where(c => c.InfrastructureType == InfrastructureTypeEnum.LineOverflow).ToList())
                {
                    sb.AppendLine(@"	<Placemark>");
                    sb.AppendLine(@"	<name>" + tvItemModelInfrastructureList.Where(c => c.TVItemID == infrastructureModel.InfrastructureTVItemID).Select(c => c.TVText).FirstOrDefault() + "</name>");
                    sb.AppendLine(@"<styleUrl>#msn_grn-pushpin</styleUrl>");

                    sb.AppendLine(@"	<description>");
                    sb.AppendLine(@"<![CDATA[");
                    sb.AppendLine(@"<pre>");
                    sb.AppendLine(@"Alarm System Type: " + _BaseEnumService.GetEnumText_AlarmSystemTypeEnum(infrastructureModel.AlarmSystemType) + "\r\n");
                    sb.AppendLine(@"Can overflow: " + infrastructureModel.CanOverflow.ToString() + "\r\n");
                    sb.AppendLine(@"Category: " + infrastructureModel.InfrastructureCategory + "\r\n");
                    sb.AppendLine(@"Collection System Type: " + _BaseEnumService.GetEnumText_CollectionSystemTypeEnum(infrastructureModel.CollectionSystemType) + "\r\n");
                    sb.AppendLine(@"Comments: " + infrastructureModel.Comment + "\r\n");
                    sb.AppendLine(@"Infrastructure Type: " + _BaseEnumService.GetEnumText_InfrastructureTypeEnum(infrastructureModel.InfrastructureType) + "\r\n");
                    sb.AppendLine(@"Percent Flow Of Total (%): " + infrastructureModel.PercFlowOfTotal + "\r\n");

                    List<MapInfoPointModel> mapInfoPointModelInfrastructureList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(infrastructureModel.InfrastructureTVItemID, TVTypeEnum.LineOverflow, MapInfoDrawTypeEnum.Point);

                    if (mapInfoPointModelInfrastructureList.Count > 0)
                    {
                        sb.AppendLine(@"Latitude Longitude: " + mapInfoPointModelInfrastructureList[0].Lat + " " + mapInfoPointModelInfrastructureList[0].Lng + "\r\n");
                    }
                    else
                    {
                        sb.AppendLine("Latitude Longitude: \r\n");
                    }

                    List<MapInfoPointModel> mapInfoPointModelInfrastructureOutfallList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(infrastructureModel.InfrastructureTVItemID, TVTypeEnum.Outfall, MapInfoDrawTypeEnum.Point);

                    if (mapInfoPointModelInfrastructureOutfallList.Count > 0)
                    {
                        sb.AppendLine(@"Outfall: Latitude Longitude: " + mapInfoPointModelInfrastructureOutfallList[0].Lat + " " + mapInfoPointModelInfrastructureOutfallList[0].Lng + "\r\n");
                    }
                    else
                    {
                        sb.AppendLine("Outfall: Latitude Longitude: \r\n");
                    }
                    sb.AppendLine("\r\n\r\n");
                    sb.AppendLine("Outfall Information\r\n\r\n");
                    sb.AppendLine(@"Average Depth (m): " + infrastructureModel.AverageDepth_m + "\r\n");
                    sb.AppendLine(@"Decay Rate (/day): " + infrastructureModel.DecayRate_per_day + "\r\n");
                    sb.AppendLine(@"Distance From Shore (m): " + infrastructureModel.DistanceFromShore_m + "\r\n");
                    sb.AppendLine(@"Far Field Velocity (m/s): " + infrastructureModel.FarFieldVelocity_m_s + "\r\n");
                    sb.AppendLine(@"Horizontal Angle (deg): " + infrastructureModel.HorizontalAngle_deg + "\r\n");
                    sb.AppendLine(@"Near Field Velocity (m/s): " + infrastructureModel.NearFieldVelocity_m_s + "\r\n");
                    sb.AppendLine(@"Number Of Ports: " + infrastructureModel.NumberOfPorts + "\r\n");
                    sb.AppendLine(@"Port Diameter (m): " + infrastructureModel.PortDiameter_m + "\r\n");
                    sb.AppendLine(@"Port Elevation (m): " + infrastructureModel.PortElevation_m + "\r\n");
                    sb.AppendLine(@"Port Spacing (m): " + infrastructureModel.PortSpacing_m + "\r\n");
                    sb.AppendLine(@"Receiving Water Concentration (FC /100 ml): " + infrastructureModel.ReceivingWater_MPN_per_100ml + "\r\n");
                    sb.AppendLine(@"Receiving Water Salinity (PSU): " + infrastructureModel.ReceivingWaterSalinity_PSU + "\r\n");
                    sb.AppendLine(@"Receiving Water Temperature (ºC): " + infrastructureModel.ReceivingWaterTemperature_C + "\r\n");
                    sb.AppendLine(@"Vertical Angle (deg): " + infrastructureModel.VerticalAngle_deg + "\r\n");
                    sb.AppendLine(@"</pre>");
                    sb.AppendLine(@"]]>");
                    sb.AppendLine(@"	</description>");

                    sb.AppendLine(@"		<Point>");
                    sb.AppendLine(@"			<coordinates>" + mapInfoPointModelInfrastructureList[0].Lng + "," + mapInfoPointModelInfrastructureList[0].Lat + ",0</coordinates>");
                    sb.AppendLine(@"		</Point>");
                    sb.AppendLine(@"	</Placemark>");

                    sb.AppendLine(@"	<Placemark>");
                    sb.AppendLine(@"	<name>Outfall " + tvItemModelInfrastructureList.Where(c => c.TVItemID == infrastructureModel.InfrastructureTVItemID).Select(c => c.TVText).FirstOrDefault() + "</name>");
                    sb.AppendLine(@"<styleUrl>#msn_grn-pushpin</styleUrl>");

                    sb.AppendLine(@"		<Point>");
                    sb.AppendLine(@"			<coordinates>" + mapInfoPointModelInfrastructureOutfallList[0].Lng + "," + mapInfoPointModelInfrastructureOutfallList[0].Lat + ",0</coordinates>");
                    sb.AppendLine(@"		</Point>");
                    sb.AppendLine(@"	</Placemark>");
                }
                sb.AppendLine(@" </Folder>");
            }
            sb.AppendLine(@" </Folder>");

            // Doing Short Pollution Source Site
            sb.AppendLine(@" <Folder>");
            sb.AppendLine(@"	<name>Short Pollution Source Sites</name>");
            List<PolSourceSiteModel> polSourceSiteModelList = _PolSourceSiteService.GetPolSourceSiteModelListWithSubsectorTVItemIDDB(tvItemModelSubsector.TVItemID);

            foreach (PolSourceSiteModel polSourceSiteModel in polSourceSiteModelList.OrderBy(c => c.Site).ToList())
            {
                // Doing point
                sb.AppendLine(@"	<Placemark>");
                sb.AppendLine(@"	<name>" + polSourceSiteModel.Site.ToString() + "</name>");
                sb.AppendLine(@"<styleUrl>#msn_ylw-pushpin</styleUrl>");
                sb.AppendLine(@"	<description>");
                sb.AppendLine(@"<![CDATA[");
                sb.AppendLine(@"<pre>");

                PolSourceObservationModel polSourceObservationModel = _PolSourceSiteService._PolSourceObservationService.GetPolSourceObservationModelLatestWithPolSourceSiteIDDB(polSourceSiteModel.PolSourceSiteID);
                if (polSourceObservationModel != null)
                {
                    List<PolSourceObservationIssueModel> polSourceObservationIssueModelList = _PolSourceSiteService._PolSourceObservationService._PolSourceObservationIssueService.GetPolSourceObservationIssueModelListWithPolSourceObservationIDDB(polSourceObservationModel.PolSourceObservationID);

                    string SelectedObservation = "Selected: \r\n";
                    foreach (PolSourceObservationIssueModel polSourceObservationIssueModel in polSourceObservationIssueModelList)
                    {
                        foreach (PolSourceObsInfoEnum polSourceObsInfo in polSourceObservationIssueModel.PolSourceObsInfoList)
                        {
                            SelectedObservation += _BaseEnumService.GetEnumText_PolSourceObsInfoReportEnum(polSourceObsInfo);
                        }
                        SelectedObservation += "\r\n\r\n";
                    }

                    sb.AppendLine("Written: \r\n" + (string.IsNullOrWhiteSpace(polSourceObservationModel.Observation_ToBeDeleted) ? "" : polSourceObservationModel.Observation_ToBeDeleted.ToString()) + "\r\n\r\n" + SelectedObservation);
                }
                else
                {
                    string SelectedObservation = "Selected: \r\n";
                    sb.AppendLine("Written: \r\n\r\n" + SelectedObservation);
                }

                sb.AppendLine(@"</pre>");
                sb.AppendLine(@"]]>");
                sb.AppendLine(@"	</description>");

                mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(polSourceSiteModel.PolSourceSiteTVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);

                sb.AppendLine(@"		<Point>");
                sb.AppendLine(@"			<coordinates>" + mapInfoPointModelList[0].Lng + "," + mapInfoPointModelList[0].Lat + ",0</coordinates>");
                sb.AppendLine(@"		</Point>");
                sb.AppendLine(@"	</Placemark>");

            }
            sb.AppendLine(@" </Folder>");

            // Doing Long Pollution Source Site
            sb.AppendLine(@" <Folder>");
            sb.AppendLine(@"	<name>Long Pollution Source Sites</name>");
            polSourceSiteModelList = _PolSourceSiteService.GetPolSourceSiteModelListWithSubsectorTVItemIDDB(tvItemModelSubsector.TVItemID);

            foreach (PolSourceSiteModel polSourceSiteModel in polSourceSiteModelList.OrderBy(c => c.Site).ToList())
            {
                TVItemModel tvItemModelPolSourceSite = _TVItemService.GetTVItemModelWithTVItemIDDB(polSourceSiteModel.PolSourceSiteTVItemID);

                // Doing point
                sb.AppendLine(@"	<Placemark>");
                sb.AppendLine(@"	<name>" + tvItemModelPolSourceSite.TVText + "</name>");
                sb.AppendLine(@"<styleUrl>#msn_ltblu-pushpin</styleUrl>");
                sb.AppendLine(@"	<description>");
                sb.AppendLine(@"<![CDATA[");
                sb.AppendLine(@"<pre>");

                PolSourceObservationModel polSourceObservationModel = _PolSourceSiteService._PolSourceObservationService.GetPolSourceObservationModelLatestWithPolSourceSiteIDDB(polSourceSiteModel.PolSourceSiteID);
                if (polSourceObservationModel != null)
                {
                    List<PolSourceObservationIssueModel> polSourceObservationIssueModelList = _PolSourceSiteService._PolSourceObservationService._PolSourceObservationIssueService.GetPolSourceObservationIssueModelListWithPolSourceObservationIDDB(polSourceObservationModel.PolSourceObservationID);

                    string SelectedObservation = "Selected: \r\n";
                    foreach (PolSourceObservationIssueModel polSourceObservationIssueModel in polSourceObservationIssueModelList)
                    {
                        foreach (PolSourceObsInfoEnum polSourceObsInfo in polSourceObservationIssueModel.PolSourceObsInfoList)
                        {
                            SelectedObservation += _BaseEnumService.GetEnumText_PolSourceObsInfoReportEnum(polSourceObsInfo);
                        }
                        SelectedObservation += "\r\n\r\n";
                    }

                    sb.AppendLine("Written: \r\n" + (string.IsNullOrWhiteSpace(polSourceObservationModel.Observation_ToBeDeleted) ? "" : polSourceObservationModel.Observation_ToBeDeleted.ToString()) + "\r\n\r\n" + SelectedObservation);
                }
                else
                {
                    string SelectedObservation = "Selected: \r\n";
                    sb.AppendLine("Written: \r\n\r\n" + SelectedObservation);
                }

                sb.AppendLine(@"</pre>");
                sb.AppendLine(@"]]>");
                sb.AppendLine(@"	</description>");

                mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelPolSourceSite.TVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);

                sb.AppendLine(@"		<Point>");
                sb.AppendLine(@"			<coordinates>" + mapInfoPointModelList[0].Lng + "," + mapInfoPointModelList[0].Lat + ",0</coordinates>");
                sb.AppendLine(@"		</Point>");
                sb.AppendLine(@"	</Placemark>");

            }
            sb.AppendLine(@" </Folder>");


            // Doing MWQM Site
            sb.AppendLine(@" <Folder>");
            sb.AppendLine(@"	<name>MWQM Sites</name>");

            foreach (TVItemModel tvItemModelMWQMSite in tvItemModelMWQMSiteList)
            {
                // Doing point
                sb.AppendLine(@"	<Placemark>");
                sb.AppendLine(@"	<name>" + tvItemModelMWQMSite.TVText + "</name>");
                sb.AppendLine(@"<styleUrl>#msn_wht-blank</styleUrl>");
                sb.AppendLine(@"	<description>");
                sb.AppendLine(@"<![CDATA[");
                sb.AppendLine(@"                                <p><a href=""" + _TaskRunnerBaseService.GetUrlFromTVItem(tvItemModelMWQMSite) + @""">" + tvItemModelMWQMSite.TVText + "</a></p>");
                sb.AppendLine(@"]]>");
                sb.AppendLine(@"	</description>");

                mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelMWQMSite.TVItemID, TVTypeEnum.MWQMSite, MapInfoDrawTypeEnum.Point);

                sb.AppendLine(@"		<Point>");
                if (mapInfoPointModelList.Count == 0)
                {
                    sb.AppendLine(@"			<coordinates>0,0,0</coordinates>");
                }
                else
                {
                    sb.AppendLine(@"			<coordinates>" + mapInfoPointModelList[0].Lng + "," + mapInfoPointModelList[0].Lat + ",0</coordinates>");
                }
                sb.AppendLine(@"		</Point>");
                sb.AppendLine(@"	</Placemark>");

            }
            sb.AppendLine(@" </Folder>");

            sb.AppendLine(@"	</Folder>");

            if (!WriteKMLBottom(this.sb))
            {
                ErrorInDoc = true;
            }

            if (ErrorInDoc)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.ErrorOccuredWhileCreatingKMZDocument_, fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ErrorOccuredWhileCreatingKMZDocument_", fi.FullName);
                return false;
            }

            return true;
        }
        private bool GenerateKMZSubsector_NotImplementedKMZ()
        {
            string NotUsed = "";
            bool ErrorInDoc = false;

            if (!WriteKMLTop(sb))
            {
                ErrorInDoc = true;
            }

            sb.AppendLine(@"	<name>" + fi.FullName.Replace(".kml", ".kmz") + "</name>");
            sb.AppendLine(@"	<Placemark>");
            sb.AppendLine(@"		<name>Not Implemented</name> ");
            sb.AppendLine(@"		<Point>");
            sb.AppendLine(@"			<coordinates>-90,50,0</coordinates>");
            sb.AppendLine(@"		</Point> ");
            sb.AppendLine(@"	</Placemark>");

            if (!WriteKMLBottom(sb))
            {
                ErrorInDoc = true;
            }

            if (ErrorInDoc)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.ErrorOccuredWhileCreatingKMZDocument_, fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ErrorOccuredWhileCreatingKMZDocument_", fi.FullName);
                return false;
            }

            return true;
        }



    }
}
