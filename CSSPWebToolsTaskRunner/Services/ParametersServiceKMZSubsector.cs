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
using DHI.Generic.MikeZero.DFS.dfsu;
using CSSPEnumsDLL.Services;
using System.Threading;
using System.Globalization;
using CSSPDBDLL;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        private bool GenerateKMZSubsector()
        {
            switch (reportTypeModel.UniqueCode)
            {
                case "SubsectorInformationFRKMZ":
                case "SubsectorInformationENKMZ":
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

            // red polyline
            sb.AppendLine(@"	<Style id=""red-polyline"">");
            sb.AppendLine(@"        <LineStyle>");
            sb.AppendLine(@"			<color>ff0000ff</color>");
            sb.AppendLine(@"			<width>2</width>");
            sb.AppendLine(@"		</LineStyle>");
            sb.AppendLine(@"		<PolyStyle>");
            sb.AppendLine(@"			<color>000000ff</color>");
            sb.AppendLine(@"		</PolyStyle>");
            sb.AppendLine(@"	</Style>");

            // Municipality
            sb.AppendLine(@"	<Style id=""municipality"">");
            sb.AppendLine(@"	    <IconStyle>");
            sb.AppendLine(@"	    	<color>ff00ff00</color>");
            sb.AppendLine(@"	    	<Icon>");
            sb.AppendLine(@"            <href>http://maps.google.com/mapfiles/kml/shapes/flag.png</href>");
            sb.AppendLine(@"	    	</Icon>");
            sb.AppendLine(@"        </IconStyle>");
            sb.AppendLine(@"	</Style>");

            // WWTP
            sb.AppendLine(@"	<Style id=""wwtp"">");
            sb.AppendLine(@"	    <IconStyle>");
            sb.AppendLine(@"	    	<color>ffff0000</color>");
            sb.AppendLine(@"	    	<Icon>");
            sb.AppendLine(@"            <href>http://maps.google.com/mapfiles/kml/shapes/donut.png</href>");
            sb.AppendLine(@"	    	</Icon>");
            sb.AppendLine(@"        </IconStyle>");
            sb.AppendLine(@"	</Style>");

            // Outfall can overflow
            sb.AppendLine(@"	<Style id=""outfallcanoverflow"">");
            sb.AppendLine(@"	    <IconStyle>");
            sb.AppendLine(@"	    	<color>ffff0000</color>");
            sb.AppendLine(@"	    	<Icon>");
            sb.AppendLine(@"            <href>http://maps.google.com/mapfiles/kml/shapes/target.png</href>");
            sb.AppendLine(@"	    	</Icon>");
            sb.AppendLine(@"        </IconStyle>");
            sb.AppendLine(@"	</Style>");

            // Outfall can not overflow
            sb.AppendLine(@"	<Style id=""outfallcannotoverflow"">");
            sb.AppendLine(@"	    <IconStyle>");
            sb.AppendLine(@"	    	<color>ff00ff00</color>");
            sb.AppendLine(@"	    	<Icon>");
            sb.AppendLine(@"            <href>http://maps.google.com/mapfiles/kml/shapes/target.png</href>");
            sb.AppendLine(@"	    	</Icon>");
            sb.AppendLine(@"        </IconStyle>");
            sb.AppendLine(@"	</Style>");

            // Outfall unknown
            sb.AppendLine(@"	<Style id=""outfallunknown"">");
            sb.AppendLine(@"	    <IconStyle>");
            sb.AppendLine(@"	    	<color>ffffff00</color>");
            sb.AppendLine(@"	    	<Icon>");
            sb.AppendLine(@"            <href>http://maps.google.com/mapfiles/kml/shapes/target.png</href>");
            sb.AppendLine(@"	    	</Icon>");
            sb.AppendLine(@"        </IconStyle>");
            sb.AppendLine(@"	</Style>");

            // Lift station
            sb.AppendLine(@"	<Style id=""liftstation"">");
            sb.AppendLine(@"	    <IconStyle>");
            sb.AppendLine(@"	    	<color>ff0000ff</color>");
            sb.AppendLine(@"	    	<Icon>");
            sb.AppendLine(@"            <href>http://maps.google.com/mapfiles/kml/shapes/donut.png</href>");
            sb.AppendLine(@"	    	</Icon>");
            sb.AppendLine(@"        </IconStyle>");
            sb.AppendLine(@"	</Style>");

            // Line overflow
            sb.AppendLine(@"	<Style id=""lineoverflow"">");
            sb.AppendLine(@"	    <IconStyle>");
            sb.AppendLine(@"	    	<color>ff00ff00</color>");
            sb.AppendLine(@"	    	<Icon>");
            sb.AppendLine(@"            <href>http://maps.google.com/mapfiles/kml/shapes/donut.png</href>");
            sb.AppendLine(@"	    	</Icon>");
            sb.AppendLine(@"        </IconStyle>");
            sb.AppendLine(@"	</Style>");

            // Default Infrastructure
            sb.AppendLine(@"	<Style id=""defaultinfrastructure"">");
            sb.AppendLine(@"	    <IconStyle>");
            sb.AppendLine(@"	    	<color>ffffff00</color>");
            sb.AppendLine(@"	    	<Icon>");
            sb.AppendLine(@"            <href>http://maps.google.com/mapfiles/kml/shapes/donut.png</href>");
            sb.AppendLine(@"	    	</Icon>");
            sb.AppendLine(@"        </IconStyle>");
            sb.AppendLine(@"	</Style>");

            // Pollution source site
            sb.AppendLine(@"	<Style id=""pollutionsourcesite"">");
            sb.AppendLine(@"	    <IconStyle>");
            sb.AppendLine(@"	    	<color>ff00ff00</color>");
            sb.AppendLine(@"	    	<Icon>");
            sb.AppendLine(@"            <href>http://maps.google.com/mapfiles/kml/shapes/square.png</href>");
            sb.AppendLine(@"	    	</Icon>");
            sb.AppendLine(@"        </IconStyle>");
            sb.AppendLine(@"	</Style>");

            // Monitoring site
            sb.AppendLine(@"	<Style id=""monitoringsite"">");
            sb.AppendLine(@"	    <IconStyle>");
            sb.AppendLine(@"	    	<color>ff00ff00</color>");
            sb.AppendLine(@"	    	<Icon>");
            sb.AppendLine(@"            <href>http://maps.google.com/mapfiles/kml/shapes/arrow.png</href>");
            sb.AppendLine(@"	    	</Icon>");
            sb.AppendLine(@"        </IconStyle>");
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


            // Doing Polygon
            sb.AppendLine(@"	<Placemark>");
            sb.AppendLine(@"		<name>" + TVText + "</name>");
            sb.AppendLine(@"        <styleUrl>#red-polyline</styleUrl>");
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

            List<UseOfSiteModel> useOfSiteModelList = _UseOfSiteService.GetUseOfSiteModelListWithSubsectorTVItemIDDB(tvItemModelSubsector.TVItemID);

            List<TVItemModel> tvItemModelMunicipalityList = new List<TVItemModel>();

            foreach (UseOfSiteModel useOfSiteModel in useOfSiteModelList)
            {
                if (useOfSiteModel.TVType == TVTypeEnum.Municipality)
                {
                    tvItemModelMunicipalityList.Add(_TVItemService.GetTVItemModelWithTVItemIDDB(useOfSiteModel.SiteTVItemID));
                }
            }

            // Doing Municipalities
            sb.AppendLine(@" <Folder>");
            sb.AppendLine(@"	<name>Municipalities</name>");

            foreach (TVItemModel tvItemModelMunicipality in tvItemModelMunicipalityList.OrderBy(c => c.TVText))
            {
                sb.AppendLine(@" <Folder>");
                sb.AppendLine(@"	<name>" + tvItemModelMunicipality.TVText + "</name>");
                // Doing point
                sb.AppendLine(@"	<Placemark>");
                sb.AppendLine(@"	<name>" + tvItemModelMunicipality.TVText + " ( Point)</name>");
                sb.AppendLine(@"    <styleUrl>#municipality</styleUrl>");

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

                    switch (infrastructureModel.InfrastructureType)
                    {
                        case InfrastructureTypeEnum.WWTP:
                            sb.AppendLine(@"    <styleUrl>#wwtp</styleUrl>");
                            break;
                        case InfrastructureTypeEnum.LiftStation:
                            sb.AppendLine(@"    <styleUrl>#liftstation</styleUrl>");
                            break;
                        case InfrastructureTypeEnum.LineOverflow:
                            sb.AppendLine(@"    <styleUrl>#lineoverflow</styleUrl>");
                            break;
                        default:
                            sb.AppendLine(@"    <styleUrl>#defaultinfrastructure</styleUrl>");
                            break;
                    }

                    sb.AppendLine(@"	<description>");
                    sb.AppendLine(@"<![CDATA[");
                    //sb.AppendLine(@"<pre>");
                    sb.AppendLine(@"<strong>Alarm System Type:</strong> " + (infrastructureModel.AlarmSystemType != null && infrastructureModel.AlarmSystemType != 0 ? _BaseEnumService.GetEnumText_AlarmSystemTypeEnum(infrastructureModel.AlarmSystemType) : "") + "<br>");
                    sb.AppendLine(@"<strong>Can overflow:</strong> " + infrastructureModel.CanOverflow.ToString() + "<br>");
                    sb.AppendLine(@"<strong>Category:</strong> " + infrastructureModel.InfrastructureCategory + "<br>");
                    sb.AppendLine(@"<strong>Collection System Type:</strong> " + (infrastructureModel.CollectionSystemType != null && infrastructureModel.CollectionSystemType != 0 ? _BaseEnumService.GetEnumText_CollectionSystemTypeEnum(infrastructureModel.CollectionSystemType) : "") + "<br>");
                    sb.AppendLine(@"<strong>Comments:</strong> " + infrastructureModel.Comment + "<br>");
                    sb.AppendLine(@"<strong>DesignFlow (m3/day):</strong> " + infrastructureModel.DesignFlow_m3_day + "<br>");
                    sb.AppendLine(@"<strong>Disinfection Type:</strong> " + (infrastructureModel.DisinfectionType != null && infrastructureModel.DisinfectionType != 0 ? _BaseEnumService.GetEnumText_DisinfectionTypeEnum(infrastructureModel.DisinfectionType) : "") + "<br>");
                    sb.AppendLine(@"<strong>Infrastructure Type:</strong> " + (infrastructureModel.InfrastructureType != null && infrastructureModel.InfrastructureType != 0 ? _BaseEnumService.GetEnumText_InfrastructureTypeEnum(infrastructureModel.InfrastructureType) : "") + "<br>");
                    sb.AppendLine(@"<strong>Average Flow (m3/day):</strong> " + infrastructureModel.AverageFlow_m3_day + "<br>");
                    sb.AppendLine(@"<strong>AverageFlow_m3_day:</strong> " + infrastructureModel.PeakFlow_m3_day + "<br>");
                    sb.AppendLine(@"<strong>Percent Flow Of Total (%):</strong> " + infrastructureModel.PercFlowOfTotal + "<br>");
                    sb.AppendLine(@"<strong>Population Served:</strong> " + infrastructureModel.PopServed + "<br>");
                    sb.AppendLine(@"<strong>Time Zone:</strong> " + infrastructureModel.TimeOffset_hour + "<br>");
                    sb.AppendLine(@"<strong>Treatment Type:</strong> " + (infrastructureModel.TreatmentType != null && infrastructureModel.TreatmentType != 0 ? _BaseEnumService.GetEnumText_TreatmentTypeEnum(infrastructureModel.TreatmentType) : "") + "<br>");

                    List<MapInfoPointModel> mapInfoPointModelInfrastructureList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(infrastructureModel.InfrastructureTVItemID, TVTypeEnum.WasteWaterTreatmentPlant, MapInfoDrawTypeEnum.Point);

                    if (mapInfoPointModelInfrastructureList.Count > 0)
                    {
                        sb.AppendLine(@"<strong>Latitude Longitude:</strong> " + mapInfoPointModelInfrastructureList[0].Lat + " " + mapInfoPointModelInfrastructureList[0].Lng + "<br>");
                    }
                    else
                    {
                        sb.AppendLine("<strong>Latitude Longitude:</strong> <br>");
                    }

                    List<MapInfoPointModel> mapInfoPointModelInfrastructureOutfallList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(infrastructureModel.InfrastructureTVItemID, TVTypeEnum.Outfall, MapInfoDrawTypeEnum.Point);

                    if (mapInfoPointModelInfrastructureOutfallList.Count > 0)
                    {
                        sb.AppendLine(@"<strong>Outfall:</strong> Latitude Longitude: " + mapInfoPointModelInfrastructureOutfallList[0].Lat + " " + mapInfoPointModelInfrastructureOutfallList[0].Lng + "<br>");
                    }
                    else
                    {
                        sb.AppendLine("<strong>Outfall:</strong> Latitude Longitude: <br>");
                    }
                    sb.AppendLine("<br><br>");
                    sb.AppendLine("<h2>Outfall Information</h2>");
                    sb.AppendLine(@"<strong>Average Depth (m):</strong> " + infrastructureModel.AverageDepth_m + "<br>");
                    sb.AppendLine(@"<strong>Decay Rate (/day):</strong> " + infrastructureModel.DecayRate_per_day + "<br>");
                    sb.AppendLine(@"<strong>Distance From Shore (m):</strong> " + infrastructureModel.DistanceFromShore_m + "<br>");
                    sb.AppendLine(@"<strong>Far Field Velocity (m/s):</strong> " + infrastructureModel.FarFieldVelocity_m_s + "<br>");
                    sb.AppendLine(@"<strong>Horizontal Angle (deg):</strong> " + infrastructureModel.HorizontalAngle_deg + "<br>");
                    sb.AppendLine(@"<strong>Near Field Velocity (m/s):</strong> " + infrastructureModel.NearFieldVelocity_m_s + "<br>");
                    sb.AppendLine(@"<strong>Number Of Ports:</strong> " + infrastructureModel.NumberOfPorts + "<br>");
                    sb.AppendLine(@"<strong>Port Diameter (m):</strong> " + infrastructureModel.PortDiameter_m + "<br>");
                    sb.AppendLine(@"<strong>Port Elevation (m):</strong> " + infrastructureModel.PortElevation_m + "<br>");
                    sb.AppendLine(@"<strong>Port Spacing (m):</strong> " + infrastructureModel.PortSpacing_m + "<br>");
                    sb.AppendLine(@"<strong>Receiving Water Concentration (FC /100 ml):</strong> " + infrastructureModel.ReceivingWater_MPN_per_100ml + "<br>");
                    sb.AppendLine(@"<strong>Receiving Water Salinity (PSU):</strong> " + infrastructureModel.ReceivingWaterSalinity_PSU + "<br>");
                    sb.AppendLine(@"<strong>Receiving Water Temperature (ºC):</strong> " + infrastructureModel.ReceivingWaterTemperature_C + "<br>");
                    sb.AppendLine(@"<strong>Vertical Angle (deg):</strong> " + infrastructureModel.VerticalAngle_deg + "<br>");
                    //sb.AppendLine(@"</pre>");
                    sb.AppendLine(@"]]>");
                    sb.AppendLine(@"	</description>");

                    sb.AppendLine(@"		<Point>");
                    sb.AppendLine(@"			<coordinates>" + mapInfoPointModelInfrastructureList[0].Lng + "," + mapInfoPointModelInfrastructureList[0].Lat + ",0</coordinates>");
                    sb.AppendLine(@"		</Point>");
                    sb.AppendLine(@"	</Placemark>");

                    sb.AppendLine(@"	<Placemark>");
                    sb.AppendLine(@"	<name>Outfall " + tvItemModelInfrastructureList.Where(c => c.TVItemID == infrastructureModel.InfrastructureTVItemID).Select(c => c.TVText).FirstOrDefault() + "</name>");
                    
                    if (infrastructureModel.CanOverflow == null)
                    {
                        sb.AppendLine(@"    <styleUrl>#outfallunknown</styleUrl>");
                    }
                    else
                    {
                        if (!(bool)infrastructureModel.CanOverflow)
                        {
                            sb.AppendLine(@"    <styleUrl>#outfallcannotoverflow</styleUrl>");
                        }
                        else
                        {
                            sb.AppendLine(@"    <styleUrl>#outfallcanoverflow</styleUrl>");
                        }
                    }

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
                    sb.AppendLine(@"    <styleUrl>#liftstation</styleUrl>");

                    sb.AppendLine(@"	<description>");
                    sb.AppendLine(@"<![CDATA[");
                    //sb.AppendLine(@"<pre>");
                    sb.AppendLine(@"<strong>Alarm System Type:</strong> " + (infrastructureModel.AlarmSystemType != null && infrastructureModel.AlarmSystemType != 0 ? _BaseEnumService.GetEnumText_AlarmSystemTypeEnum(infrastructureModel.AlarmSystemType) : "") + "<br>");
                    sb.AppendLine(@"<strong>Can overflow:</strong> " + infrastructureModel.CanOverflow.ToString() + "<br>");
                    sb.AppendLine(@"<strong>Category:</strong> " + infrastructureModel.InfrastructureCategory + "<br>");
                    sb.AppendLine(@"<strong>Collection System Type:</strong> " + (infrastructureModel.CollectionSystemType != null && infrastructureModel.CollectionSystemType != 0 ? _BaseEnumService.GetEnumText_CollectionSystemTypeEnum(infrastructureModel.CollectionSystemType) : "") + "<br>");
                    sb.AppendLine(@"<strong>Comments:</strong> " + infrastructureModel.Comment + "<br>");
                    sb.AppendLine(@"<strong>Infrastructure Type:</strong> " + (infrastructureModel.InfrastructureType != null && infrastructureModel.InfrastructureType != 0 ? _BaseEnumService.GetEnumText_InfrastructureTypeEnum(infrastructureModel.InfrastructureType) : "") + "<br>");
                    sb.AppendLine(@"<strong>Percent Flow Of Total (%):</strong> " + infrastructureModel.PercFlowOfTotal + "<br>");

                    List<MapInfoPointModel> mapInfoPointModelInfrastructureList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(infrastructureModel.InfrastructureTVItemID, TVTypeEnum.LiftStation, MapInfoDrawTypeEnum.Point);

                    if (mapInfoPointModelInfrastructureList.Count > 0)
                    {
                        sb.AppendLine(@"<strong>Latitude Longitude:</strong> " + mapInfoPointModelInfrastructureList[0].Lat + " " + mapInfoPointModelInfrastructureList[0].Lng + "<br>");
                    }
                    else
                    {
                        sb.AppendLine("<strong>Latitude Longitude:</strong> <br>");
                    }

                    List<MapInfoPointModel> mapInfoPointModelInfrastructureOutfallList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(infrastructureModel.InfrastructureTVItemID, TVTypeEnum.Outfall, MapInfoDrawTypeEnum.Point);

                    if (mapInfoPointModelInfrastructureOutfallList.Count > 0)
                    {
                        sb.AppendLine(@"<strong>Outfall:</strong> Latitude Longitude: " + mapInfoPointModelInfrastructureOutfallList[0].Lat + " " + mapInfoPointModelInfrastructureOutfallList[0].Lng + "<br>");
                    }
                    else
                    {
                        sb.AppendLine("<strong>Outfall:</strong> Latitude Longitude: <br>");
                    }
                    sb.AppendLine("<br><br>");
                    sb.AppendLine("<h2>Outfall Information</h2>");
                    sb.AppendLine(@"<strong>Average Depth (m):</strong> " + infrastructureModel.AverageDepth_m + "<br>");
                    sb.AppendLine(@"<strong>Decay Rate (/day):</strong> " + infrastructureModel.DecayRate_per_day + "<br>");
                    sb.AppendLine(@"<strong>Distance From Shore (m):</strong> " + infrastructureModel.DistanceFromShore_m + "<br>");
                    sb.AppendLine(@"<strong>Far Field Velocity (m/s):</strong> " + infrastructureModel.FarFieldVelocity_m_s + "<br>");
                    sb.AppendLine(@"<strong>Horizontal Angle (deg):</strong> " + infrastructureModel.HorizontalAngle_deg + "<br>");
                    sb.AppendLine(@"<strong>Near Field Velocity (m/s):</strong> " + infrastructureModel.NearFieldVelocity_m_s + "<br>");
                    sb.AppendLine(@"<strong>Number Of Ports:</strong> " + infrastructureModel.NumberOfPorts + "<br>");
                    sb.AppendLine(@"<strong>Port Diameter (m):</strong> " + infrastructureModel.PortDiameter_m + "<br>");
                    sb.AppendLine(@"<strong>Port Elevation (m):</strong> " + infrastructureModel.PortElevation_m + "<br>");
                    sb.AppendLine(@"<strong>Port Spacing (m):</strong> " + infrastructureModel.PortSpacing_m + "<br>");
                    sb.AppendLine(@"<strong>Receiving Water Concentration (FC /100 ml):</strong> " + infrastructureModel.ReceivingWater_MPN_per_100ml + "<br>");
                    sb.AppendLine(@"<strong>Receiving Water Salinity (PSU):</strong> " + infrastructureModel.ReceivingWaterSalinity_PSU + "<br>");
                    sb.AppendLine(@"<strong>Receiving Water Temperature (ºC):</strong> " + infrastructureModel.ReceivingWaterTemperature_C + "<br>");
                    sb.AppendLine(@"<strong>Vertical Angle (deg):</strong> " + infrastructureModel.VerticalAngle_deg + "<br>");
                    //sb.AppendLine(@"</pre>");
                    sb.AppendLine(@"]]>");
                    sb.AppendLine(@"	</description>");

                    sb.AppendLine(@"		<Point>");
                    sb.AppendLine(@"			<coordinates>" + mapInfoPointModelInfrastructureList[0].Lng + "," + mapInfoPointModelInfrastructureList[0].Lat + ",0</coordinates>");
                    sb.AppendLine(@"		</Point>");
                    sb.AppendLine(@"	</Placemark>");

                    sb.AppendLine(@"	<Placemark>");
                    sb.AppendLine(@"	<name>Outfall " + tvItemModelInfrastructureList.Where(c => c.TVItemID == infrastructureModel.InfrastructureTVItemID).Select(c => c.TVText).FirstOrDefault() + "</name>");
                    
                    if (infrastructureModel.CanOverflow == null)
                    {
                        sb.AppendLine(@"    <styleUrl>#outfallunknown</styleUrl>");
                    }
                    else
                    {
                        if (!(bool)infrastructureModel.CanOverflow)
                        {
                            sb.AppendLine(@"    <styleUrl>#outfallcannotoverflow</styleUrl>");
                        }
                        else
                        {
                            sb.AppendLine(@"    <styleUrl>#outfallcanoverflow</styleUrl>");
                        }
                    }
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
                    sb.AppendLine(@"    <styleUrl>#lineoverflow</styleUrl>");

                    sb.AppendLine(@"	<description>");
                    sb.AppendLine(@"<![CDATA[");
                    //sb.AppendLine(@"<pre>");
                    sb.AppendLine(@"<strong>Alarm System Type:</strong> " + _BaseEnumService.GetEnumText_AlarmSystemTypeEnum(infrastructureModel.AlarmSystemType) + "<br>");
                    sb.AppendLine(@"<strong>Can overflow:</strong> " + infrastructureModel.CanOverflow.ToString() + "<br>");
                    sb.AppendLine(@"<strong>Category:</strong> " + infrastructureModel.InfrastructureCategory + "<br>");
                    sb.AppendLine(@"<strong>Collection System Type:</strong> " + _BaseEnumService.GetEnumText_CollectionSystemTypeEnum(infrastructureModel.CollectionSystemType) + "<br>");
                    sb.AppendLine(@"<strong>Comments:</strong> " + infrastructureModel.Comment + "<br>");
                    sb.AppendLine(@"<strong>Infrastructure Type:</strong> " + _BaseEnumService.GetEnumText_InfrastructureTypeEnum(infrastructureModel.InfrastructureType) + "<br>");
                    sb.AppendLine(@"<strong>Percent Flow Of Total (%):</strong> " + infrastructureModel.PercFlowOfTotal + "<br>");

                    List<MapInfoPointModel> mapInfoPointModelInfrastructureList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(infrastructureModel.InfrastructureTVItemID, TVTypeEnum.LineOverflow, MapInfoDrawTypeEnum.Point);

                    if (mapInfoPointModelInfrastructureList.Count > 0)
                    {
                        sb.AppendLine(@"<strong>Latitude Longitude:</strong> " + mapInfoPointModelInfrastructureList[0].Lat + " " + mapInfoPointModelInfrastructureList[0].Lng + "<br>");
                    }
                    else
                    {
                        sb.AppendLine("<strong>Latitude Longitude:</strong> <br>");
                    }

                    List<MapInfoPointModel> mapInfoPointModelInfrastructureOutfallList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(infrastructureModel.InfrastructureTVItemID, TVTypeEnum.Outfall, MapInfoDrawTypeEnum.Point);

                    if (mapInfoPointModelInfrastructureOutfallList.Count > 0)
                    {
                        sb.AppendLine(@"<strong>Outfall:</strong> Latitude Longitude: " + mapInfoPointModelInfrastructureOutfallList[0].Lat + " " + mapInfoPointModelInfrastructureOutfallList[0].Lng + "<br>");
                    }
                    else
                    {
                        sb.AppendLine("<strong>Outfall:</strong> Latitude Longitude: <br>");
                    }
                    sb.AppendLine("<br><br>");
                    sb.AppendLine(@"<h2>Outfall Information</h2>");
                    sb.AppendLine(@"<strong>Average Depth (m):</strong> " + infrastructureModel.AverageDepth_m + "<br>");
                    sb.AppendLine(@"<strong>Decay Rate (/day):</strong> " + infrastructureModel.DecayRate_per_day + "<br>");
                    sb.AppendLine(@"<strong>Distance From Shore (m):</strong> " + infrastructureModel.DistanceFromShore_m + "<br>");
                    sb.AppendLine(@"<strong>Far Field Velocity (m/s):</strong> " + infrastructureModel.FarFieldVelocity_m_s + "<br>");
                    sb.AppendLine(@"<strong>Horizontal Angle (deg):</strong> " + infrastructureModel.HorizontalAngle_deg + "<br>");
                    sb.AppendLine(@"<strong>Near Field Velocity (m/s):</strong> " + infrastructureModel.NearFieldVelocity_m_s + "<br>");
                    sb.AppendLine(@"<strong>Number Of Ports:</strong> " + infrastructureModel.NumberOfPorts + "<br>");
                    sb.AppendLine(@"<strong>Port Diameter (m):</strong> " + infrastructureModel.PortDiameter_m + "<br>");
                    sb.AppendLine(@"<strong>Port Elevation (m):</strong> " + infrastructureModel.PortElevation_m + "<br>");
                    sb.AppendLine(@"<strong>Port Spacing (m):</strong> " + infrastructureModel.PortSpacing_m + "<br>");
                    sb.AppendLine(@"<strong>Receiving Water Concentration (FC/100 ml):</strong> " + infrastructureModel.ReceivingWater_MPN_per_100ml + "<br>");
                    sb.AppendLine(@"<strong>Receiving Water Salinity (PSU):</strong> " + infrastructureModel.ReceivingWaterSalinity_PSU + "<br>");
                    sb.AppendLine(@"<strong>Receiving Water Temperature (ºC):</strong> " + infrastructureModel.ReceivingWaterTemperature_C + "<br>");
                    sb.AppendLine(@"<strong>Vertical Angle (deg):</strong> " + infrastructureModel.VerticalAngle_deg + "<br>");
                    //sb.AppendLine(@"</pre>");
                    sb.AppendLine(@"]]>");
                    sb.AppendLine(@"	</description>");

                    sb.AppendLine(@"		<Point>");
                    sb.AppendLine(@"			<coordinates>" + mapInfoPointModelInfrastructureList[0].Lng + "," + mapInfoPointModelInfrastructureList[0].Lat + ",0</coordinates>");
                    sb.AppendLine(@"		</Point>");
                    sb.AppendLine(@"	</Placemark>");

                    sb.AppendLine(@"	<Placemark>");
                    sb.AppendLine(@"	<name>Outfall " + tvItemModelInfrastructureList.Where(c => c.TVItemID == infrastructureModel.InfrastructureTVItemID).Select(c => c.TVText).FirstOrDefault() + "</name>");
                    
                    if (infrastructureModel.CanOverflow == null)
                    {
                        sb.AppendLine(@"    <styleUrl>#outfallunknown</styleUrl>");
                    }
                    else
                    {
                        if (!(bool)infrastructureModel.CanOverflow)
                        {
                            sb.AppendLine(@"    <styleUrl>#outfallcannotoverflow</styleUrl>");
                        }
                        else
                        {
                            sb.AppendLine(@"    <styleUrl>#outfallcanoverflow</styleUrl>");
                        }
                    }
                    sb.AppendLine(@"		<Point>");
                    sb.AppendLine(@"			<coordinates>" + mapInfoPointModelInfrastructureOutfallList[0].Lng + "," + mapInfoPointModelInfrastructureOutfallList[0].Lat + ",0</coordinates>");
                    sb.AppendLine(@"		</Point>");
                    sb.AppendLine(@"	</Placemark>");
                }
                sb.AppendLine(@" </Folder>");
            }
            sb.AppendLine(@" </Folder>");

            // Doing Long Pollution Source Site
            sb.AppendLine(@" <Folder>");
            sb.AppendLine(@"	<name>Pollution Source Sites</name>");
            List<PolSourceSiteModel> polSourceSiteModelList = _PolSourceSiteService.GetPolSourceSiteModelListWithSubsectorTVItemIDDB(tvItemModelSubsector.TVItemID);
            List<TVItemModel> tvItemModelListPolSourceSite = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.PolSourceSite);

            foreach (PolSourceSiteModel polSourceSiteModel in polSourceSiteModelList.OrderBy(c => c.Site).ToList())
            {
                TVItemModel tvItemModelPolSourceSite = (from c in tvItemModelListPolSourceSite
                                                        where c.TVItemID == polSourceSiteModel.PolSourceSiteTVItemID
                                                        select c).FirstOrDefault();

                string labelPolSourceSite = tvItemModelPolSourceSite.TVText;
                if (labelPolSourceSite.Contains(" "))
                {
                    labelPolSourceSite = labelPolSourceSite.Substring(0, labelPolSourceSite.IndexOf(" "));
                }

                // Doing point
                sb.AppendLine(@"	<Placemark>");
                sb.AppendLine(@"	<name>" + labelPolSourceSite + "</name>");
                sb.AppendLine(@"    <styleUrl>#pollutionsourcesite</styleUrl>");
                sb.AppendLine(@"	<description>");
                sb.AppendLine(@"<![CDATA[");
                //sb.AppendLine(@"<pre>");

                PolSourceObservationModel polSourceObservationModel = _PolSourceSiteService._PolSourceObservationService.GetPolSourceObservationModelLatestWithPolSourceSiteIDDB(polSourceSiteModel.PolSourceSiteID);
                if (polSourceObservationModel != null)
                {
                    List<PolSourceObservationIssueModel> polSourceObservationIssueModelList = _PolSourceSiteService._PolSourceObservationService._PolSourceObservationIssueService.GetPolSourceObservationIssueModelListWithPolSourceObservationIDDB(polSourceObservationModel.PolSourceObservationID);

                    string SelectedObservation = "<strong>Selected:</strong> <br>";
                    foreach (PolSourceObservationIssueModel polSourceObservationIssueModel in polSourceObservationIssueModelList)
                    {
                        foreach (PolSourceObsInfoEnum polSourceObsInfo in polSourceObservationIssueModel.PolSourceObsInfoList)
                        {
                            SelectedObservation += _BaseEnumService.GetEnumText_PolSourceObsInfoReportEnum(polSourceObsInfo);
                        }
                        SelectedObservation += "<br><br>";
                    }

                    sb.AppendLine("<strong>Written:</strong> <br>" + (string.IsNullOrWhiteSpace(polSourceObservationModel.Observation_ToBeDeleted) ? "" : polSourceObservationModel.Observation_ToBeDeleted.ToString()) + "<br><br>" + SelectedObservation);
                }
                else
                {
                    string SelectedObservation = "<strong>Selected:</strong> <br>";
                    sb.AppendLine("<strong>Written:</strong> <br><br>" + SelectedObservation);
                }

                //sb.AppendLine(@"</pre>");
                sb.AppendLine(@"]]>");
                sb.AppendLine(@"	</description>");

                mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelPolSourceSite.TVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);

                sb.AppendLine(@"		<Point>");
                sb.AppendLine(@"			<coordinates>" + mapInfoPointModelList[0].Lng + "," + mapInfoPointModelList[0].Lat + ",0</coordinates>");
                sb.AppendLine(@"		</Point>");
                sb.AppendLine(@"	</Placemark>");

            }
            sb.AppendLine(@" </Folder>");

            List<TVItemModel> tvItemModelMWQMSiteList = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.MWQMSite);

            // Doing MWQM Site
            sb.AppendLine(@" <Folder>");
            sb.AppendLine(@"	<name>MWQM Sites</name>");

            foreach (TVItemModel tvItemModelMWQMSite in tvItemModelMWQMSiteList.Where(c => c.IsActive))
            {
                // Doing point
                sb.AppendLine(@"	<Placemark>");
                sb.AppendLine(@"	<name>" + tvItemModelMWQMSite.TVText + "</name>");
                sb.AppendLine(@"    <styleUrl>#monitoringsite</styleUrl>");
                //sb.AppendLine(@"	<description>");
                //sb.AppendLine(@"<![CDATA[");
                //sb.AppendLine(@"                                <p><a href=""" + _TaskRunnerBaseService.GetUrlFromTVItem(tvItemModelMWQMSite) + @""">" + tvItemModelMWQMSite.TVText + "</a></p>");
                //sb.AppendLine(@"]]>");
                //sb.AppendLine(@"	</description>");

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
                NotUsed = string.Format(TaskRunnerServiceRes.ErrorOccuredWhileCreating_Document_, "KMZ", fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("ErrorOccuredWhileCreating_Document_", "KMZ", fi.FullName);
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
                NotUsed = string.Format(TaskRunnerServiceRes.ErrorOccuredWhileCreating_Document_, "KMZ", fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("ErrorOccuredWhileCreating_Document_", "KMZ", fi.FullName);
                return false;
            }

            return true;
        }



    }
}
