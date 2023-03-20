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

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 10);
            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("DoingStyle");
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);

            sb.AppendLine(@"	<name>" + fi.Name + "</name>");

            #region style
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

            // red polygon
            sb.AppendLine(@"	<Style id=""red-polygon"">");
            sb.AppendLine(@"        <LineStyle>");
            sb.AppendLine(@"			<color>ff0000ff</color>");
            sb.AppendLine(@"			<width>2</width>");
            sb.AppendLine(@"		</LineStyle>");
            sb.AppendLine(@"		<PolyStyle>");
            sb.AppendLine(@"			<color>ff0000ff</color>");
            sb.AppendLine(@"		</PolyStyle>");
            sb.AppendLine(@"	</Style>");

            // green polygon
            sb.AppendLine(@"	<Style id=""green-polygon"">");
            sb.AppendLine(@"        <LineStyle>");
            sb.AppendLine(@"			<color>00ff00ff</color>");
            sb.AppendLine(@"			<width>2</width>");
            sb.AppendLine(@"		</LineStyle>");
            sb.AppendLine(@"		<PolyStyle>");
            sb.AppendLine(@"			<color>00ff00ff</color>");
            sb.AppendLine(@"		</PolyStyle>");
            sb.AppendLine(@"	</Style>");

            // blue polygon
            sb.AppendLine(@"	<Style id=""blue-polygon"">");
            sb.AppendLine(@"        <LineStyle>");
            sb.AppendLine(@"			<color>0000ffff</color>");
            sb.AppendLine(@"			<width>2</width>");
            sb.AppendLine(@"		</LineStyle>");
            sb.AppendLine(@"		<PolyStyle>");
            sb.AppendLine(@"			<color>0000ffff</color>");
            sb.AppendLine(@"		</PolyStyle>");
            sb.AppendLine(@"	</Style>");

            // Municipality
            sb.AppendLine(@"	<Style id=""municipality"">");
            sb.AppendLine(@"	    <IconStyle>");
            sb.AppendLine(@"	    	<Icon>");
            sb.AppendLine(@"            <href>https://maps.google.com/mapfiles/kml/pal3/icon28.png</href>");
            sb.AppendLine(@"	    	</Icon>");
            sb.AppendLine(@"        </IconStyle>");
            sb.AppendLine(@"	</Style>");

            // WWTP
            sb.AppendLine(@"	<Style id=""wwtp"">");
            sb.AppendLine(@"	    <IconStyle>");
            sb.AppendLine(@"	    	<Icon>");
            sb.AppendLine(@"            <href>https://maps.google.com/mapfiles/kml/pal5/icon22.png</href>");
            sb.AppendLine(@"	    	</Icon>");
            sb.AppendLine(@"        </IconStyle>");
            sb.AppendLine(@"	</Style>");

            // Outfall can overflow
            sb.AppendLine(@"	<Style id=""outfall_can_overflow"">");
            sb.AppendLine(@"	    <IconStyle>");
            sb.AppendLine(@"	    	<Icon>");
            sb.AppendLine(@"            <href>https://maps.google.com/mapfiles/kml/pal5/icon38.png</href>");
            sb.AppendLine(@"	    	</Icon>");
            sb.AppendLine(@"        </IconStyle>");
            sb.AppendLine(@"	</Style>");

            // Outfall can not overflow
            sb.AppendLine(@"	<Style id=""outfall_cannot_overflow"">");
            sb.AppendLine(@"	    <IconStyle>");
            sb.AppendLine(@"	    	<Icon>");
            sb.AppendLine(@"            <href>https://maps.google.com/mapfiles/kml/pal4/icon17.png</href>");
            sb.AppendLine(@"	    	</Icon>");
            sb.AppendLine(@"        </IconStyle>");
            sb.AppendLine(@"	</Style>");

            // Outfall unknown
            sb.AppendLine(@"	<Style id=""outfall_unknown"">");
            sb.AppendLine(@"	    <IconStyle>");
            sb.AppendLine(@"	    	<Icon>");
            sb.AppendLine(@"            <href>	https://maps.google.com/mapfiles/kml/pal3/icon49.png</href>");
            sb.AppendLine(@"	    	</Icon>");
            sb.AppendLine(@"        </IconStyle>");
            sb.AppendLine(@"	</Style>");

            // Lift station
            sb.AppendLine(@"	<Style id=""lift_station"">");
            sb.AppendLine(@"	    <IconStyle>");
            sb.AppendLine(@"	    	<Icon>");
            sb.AppendLine(@"            <href>	https://maps.google.com/mapfiles/kml/pal5/icon35.png</href>");
            sb.AppendLine(@"	    	</Icon>");
            sb.AppendLine(@"        </IconStyle>");
            sb.AppendLine(@"	</Style>");

            // Line overflow
            sb.AppendLine(@"	<Style id=""line_overflow"">");
            sb.AppendLine(@"	    <IconStyle>");
            sb.AppendLine(@"	    	<Icon>");
            sb.AppendLine(@"            <href>https://maps.google.com/mapfiles/kml/pal4/icon58.png</href>");
            sb.AppendLine(@"	    	</Icon>");
            sb.AppendLine(@"        </IconStyle>");
            sb.AppendLine(@"	</Style>");

            // Default Infrastructure
            sb.AppendLine(@"	<Style id=""default_infrastructure"">");
            sb.AppendLine(@"	    <IconStyle>");
            sb.AppendLine(@"	    	<Icon>");
            sb.AppendLine(@"            <href>https://maps.google.com/mapfiles/kml/pal5/icon32.png</href>");
            sb.AppendLine(@"	    	</Icon>");
            sb.AppendLine(@"        </IconStyle>");
            sb.AppendLine(@"	</Style>");

            // Pollution source site (High Risk)
            sb.AppendLine(@"	<Style id=""pollutionsourcesite_high_risk"">");
            sb.AppendLine(@"	    <IconStyle>");
            sb.AppendLine(@"	    	<Icon>");
            sb.AppendLine(@"            <href>https://maps.google.com/mapfiles/kml/pal3/icon39.png</href>");
            sb.AppendLine(@"	    	</Icon>");
            sb.AppendLine(@"        </IconStyle>");
            sb.AppendLine(@"	</Style>");      

            // Pollution source site (Moderate Risk)
            sb.AppendLine(@"	<Style id=""pollutionsourcesite_moderate_risk"">");
            sb.AppendLine(@"	    <IconStyle>");
            sb.AppendLine(@"	    	<Icon>");
            sb.AppendLine(@"            <href>https://maps.google.com/mapfiles/kml/pal3/icon33.png</href>");
            sb.AppendLine(@"	    	</Icon>");
            sb.AppendLine(@"        </IconStyle>");
            sb.AppendLine(@"	</Style>");

            // Pollution source site (Low Risk)
            sb.AppendLine(@"	<Style id=""pollutionsourcesite_low_risk"">");
            sb.AppendLine(@"	    <IconStyle>");
            sb.AppendLine(@"	    	<Icon>");
            sb.AppendLine(@"            <href>https://maps.google.com/mapfiles/kml/pal3/icon42.png</href>");
            sb.AppendLine(@"	    	</Icon>");
            sb.AppendLine(@"        </IconStyle>");
            sb.AppendLine(@"	</Style>");

            // Pollution source site (Unknown Risk)
            sb.AppendLine(@"	<Style id=""pollutionsourcesite_unknown_risk"">");
            sb.AppendLine(@"	    <IconStyle>");
            sb.AppendLine(@"	    	<Icon>");
            sb.AppendLine(@"            <href>https://maps.google.com/mapfiles/kml/pal3/icon35.png</href>");
            sb.AppendLine(@"	    	</Icon>");
            sb.AppendLine(@"        </IconStyle>");
            sb.AppendLine(@"	</Style>");

            sb.AppendLine(@"  <Style id=""A"">");
            sb.AppendLine(@"	<IconStyle>");
            sb.AppendLine(@"		<Icon>");
            sb.AppendLine(@"        <href>https://maps.google.com/mapfiles/kml/paddle/grn-blank.png</href>");
            sb.AppendLine(@"		</Icon>");
            sb.AppendLine(@"    </IconStyle>");
            sb.AppendLine(@"    <LineStyle>");
            sb.AppendLine(@"      <color>ff00ff00</color>");
            sb.AppendLine(@"      <width>1</width>");
            sb.AppendLine(@"  	</LineStyle>");
            sb.AppendLine(@"    <PolyStyle>");
            sb.AppendLine(@"  	  <color>ff00ff00</color>");
            sb.AppendLine(@"    </PolyStyle>");
            sb.AppendLine(@"  </Style>");

            sb.AppendLine(@"  <Style id=""R"">");
            sb.AppendLine(@"	<IconStyle>");
            sb.AppendLine(@"		<Icon>");
            sb.AppendLine(@"        <href>https://maps.google.com/mapfiles/kml/paddle/red-blank.png</href>");
            sb.AppendLine(@"		</Icon>");
            sb.AppendLine(@"    </IconStyle>");
            sb.AppendLine(@"    <LineStyle>");
            sb.AppendLine(@"      <color>ff0000ff</color>");
            sb.AppendLine(@"      <width>1</width>");
            sb.AppendLine(@"    </LineStyle>");
            sb.AppendLine(@"    <PolyStyle>");
            sb.AppendLine(@"      <color>ff0000ff</color>");
            sb.AppendLine(@"    </PolyStyle>");
            sb.AppendLine(@"  </Style>");

            sb.AppendLine(@"  <Style id=""P"">");
            sb.AppendLine(@"	<IconStyle>");
            sb.AppendLine(@"		<Icon>");
            sb.AppendLine(@"        <href>https://maps.google.com/mapfiles/kml/paddle/purple-blank.png</href>");
            sb.AppendLine(@"		</Icon>");
            sb.AppendLine(@"    </IconStyle>");
            sb.AppendLine(@"    <LineStyle>");
            sb.AppendLine(@"      <color>ffcccccc</color>");
            sb.AppendLine(@"      <width>1</width>");
            sb.AppendLine(@"    </LineStyle>");
            sb.AppendLine(@"    <PolyStyle>");
            sb.AppendLine(@"      <color>ffcccccc</color>");
            sb.AppendLine(@"    </PolyStyle>");
            sb.AppendLine(@"  </Style>");

            sb.AppendLine(@"  <Style id=""CA"">");
            sb.AppendLine(@"	<IconStyle>");
            sb.AppendLine(@"		<Icon>");
            sb.AppendLine(@"        <href>https://maps.google.com/mapfiles/kml/paddle/grn-blank.png</href>");
            sb.AppendLine(@"		</Icon>");
            sb.AppendLine(@"    </IconStyle>");
            sb.AppendLine(@"    <LineStyle>");
            sb.AppendLine(@"      <color>ff00ffff</color>");
            sb.AppendLine(@"      <width>1</width>");
            sb.AppendLine(@"    </LineStyle>");
            sb.AppendLine(@"    <PolyStyle>");
            sb.AppendLine(@"     <color>ff00ffff</color>");
            sb.AppendLine(@"    </PolyStyle>");
            sb.AppendLine(@"  </Style>");

            sb.AppendLine(@"  <Style id=""CR"">");
            sb.AppendLine(@"	<IconStyle>");
            sb.AppendLine(@"		<Icon>");
            sb.AppendLine(@"        <href>https://maps.google.com/mapfiles/kml/paddle/red-blank.png</href>");
            sb.AppendLine(@"		</Icon>");
            sb.AppendLine(@"    </IconStyle>");
            sb.AppendLine(@"    <LineStyle>");
            sb.AppendLine(@"      <color>ffff00aa</color>");
            sb.AppendLine(@"      <width>1</width>");
            sb.AppendLine(@"    </LineStyle>");
            sb.AppendLine(@"    <PolyStyle>");
            sb.AppendLine(@"      <color>ffff00aa</color>");
            sb.AppendLine(@"    </PolyStyle>");
            sb.AppendLine(@"  </Style>");

            #endregion style

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 15);
            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("DoingSubsectorPolygon");
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);



            sb.AppendLine(@"	<Folder>");

            #region subsector

            TVItemModel tvItemModelSubsector = _TVItemService.GetTVItemModelWithTVItemIDDB(TVItemID);
            List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.Subsector, MapInfoDrawTypeEnum.Point);
            double latSSCenter = 0;
            double lngSSCenter = 0;
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
            latSSCenter = (from c in mapInfoPointModelList
                           select c.Lat).Average();
            lngSSCenter = (from c in mapInfoPointModelList
                           select c.Lng).Average();
            foreach (MapInfoPointModel mapInfoPointModel in mapInfoPointModelList)
            {
                sb.AppendLine(mapInfoPointModel.Lng + "," + mapInfoPointModel.Lat + ",0 ");
            }

            sb.AppendLine(@"					</coordinates>");
            sb.AppendLine(@"				</LinearRing>");
            sb.AppendLine(@"			</outerBoundaryIs>");
            sb.AppendLine(@"		</Polygon>");
            sb.AppendLine(@"	</Placemark>");

            #endregion subsector

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 20);
            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("DoingClassification");
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);

            #region classification

            // Doing Long Pollution Source Site
            sb.AppendLine("     <Folder>");
            sb.AppendLine($"        <name>Classification</name>");

            var ClassificationList = (from t in _MapInfoService.db.TVItems
                                      from c in _MapInfoService.db.Classifications
                                      from mi in _MapInfoService.db.MapInfos
                                      let mipList = (from mip in _MapInfoService.db.MapInfoPoints
                                                     where mip.MapInfoID == mi.MapInfoID
                                                     select mip).ToList()
                                      where c.ClassificationTVItemID == t.TVItemID
                                      && mi.TVItemID == t.TVItemID
                                      && t.TVPath.StartsWith(tvItemModelSubsector.TVPath + "p")
                                      && t.TVType == (int)TVTypeEnum.Classification
                                      && mi.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Polyline
                                      select new { t, mipList, c }).ToList();


            foreach (var classification in ClassificationList)
            {
                string TVTextClass = "";

                switch (((ClassificationTypeEnum)classification.c.ClassificationType))
                {
                    case ClassificationTypeEnum.Approved:
                        {
                            TVTextClass = "A";
                        }
                        break;
                    case ClassificationTypeEnum.Restricted:
                        {
                            TVTextClass = "R";
                        }
                        break;
                    case ClassificationTypeEnum.Prohibited:
                        {
                            TVTextClass = "P";
                        }
                        break;
                    case ClassificationTypeEnum.ConditionallyApproved:
                        {
                            TVTextClass = "CA";
                        }
                        break;
                    case ClassificationTypeEnum.ConditionallyRestricted:
                        {
                            TVTextClass = "CR";
                        }
                        break;
                    default:
                        {
                            TVTextClass = "E";
                        }
                        break;
                }

                //string style = "msn_ylw-pushpin";

                List<MapInfoPoint> mapInfoPointList = classification.mipList.OrderBy(c => c.Ordinal).ToList();

                for (int i = 0, count = mapInfoPointList.Count - 1; i < count; i++)
                {
                    GeoLocation geo1 = new GeoLocation() { Latitude = (float)classification.mipList[i].Lat, Longitude = (float)classification.mipList[i].Lng };
                    GeoLocation geo2 = new GeoLocation() { Latitude = (float)classification.mipList[i + 1].Lat, Longitude = (float)classification.mipList[i + 1].Lng };
                    GeoLocation geo3 = new GeoLocation();
                    GeoLocation geo4 = new GeoLocation();

                    geo3.Longitude = (geo1.Longitude + geo2.Longitude) / 2.0D;
                    geo3.Latitude = (geo1.Latitude + geo2.Latitude) / 2.0D;

                    double xDiff = geo2.Longitude - geo1.Longitude;
                    double yDiff = geo2.Latitude - geo1.Latitude;
                    double rad = Math.Atan2(yDiff, xDiff);

                    double ABY = geo2.Latitude - geo1.Latitude;
                    double ABX = geo2.Longitude - geo1.Longitude;
                    double Len = Math.Sqrt(ABY * ABY + ABX * ABX);
                    geo4.Longitude = geo3.Longitude + 0.001D * ABY / Len;
                    geo4.Latitude = geo3.Latitude - 0.001D * ABX / Len;

                    sb.AppendLine(@"        <Placemark>");
                    sb.AppendLine($@"            <name>{TVTextClass}</name>");
                    sb.AppendLine($@"            <styleUrl>#{TVTextClass}</styleUrl>");
                    sb.AppendLine(@"            <Polygon>");
                    sb.AppendLine(@"		    	<tessellate>1</tessellate>");
                    sb.AppendLine(@"		    	<outerBoundaryIs>");
                    sb.AppendLine(@"		    		<LinearRing>");
                    sb.AppendLine(@"		    			<coordinates>");
                    sb.AppendLine($@"		    				{geo1.Longitude},{geo1.Latitude},0 {geo2.Longitude},{geo2.Latitude},0 {geo4.Longitude},{geo4.Latitude},0 {geo1.Longitude},{geo1.Latitude},0 ");
                    sb.AppendLine(@"		    			</coordinates>");
                    sb.AppendLine(@"		    		</LinearRing>");
                    sb.AppendLine(@"		    	</outerBoundaryIs>");
                    sb.AppendLine(@"		    </Polygon>");
                    sb.AppendLine(@"        </Placemark>");

                }
            }
            sb.AppendLine("     </Folder>");

            #endregion classification

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 40);
            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("DoingPollutionSourceSite");
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);

            #region pollution source site

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

                if (tvItemModelPolSourceSite == null) continue;
                if (!tvItemModelPolSourceSite.IsActive) continue;

                string labelPolSourceSite = tvItemModelPolSourceSite.TVText;
                if (labelPolSourceSite.Contains(" "))
                {
                    labelPolSourceSite = labelPolSourceSite.Substring(0, labelPolSourceSite.IndexOf(" "));
                }

                // Doing point

                PolSourceObservationModel polSourceObservationModel = _PolSourceSiteService._PolSourceObservationService.GetPolSourceObservationModelLatestWithPolSourceSiteIDDB(polSourceSiteModel.PolSourceSiteID);
                if (polSourceObservationModel != null)
                {
                    List<PolSourceObservationIssueModel> polSourceObservationIssueModelList = _PolSourceSiteService._PolSourceObservationService._PolSourceObservationIssueService.GetPolSourceObservationIssueModelListWithPolSourceObservationIDDB(polSourceObservationModel.PolSourceObservationID);

                    string risk = "unknown";

                    if (polSourceObservationIssueModelList.Count > 0)
                    {
                        if (polSourceObservationIssueModelList[0].PolSourceObsInfoList.Contains(PolSourceObsInfoEnum.RiskHigh))
                        {
                            risk = "high";
                        }

                        if (polSourceObservationIssueModelList[0].PolSourceObsInfoList.Contains(PolSourceObsInfoEnum.RiskModerate))
                        {
                            risk = "moderate";
                        }

                        if (polSourceObservationIssueModelList[0].PolSourceObsInfoList.Contains(PolSourceObsInfoEnum.RiskLow))
                        {
                            risk = "low";
                        }
                    }

                    sb.AppendLine(@"	<Placemark>");
                    sb.AppendLine($@"	<name>{ labelPolSourceSite } ({ risk })</name>");
                    sb.AppendLine($@"    <styleUrl>#pollutionsourcesite_{risk}_risk</styleUrl>");
                    sb.AppendLine(@"	<description>");
                    sb.AppendLine(@"<![CDATA[");
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
                    sb.AppendLine(@"]]>");
                    sb.AppendLine(@"	</description>");
                }
                else
                {
                    sb.AppendLine(@"	<Placemark>");
                    sb.AppendLine($@"	<name>{labelPolSourceSite} (unknown)</name>");
                    sb.AppendLine(@"    <styleUrl>#pollutionsourcesite_unknown_risk</styleUrl>");
                    sb.AppendLine(@"	<description>");
                    sb.AppendLine(@"<![CDATA[");
                    string SelectedObservation = "<strong>Selected:</strong> <br>";
                    sb.AppendLine("<strong>Written:</strong> <br><br>" + SelectedObservation);
                    sb.AppendLine(@"]]>");
                    sb.AppendLine(@"	</description>");
                }


                mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelPolSourceSite.TVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);

                sb.AppendLine(@"		<Point>");
                sb.AppendLine(@"			<coordinates>" + mapInfoPointModelList[0].Lng + "," + mapInfoPointModelList[0].Lat + ",0</coordinates>");
                sb.AppendLine(@"		</Point>");
                sb.AppendLine(@"	</Placemark>");

            }
            sb.AppendLine(@" </Folder>");

            #endregion pollution source site

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 60);
            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("DoingMonitoringSites");
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);

            #region MWQM site

            List<TVItemModel> tvItemModelMWQMSiteList = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.MWQMSite);
            List<MWQMSiteModel> mwqmSiteModelList = _MWQMSiteService.GetMWQMSiteModelListWithSubsectorTVItemIDDB(tvItemModelSubsector.TVItemID);


            // Doing MWQM Site
            sb.AppendLine(@" <Folder>");
            sb.AppendLine(@"	<name>MWQM Sites</name>");

            foreach (TVItemModel tvItemModelMWQMSite in tvItemModelMWQMSiteList.Where(c => c.IsActive))
            {
                MWQMSiteModel mwqmSiteModel = (from c in mwqmSiteModelList
                                               where c.MWQMSiteTVItemID == tvItemModelMWQMSite.TVItemID
                                               select c).FirstOrDefault();

                string TVTextCurrentClass = "E";

                switch ((MWQMSiteLatestClassificationEnum)mwqmSiteModel.MWQMSiteLatestClassification)
                {
                    case MWQMSiteLatestClassificationEnum.Approved:
                        {
                            TVTextCurrentClass = "A";
                        }
                        break;
                    case MWQMSiteLatestClassificationEnum.Restricted:
                        {
                            TVTextCurrentClass = "R";
                        }
                        break;
                    case MWQMSiteLatestClassificationEnum.Prohibited:
                        {
                            TVTextCurrentClass = "P";
                        }
                        break;
                    case MWQMSiteLatestClassificationEnum.ConditionallyApproved:
                        {
                            TVTextCurrentClass = "CA";
                        }
                        break;
                    case MWQMSiteLatestClassificationEnum.ConditionallyRestricted:
                        {
                            TVTextCurrentClass = "CR";
                        }
                        break;
                    default:
                        {
                            TVTextCurrentClass = "E";
                        }
                        break;
                }

                // Doing point
                sb.AppendLine(@"	<Placemark>");
                sb.AppendLine($@"	<name>{tvItemModelMWQMSite.TVText} ({ TVTextCurrentClass})</name>");
                sb.AppendLine($@"    <styleUrl>#{ TVTextCurrentClass}</styleUrl>");
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

            #endregion MWQM site

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 80);
            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("DoingMunicipalities");
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);

            #region municipality

            // Doing Municipalities
            sb.AppendLine(@" <Folder>");
            sb.AppendLine(@"	<name>Municipalities</name>");

            List<UseOfSiteModel> useOfSiteModelList = _UseOfSiteService.GetUseOfSiteModelListWithSubsectorTVItemIDDB(tvItemModelSubsector.TVItemID);

            List<TVItemModel> tvItemModelMunicipalityList = new List<TVItemModel>();

            foreach (UseOfSiteModel useOfSiteModel in useOfSiteModelList)
            {
                if (useOfSiteModel.TVType == TVTypeEnum.Municipality)
                {
                    tvItemModelMunicipalityList.Add(_TVItemService.GetTVItemModelWithTVItemIDDB(useOfSiteModel.SiteTVItemID));
                }
            }

            int notFoundCount = 0;
            foreach (TVItemModel tvItemModelMunicipality in tvItemModelMunicipalityList.OrderBy(c => c.TVText))
            {
                sb.AppendLine(@" <Folder>");
                sb.AppendLine(@"	<name>" + tvItemModelMunicipality.TVText + "</name>");
                // Doing point
                sb.AppendLine(@"	<Placemark>");
                sb.AppendLine(@"	<name>" + tvItemModelMunicipality.TVText + " ( Point)</name>");
                sb.AppendLine(@"    <styleUrl>#municipality</styleUrl>");

                mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelMunicipality.TVItemID, TVTypeEnum.Municipality, MapInfoDrawTypeEnum.Point);

                if (mapInfoPointModelList.Count == 0)
                {
                    notFoundCount += 1;
                    sb.AppendLine(@"		<Point>");
                    sb.AppendLine(@"			<coordinates>" + (lngSSCenter + (notFoundCount*0.001)) + "," + (latSSCenter + (notFoundCount * 0.001)) + ",0</coordinates>");
                    sb.AppendLine(@"		</Point>");
                    sb.AppendLine(@"	</Placemark>");

                }
                else
                {
                    sb.AppendLine(@"		<Point>");
                    sb.AppendLine(@"			<coordinates>" + mapInfoPointModelList[0].Lng + "," + mapInfoPointModelList[0].Lat + ",0</coordinates>");
                    sb.AppendLine(@"		</Point>");
                    sb.AppendLine(@"	</Placemark>");
                }

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
                            sb.AppendLine(@"    <styleUrl>#lift_station</styleUrl>");
                            break;
                        case InfrastructureTypeEnum.LineOverflow:
                            sb.AppendLine(@"    <styleUrl>#line_overflow</styleUrl>");
                            break;
                        default:
                            sb.AppendLine(@"    <styleUrl>#default_infrastructure</styleUrl>");
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
                        sb.AppendLine(@"    <styleUrl>#outfall_unknown</styleUrl>");
                    }
                    else
                    {
                        if (!(bool)infrastructureModel.CanOverflow)
                        {
                            sb.AppendLine(@"    <styleUrl>#outfall_cannot_overflow</styleUrl>");
                        }
                        else
                        {
                            sb.AppendLine(@"    <styleUrl>#outfall_can_overflow</styleUrl>");
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
                    sb.AppendLine(@"    <styleUrl>#lift_station</styleUrl>");

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
                        sb.AppendLine(@"    <styleUrl>#outfall_unknown</styleUrl>");
                    }
                    else
                    {
                        if (!(bool)infrastructureModel.CanOverflow)
                        {
                            sb.AppendLine(@"    <styleUrl>#outfall_cannot_overflow</styleUrl>");
                        }
                        else
                        {
                            sb.AppendLine(@"    <styleUrl>#outfall_can_overflow</styleUrl>");
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
                    sb.AppendLine(@"    <styleUrl>#line_overflow</styleUrl>");

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
                        sb.AppendLine(@"    <styleUrl>#outfall_unknown</styleUrl>");
                    }
                    else
                    {
                        if (!(bool)infrastructureModel.CanOverflow)
                        {
                            sb.AppendLine(@"    <styleUrl>#outfall_cannot_overflow</styleUrl>");
                        }
                        else
                        {
                            sb.AppendLine(@"    <styleUrl>#outfall_can_overflow</styleUrl>");
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

            #endregion municipality

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 80);
            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("SavingDocument");
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);

            sb.AppendLine(@"	</Folder>");

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

            _TaskRunnerBaseService._BWObj.TextLanguageList = new List<TextLanguage>();

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

    public struct GeoLocation
    {
        public double Latitude { get; set; }
        public double Longitude { get; set; }
    }
}
