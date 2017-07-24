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
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;

namespace CSSPWebToolsTaskRunner.Services
{
    public class KmzServiceMunicipality
    {
        #region Variables
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        #endregion Variables

        #region Constructors
        public KmzServiceMunicipality(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
        }
        #endregion Constructors

        #region Functions public
        public void Generate(FileInfo fi)
        {
            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
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

            StringBuilder sbKMZ = new StringBuilder();
            StringBuilder sbKMZWWTP = new StringBuilder();
            StringBuilder sbKMZLS = new StringBuilder();
            StringBuilder sbKMZOutfall = new StringBuilder();
            StringBuilder sbKMZOther = new StringBuilder();

            sbKMZ.AppendLine(@"<?xml version=""1.0"" encoding=""UTF-8""?>");
            sbKMZ.AppendLine(@"<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom"">");
            sbKMZ.AppendLine(@"<Document>");
            sbKMZ.AppendLine(@"	<name>" + _TaskRunnerBaseService.generateDocParams.FileName + "</name>");

            sbKMZ.AppendLine(@"    <StyleMap id=""msn_ylw-pushpin"">");
            sbKMZ.AppendLine(@"		<Pair>");
            sbKMZ.AppendLine(@"			<key>normal</key>");
            sbKMZ.AppendLine(@"			<styleUrl>#sn_ylw-pushpin</styleUrl>");
            sbKMZ.AppendLine(@"		</Pair>");
            sbKMZ.AppendLine(@"		<Pair>");
            sbKMZ.AppendLine(@"			<key>highlight</key>");
            sbKMZ.AppendLine(@"			<styleUrl>#sh_ylw-pushpin</styleUrl>");
            sbKMZ.AppendLine(@"		</Pair>");
            sbKMZ.AppendLine(@"	</StyleMap>");
            sbKMZ.AppendLine(@"	<Style id=""sh_ylw-pushpin"">");
            sbKMZ.AppendLine(@"		<IconStyle>");
            sbKMZ.AppendLine(@"			<scale>1.2</scale>");
            sbKMZ.AppendLine(@"		</IconStyle>");
            sbKMZ.AppendLine(@"		<LineStyle>");
            sbKMZ.AppendLine(@"			<color>ff00ff00</color>");
            sbKMZ.AppendLine(@"			<width>1.5</width>");
            sbKMZ.AppendLine(@"		</LineStyle>");
            sbKMZ.AppendLine(@"		<PolyStyle>");
            sbKMZ.AppendLine(@"			<color>0000ff00</color>");
            sbKMZ.AppendLine(@"		</PolyStyle>");
            sbKMZ.AppendLine(@"	</Style>");
            sbKMZ.AppendLine(@"	<Style id=""sn_ylw-pushpin"">");
            sbKMZ.AppendLine(@"		<LineStyle>");
            sbKMZ.AppendLine(@"			<color>ff00ff00</color>");
            sbKMZ.AppendLine(@"			<width>1.5</width>");
            sbKMZ.AppendLine(@"		</LineStyle>");
            sbKMZ.AppendLine(@"		<PolyStyle>");
            sbKMZ.AppendLine(@"			<color>0000ff00</color>");
            sbKMZ.AppendLine(@"		</PolyStyle>");
            sbKMZ.AppendLine(@"	</Style>");

            TVItemModel tvItemModelMunicipality = tvItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            List<MapInfoPointModel> mapInfoPointModelMuniList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelMunicipality.TVItemID, TVTypeEnum.Municipality, MapInfoDrawTypeEnum.Point);

            sbKMZ.AppendLine(@"	<Folder>");
            sbKMZ.AppendLine(@"	<name>Municipality</name>");

            // Doing Point
            sbKMZ.AppendLine(@"	<Placemark>");
            sbKMZ.AppendLine(@"	<name>" + tvItemModelMunicipality.TVText + "</name>");
            sbKMZ.AppendLine(@"<styleUrl>#msn_ylw-pushpin</styleUrl>");
            //sbKMZ.AppendLine(@"	<description>");
            //sbKMZ.AppendLine(@"<![CDATA[");
            //sbKMZ.AppendLine(@"<a href=""" + _TaskRunnerBaseService.GetUrlFromTVItem(tvItemModelRoot) + @""">" + tvItemModelRoot.TVText + "</a>");
            //sbKMZ.AppendLine(@"]]>");
            //sbKMZ.AppendLine(@"	</description>");
            sbKMZ.AppendLine(@"	<Point>");
            sbKMZ.AppendLine(@"		<coordinates>" + mapInfoPointModelMuniList[0].Lng + "," + mapInfoPointModelMuniList[0].Lat + ",0</coordinates>");
            sbKMZ.AppendLine(@"	</Point>");
            sbKMZ.AppendLine(@"	</Placemark>");

            // Doing Infrastructures
            sbKMZ.AppendLine(@" <Folder>");
            sbKMZ.AppendLine(@"	<name>Infrastructures</name>");

            sbKMZWWTP.AppendLine(@" <Folder>");
            sbKMZWWTP.AppendLine(@"	<name>Waste Water Treatment Plants</name>");
            sbKMZLS.AppendLine(@" <Folder>");
            sbKMZLS.AppendLine(@"	<name>Lift Stations</name>");
            sbKMZOutfall.AppendLine(@" <Folder>");
            sbKMZOutfall.AppendLine(@"	<name>Outfalls</name>");
            sbKMZOther.AppendLine(@" <Folder>");
            sbKMZOther.AppendLine(@"	<name>Others</name>");

            List<TVItemModel> tvItemModelInfrastructuresList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelMunicipality.TVItemID, TVTypeEnum.Infrastructure);

            foreach (TVItemModel tvItemModelInfrastructure in tvItemModelInfrastructuresList)
            {
                InfrastructureModel infrastructureModel = infrastructureService.GetInfrastructureModelWithInfrastructureTVItemIDDB(tvItemModelInfrastructure.TVItemID);

                List<MapInfoPointModel> mapInfoPointModelInfList = new List<MapInfoPointModel>();
                List<MapInfoPointModel> mapInfoPointModelInfOutList = new List<MapInfoPointModel>();

                if (infrastructureModel.InfrastructureType == InfrastructureTypeEnum.WWTP)
                {
                    mapInfoPointModelInfList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelInfrastructure.TVItemID, TVTypeEnum.WasteWaterTreatmentPlant, MapInfoDrawTypeEnum.Point);
                    mapInfoPointModelInfOutList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelInfrastructure.TVItemID, TVTypeEnum.Outfall, MapInfoDrawTypeEnum.Point);
                }
                else if (infrastructureModel.InfrastructureType == InfrastructureTypeEnum.LiftStation)
                {
                    mapInfoPointModelInfList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelInfrastructure.TVItemID, TVTypeEnum.LiftStation, MapInfoDrawTypeEnum.Point);
                    mapInfoPointModelInfOutList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelInfrastructure.TVItemID, TVTypeEnum.Outfall, MapInfoDrawTypeEnum.Point);
                }
                else if (infrastructureModel.InfrastructureType == InfrastructureTypeEnum.Other)
                {
                    mapInfoPointModelInfList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelInfrastructure.TVItemID, TVTypeEnum.OtherInfrastructure, MapInfoDrawTypeEnum.Point);
                    mapInfoPointModelInfOutList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelInfrastructure.TVItemID, TVTypeEnum.Outfall, MapInfoDrawTypeEnum.Point);
                }


                Coord coord = new Coord();
                if (mapInfoPointModelInfList.Count == 0)
                {
                    coord.Lat = (float)mapInfoPointModelMuniList[0].Lat + 0.01f;
                    coord.Lng = (float)mapInfoPointModelMuniList[0].Lng + 0.01f;
                    coord.Ordinal = 0;
                }
                else
                {
                    coord.Lat = (float)mapInfoPointModelInfList[0].Lat;
                    coord.Lng = (float)mapInfoPointModelInfList[0].Lng;
                    coord.Ordinal = 0;
                }

                Coord coord2 = new Coord();
                if (mapInfoPointModelInfOutList.Count == 0)
                {
                    coord2.Lat = (float)mapInfoPointModelMuniList[0].Lat + 0.01f;
                    coord2.Lng = (float)mapInfoPointModelMuniList[0].Lng + 0.01f;
                    coord2.Ordinal = 0;
                }
                else
                {
                    coord2.Lat = (float)mapInfoPointModelInfOutList[0].Lat;
                    coord2.Lng = (float)mapInfoPointModelInfOutList[0].Lng;
                    coord2.Ordinal = 0;
                }

                if (infrastructureModel.InfrastructureType == InfrastructureTypeEnum.WWTP)
                {
                    // Doing point Infr. WWTP
                    sbKMZWWTP.AppendLine(@"	<Placemark>");
                    sbKMZWWTP.AppendLine(@"	<name>" + tvItemModelInfrastructure.TVText + "</name>");
                    sbKMZWWTP.AppendLine(@"<styleUrl>#msn_ylw-pushpin</styleUrl>");
                    //sbKMZWWTP.AppendLine(@"	<description>");
                    //sbKMZWWTP.AppendLine(@"<![CDATA[");
                    //sbKMZWWTP.AppendLine(@"<a href=""" + _TaskRunnerBaseService.GetUrlFromTVItem(tvItemModelCountry) + @""">" + tvItemModelCountry.TVText + "</a>");
                    //sbKMZWWTP.AppendLine(@"]]>");
                    //sbKMZWWTP.AppendLine(@"	</description>");

                    sbKMZWWTP.AppendLine(@"		<Point>");
                    sbKMZWWTP.AppendLine(@"			<coordinates>" + coord.Lng + "," + coord.Lat + ",0</coordinates>");
                    sbKMZWWTP.AppendLine(@"		</Point>");
                    sbKMZWWTP.AppendLine(@"	</Placemark>");

                }
                else if (infrastructureModel.InfrastructureType == InfrastructureTypeEnum.LiftStation)
                {
                    // Doing point Infr. LS
                    sbKMZLS.AppendLine(@"	<Placemark>");
                    sbKMZLS.AppendLine(@"	<name>" + tvItemModelInfrastructure.TVText + "</name>");
                    sbKMZLS.AppendLine(@"<styleUrl>#msn_ylw-pushpin</styleUrl>");
                    //sbKMZLS.AppendLine(@"	<description>");
                    //sbKMZLS.AppendLine(@"<![CDATA[");
                    //sbKMZLS.AppendLine(@"<a href=""" + _TaskRunnerBaseService.GetUrlFromTVItem(tvItemModelCountry) + @""">" + tvItemModelCountry.TVText + "</a>");
                    //sbKMZLS.AppendLine(@"]]>");
                    //sbKMZLS.AppendLine(@"	</description>");

                    sbKMZLS.AppendLine(@"		<Point>");
                    sbKMZLS.AppendLine(@"			<coordinates>" + coord.Lng + "," + coord.Lat + ",0</coordinates>");
                    sbKMZLS.AppendLine(@"		</Point>");
                    sbKMZLS.AppendLine(@"	</Placemark>");
                }
                else if (infrastructureModel.InfrastructureType == InfrastructureTypeEnum.Other)
                {
                    // Doing point Infr. WWTP
                    sbKMZOther.AppendLine(@"	<Placemark>");
                    sbKMZOther.AppendLine(@"	<name>" + tvItemModelInfrastructure.TVText + "</name>");
                    sbKMZOther.AppendLine(@"<styleUrl>#msn_ylw-pushpin</styleUrl>");
                    //sbKMZOther.AppendLine(@"	<description>");
                    //sbKMZOther.AppendLine(@"<![CDATA[");
                    //sbKMZOther.AppendLine(@"<a href=""" + _TaskRunnerBaseService.GetUrlFromTVItem(tvItemModelCountry) + @""">" + tvItemModelCountry.TVText + "</a>");
                    //sbKMZOther.AppendLine(@"]]>");
                    //sbKMZOther.AppendLine(@"	</description>");

                    sbKMZOther.AppendLine(@"		<Point>");
                    sbKMZOther.AppendLine(@"			<coordinates>" + coord.Lng + "," + coord.Lat + ",0</coordinates>");
                    sbKMZOther.AppendLine(@"		</Point>");
                    sbKMZOther.AppendLine(@"	</Placemark>");
                }
                else
                {
                }

                // Doing point Infr. outfall
                sbKMZOutfall.AppendLine(@"	<Placemark>");
                sbKMZOutfall.AppendLine(@"	<name>" + tvItemModelInfrastructure.TVText + "</name>");
                sbKMZOutfall.AppendLine(@"<styleUrl>#msn_ylw-pushpin</styleUrl>");
                //sbKMZOutfall.AppendLine(@"	<description>");
                //sbKMZOutfall.AppendLine(@"<![CDATA[");
                //sbKMZOutfall.AppendLine(@"<a href=""" + _TaskRunnerBaseService.GetUrlFromTVItem(tvItemModelCountry) + @""">" + tvItemModelCountry.TVText + "</a>");
                //sbKMZOutfall.AppendLine(@"]]>");
                //sbKMZOutfall.AppendLine(@"	</description>");

                sbKMZOutfall.AppendLine(@"		<Point>");
                sbKMZOutfall.AppendLine(@"			<coordinates>" + coord2.Lng + "," + coord2.Lat + ",0</coordinates>");
                sbKMZOutfall.AppendLine(@"		</Point>");
                sbKMZOutfall.AppendLine(@"	</Placemark>");

               

            }
            sbKMZWWTP.AppendLine(@" </Folder>");
            sbKMZLS.AppendLine(@" </Folder>");
            sbKMZOutfall.AppendLine(@" </Folder>");
            sbKMZOther.AppendLine(@" </Folder>");

            sbKMZ.Append(sbKMZWWTP.ToString());
            sbKMZ.Append(sbKMZLS.ToString());
            sbKMZ.Append(sbKMZOutfall.ToString());
            sbKMZ.Append(sbKMZOther.ToString());          
            
            sbKMZ.AppendLine(@" </Folder>");

            sbKMZ.AppendLine(@"	</Folder>");

            sbKMZ.AppendLine(@"</Document>");
            sbKMZ.AppendLine(@"</kml>");

            StreamWriter sw = fi.CreateText();

            sw.Write(sbKMZ.ToString());

            sw.Close();

        }
        #endregion Functions public
    }
}
