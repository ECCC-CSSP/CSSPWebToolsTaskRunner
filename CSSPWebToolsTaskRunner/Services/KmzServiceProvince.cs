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
    public class KmzServiceProvince
    {
        #region Variables
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        #endregion Variables

        #region Constructors
        public KmzServiceProvince(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
        }
        #endregion Constructors

        #region Functions public
        public void Generate(FileInfo fi)
        {
            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

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

            TVItemModel tvItemModelProvince = tvItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            List<MapInfoPointModel> mapInfoPointModelList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelProvince.TVItemID, TVTypeEnum.Province, MapInfoDrawTypeEnum.Point);

            sbKMZ.AppendLine(@"	<Folder>");
            sbKMZ.AppendLine(@"	<name>Province</name>");

            // Doing Point
            sbKMZ.AppendLine(@"	<Placemark>");
            sbKMZ.AppendLine(@"	<name>" + tvItemModelProvince.TVText + "</name>");
            sbKMZ.AppendLine(@"<styleUrl>#msn_ylw-pushpin</styleUrl>");
            //sbKMZ.AppendLine(@"	<description>");
            //sbKMZ.AppendLine(@"<![CDATA[");
            //sbKMZ.AppendLine(@"<a href=""" + _TaskRunnerBaseService.GetUrlFromTVItem(tvItemModelRoot) + @""">" + tvItemModelRoot.TVText + "</a>");
            //sbKMZ.AppendLine(@"]]>");
            //sbKMZ.AppendLine(@"	</description>");
            sbKMZ.AppendLine(@"	<Point>");
            sbKMZ.AppendLine(@"		<coordinates>" + mapInfoPointModelList[0].Lng + "," + mapInfoPointModelList[0].Lat + ",0</coordinates>");
            sbKMZ.AppendLine(@"	</Point>");
            sbKMZ.AppendLine(@"	</Placemark>");

            // Doing Polygon
            sbKMZ.AppendLine(@"	<Placemark>");
            sbKMZ.AppendLine(@"		<name>" + tvItemModelProvince.TVText + " (poly)</name>");
            sbKMZ.AppendLine(@"<styleUrl>#msn_ylw-pushpin</styleUrl>");
            //sbKMZ.AppendLine(@"	<description>");
            //sbKMZ.AppendLine(@"<![CDATA[");
            //sbKMZ.AppendLine(@"<a href=""" + _TaskRunnerBaseService.GetUrlFromTVItem(tvItemModelRoot) + @""">" + tvItemModelRoot.TVText + "</a>");
            //sbKMZ.AppendLine(@"]]>");
            //sbKMZ.AppendLine(@"	</description>");
            sbKMZ.AppendLine(@"		<Polygon>");
            sbKMZ.AppendLine(@"			<outerBoundaryIs>");
            sbKMZ.AppendLine(@"				<LinearRing>");
            sbKMZ.AppendLine(@"					<coordinates>");
            sbKMZ.AppendLine(@"		" + (mapInfoPointModelList[0].Lng - 0.1) + "," + (mapInfoPointModelList[0].Lat - 0.1) + ",0");
            sbKMZ.AppendLine(@"		" + (mapInfoPointModelList[0].Lng + 0.1) + "," + (mapInfoPointModelList[0].Lat - 0.1) + ",0");
            sbKMZ.AppendLine(@"		" + (mapInfoPointModelList[0].Lng + 0.1) + "," + (mapInfoPointModelList[0].Lat + 0.1) + ",0");
            sbKMZ.AppendLine(@"		" + (mapInfoPointModelList[0].Lng - 0.1) + "," + (mapInfoPointModelList[0].Lat + 0.1) + ",0");
            sbKMZ.AppendLine(@"		" + (mapInfoPointModelList[0].Lng - 0.1) + "," + (mapInfoPointModelList[0].Lat - 0.1) + ",0");
            sbKMZ.AppendLine(@"					</coordinates>");
            sbKMZ.AppendLine(@"				</LinearRing>");
            sbKMZ.AppendLine(@"			</outerBoundaryIs>");
            sbKMZ.AppendLine(@"		</Polygon>");
            sbKMZ.AppendLine(@"	</Placemark>");

            sbKMZ.AppendLine(@" <Folder>");
            sbKMZ.AppendLine(@"	<name>Areas</name>");
            List<TVItemModel> tvItemModelAreaList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelProvince.TVItemID, TVTypeEnum.Area);

            foreach (TVItemModel tvItemModelArea in tvItemModelAreaList)
            {
                // Doing point
                sbKMZ.AppendLine(@"	<Placemark>");
                sbKMZ.AppendLine(@"	<name>" + tvItemModelArea.TVText + "</name>");
                sbKMZ.AppendLine(@"<styleUrl>#msn_ylw-pushpin</styleUrl>");
                //sbKMZ.AppendLine(@"	<description>");
                //sbKMZ.AppendLine(@"<![CDATA[");
                //sbKMZ.AppendLine(@"<a href=""" + _TaskRunnerBaseService.GetUrlFromTVItem(tvItemModelCountry) + @""">" + tvItemModelCountry.TVText + "</a>");
                //sbKMZ.AppendLine(@"]]>");
                //sbKMZ.AppendLine(@"	</description>");

                mapInfoPointModelList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelArea.TVItemID, TVTypeEnum.Area, MapInfoDrawTypeEnum.Point);

                sbKMZ.AppendLine(@"		<Point>");
                sbKMZ.AppendLine(@"			<coordinates>" + mapInfoPointModelList[0].Lng + "," + mapInfoPointModelList[0].Lat + ",0</coordinates>");
                sbKMZ.AppendLine(@"		</Point>");
                sbKMZ.AppendLine(@"	</Placemark>");

                // Doing Polygon
                sbKMZ.AppendLine(@"	<Placemark>");
                sbKMZ.AppendLine(@"		<name>" + tvItemModelArea.TVText + " (poly)</name>");
                sbKMZ.AppendLine(@"<styleUrl>#msn_ylw-pushpin</styleUrl>");
                //sbKMZ.AppendLine(@"	<description>");
                //sbKMZ.AppendLine(@"<![CDATA[");
                //sbKMZ.AppendLine(@"<a href=""" + _TaskRunnerBaseService.GetUrlFromTVItem(tvItemModelCountry) + @""">" + tvItemModelCountry.TVText + "</a>");
                //sbKMZ.AppendLine(@"]]>");
                //sbKMZ.AppendLine(@"	</description>");
                sbKMZ.AppendLine(@"		<Polygon>");
                sbKMZ.AppendLine(@"			<outerBoundaryIs>");
                sbKMZ.AppendLine(@"				<LinearRing>");
                sbKMZ.AppendLine(@"					<coordinates>");

                mapInfoPointModelList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelArea.TVItemID, TVTypeEnum.Area, MapInfoDrawTypeEnum.Polygon);
                foreach (MapInfoPointModel mapInfoPointModel in mapInfoPointModelList)
                {
                    sbKMZ.AppendLine(mapInfoPointModel.Lng + "," + mapInfoPointModel.Lat + ",0 ");

                }

                sbKMZ.AppendLine(@"					</coordinates>");
                sbKMZ.AppendLine(@"				</LinearRing>");
                sbKMZ.AppendLine(@"			</outerBoundaryIs>");
                sbKMZ.AppendLine(@"		</Polygon>");
                sbKMZ.AppendLine(@"	</Placemark>");
            }
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
