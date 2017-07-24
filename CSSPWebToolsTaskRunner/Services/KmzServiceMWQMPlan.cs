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
using System.Diagnostics;
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;

namespace CSSPWebToolsTaskRunner.Services
{
    public class KmzServiceMWQMPlan
    {
        #region Variables
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        public MapInfoService _MapInfoService { get; private set; }
        public MWQMPlanService _MWQMPlanService { get; private set; }
        public MWQMPlanSubsectorService _MWQMPlanSubsectorService { get; private set; }
        public MWQMPlanSubsectorSiteService _MWQMPlanSubsectorSiteService { get; private set; }
        #endregion Variables

        #region Constructors
        public KmzServiceMWQMPlan(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
            _MapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _MWQMPlanService = new MWQMPlanService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _MWQMPlanSubsectorService = new MWQMPlanSubsectorService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _MWQMPlanSubsectorSiteService = new MWQMPlanSubsectorSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        }
        #endregion Constructors

        #region Functions public
        public void Generate(FileInfo fi)
        {
            string NotUsed = "";

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

            MWQMPlanModel mwqmPlanModel = _MWQMPlanService.GetMWQMPlanModelWithMWQMPlanIDDB(_TaskRunnerBaseService._BWObj.MWQMPlanID);
            if (!string.IsNullOrWhiteSpace(mwqmPlanModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, "MWQMPlan", "MWQMPlanID", _TaskRunnerBaseService._BWObj.MWQMPlanID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", "MWQMPlan", "MWQMPlanID", _TaskRunnerBaseService._BWObj.MWQMPlanID.ToString());
                return;
            }

            StringBuilder sbKMZ = new StringBuilder();

            sbKMZ.AppendLine(@"<?xml version=""1.0"" encoding=""UTF-8""?>");
            sbKMZ.AppendLine(@"<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom"">");
            sbKMZ.AppendLine(@"<Document>");
            sbKMZ.AppendLine(@"	<name>KmlFile</name>");
            sbKMZ.AppendLine(@"	<StyleMap id=""msn_ylw-pushpin1"">");
            sbKMZ.AppendLine(@"		<Pair>");
            sbKMZ.AppendLine(@"			<key>normal</key>");
            sbKMZ.AppendLine(@"			<styleUrl>#sn_ylw-pushpin1</styleUrl>");
            sbKMZ.AppendLine(@"		</Pair>");
            sbKMZ.AppendLine(@"		<Pair>");
            sbKMZ.AppendLine(@"			<key>highlight</key>");
            sbKMZ.AppendLine(@"			<styleUrl>#sh_ylw-pushpin1</styleUrl>");
            sbKMZ.AppendLine(@"		</Pair>");
            sbKMZ.AppendLine(@"	</StyleMap>");
            sbKMZ.AppendLine(@"	<Style id=""sn_ylw-pushpin"">");
            sbKMZ.AppendLine(@"		<IconStyle>");
            sbKMZ.AppendLine(@"			<color>ff00ff00</color>");
            sbKMZ.AppendLine(@"			<scale>1.1</scale>");
            sbKMZ.AppendLine(@"			<Icon>");
            sbKMZ.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>");
            sbKMZ.AppendLine(@"			</Icon>");
            sbKMZ.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbKMZ.AppendLine(@"		</IconStyle>");
            sbKMZ.AppendLine(@"		<LabelStyle>");
            sbKMZ.AppendLine(@"			<color>ff00ff00</color>");
            sbKMZ.AppendLine(@"		</LabelStyle>");
            sbKMZ.AppendLine(@"		<LineStyle>");
            sbKMZ.AppendLine(@"			<color>ff00ff00</color>");
            sbKMZ.AppendLine(@"		</LineStyle>");
            sbKMZ.AppendLine(@"		<PolyStyle>");
            sbKMZ.AppendLine(@"			<color>ff00ff00</color>");
            sbKMZ.AppendLine(@"			<fill>0</fill>");
            sbKMZ.AppendLine(@"		</PolyStyle>");
            sbKMZ.AppendLine(@"	</Style>");
            sbKMZ.AppendLine(@"	<Style id=""sn_ylw-pushpin0"">");
            sbKMZ.AppendLine(@"		<IconStyle>");
            sbKMZ.AppendLine(@"			<scale>1.1</scale>");
            sbKMZ.AppendLine(@"			<Icon>");
            sbKMZ.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>");
            sbKMZ.AppendLine(@"			</Icon>");
            sbKMZ.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbKMZ.AppendLine(@"		</IconStyle>");
            sbKMZ.AppendLine(@"		<LineStyle>");
            sbKMZ.AppendLine(@"			<color>ff0000ff</color>");
            sbKMZ.AppendLine(@"		</LineStyle>");
            sbKMZ.AppendLine(@"		<PolyStyle>");
            sbKMZ.AppendLine(@"			<color>ff0000ff</color>");
            sbKMZ.AppendLine(@"			<fill>0</fill>");
            sbKMZ.AppendLine(@"		</PolyStyle>");
            sbKMZ.AppendLine(@"	</Style>");
            sbKMZ.AppendLine(@"	<StyleMap id=""msn_ylw-pushpin10"">");
            sbKMZ.AppendLine(@"		<Pair>");
            sbKMZ.AppendLine(@"			<key>normal</key>");
            sbKMZ.AppendLine(@"			<styleUrl>#sn_ylw-pushpin0</styleUrl>");
            sbKMZ.AppendLine(@"		</Pair>");
            sbKMZ.AppendLine(@"		<Pair>");
            sbKMZ.AppendLine(@"			<key>highlight</key>");
            sbKMZ.AppendLine(@"			<styleUrl>#sh_ylw-pushpin0</styleUrl>");
            sbKMZ.AppendLine(@"		</Pair>");
            sbKMZ.AppendLine(@"	</StyleMap>");
            sbKMZ.AppendLine(@"	<Style id=""sh_ylw-pushpin"">");
            sbKMZ.AppendLine(@"		<IconStyle>");
            sbKMZ.AppendLine(@"			<color>ff00ff00</color>");
            sbKMZ.AppendLine(@"			<scale>1.3</scale>");
            sbKMZ.AppendLine(@"			<Icon>");
            sbKMZ.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>");
            sbKMZ.AppendLine(@"			</Icon>");
            sbKMZ.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbKMZ.AppendLine(@"		</IconStyle>");
            sbKMZ.AppendLine(@"		<LabelStyle>");
            sbKMZ.AppendLine(@"			<color>ff00ff00</color>");
            sbKMZ.AppendLine(@"		</LabelStyle>");
            sbKMZ.AppendLine(@"		<LineStyle>");
            sbKMZ.AppendLine(@"			<color>ff00ff00</color>");
            sbKMZ.AppendLine(@"		</LineStyle>");
            sbKMZ.AppendLine(@"		<PolyStyle>");
            sbKMZ.AppendLine(@"			<color>ff00ff00</color>");
            sbKMZ.AppendLine(@"			<fill>0</fill>");
            sbKMZ.AppendLine(@"		</PolyStyle>");
            sbKMZ.AppendLine(@"	</Style>");
            sbKMZ.AppendLine(@"	<Style id=""sh_ylw-pushpin0"">");
            sbKMZ.AppendLine(@"		<IconStyle>");
            sbKMZ.AppendLine(@"			<scale>1.3</scale>");
            sbKMZ.AppendLine(@"			<Icon>");
            sbKMZ.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>");
            sbKMZ.AppendLine(@"			</Icon>");
            sbKMZ.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbKMZ.AppendLine(@"		</IconStyle>");
            sbKMZ.AppendLine(@"		<LineStyle>");
            sbKMZ.AppendLine(@"			<color>ff0000ff</color>");
            sbKMZ.AppendLine(@"		</LineStyle>");
            sbKMZ.AppendLine(@"		<PolyStyle>");
            sbKMZ.AppendLine(@"			<color>ff0000ff</color>");
            sbKMZ.AppendLine(@"			<fill>0</fill>");
            sbKMZ.AppendLine(@"		</PolyStyle>");
            sbKMZ.AppendLine(@"	</Style>");
            sbKMZ.AppendLine(@"	<Style id=""sh_ylw-pushpin1"">");
            sbKMZ.AppendLine(@"		<IconStyle>");
            sbKMZ.AppendLine(@"			<scale>1.3</scale>");
            sbKMZ.AppendLine(@"			<Icon>");
            sbKMZ.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>");
            sbKMZ.AppendLine(@"			</Icon>");
            sbKMZ.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbKMZ.AppendLine(@"		</IconStyle>");
            sbKMZ.AppendLine(@"		<LineStyle>");
            sbKMZ.AppendLine(@"			<color>ff00ff00</color>");
            sbKMZ.AppendLine(@"		</LineStyle>");
            sbKMZ.AppendLine(@"		<PolyStyle>");
            sbKMZ.AppendLine(@"			<color>ff00ff00</color>");
            sbKMZ.AppendLine(@"			<fill>0</fill>");
            sbKMZ.AppendLine(@"		</PolyStyle>");
            sbKMZ.AppendLine(@"	</Style>");
            sbKMZ.AppendLine(@"	<Style id=""sn_ylw-pushpin1"">");
            sbKMZ.AppendLine(@"		<IconStyle>");
            sbKMZ.AppendLine(@"			<scale>1.1</scale>");
            sbKMZ.AppendLine(@"			<Icon>");
            sbKMZ.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>");
            sbKMZ.AppendLine(@"			</Icon>");
            sbKMZ.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbKMZ.AppendLine(@"		</IconStyle>");
            sbKMZ.AppendLine(@"		<LineStyle>");
            sbKMZ.AppendLine(@"			<color>ff00ff00</color>");
            sbKMZ.AppendLine(@"		</LineStyle>");
            sbKMZ.AppendLine(@"		<PolyStyle>");
            sbKMZ.AppendLine(@"			<color>ff00ff00</color>");
            sbKMZ.AppendLine(@"			<fill>0</fill>");
            sbKMZ.AppendLine(@"		</PolyStyle>");
            sbKMZ.AppendLine(@"	</Style>");
            sbKMZ.AppendLine(@"	<StyleMap id=""msn_ylw-pushpin11"">");
            sbKMZ.AppendLine(@"		<Pair>");
            sbKMZ.AppendLine(@"			<key>normal</key>");
            sbKMZ.AppendLine(@"			<styleUrl>#sn_ylw-pushpin</styleUrl>");
            sbKMZ.AppendLine(@"		</Pair>");
            sbKMZ.AppendLine(@"		<Pair>");
            sbKMZ.AppendLine(@"			<key>highlight</key>");
            sbKMZ.AppendLine(@"			<styleUrl>#sh_ylw-pushpin</styleUrl>");
            sbKMZ.AppendLine(@"		</Pair>");
            sbKMZ.AppendLine(@"	</StyleMap>");

            sbKMZ.AppendLine(@"	<Folder>");
            sbKMZ.AppendLine(@"		<name>" + mwqmPlanModel.ConfigFileName.Replace(".txt", ".kmz") + "</name>");
            sbKMZ.AppendLine(@"		<open>1</open>");

            List<MWQMPlanSubsectorModel> mwqmPlanSubsectorModelList = _MWQMPlanSubsectorService.GetMWQMPlanSubsectorModelListWithMWQMPlanIDDB(_TaskRunnerBaseService._BWObj.MWQMPlanID);

            foreach (MWQMPlanSubsectorModel MWQMPlanSubsectorModel in mwqmPlanSubsectorModelList)
            {
                sbKMZ.AppendLine(@"		<Folder>");
                sbKMZ.AppendLine(@"			<name>" + MWQMPlanSubsectorModel.SubsectorTVText + "</name>");
                sbKMZ.AppendLine(@"			<open>1</open>");
                sbKMZ.AppendLine(@"			<Placemark>");
                sbKMZ.AppendLine(@"				<name>Poly</name>");
                sbKMZ.AppendLine(@"				<styleUrl>#msn_ylw-pushpin10</styleUrl>");
                sbKMZ.AppendLine(@"				<Polygon>");
                sbKMZ.AppendLine(@"					<tessellate>1</tessellate>");
                sbKMZ.AppendLine(@"					<outerBoundaryIs>");
                sbKMZ.AppendLine(@"						<LinearRing>");
                sbKMZ.AppendLine(@"							<coordinates>");

                List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(MWQMPlanSubsectorModel.SubsectorTVItemID, TVTypeEnum.Subsector, MapInfoDrawTypeEnum.Polygon);

                foreach (MapInfoPointModel mapInfoPointModel in mapInfoPointModelList)
                {
                    sbKMZ.Append(mapInfoPointModel.Lng + "," + mapInfoPointModel.Lat + ",0 ");

                }
                sbKMZ.AppendLine("");
                sbKMZ.AppendLine(@"							</coordinates>");
                sbKMZ.AppendLine(@"						</LinearRing>");
                sbKMZ.AppendLine(@"					</outerBoundaryIs>");
                sbKMZ.AppendLine(@"				</Polygon>");
                sbKMZ.AppendLine(@"			</Placemark>");

                mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(MWQMPlanSubsectorModel.SubsectorTVItemID, TVTypeEnum.Subsector, MapInfoDrawTypeEnum.Point);

                if (mapInfoPointModelList.Count > 0)
                {
                    sbKMZ.AppendLine(@"			<Placemark>");
                    sbKMZ.AppendLine(@"				<name>" + MWQMPlanSubsectorModel.SubsectorTVText + "</name>");
                    sbKMZ.AppendLine(@"				<styleUrl>#msn_ylw-pushpin1</styleUrl>");
                    sbKMZ.AppendLine(@"				<Point>");
                    sbKMZ.AppendLine(@"					<coordinates>" + mapInfoPointModelList[0].Lng + "," + mapInfoPointModelList[0].Lat + ",0</coordinates>");
                    sbKMZ.AppendLine(@"				</Point>");
                    sbKMZ.AppendLine(@"			</Placemark>");
                }

                List<MWQMPlanSubsectorSiteModel> mwqmPlanSubsectorSiteModelList = _MWQMPlanSubsectorSiteService.GetMWQMPlanSubsectorSiteModelListWithMWQMPlanSubsectorIDDB(MWQMPlanSubsectorModel.MWQMPlanSubsectorID);

                sbKMZ.AppendLine(@"			<Folder>");
                sbKMZ.AppendLine(@"				<name>MWQMSites</name>");
                sbKMZ.AppendLine(@"				<open>1</open>");

                foreach (MWQMPlanSubsectorSiteModel MWQMPlanSubsectorSiteModel in mwqmPlanSubsectorSiteModelList)
                {
                    mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(MWQMPlanSubsectorSiteModel.MWQMSiteTVItemID, TVTypeEnum.MWQMSite, MapInfoDrawTypeEnum.Point);
                    if (mapInfoPointModelList.Count > 0)
                    {
                        sbKMZ.AppendLine(@"    			<Placemark>");
                        sbKMZ.AppendLine(@"    				<name>" + MWQMPlanSubsectorSiteModel.MWQMSiteText + "</name>");
                        sbKMZ.AppendLine(@"    				<styleUrl>#msn_ylw-pushpin11</styleUrl>");
                        sbKMZ.AppendLine(@"    				<Point>");
                        sbKMZ.AppendLine(@"    					<coordinates>" + mapInfoPointModelList[0].Lng + "," + mapInfoPointModelList[0].Lat + ",0</coordinates>");
                        sbKMZ.AppendLine(@"    				</Point>");
                        sbKMZ.AppendLine(@"    			</Placemark>");
                    }

                }
                sbKMZ.AppendLine(@"			</Folder>");
                sbKMZ.AppendLine(@"		</Folder>");
            }
            sbKMZ.AppendLine(@"	</Folder>");
            sbKMZ.AppendLine(@"</Document>");
            sbKMZ.AppendLine(@"</kml>");

            StreamWriter sw = fi.CreateText();
            sw.Write(sbKMZ.ToString());
            sw.Close();

            //string KMLFileName = fi.FullName;
            //string KMZFileName = fi.FullName.Replace(".kml", ".kmz");

            //// make sure directory exist if not create it
            //di = new DirectoryInfo(KMZFileName.Substring(0, KMZFileName.LastIndexOf("\\")));
            //if (!di.Exists)
            //    di.Create();

            //ProcessStartInfo pZip = new ProcessStartInfo();
            //pZip.Arguments = "a -tzip \"" + KMZFileName + "\" \"" + KMLFileName + "\"";
            //pZip.RedirectStandardInput = true;
            //pZip.UseShellExecute = false;
            //pZip.CreateNoWindow = true;
            //pZip.WindowStyle = ProcessWindowStyle.Hidden;

            //Process processZip = new Process();
            //processZip.StartInfo = pZip;
            //try
            //{
            //    pZip.FileName = @"C:\Program Files\7-Zip\7z.exe";
            //    processZip.Start();
            //}
            //catch (Exception ex)
            //{
            //    NotUsed = string.Format(TaskRunnerServiceRes.CompressKMLDidNotWorkWith7zError_, ex.Message);
            //    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CompressKMLDidNotWorkWith7zError_", ex.Message);
            //    return;
            //}

            //while (!processZip.HasExited)
            //{
            //    // waiting for the processZip to finish then continue
            //}

            //File.Delete(KMLFileName);

        }
        #endregion Functions public
    }
}
