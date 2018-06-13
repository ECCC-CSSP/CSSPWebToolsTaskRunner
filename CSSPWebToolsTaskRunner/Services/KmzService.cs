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
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;
using System.Data.OleDb;

namespace CSSPWebToolsTaskRunner.Services
{
    public class KmzService
    {
        #region Variables
        #endregion Variables

        #region Properties
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        public SamplingPlanService _SamplingPlanService { get; private set; }
        public LabSheetService _LabSheetService { get; private set; }
        public LabSheetDetailService _LabSheetDetailService { get; private set; }
        public LabSheetTubeMPNDetailService _LabSheetTubeMPNDetailService { get; private set; }
        public TVFileService _TVFileService { get; private set; }
        public AppTaskService _AppTaskService { get; private set; }
        public TVItemService _TVItemService { get; private set; }
        public ContactService _ContactService { get; private set; }
        public MikeScenarioService _MikeScenarioService { get; private set; }
        public MapInfoService _MapInfoService { get; private set; }
        #endregion Properties

        #region Constructors
        public KmzService(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
            _SamplingPlanService = new SamplingPlanService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _LabSheetService = new LabSheetService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _LabSheetDetailService = new LabSheetDetailService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _LabSheetTubeMPNDetailService = new LabSheetTubeMPNDetailService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _TVFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _AppTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _TVItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _ContactService = new ContactService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _MikeScenarioService = new MikeScenarioService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _MapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
        }
        #endregion Constructors

        #region Functions public
        public void CreateDocumentFromTemplateKml(DocTemplateModel docTemplateModel)
        {
            string NotUsed = "";

            TVFileModel tvFileModelTemplate = _TVFileService.GetTVFileModelWithTVFileTVItemIDDB(docTemplateModel.TVFileTVItemID);
            if (!string.IsNullOrWhiteSpace(tvFileModelTemplate.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.TVFileTVItemID, docTemplateModel.TVFileTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.TVFileTVItemID, docTemplateModel.TVFileTVItemID.ToString());
                return;
            }

            string ServerNewFilePath = _TVFileService.ChoseEDriveOrCDrive(_TVFileService.GetServerFilePath(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID));

            DateTime CD = DateTime.Now;
            string Language = "_" + _TaskRunnerBaseService._BWObj.appTaskModel.Language;

            string DateText = "_" + CD.Year.ToString() +
                "_" + (CD.Month > 9 ? CD.Month.ToString() : "0" + CD.Month.ToString()) +
                "_" + (CD.Day > 9 ? CD.Day.ToString() : "0" + CD.Day.ToString()) +
                "_" + (CD.Hour > 9 ? CD.Hour.ToString() : "0" + CD.Hour.ToString()) +
                "_" + (CD.Minute > 9 ? CD.Minute.ToString() : "0" + CD.Minute.ToString());

            FileInfo fi = new FileInfo(docTemplateModel.FileName);
            string NewFileName = docTemplateModel.FileName.Replace(fi.Extension.ToLower(), DateText + Language + fi.Extension.ToLower());

            fi = new FileInfo(ServerNewFilePath + NewFileName);

            if (fi.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.FileAlreadyExist_, fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("FileAlreadyExist_", fi.FullName);
                return;
            }

            try
            {
                File.Copy(_TVFileService.ChoseEDriveOrCDrive(tvFileModelTemplate.ServerFilePath) + tvFileModelTemplate.ServerFileName, fi.FullName);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateDocumentFromTemplateError_, ex.Message + " - " + (ex.InnerException != null ? ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotCreateDocumentFromTemplateError_", ex.Message + " - " + (ex.InnerException != null ? ex.InnerException.Message : ""));
                return;
            }

            TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            string retStr = ParseKml(fi);
            if (!string.IsNullOrWhiteSpace(retStr))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateDocumentFromTemplateError_, retStr);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotCreateDocumentFromTemplateError_", retStr);
                return;
            }

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.FileGeneratedFromTemplate, FilePurposeEnum.TemplateGenerated);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;
        }
        public void CreateKMZOfMWQMSites()
        {
            string NotUsed = "";
            string ProvInit = "";
            List<string> ProvInitList = new List<string>()
            {
                "BC", "ME", "NB", "NL", "NS", "PE", "QC",
            };
            List<string> ProvList = new List<string>()
            {
                "British Columbia", "Maine", "New Brunswick", "Newfoundland and Labrador", "Nova Scotia", "Prince Edward Island", "Québec",
            };
            int TVItemID = _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID;

            if (TVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.Parameter_NotFound, "TVItemID");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Parameter_NotFound", "TVItemID");
                return;
            }

            TVItemModel tvItemModel = _TVItemService.GetTVItemModelWithTVItemIDDB(TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotDeleteFile_Error_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                return;
            }

            for (int i = 0, countProv = ProvList.Count; i < countProv; i++)
            {
                if (ProvList[i] == tvItemModel.TVText)
                {
                    ProvInit = ProvInitList[i];
                    break;
                }
            }

            string ServerFilePath = _TVFileService.GetServerFilePath(TVItemID);

            FileInfo fi = new FileInfo(_TVFileService.ChoseEDriveOrCDrive(ServerFilePath) + $"Sites_{ProvInit}.kml");

            TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            // loop through all the MWQMSites etc...


            if (fi.Exists)
            {
                try
                {
                    fi.Delete();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDeleteFile_Error_, fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDeleteFile_Error_", fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    return;
                }
            }

            List<SensitiveSite> sensitiveSiteList = new List<SensitiveSite>();

            FileInfo fiBCSensitiveSite = new FileInfo(fi.FullName.Replace($"Sites_{ProvInit}.kml", "SensitiveSites.xlsx"));

            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fiBCSensitiveSite.FullName + ";Extended Properties=Excel 12.0";

            OleDbConnection conn = new OleDbConnection(connectionString);

            conn.Open();
            OleDbDataReader reader;
            OleDbCommand comm = new OleDbCommand("Select * from [Sheet1$];");

            comm.Connection = conn;
            reader = comm.ExecuteReader();

            while (reader.Read())
            {
                string SiteName = "";
                int Height_m = 0;
                int Width_m = 0;

                // SiteName
                if (reader.GetValue(0).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(0).ToString()))
                {
                    SiteName = "";
                }
                else
                {
                    SiteName = reader.GetValue(0).ToString().Trim();
                }

                // Height_m
                if (reader.GetValue(1).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(1).ToString()))
                {
                    Height_m = 0;
                }
                else
                {
                    Height_m = int.Parse(reader.GetValue(1).ToString().Trim());
                }

                // Width_m
                if (reader.GetValue(2).GetType() == typeof(DBNull) || string.IsNullOrEmpty(reader.GetValue(2).ToString()))
                {
                    Width_m = 0;
                }
                else
                {
                    Width_m = int.Parse(reader.GetValue(2).ToString().Trim());
                }

                if (string.IsNullOrWhiteSpace(SiteName))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes._WithinDocument_ShouldNotBeEmpty, "SiteName", "SensitiveSites.xlsx");
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("_WithinDocument_ShouldNotBeEmpty", "SiteName", "BCSensitiveSites.xlsx");
                    return;
                }

                if (Height_m == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes._WithinDocument_ShouldNotBeEmpty, "Height_m", "SensitiveSites.xlsx");
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("_WithinDocument_ShouldNotBeEmpty", "Height_m", "BCSensitiveSites.xlsx");
                    return;
                }

                if (Width_m == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes._WithinDocument_ShouldNotBeEmpty, "Width_m", "SensitiveSites.xlsx");
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("_WithinDocument_ShouldNotBeEmpty", "Width_m", "BCSensitiveSites.xlsx");
                    return;
                }

                SensitiveSite sensitiveSite = new SensitiveSite()
                {
                    SiteName = SiteName,
                    Height_m = Height_m,
                    Width_m = Width_m,
                };

                sensitiveSiteList.Add(sensitiveSite);
            }

            conn.Close();


            StringBuilder sb = new StringBuilder();

            sb.AppendLine(@"<?xml version=""1.0"" encoding=""UTF-8""?>");
            sb.AppendLine(@"<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom"">");
            sb.AppendLine(@"<Document>");
            sb.AppendLine(@"	<name>" + fi.Name + "</name>");
            sb.AppendLine(@"	<Style id=""s_ylw-pushpin_hl"">");
            sb.AppendLine(@"		<IconStyle>");
            sb.AppendLine(@"			<scale>1.3</scale>");
            sb.AppendLine(@"			<Icon>");
            sb.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>");
            sb.AppendLine(@"			</Icon>");
            sb.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sb.AppendLine(@"		</IconStyle>");
            sb.AppendLine(@"	</Style>");
            sb.AppendLine(@"	<StyleMap id=""m_ylw-pushpin"">");
            sb.AppendLine(@"		<Pair>");
            sb.AppendLine(@"			<key>normal</key>");
            sb.AppendLine(@"			<styleUrl>#s_ylw-pushpin</styleUrl>");
            sb.AppendLine(@"		</Pair>");
            sb.AppendLine(@"		<Pair>");
            sb.AppendLine(@"			<key>highlight</key>");
            sb.AppendLine(@"			<styleUrl>#s_ylw-pushpin_hl</styleUrl>");
            sb.AppendLine(@"		</Pair>");
            sb.AppendLine(@"	</StyleMap>");
            sb.AppendLine(@"	<Style id=""s_ylw-pushpin"">");
            sb.AppendLine(@"		<IconStyle>");
            sb.AppendLine(@"			<scale>1.1</scale>");
            sb.AppendLine(@"			<Icon>");
            sb.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>");
            sb.AppendLine(@"			</Icon>");
            sb.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sb.AppendLine(@"		</IconStyle>");
            sb.AppendLine(@"	</Style>");
            sb.AppendLine(@"	<Style id=""sn_ylw-pushpin"">");
            sb.AppendLine(@"		<IconStyle>");
            sb.AppendLine(@"			<scale>1.1</scale>");
            sb.AppendLine(@"			<Icon>");
            sb.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>");
            sb.AppendLine(@"			</Icon>");
            sb.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sb.AppendLine(@"		</IconStyle>");
            sb.AppendLine(@"		<LineStyle>");
            sb.AppendLine(@"			<color>ff00ff00</color>");
            sb.AppendLine(@"			<width>2</width>");
            sb.AppendLine(@"		</LineStyle>");
            sb.AppendLine(@"		<PolyStyle>");
            sb.AppendLine(@"			<color>00ffffff</color>");
            sb.AppendLine(@"		</PolyStyle>");
            sb.AppendLine(@"	</Style>");
            sb.AppendLine(@"	<StyleMap id=""msn_ylw-pushpin"">");
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
            sb.AppendLine(@"			<scale>1.3</scale>");
            sb.AppendLine(@"			<Icon>");
            sb.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>");
            sb.AppendLine(@"			</Icon>");
            sb.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sb.AppendLine(@"		</IconStyle>");
            sb.AppendLine(@"		<LineStyle>");
            sb.AppendLine(@"			<color>ff00ff00</color>");
            sb.AppendLine(@"			<width>2</width>");
            sb.AppendLine(@"		</LineStyle>");
            sb.AppendLine(@"		<PolyStyle>");
            sb.AppendLine(@"			<color>00ffffff</color>");
            sb.AppendLine(@"		</PolyStyle>");
            sb.AppendLine(@" </Style>");


            using (CSSPWebToolsDBEntities db = new CSSPWebToolsDBEntities())
            {
                var tvItemProv = (from c in db.TVItems
                                  from cl in db.TVItemLanguages
                                  where c.TVItemID == cl.TVItemID
                                  && c.TVItemID == TVItemID
                                  && cl.Language == (int)LanguageEnum.en
                                  && c.TVType == (int)TVTypeEnum.Province
                                  select new { c, cl }).FirstOrDefault();

                if (tvItemProv == null)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                    return;
                }

                for (int i = 0, countProv = ProvList.Count; i < countProv; i++)
                {
                    if (ProvList[i] == tvItemProv.cl.TVText)
                    {
                        ProvInit = ProvInitList[i];
                        break;
                    }
                }

                NotUsed = string.Format(TaskRunnerServiceRes.Creating_, fi.Name);
                List<TextLanguage> TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", fi.Name);

                _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);

                var tvItemSSList = (from t in db.TVItems
                                    from tl in db.TVItemLanguages
                                    where t.TVItemID == tl.TVItemID
                                    && tl.Language == (int)LanguageEnum.en
                                    && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                    && t.TVType == (int)TVTypeEnum.Subsector
                                    orderby tl.TVText
                                    select new { t, tl }).ToList();

                var MonitoringSiteList = (from t in db.TVItems
                                          from tl in db.TVItemLanguages
                                          from mi in db.MapInfos
                                          from mip in db.MapInfoPoints
                                          let hasSample = (from c in db.MWQMSamples
                                                           where c.MWQMSiteTVItemID == t.TVItemID
                                                           && c.UseForOpenData == true
                                                           select c).Any()
                                          where t.TVItemID == tl.TVItemID
                                          && mi.TVItemID == t.TVItemID
                                          && mip.MapInfoID == mi.MapInfoID
                                          && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                          && t.TVType == (int)TVTypeEnum.MWQMSite
                                          && mi.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Point
                                          && mi.TVType == (int)TVTypeEnum.MWQMSite
                                          && hasSample == true
                                          && tl.Language == (int)LanguageEnum.en
                                          select new { t, tl, mip, hasSample }).ToList();


                int TotalCount2 = tvItemSSList.Count;
                int Count2 = 0;
                foreach (var tvItemSS in tvItemSSList)
                {
                    if (Count2 % 20 == 0)
                    {
                        _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)(100.0f * ((float)Count2 / (float)TotalCount2)));

                        NotUsed = string.Format(TaskRunnerServiceRes.Creating_, fi.Name + " --- doing " + tvItemSS.tl.TVText + "");
                        TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", fi.Name + " --- doing " + tvItemSS.tl.TVText + "");

                        _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);
                    }

                    Count2 += 1;

                    foreach (var mwqmSite in MonitoringSiteList.Where(c => c.t.ParentID == tvItemSS.t.TVItemID))
                    {
                        string TVText = mwqmSite.tl.TVText;
                        string MS = mwqmSite.t.TVItemID.ToString();
                        string Lat = (mwqmSite.mip != null ? mwqmSite.mip.Lat.ToString("F6").Replace(",", ".") : "");
                        string Lng = (mwqmSite.mip != null ? mwqmSite.mip.Lng.ToString("F6").Replace(",", ".") : "");

                        SensitiveSite sensitiveSiteExist = (from c in sensitiveSiteList
                                                            where c.SiteName == TVText
                                                            select c).FirstOrDefault();

                        if (sensitiveSiteExist != null)
                        {
                            sb.AppendLine(@"    <Placemark>");
                            sb.AppendLine($@"		<name>{MS}</name>");
                            sb.AppendLine(@"		<styleUrl>#msn_ylw-pushpin</styleUrl>");
                            sb.AppendLine(@"		<Polygon>");
                            sb.AppendLine(@"			<tessellate>1</tessellate>");
                            sb.AppendLine(@"			<outerBoundaryIs>");
                            sb.AppendLine(@"				<LinearRing>");
                            sb.AppendLine(@"					<coordinates>");

                            int height = sensitiveSiteExist.Height_m;

                            double dy = 100000000;
                            double dyLat = 0.001;
                            while (!(dy < height * 1.1 && dy > height * 0.9))
                            {
                                if (dy > height * 1.1)
                                {
                                    dyLat = dyLat / 1.987634;
                                }
                                else if (dy < height * 0.9)
                                {
                                    dyLat = dyLat * 1.19384937;
                                }

                                dy = _MapInfoService.CalculateDistance(mwqmSite.mip.Lat * _MapInfoService.d2r, mwqmSite.mip.Lng * _MapInfoService.d2r, (mwqmSite.mip.Lat + dyLat) * _MapInfoService.d2r, mwqmSite.mip.Lng * _MapInfoService.d2r, _MapInfoService.R);
                            }

                            int width = sensitiveSiteExist.Width_m;
                            double dx = 100000000;
                            double dxLng = 0.001;
                            while (!(dx < width * 1.1 && dx > width * 0.9))
                            {
                                if (dx > width * 1.1)
                                {
                                    dxLng = dxLng / 1.987634;
                                }
                                else if (dx < width * 0.9)
                                {
                                    dxLng = dxLng * 1.19384937;
                                }

                                dx = _MapInfoService.CalculateDistance(mwqmSite.mip.Lat * _MapInfoService.d2r, mwqmSite.mip.Lng * _MapInfoService.d2r, mwqmSite.mip.Lat * _MapInfoService.d2r, (mwqmSite.mip.Lng + dxLng) * _MapInfoService.d2r, _MapInfoService.R);
                            }

                            Random random = new Random((int)DateTime.Now.Ticks);

                            double PercLat = random.Next(10, 90);
                            double PercLng = random.Next(10, 90);

                            double BottomLeftLat = mwqmSite.mip.Lat + (dyLat * ((double)PercLat/100.0D));
                            double TopRightLat = mwqmSite.mip.Lat - (dyLat * ((100.0D - (double)PercLat) / 100.0D));
                            double BottomLeftLng = mwqmSite.mip.Lng + (dxLng * ((double)PercLng / 100.0D));
                            double TopRightLng = mwqmSite.mip.Lng - (dxLng * ((100.0D - (double)PercLng) / 100.0D));

                            sb.AppendLine($@"						{BottomLeftLng.ToString("F7")},{BottomLeftLat.ToString("F7")},0 {TopRightLng.ToString("F7")},{BottomLeftLat.ToString("F7")},0 {TopRightLng.ToString("F7")},{TopRightLat.ToString("F7")},0 {BottomLeftLng.ToString("F7")},{TopRightLat.ToString("F7")},0 {BottomLeftLng.ToString("F7")},{BottomLeftLat.ToString("F7")},0 ");
                            sb.AppendLine(@"					</coordinates>");
                            sb.AppendLine(@"				</LinearRing>");
                            sb.AppendLine(@"			</outerBoundaryIs>");
                            sb.AppendLine(@"		</Polygon>");
                            sb.AppendLine(@" </Placemark>");

                            //sb.AppendLine(@"	<Placemark>");
                            //sb.AppendLine($@"		<name>{MS} Would be removed</name>");
                            //sb.AppendLine(@"		<styleUrl>#m_ylw-pushpin</styleUrl>");
                            //sb.AppendLine(@"		<Point>");
                            //sb.AppendLine($@"			<coordinates>{Lng},{Lat},0</coordinates>");
                            //sb.AppendLine(@"		</Point>");
                            //sb.AppendLine(@"	</Placemark>");
                        }
                        else
                        {
                            sb.AppendLine(@"	<Placemark>");
                            sb.AppendLine($@"		<name>{MS}</name>");
                            sb.AppendLine(@"		<styleUrl>#m_ylw-pushpin</styleUrl>");
                            sb.AppendLine(@"		<Point>");
                            sb.AppendLine($@"			<coordinates>{Lng},{Lat},0</coordinates>");
                            sb.AppendLine(@"		</Point>");
                            sb.AppendLine(@"	</Placemark>");
                        }
                    }
                }
            }

            sb.AppendLine(@"</Document>");
            sb.AppendLine(@"</kml>");


            UnicodeEncoding encoding = new UnicodeEncoding();

            FileStream fs = fi.Create();
            byte[] bytes = encoding.GetBytes(sb.ToString());
            fs.Write(bytes, 0, bytes.Length);
            fs.Close();

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.KMZOfMWQMSites, FilePurposeEnum.OpenData);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;
        }
        public string ParseKml(FileInfo fi)
        {

            return "Need to finalize ParseKML";
        }
        #endregion Functions public

        #region Functions private
        #endregion Functions private

    }

    public class SensitiveSite
    {
        public string SiteName { get; set; }
        public int Height_m { get; set; }
        public int Width_m { get; set; }
    }
}
