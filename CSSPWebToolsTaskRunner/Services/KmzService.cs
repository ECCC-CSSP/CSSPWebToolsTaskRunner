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
using CSSPEnumsDLL.Services;

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
        public ProvinceToolsService _ProvinceToolsService { get; private set; }
        public PolSourceSiteService _PolSourceSiteService { get; private set; }
        public PolSourceObservationService _PolSourceObservationService { get; private set; }
        public PolSourceObservationIssueService _PolSourceObservationIssueService { get; private set; }

        public BaseEnumService _BaseEnumService { get; set; }
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
            _ProvinceToolsService = new ProvinceToolsService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _PolSourceSiteService = new PolSourceSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _PolSourceObservationService = new PolSourceObservationService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _PolSourceObservationIssueService = new PolSourceObservationIssueService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            _BaseEnumService = new BaseEnumService(_TaskRunnerBaseService._BWObj.appTaskModel.Language);
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

            sb.AppendLine(@"</Document>");
            sb.AppendLine(@"</kml>");


            StreamWriter sw = fi.CreateText();
            sw.Write(sb.ToString());
            sw.Close();
            
            //UnicodeEncoding encoding = new UnicodeEncoding();

            //FileStream fs = fi.Create();
            //byte[] bytes = encoding.GetBytes(sb.ToString());
            //fs.Write(bytes, 0, bytes.Length);
            //fs.Close();

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.KMZOfMWQMSites, FilePurposeEnum.OpenData);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;
        }
        public void ProvinceToolsCreateClassificationInputsKML()
        {
            string NotUsed = "";
            int ProvinceTVItemID = _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID;

            if (ProvinceTVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.Parameter_NotFound, "ProvinceTVItemID");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Parameter_NotFound", "ProvinceTVItemID");
                return;
            }

            TVItemModel tvItemModelProv = _TVItemService.GetTVItemModelWithTVItemIDDB(ProvinceTVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelProv.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, ProvinceTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotDeleteFile_Error_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, ProvinceTVItemID.ToString());
                return;
            }

            string Init = _ProvinceToolsService.GetInit(ProvinceTVItemID);

            string ServerPath = _TVFileService.GetServerFilePath(ProvinceTVItemID);

            string FileName = $"ClassificationInputs_{Init}.kml";

            FileInfo fi = new FileInfo(ServerPath + FileName);

            if (fi.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.FileAlreadyExist_, fi.Name);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("FileAlreadyExist_", fi.Name);
                return;
            }

            TVItemModel tvItemModelTVFile = _TVItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelProv.TVItemID, fi.Name, TVTypeEnum.File);
            if (!string.IsNullOrWhiteSpace(tvItemModelTVFile.Error))
            {
                tvItemModelTVFile = _TVItemService.PostAddChildTVItemDB(tvItemModelProv.TVItemID, fi.Name, TVTypeEnum.File);
                if (!string.IsNullOrWhiteSpace(tvItemModelTVFile.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVItem, tvItemModelTVFile.Error);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVItem, tvItemModelTVFile.Error);
                    return;
                }
            }

            StringBuilder sb = new StringBuilder();

            sb.AppendLine($@"<?xml version=""1.0"" encoding=""UTF-8""?>");
            sb.AppendLine($@"<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom"">");
            sb.AppendLine($@"<Document>");
            sb.AppendLine($@"	<name>{FileName}</name>");
            sb.AppendLine($@"	<StyleMap id=""msn_ylw-pushpin"">");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>normal</key>");
            sb.AppendLine($@"			<styleUrl>#sn_ylw-pushpin</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>highlight</key>");
            sb.AppendLine($@"			<styleUrl>#sh_ylw-pushpin</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"	</StyleMap>");
            sb.AppendLine($@"	<Style id=""sn_ylw-pushpin"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<scale>1.1</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ffff00ff</color>");
            sb.AppendLine($@"			<width>2</width>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@"		<PolyStyle>");
            sb.AppendLine($@"			<color>00ffffff</color>");
            sb.AppendLine($@"		</PolyStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<Style id=""sh_ylw-pushpin"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<scale>1.3</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ff00ff00</color>");
            sb.AppendLine($@"			<width>2</width>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@"		<PolyStyle>");
            sb.AppendLine($@"			<color>00ffffff</color>");
            sb.AppendLine($@"		</PolyStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<Folder>");
            sb.AppendLine($@"		<name>NB Subsectors</name>");
            sb.AppendLine($@"		<open>1</open>");

            List<TVItemModel> tvitemModelSSList = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelProv.TVItemID, TVTypeEnum.Subsector);

            foreach (TVItemModel tvItemModelSS in tvitemModelSSList)
            {
                string TVText = tvItemModelSS.TVText.Substring(0, tvItemModelSS.TVText.IndexOf(" "));

                sb.AppendLine($@"		<Folder>");
                sb.AppendLine($@"			<name>{TVText}</name> ");
                sb.AppendLine($@"			<open>1</open>");
                sb.AppendLine($@"			<Placemark>");
                sb.AppendLine($@"				<name>Subsector Polygon</name>");
                sb.AppendLine($@"				<styleUrl>#msn_ylw-pushpin</styleUrl>");
                sb.AppendLine($@"				<Polygon>");
                sb.AppendLine($@"					<tessellate>1</tessellate>");
                sb.AppendLine($@"					<outerBoundaryIs>");
                sb.AppendLine($@"						<LinearRing>");
                sb.AppendLine($@"							<coordinates>");
                using (CSSPWebToolsDBEntities db = new CSSPWebToolsDBEntities())
                {
                    List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelSS.TVItemID, TVTypeEnum.Subsector, MapInfoDrawTypeEnum.Polygon);

                    foreach (MapInfoPointModel mapInfoPointModel in mapInfoPointModelList)
                    {
                        sb.AppendLine($@"{mapInfoPointModel.Lng.ToString("F6")},{mapInfoPointModel.Lat.ToString("F6")},0 ");
                    }
                }

                sb.AppendLine($@"							</coordinates>");
                sb.AppendLine($@"						</LinearRing>");
                sb.AppendLine($@"					</outerBoundaryIs>");
                sb.AppendLine($@"				</Polygon>");
                sb.AppendLine($@"			</Placemark>");



                sb.AppendLine($@"		</Folder>");
            }

            sb.AppendLine($@"	</Folder>");
            sb.AppendLine($@"</Document>");
            sb.AppendLine($@"</kml>");

            StreamWriter sw = fi.CreateText();
            sw.Write(sb.ToString());
            sw.Close();

            //UnicodeEncoding encoding = new UnicodeEncoding();

            //FileStream fs = fi.Create();
            //byte[] bytes = encoding.GetBytes(sb.ToString());
            //fs.Write(bytes, 0, bytes.Length);
            //fs.Close();

            fi = new FileInfo(ServerPath + FileName);

            TVFileModel tvFileModelNew = new TVFileModel();
            tvFileModelNew.FileCreatedDate_UTC = fi.CreationTimeUtc;
            tvFileModelNew.FileDescription = "Classification Inputs KML file";
            tvFileModelNew.FileInfo = "Classification Input KML file";
            tvFileModelNew.FilePurpose = FilePurposeEnum.Information;
            tvFileModelNew.FileSize_kb = (int)(fi.Length / 1024);
            tvFileModelNew.FileType = FileTypeEnum.KML;
            tvFileModelNew.FromWater = false;
            tvFileModelNew.Language = LanguageEnum.en;
            tvFileModelNew.Parameters = "";
            tvFileModelNew.ReportTypeID = null;
            tvFileModelNew.ServerFileName = FileName;
            tvFileModelNew.ServerFilePath = ServerPath;
            tvFileModelNew.TemplateTVType = null;
            tvFileModelNew.TVFileTVItemID = tvItemModelTVFile.TVItemID;
            tvFileModelNew.TVFileTVText = "";
            tvFileModelNew.Year = DateTime.Now.Year;

            TVFileModel tvFileModelRet = _TVFileService.PostAddTVFileDB(tvFileModelNew);
            if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                return;
            }

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelTVFile, TaskRunnerServiceRes.ClassificationInputsKML, FilePurposeEnum.Information);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;
        }
        public void ProvinceToolsCreateGroupingInputsKML()
        {
            string NotUsed = "";
            int ProvinceTVItemID = _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID;

            if (ProvinceTVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.Parameter_NotFound, "ProvinceTVItemID");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Parameter_NotFound", "ProvinceTVItemID");
                return;
            }

            TVItemModel tvItemModelProv = _TVItemService.GetTVItemModelWithTVItemIDDB(ProvinceTVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelProv.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, ProvinceTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotDeleteFile_Error_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, ProvinceTVItemID.ToString());
                return;
            }

            string Init = _ProvinceToolsService.GetInit(ProvinceTVItemID);

            string ServerPath = _TVFileService.GetServerFilePath(ProvinceTVItemID);

            string FileName = $"GroupingInputs_{Init}.kml";

            FileInfo fi = new FileInfo(ServerPath + FileName);

            if (fi.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.FileAlreadyExist_, fi.Name);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("FileAlreadyExist_", fi.Name);
                return;
            }

            TVItemModel tvItemModelTVFile = _TVItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelProv.TVItemID, fi.Name, TVTypeEnum.File);
            if (!string.IsNullOrWhiteSpace(tvItemModelTVFile.Error))
            {
                tvItemModelTVFile = _TVItemService.PostAddChildTVItemDB(tvItemModelProv.TVItemID, fi.Name, TVTypeEnum.File);
                if (!string.IsNullOrWhiteSpace(tvItemModelTVFile.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVItem, tvItemModelTVFile.Error);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVItem, tvItemModelTVFile.Error);
                    return;
                }
            }

            StringBuilder sb = new StringBuilder();

            sb.AppendLine($@"<?xml version=""1.0"" encoding=""UTF-8""?>");
            sb.AppendLine($@"<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom"">");
            sb.AppendLine($@"<Document>");
            sb.AppendLine($@"	<name>{ FileName }</name>");
            sb.AppendLine($@"	<open>0</open>");

            // Style for Polygon
            sb.AppendLine($@"	<StyleMap id=""msn_ylw-pushpin"">");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>normal</key>");
            sb.AppendLine($@"			<styleUrl>#sn_ylw-pushpin</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>highlight</key>");
            sb.AppendLine($@"			<styleUrl>#sh_ylw-pushpin</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"	</StyleMap>");
            sb.AppendLine($@"	<Style id=""sn_ylw-pushpin"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<scale>1.1</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ffffffff</color>");
            sb.AppendLine($@"			<width>2</width>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@"		<PolyStyle>");
            sb.AppendLine($@"			<color>00ffffff</color>");
            sb.AppendLine($@"		</PolyStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<Style id=""sh_ylw-pushpin"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<scale>1.3</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ffffffff</color>");
            sb.AppendLine($@"			<width>2</width>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@"		<PolyStyle>");
            sb.AppendLine($@"			<color>00ffffff</color>");
            sb.AppendLine($@"		</PolyStyle>");
            sb.AppendLine($@"	</Style>");

            sb.AppendLine($@"	<Style id=""s_ylw-pushpin"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<scale>1.1</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ffff00ff</color>");
            sb.AppendLine($@"			<width>2</width>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@"		<PolyStyle>");
            sb.AppendLine($@"			<color>00ffffff</color>");
            sb.AppendLine($@"		</PolyStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<Style id=""s_ylw-pushpin_hl"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<scale>1.3</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ffff00ff</color>");
            sb.AppendLine($@"			<width>2</width>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@"		<PolyStyle>");
            sb.AppendLine($@"			<color>00ffffff</color>");
            sb.AppendLine($@"		</PolyStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<StyleMap id=""m_ylw-pushpin"">");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>normal</key>");
            sb.AppendLine($@"			<styleUrl>#s_ylw-pushpin</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>highlight</key>");
            sb.AppendLine($@"			<styleUrl>#s_ylw-pushpin_hl</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@" </StyleMap>");

            List<TVItemModel> tvitemModelSSList = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelProv.TVItemID, TVTypeEnum.Subsector);

            foreach (TVItemModel tvItemModelSS in tvitemModelSSList)
            {
                //lblStatus.Text = tvItemModelSS.TVText;
                //lblStatus.Refresh();
                //Application.DoEvents();

                string TVText = tvItemModelSS.TVText.Substring(0, tvItemModelSS.TVText.IndexOf(" "));

                // ---------------------------------------------------------------
                // doing Subsector
                // ---------------------------------------------------------------
                sb.AppendLine($@"		<Folder>");
                sb.AppendLine($@"			<name>{TVText}</name> ");
                sb.AppendLine($@"		</Folder>");
            }

            sb.AppendLine($@"</Document>");
            sb.AppendLine($@"</kml>");

            StreamWriter sw = fi.CreateText();
            sw.Write(sb.ToString());
            sw.Close();
            
            //UnicodeEncoding encoding = new UnicodeEncoding();

            //FileStream fs = fi.Create();
            //byte[] bytes = encoding.GetBytes(sb.ToString());
            //fs.Write(bytes, 0, bytes.Length);
            //fs.Close();

            fi = new FileInfo(ServerPath + FileName);

            TVFileModel tvFileModelNew = new TVFileModel();
            tvFileModelNew.FileCreatedDate_UTC = fi.CreationTimeUtc;
            tvFileModelNew.FileDescription = "Grouping Inputs KML file";
            tvFileModelNew.FileInfo = "Grouping Input KML file";
            tvFileModelNew.FilePurpose = FilePurposeEnum.Information;
            tvFileModelNew.FileSize_kb = (int)(fi.Length / 1024);
            tvFileModelNew.FileType = FileTypeEnum.KML;
            tvFileModelNew.FromWater = false;
            tvFileModelNew.Language = LanguageEnum.en;
            tvFileModelNew.Parameters = "";
            tvFileModelNew.ReportTypeID = null;
            tvFileModelNew.ServerFileName = FileName;
            tvFileModelNew.ServerFilePath = ServerPath;
            tvFileModelNew.TemplateTVType = null;
            tvFileModelNew.TVFileTVItemID = tvItemModelTVFile.TVItemID;
            tvFileModelNew.TVFileTVText = "";
            tvFileModelNew.Year = DateTime.Now.Year;

            TVFileModel tvFileModelRet = _TVFileService.PostAddTVFileDB(tvFileModelNew);
            if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                return;
            }

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelTVFile, TaskRunnerServiceRes.GroupingInputsKML, FilePurposeEnum.Information);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;
        }
        public void ProvinceToolsCreateMWQMSitesAndPolSourceSitesKML()
        {
            string NotUsed = "";
            int ProvinceTVItemID = _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID;

            if (ProvinceTVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.Parameter_NotFound, "ProvinceTVItemID");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Parameter_NotFound", "ProvinceTVItemID");
                return;
            }

            TVItemModel tvItemModelProv = _TVItemService.GetTVItemModelWithTVItemIDDB(ProvinceTVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelProv.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, ProvinceTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotDeleteFile_Error_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, ProvinceTVItemID.ToString());
                return;
            }

            string Init = _ProvinceToolsService.GetInit(ProvinceTVItemID);

            string ServerPath = _TVFileService.GetServerFilePath(ProvinceTVItemID);

            string FileName = $"MWQMSitesAndPolSourceSites_{Init}.kml";

            FileInfo fi = new FileInfo(ServerPath + FileName);

            if (fi.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.FileAlreadyExist_, fi.Name);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("FileAlreadyExist_", fi.Name);
                return;
            }

            TVItemModel tvItemModelTVFile = _TVItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(tvItemModelProv.TVItemID, fi.Name, TVTypeEnum.File);
            if (!string.IsNullOrWhiteSpace(tvItemModelTVFile.Error))
            {
                tvItemModelTVFile = _TVItemService.PostAddChildTVItemDB(tvItemModelProv.TVItemID, fi.Name, TVTypeEnum.File);
                if (!string.IsNullOrWhiteSpace(tvItemModelTVFile.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVItem, tvItemModelTVFile.Error);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVItem, tvItemModelTVFile.Error);
                    return;
                }
            }

            StringBuilder sb = new StringBuilder();

            sb.AppendLine($@"<?xml version=""1.0"" encoding=""UTF-8""?>");
            sb.AppendLine($@"<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom"">");
            sb.AppendLine($@"<Document>");
            sb.AppendLine($@"	<name>{ FileName }</name>");
            sb.AppendLine($@"	<open>0</open>");
            sb.AppendLine($@"	<visibility>0</visibility>");

            // Style for Subsector
            sb.AppendLine($@"	<StyleMap id=""msn_ylw-pushpin"">");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>normal</key>");
            sb.AppendLine($@"			<styleUrl>#sn_ylw-pushpin</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>highlight</key>");
            sb.AppendLine($@"			<styleUrl>#sh_ylw-pushpin</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"	</StyleMap>");
            sb.AppendLine($@"	<Style id=""sn_ylw-pushpin"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<scale>1.1</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ffff00ff</color>");
            sb.AppendLine($@"			<width>2</width>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@"		<PolyStyle>");
            sb.AppendLine($@"			<color>00ffffff</color>");
            sb.AppendLine($@"		</PolyStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<Style id=""sh_ylw-pushpin"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<scale>1.3</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ff00ff00</color>");
            sb.AppendLine($@"			<width>2</width>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@"		<PolyStyle>");
            sb.AppendLine($@"			<color>00ffffff</color>");
            sb.AppendLine($@"		</PolyStyle>");
            sb.AppendLine($@"	</Style>");

            // Style for Pollution Source Sites
            sb.AppendLine($@"	<Style id=""s_ylw-pushpin"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ffff0000</color>");
            sb.AppendLine($@"			<scale>1.2</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<ListStyle>");
            sb.AppendLine($@"		</ListStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<StyleMap id=""m_ylw-pushpin"">");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>normal</key>");
            sb.AppendLine($@"			<styleUrl>#s_ylw-pushpin</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>highlight</key>");
            sb.AppendLine($@"			<styleUrl>#s_ylw-pushpin_hl</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"	</StyleMap>");
            sb.AppendLine($@"	<Style id=""s_ylw-pushpin_hl"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ffff0000</color>");
            sb.AppendLine($@"			<scale>1.2</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle_highlight.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<ListStyle>");
            sb.AppendLine($@"		</ListStyle>");
            sb.AppendLine($@" </Style>");

            // Style for MWQM Sites
            sb.AppendLine($@"	<StyleMap id=""msn_placemark_square"">");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>normal</key>");
            sb.AppendLine($@"			<styleUrl>#sn_placemark_square</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>highlight</key>");
            sb.AppendLine($@"			<styleUrl>#sh_placemark_square_highlight</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"	</StyleMap>");
            sb.AppendLine($@"	<Style id=""sh_placemark_square_highlight"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff00ff00</color>");
            sb.AppendLine($@"			<scale>1.2</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_square_highlight.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<ListStyle>");
            sb.AppendLine($@"		</ListStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<Style id=""sn_placemark_square"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff00ff00</color>");
            sb.AppendLine($@"			<scale>1.2</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_square.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<ListStyle>");
            sb.AppendLine($@"		</ListStyle>");
            sb.AppendLine($@" </Style>");

            // Style for Infrastructures
            sb.AppendLine($@"	<Style id=""sn_shaded_dot"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff0000ff</color>");
            sb.AppendLine($@"			<scale>1.2</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<ListStyle>");
            sb.AppendLine($@"		</ListStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<Style id=""sh_shaded_dot"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff0000ff</color>");
            sb.AppendLine($@"			<scale>1.2</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<ListStyle>");
            sb.AppendLine($@"		</ListStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<StyleMap id=""msn_shaded_dot"">");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>normal</key>");
            sb.AppendLine($@"			<styleUrl>#sn_shaded_dot</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>highlight</key>");
            sb.AppendLine($@"			<styleUrl>#sh_shaded_dot</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@" </StyleMap>");

            List<TVItemModel> tvitemModelSSList = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelProv.TVItemID, TVTypeEnum.Subsector);

            int TotalCount = tvitemModelSSList.Count;
            int Count = 0;
            foreach (TVItemModel tvItemModelSS in tvitemModelSSList)
            {
                Count += 1;

                if (Count % 5 == 0)
                {
                    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)(100.0f * ((float)Count/(float)TotalCount)));
                }

                string TVText = tvItemModelSS.TVText.Substring(0, tvItemModelSS.TVText.IndexOf(" "));

                // ---------------------------------------------------------------
                // doing Subsector
                // ---------------------------------------------------------------
                sb.AppendLine($@"		<Folder>");
                sb.AppendLine($@"			<name>{TVText}</name> ");
                sb.AppendLine($@"			<Placemark>");
                sb.AppendLine($@"				<name>Subsector Polygon</name>");
                sb.AppendLine($@"	            <visibility>0</visibility>");
                sb.AppendLine($@"				<styleUrl>#msn_ylw-pushpin</styleUrl>");
                sb.AppendLine($@"				<Polygon>");
                sb.AppendLine($@"					<outerBoundaryIs>");
                sb.AppendLine($@"						<LinearRing>");
                sb.AppendLine($@"							<coordinates>");
                using (CSSPWebToolsDBEntities db = new CSSPWebToolsDBEntities())
                {
                    List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModelSS.TVItemID, TVTypeEnum.Subsector, MapInfoDrawTypeEnum.Polygon);

                    foreach (MapInfoPointModel mapInfoPointModel in mapInfoPointModelList)
                    {
                        sb.AppendLine($@"{mapInfoPointModel.Lng.ToString("F6")},{mapInfoPointModel.Lat.ToString("F6")},0 ");
                    }
                }

                sb.AppendLine($@"							</coordinates>");
                sb.AppendLine($@"						</LinearRing>");
                sb.AppendLine($@"					</outerBoundaryIs>");
                sb.AppendLine($@"				</Polygon>");
                sb.AppendLine($@"			</Placemark>");

                // ---------------------------------------------------------------
                // doing Pollution Source Sites
                // ---------------------------------------------------------------
                sb.AppendLine($@"		    <Folder>");
                sb.AppendLine($@"			    <name>Pollution Source Sites</name> ");
                sb.AppendLine($@"	            <visibility>0</visibility>");

                List<TVItemModel> tvitemModelPSSList = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSS.TVItemID, TVTypeEnum.PolSourceSite).Where(c => c.IsActive == true).ToList();
                List<PolSourceSiteModel> polSourceSiteModelList = _PolSourceSiteService.GetPolSourceSiteModelListWithSubsectorTVItemIDDB(tvItemModelSS.TVItemID);
                List<PolSourceObservationModel> polSourceObservationModelList = _PolSourceObservationService.GetPolSourceObservationModelListWithSubsectorTVItemIDDB(tvItemModelSS.TVItemID);
                List<PolSourceObservationIssueModel> polSourceObservationIssueModelList = _PolSourceObservationIssueService.GetPolSourceObservationIssueModelListWithSubsectorTVItemIDDB(tvItemModelSS.TVItemID);

                foreach (TVItemModel tvItemModel in tvitemModelPSSList)
                {
                    PolSourceSiteModel polSourceSiteModel = polSourceSiteModelList.Where(c => c.PolSourceSiteTVItemID == tvItemModel.TVItemID).FirstOrDefault();

                    sb.AppendLine($@"			    <Placemark>");
                    sb.AppendLine($@"			    <name>P{ polSourceSiteModel.Site }</name>");
                    sb.AppendLine($@"	            <visibility>0</visibility>");
                    sb.AppendLine($@"               <description><![CDATA[");

                    if (polSourceObservationModelList.Count > 0)
                    {
                        PolSourceObservationModel polSourceObservationModel = polSourceObservationModelList.OrderByDescending(c => c.ObservationDate_Local).FirstOrDefault();

                        if (polSourceObservationModel != null)
                        {
                            sb.AppendLine($@"                <h3>Last Observation</h3>");
                            sb.AppendLine($@"                <span data-tvitemid=""{tvItemModel.TVItemID}"">&nbsp;</span>");
                            sb.AppendLine($@"                <blockquote>");
                            sb.AppendLine($@"                <p><b>Date:</b> {((DateTime)polSourceObservationModel.ObservationDate_Local).ToString("yyyy MMMM dd")}</p>");
                            sb.AppendLine($@"                <p><b>Observation Last Update (UTC):</b> {((DateTime)polSourceObservationModel.LastUpdateDate_UTC).ToString("yyyy MMMM dd HH:mm:ss")}</p>");
                            sb.AppendLine($@"                <p><b>Old Written Description:</b> {polSourceObservationModel.Observation_ToBeDeleted}</p>");

                            List<PolSourceObservationIssueModel> polSourceObsIssueModelList = polSourceObservationIssueModelList.Where(c => c.PolSourceObservationID == polSourceObservationModel.PolSourceObservationID).OrderBy(c => c.Ordinal).ToList();
                            if (polSourceObsIssueModelList.Count > 0)
                            {
                                sb.AppendLine($@"                <blockquote>");
                                sb.AppendLine($@"                <ol>");
                                foreach (PolSourceObservationIssueModel polSourceObservationIssueModel in polSourceObsIssueModelList)
                                {
                                    sb.AppendLine($@"                <li>");
                                    sb.AppendLine($@"                <p><b>Issue Last Update (UTC):</b> {((DateTime)polSourceObservationIssueModel.LastUpdateDate_UTC).ToString("yyyy MMMM dd HH:mm:ss")}</p>");

                                    string TVTextIssue = "";

                                    if (!string.IsNullOrWhiteSpace(polSourceObservationIssueModel.ObservationInfo.Trim()))
                                    {
                                        polSourceObservationIssueModel.ObservationInfo = polSourceObservationIssueModel.ObservationInfo.Trim();
                                        List<int> PolSourceObsInfoIntList = polSourceObservationIssueModel.ObservationInfo.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Select(c => int.Parse(c)).ToList();

                                        for (int i = 0, count = PolSourceObsInfoIntList.Count; i < count; i++)
                                        {
                                            string Temp = _BaseEnumService.GetEnumText_PolSourceObsInfoReportEnum((PolSourceObsInfoEnum)PolSourceObsInfoIntList[i]);
                                            switch ((PolSourceObsInfoIntList[i].ToString()).Substring(0, 3))
                                            {
                                                case "101":
                                                    {
                                                        Temp = Temp.Replace("Source", "<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>Source</strong>");
                                                    }
                                                    break;
                                                //case "153":
                                                //    {
                                                //        Temp = Temp.Replace("Dilution Analyses", "<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>Dilution Analyses</strong>");
                                                //    }
                                                //    break;
                                                case "250":
                                                    {
                                                        Temp = Temp.Replace("Pathway", "<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>Pathway</strong>");
                                                    }
                                                    break;
                                                case "900":
                                                    {
                                                        Temp = Temp.Replace("Status", "<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>Status</strong>");
                                                    }
                                                    break;
                                                case "910":
                                                    {
                                                        Temp = Temp.Replace("Risk", "<strong>Risk</strong>");
                                                    }
                                                    break;
                                                case "110":
                                                case "120":
                                                case "122":
                                                case "151":
                                                case "152":
                                                case "153":
                                                case "155":
                                                case "156":
                                                case "157":
                                                case "163":
                                                case "166":
                                                case "167":
                                                case "170":
                                                case "171":
                                                case "172":
                                                case "173":
                                                case "176":
                                                case "178":
                                                case "181":
                                                case "182":
                                                case "183":
                                                case "185":
                                                case "186":
                                                case "187":
                                                case "190":
                                                case "191":
                                                case "192":
                                                case "193":
                                                case "194":
                                                case "196":
                                                case "198":
                                                case "199":
                                                case "220":
                                                case "930":
                                                    {
                                                        Temp = @"<span>" + Temp + "</span>";
                                                    }
                                                    break;
                                                default:
                                                    break;
                                            }
                                            TVTextIssue = TVTextIssue + Temp;
                                        }
                                        sb.AppendLine($@"                <p><b>Selected:</b> {TVTextIssue}</p>");
                                    }
                                    sb.AppendLine($@"                </li>");
                                }
                                sb.AppendLine($@"                </ol>");
                                sb.AppendLine($@"                </blockquote>");
                            }
                            sb.AppendLine($@"                </blockquote>");
                        }
                    }
                    sb.AppendLine($@"                   ]]></description>");
                    sb.AppendLine($@"			    	<styleUrl>#s_ylw-pushpin</styleUrl>");
                    sb.AppendLine($@"			    	<Point>");
                    sb.AppendLine($@"		    			<coordinates>");
                    using (CSSPWebToolsDBEntities db = new CSSPWebToolsDBEntities())
                    {
                        List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModel.TVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);

                        foreach (MapInfoPointModel mapInfoPointModel in mapInfoPointModelList)
                        {
                            sb.AppendLine($@"{mapInfoPointModel.Lng.ToString("F6")},{mapInfoPointModel.Lat.ToString("F6")},0 ");
                        }
                    }

                    sb.AppendLine($@"						</coordinates>");
                    sb.AppendLine($@"		    		</Point>");
                    sb.AppendLine($@"		    	</Placemark>");
                }
                sb.AppendLine($@"		    </Folder>");

                // ---------------------------------------------------------------
                // doing MWQM Sites
                // ---------------------------------------------------------------
                sb.AppendLine($@"		    <Folder>");
                sb.AppendLine($@"			    <name>MWQM Sites</name> ");
                sb.AppendLine($@"	            <visibility>0</visibility>");

                List<TVItemModel> tvitemModelMWQMSiteList = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSS.TVItemID, TVTypeEnum.MWQMSite).Where(c => c.IsActive == true).ToList();
                foreach (TVItemModel tvItemModel in tvitemModelMWQMSiteList)
                {
                    sb.AppendLine($@"			    <Placemark>");
                    sb.AppendLine($@"			    <name>S{tvItemModel.TVText}</name>");
                    sb.AppendLine($@"	            <visibility>0</visibility>");
                    sb.AppendLine($@"               <description><![CDATA[");
                    sb.AppendLine($@"                <span data-tvitemid=""{tvItemModel.TVItemID}"">&nbsp;</span>");
                    sb.AppendLine($@"                   ]]></description>");
                    sb.AppendLine($@"			    	<styleUrl>#msn_placemark_square</styleUrl>");
                    sb.AppendLine($@"			    	<Point>");
                    sb.AppendLine($@"		    			<coordinates>");
                    using (CSSPWebToolsDBEntities db = new CSSPWebToolsDBEntities())
                    {
                        List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModel.TVItemID, TVTypeEnum.MWQMSite, MapInfoDrawTypeEnum.Point);

                        foreach (MapInfoPointModel mapInfoPointModel in mapInfoPointModelList)
                        {
                            sb.AppendLine($@"{mapInfoPointModel.Lng.ToString("F6")},{mapInfoPointModel.Lat.ToString("F6")},0 ");
                        }
                    }

                    sb.AppendLine($@"						</coordinates>");
                    sb.AppendLine($@"		    		</Point>");
                    sb.AppendLine($@"		    	</Placemark>");
                }
                sb.AppendLine($@"		    </Folder>");

                // ---------------------------------------------------------------
                // doing Municipality
                // ---------------------------------------------------------------
                sb.AppendLine($@"		    <Folder>");
                sb.AppendLine($@"			    <name>Municipalities</name> ");
                sb.AppendLine($@"	            <visibility>0</visibility>");

                List<TVItemModel> tvItemModelMuniList = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSS.TVItemID, TVTypeEnum.Municipality).Where(c => c.IsActive == true).ToList();
                foreach (TVItemModel tvItemModelMuni in tvItemModelMuniList)
                {
                    sb.AppendLine($@"		        <Folder>");
                    sb.AppendLine($@"			        <name>{tvItemModelMuni.TVText}</name> ");
                    sb.AppendLine($@"	                <visibility>0</visibility>");
                    List<TVItemModel> tvItemModelInfraList = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelMuni.TVItemID, TVTypeEnum.Infrastructure).Where(c => c.IsActive == true).ToList();
                    foreach (TVItemModel tvItemModel in tvItemModelInfraList)
                    {
                        List<TVTypeEnum> tvTypeInfList = new List<TVTypeEnum>() { TVTypeEnum.WasteWaterTreatmentPlant, TVTypeEnum.LiftStation, TVTypeEnum.LineOverflow };
                        foreach (TVTypeEnum tvTypeInf in tvTypeInfList)
                        {
                            List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModel.TVItemID, tvTypeInf, MapInfoDrawTypeEnum.Point);

                            if (mapInfoPointModelList.Count > 0)
                            {
                                sb.AppendLine($@"			        <Placemark>");
                                sb.AppendLine($@"			        	<name>{tvItemModel.TVText}</name>");
                                sb.AppendLine($@"	                    <visibility>0</visibility>");
                                sb.AppendLine($@"                       <description><![CDATA[");
                                sb.AppendLine($@"                        <span data-tvitemid=""{tvItemModel.TVItemID}"">&nbsp;</span>");
                                sb.AppendLine($@"                           ]]></description>");
                                sb.AppendLine($@"			                	<styleUrl>#sn_shaded_dot</styleUrl>");
                                sb.AppendLine($@"			                	<Point>");
                                sb.AppendLine($@"		                			<coordinates>");
                                using (CSSPWebToolsDBEntities db = new CSSPWebToolsDBEntities())
                                {
                                    foreach (MapInfoPointModel mapInfoPointModel in mapInfoPointModelList)
                                    {
                                        sb.AppendLine($@"{mapInfoPointModel.Lng.ToString("F6")},{mapInfoPointModel.Lat.ToString("F6")},0 ");
                                    }
                                }

                                sb.AppendLine($@"				            		</coordinates>");
                                sb.AppendLine($@"		    	            	</Point>");
                                sb.AppendLine($@"		    	            </Placemark>");
                            }
                        }

                    }
                    sb.AppendLine($@"		      </Folder>");
                }
                sb.AppendLine($@"		    </Folder>");

                sb.AppendLine($@"		</Folder>");
            }

            sb.AppendLine($@"</Document>");
            sb.AppendLine($@"</kml>");

            StreamWriter sw = fi.CreateText();
            sw.Write(sb.ToString());
            sw.Close();

            //UnicodeEncoding encoding = new UnicodeEncoding();

            //FileStream fs = fi.Create();
            //byte[] bytes = encoding.GetBytes(sb.ToString());
            //fs.Write(bytes, 0, bytes.Length);
            //fs.Close();

            fi = new FileInfo(ServerPath + FileName);

            TVFileModel tvFileModelNew = new TVFileModel();
            tvFileModelNew.FileCreatedDate_UTC = fi.CreationTimeUtc;
            tvFileModelNew.FileDescription = "MWQM Sites and Pollution Source Sites KML file";
            tvFileModelNew.FileInfo = "MWQM Sites and Pollution Source Sites KML file";
            tvFileModelNew.FilePurpose = FilePurposeEnum.Information;
            tvFileModelNew.FileSize_kb = (int)(fi.Length / 1024);
            tvFileModelNew.FileType = FileTypeEnum.KML;
            tvFileModelNew.FromWater = false;
            tvFileModelNew.Language = LanguageEnum.en;
            tvFileModelNew.Parameters = "";
            tvFileModelNew.ReportTypeID = null;
            tvFileModelNew.ServerFileName = FileName;
            tvFileModelNew.ServerFilePath = ServerPath;
            tvFileModelNew.TemplateTVType = null;
            tvFileModelNew.TVFileTVItemID = tvItemModelTVFile.TVItemID;
            tvFileModelNew.TVFileTVText = "";
            tvFileModelNew.Year = DateTime.Now.Year;

            TVFileModel tvFileModelRet = _TVFileService.PostAddTVFileDB(tvFileModelNew);
            if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                return;
            }

            _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelTVFile, TaskRunnerServiceRes.MWQMSitesAndPolSourceSitesKML, FilePurposeEnum.Information);
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
