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
using CSSPDBDLL.Models;
using CSSPDBDLL;
using CSSPDBDLL.Services;
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
        public MWQMSiteService _MWQMSiteService { get; private set; }
        public TVItemLinkService _TVItemLinkService { get; private set; }
        public UseOfSiteService _UseOfSiteService { get; private set; }
        public InfrastructureService _InfrastructureService { get; private set; }


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
            _MWQMSiteService = new MWQMSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _TVItemLinkService = new TVItemLinkService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _UseOfSiteService = new UseOfSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _InfrastructureService = new InfrastructureService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

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

            sb.AppendLine(@"<StyleMap id=""m_ylw-pushpin"">");
            sb.AppendLine(@"	<Pair>");
            sb.AppendLine(@"		<key>normal</key>");
            sb.AppendLine(@"		<styleUrl>#s_ylw-pushpin</styleUrl>");
            sb.AppendLine(@"	</Pair>");
            sb.AppendLine(@"	<Pair>");
            sb.AppendLine(@"		<key>highlight</key>");
            sb.AppendLine(@"		<styleUrl>#s_ylw-pushpin_hl</styleUrl>");
            sb.AppendLine(@"	</Pair>");
            sb.AppendLine(@"</StyleMap>");
            sb.AppendLine(@"<Style id=""s_ylw-pushpin"">");
            sb.AppendLine(@"	<IconStyle>");
            sb.AppendLine(@"		<color>ff00ff00</color>");
            sb.AppendLine(@"		<scale>0.8</scale>");
            sb.AppendLine(@"		<Icon>");
            sb.AppendLine(@"			<href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png</href>");
            sb.AppendLine(@"		</Icon>");
            sb.AppendLine(@"	</IconStyle>");
            sb.AppendLine(@"	<LabelStyle>");
            sb.AppendLine(@"		<scale>0.8</scale>");
            sb.AppendLine(@"	</LabelStyle>");
            sb.AppendLine(@"	<LineStyle>");
            sb.AppendLine(@"		<color>ff00ff00</color>");
            sb.AppendLine(@"		<width>2</width>");
            sb.AppendLine(@"	</LineStyle>");
            sb.AppendLine(@"	<PolyStyle>");
            sb.AppendLine(@"		<color>00ffffff</color>");
            sb.AppendLine(@"	</PolyStyle>");
            sb.AppendLine(@"</Style>");
            sb.AppendLine(@"<Style id=""s_ylw-pushpin_hl"">");
            sb.AppendLine(@"	<IconStyle>");
            sb.AppendLine(@"		<color>ff00ff00</color>");
            sb.AppendLine(@"		<scale>0.8</scale>");
            sb.AppendLine(@"		<Icon>");
            sb.AppendLine(@"			<href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png</href>");
            sb.AppendLine(@"		</Icon>");
            sb.AppendLine(@"	</IconStyle>");
            sb.AppendLine(@"	<LabelStyle>");
            sb.AppendLine(@"		<scale>0.8</scale>");
            sb.AppendLine(@"	</LabelStyle>");
            sb.AppendLine(@"	<LineStyle>");
            sb.AppendLine(@"		<color>ff00ff00</color>");
            sb.AppendLine(@"		<width>2</width>");
            sb.AppendLine(@"	</LineStyle>");
            sb.AppendLine(@"	<PolyStyle>");
            sb.AppendLine(@"		<color>00ffffff</color>");
            sb.AppendLine(@"	</PolyStyle>");
            sb.AppendLine(@"</Style>");


            using (CSSPDBEntities db = new CSSPDBEntities())
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
                                    let mip = (from mi in db.MapInfos
                                               from mip in db.MapInfoPoints
                                               orderby mip.Ordinal ascending
                                               where mi.TVItemID == t.TVItemID
                                               && mi.MapInfoID == mip.MapInfoID
                                               && mi.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Polygon
                                               && mi.TVType == (int)TVTypeEnum.Subsector
                                               select mip).ToList()
                                    where t.TVItemID == tl.TVItemID
                                    && tl.Language == (int)LanguageEnum.en
                                    && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                    && t.TVType == (int)TVTypeEnum.Subsector
                                    orderby tl.TVText
                                    select new { t, mip, tl }).ToList();

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
                foreach (var tvItemSS in tvItemSSList.OrderBy(c => c.tl.TVText))
                {
                    bool HasSample = false;
                    foreach (var mwqmSite in MonitoringSiteList.Where(c => c.t.ParentID == tvItemSS.t.TVItemID))
                    {
                        if (mwqmSite.hasSample)
                        {
                            HasSample = true;
                            break;
                        }
                    }

                    string More = "";
                    if (HasSample)
                    {
                        More = " (+)";
                    }

                    string Subsector = tvItemSS.tl.TVText;
                    if (Subsector.Contains(" "))
                    {
                        Subsector = Subsector.Substring(0, Subsector.IndexOf(" "));
                    }

                    sb.AppendLine(@"	<Folder>");
                    sb.AppendLine($@"		<name>{Subsector}{More}</name>");
                    sb.AppendLine($@"		<visibility>0</visibility>");
                    if (Count2 % 20 == 0)
                    {
                        _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)(100.0f * ((float)Count2 / (float)TotalCount2)));

                        NotUsed = string.Format(TaskRunnerServiceRes.Creating_, fi.Name + " --- doing " + tvItemSS.tl.TVText + "");
                        TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", fi.Name + " --- doing " + tvItemSS.tl.TVText + "");

                        _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);
                    }

                    Count2 += 1;

                    if (tvItemSS.mip.Count > 0)
                    {

                        sb.AppendLine(@"	    <Placemark>");
                        sb.AppendLine($@"	    	<name>{Subsector}</name>");
                        sb.AppendLine($@"	    	<visibility>0</visibility>");
                        sb.AppendLine(@"	    	<styleUrl>#m_ylw-pushpin</styleUrl>");
                        sb.AppendLine(@"	    	<Polygon>");
                        sb.AppendLine(@"	    	    <outerBoundaryIs>");
                        sb.AppendLine(@"	    	        <LinearRing>");
                        sb.AppendLine($@"	    		        <coordinates>");
                        foreach (MapInfoPoint mipm in tvItemSS.mip)
                        {
                            sb.AppendLine($@"{mipm.Lng},{mipm.Lat},0 ");
                        }
                        sb.AppendLine($@"{tvItemSS.mip[0].Lng},{tvItemSS.mip[0].Lat},0 ");

                        sb.AppendLine($@"	    		        </coordinates>");
                        sb.AppendLine(@"	    	        </LinearRing>");
                        sb.AppendLine(@"	    	    </outerBoundaryIs>");
                        sb.AppendLine(@"	    	</Polygon>");
                        sb.AppendLine(@"	    </Placemark>");

                    }

                    foreach (var mwqmSite in MonitoringSiteList.Where(c => c.t.ParentID == tvItemSS.t.TVItemID).OrderBy(c => c.tl.TVText))
                    {
                        string TVText = mwqmSite.tl.TVText;
                        //string MS = mwqmSite.t.TVItemID.ToString();
                        string Lat = (mwqmSite.mip != null ? mwqmSite.mip.Lat.ToString("F6").Replace(",", ".") : "");
                        string Lng = (mwqmSite.mip != null ? mwqmSite.mip.Lng.ToString("F6").Replace(",", ".") : "");

                        sb.AppendLine(@"	    <Placemark>");
                        sb.AppendLine($@"	    	<name>{ProvInit}_{TVText} - [{mwqmSite.t.TVItemID}]</name>");
                        sb.AppendLine($@"	    	<visibility>0</visibility>");
                        sb.AppendLine(@"	    	<styleUrl>#m_ylw-pushpin</styleUrl>");
                        sb.AppendLine(@"	    	<Point>");
                        sb.AppendLine($@"	    		<coordinates>{Lng},{Lat},0</coordinates>");
                        sb.AppendLine(@"	    	</Point>");
                        sb.AppendLine(@"	    </Placemark>");
                    }

                    sb.AppendLine(@"	</Folder>");
                }
            }

            sb.AppendLine(@"</Document>");
            sb.AppendLine(@"</kml>");


            try
            {
                File.WriteAllText(fi.FullName, sb.ToString(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateFile_Error_, fi.FullName, ex.Message);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateFile_Error_", fi.FullName, ex.Message);
                return;
            }

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
                using (CSSPDBEntities db = new CSSPDBEntities())
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

            try
            {
                File.WriteAllText(fi.FullName, sb.ToString(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateFile_Error_, fi.FullName, ex.Message);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateFile_Error_", fi.FullName, ex.Message);
                return;
            }

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

            try
            {
                File.WriteAllText(fi.FullName, sb.ToString(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateFile_Error_, fi.FullName, ex.Message);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateFile_Error_", fi.FullName, ex.Message);
                return;
            }

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
            sb.AppendLine($@"			<scale>0.8</scale>");
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
            sb.AppendLine($@"			<scale>1.0</scale>");
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

            // Style for Pollution Source Sites Active --- blue circle
            sb.AppendLine($@"	<Style id=""s_ylw-pushpin"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ffff0000</color>");
            sb.AppendLine($@"			<scale>0.8</scale>");
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
            sb.AppendLine($@"			<scale>1.0</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle_highlight.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<ListStyle>");
            sb.AppendLine($@"		</ListStyle>");
            sb.AppendLine($@" </Style>");

            // Style for Pollution Source Sites Inactive --- white circle
            sb.AppendLine($@"	<Style id=""s_ylwI-pushpin"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ffffffff</color>");
            sb.AppendLine($@"			<scale>0.8</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<ListStyle>");
            sb.AppendLine($@"		</ListStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<StyleMap id=""m_ylwI-pushpin"">");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>normal</key>");
            sb.AppendLine($@"			<styleUrl>#s_ylwI-pushpin</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>highlight</key>");
            sb.AppendLine($@"			<styleUrl>#s_ylwI-pushpin_hl</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"	</StyleMap>");
            sb.AppendLine($@"	<Style id=""s_ylwI-pushpin_hl"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ffffffff</color>");
            sb.AppendLine($@"			<scale>1.0</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle_highlight.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<ListStyle>");
            sb.AppendLine($@"		</ListStyle>");
            sb.AppendLine($@" </Style>");

            // Style for MWQM Sites white square --- inactive
            sb.AppendLine($@"	<StyleMap id=""msn_placemark_square_white"">");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>normal</key>");
            sb.AppendLine($@"			<styleUrl>#sn_placemark_square_white</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>highlight</key>");
            sb.AppendLine($@"			<styleUrl>#sh_placemark_square_highlight_white</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"	</StyleMap>");
            sb.AppendLine($@"	<Style id=""sh_placemark_square_highlight_white"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ffffffff</color>");
            sb.AppendLine($@"			<scale>0.8</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_square_highlight.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<ListStyle>");
            sb.AppendLine($@"		</ListStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<Style id=""sn_placemark_square_white"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ffffffff</color>");
            sb.AppendLine($@"			<scale>1.0</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_square.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<ListStyle>");
            sb.AppendLine($@"		</ListStyle>");
            sb.AppendLine($@" </Style>");

            // Style for MWQM Sites green square --- active
            sb.AppendLine($@"	<StyleMap id=""msn_placemark_square_green"">");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>normal</key>");
            sb.AppendLine($@"			<styleUrl>#sn_placemark_square_green</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>highlight</key>");
            sb.AppendLine($@"			<styleUrl>#sh_placemark_square_highlight_green</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"	</StyleMap>");
            sb.AppendLine($@"	<Style id=""sh_placemark_square_highlight_green"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff00ff00</color>");
            sb.AppendLine($@"			<scale>0.8</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_square_highlight.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<ListStyle>");
            sb.AppendLine($@"		</ListStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<Style id=""sn_placemark_square_green"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff00ff00</color>");
            sb.AppendLine($@"			<scale>1.0</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_square.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<ListStyle>");
            sb.AppendLine($@"		</ListStyle>");
            sb.AppendLine($@" </Style>");

            // Style for MWQM Sites red square -- active
            sb.AppendLine($@"	<StyleMap id=""msn_placemark_square_red"">");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>normal</key>");
            sb.AppendLine($@"			<styleUrl>#sn_placemark_square_red</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>highlight</key>");
            sb.AppendLine($@"			<styleUrl>#sh_placemark_square_highlight_red</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"	</StyleMap>");
            sb.AppendLine($@"	<Style id=""sh_placemark_square_highlight_red"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff0000ff</color>");
            sb.AppendLine($@"			<scale>0.8</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_square_highlight.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<ListStyle>");
            sb.AppendLine($@"		</ListStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<Style id=""sn_placemark_square_red"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff0000ff</color>");
            sb.AppendLine($@"			<scale>1.0</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_square.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<ListStyle>");
            sb.AppendLine($@"		</ListStyle>");
            sb.AppendLine($@" </Style>");

            // Style for MWQM Sites purple square -- active
            sb.AppendLine($@"	<StyleMap id=""msn_placemark_square_purple"">");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>normal</key>");
            sb.AppendLine($@"			<styleUrl>#sn_placemark_square_purple</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>highlight</key>");
            sb.AppendLine($@"			<styleUrl>#sh_placemark_square_highlight_purple</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"	</StyleMap>");
            sb.AppendLine($@"	<Style id=""sh_placemark_square_highlight_purple"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff800080</color>");
            sb.AppendLine($@"			<scale>0.8</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_square_highlight.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<ListStyle>");
            sb.AppendLine($@"		</ListStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<Style id=""sn_placemark_square_purple"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff800080</color>");
            sb.AppendLine($@"			<scale>1.0</scale>");
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
            sb.AppendLine($@"			<scale>0.8</scale>");
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
            sb.AppendLine($@"			<scale>1.0</scale>");
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
                //if (!tvItemModelSS.TVText.StartsWith("NB-06-020"))
                //{
                //    continue;
                //}

                Count += 1;

                if (Count % 5 == 0)
                {
                    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)(100.0f * ((float)Count / (float)TotalCount)));
                }

                string TVText = "";
                if (tvItemModelSS.TVText.Contains(" "))
                {
                    TVText = tvItemModelSS.TVText.Substring(0, tvItemModelSS.TVText.IndexOf(" "));
                }
                else
                {
                    TVText = tvItemModelSS.TVText;
                }

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
                using (CSSPDBEntities db = new CSSPDBEntities())
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

                List<TVItemModel> tvitemModelPSSList = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSS.TVItemID, TVTypeEnum.PolSourceSite).ToList();
                List<PolSourceSiteModel> polSourceSiteModelList = _PolSourceSiteService.GetPolSourceSiteModelListWithSubsectorTVItemIDDB(tvItemModelSS.TVItemID);
                List<PolSourceObservationModel> polSourceObservationModelList = _PolSourceObservationService.GetPolSourceObservationModelListWithSubsectorTVItemIDDB(tvItemModelSS.TVItemID);
                List<PolSourceObservationIssueModel> polSourceObservationIssueModelList = _PolSourceObservationIssueService.GetPolSourceObservationIssueModelListWithSubsectorTVItemIDDB(tvItemModelSS.TVItemID);

                // ---------------------------------------------------------------
                // doing Active Pollution Source Sites
                // ---------------------------------------------------------------
                sb.AppendLine($@"		        <Folder>");
                sb.AppendLine($@"			        <name>Active</name> ");
                sb.AppendLine($@"	                <visibility>0</visibility>");
                foreach (TVItemModel tvItemModel in tvitemModelPSSList.Where(c => c.IsActive == true))
                {
                    PolSourceSiteModel polSourceSiteModel = polSourceSiteModelList.Where(c => c.PolSourceSiteTVItemID == tvItemModel.TVItemID).FirstOrDefault();

                    if (polSourceSiteModel == null)
                    {
                        continue;
                    }

                    sb.AppendLine($@"			    <Placemark>");
                    sb.AppendLine($@"			    <name>P{ polSourceSiteModel.Site }</name>");
                    sb.AppendLine($@"	            <visibility>0</visibility>");
                    sb.AppendLine($@"               <description><![CDATA[");

                    if (polSourceObservationModelList.Count > 0)
                    {
                        PolSourceObservationModel polSourceObservationModel = polSourceObservationModelList
                            .Where(c => c.PolSourceSiteID == polSourceSiteModel.PolSourceSiteID)
                            .OrderByDescending(c => c.ObservationDate_Local).FirstOrDefault();

                        if (polSourceObservationModel != null)
                        {

                            sb.AppendLine($@"                <h3>Last Observation</h3>");
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

                        ShowPictures(sb, tvItemModel.TVItemID, false);

                    }
                    sb.AppendLine($@"                   ]]></description>");
                    sb.AppendLine($@"			    	<styleUrl>#s_ylw-pushpin</styleUrl>");
                    sb.AppendLine($@"			    	<Point>");
                    sb.AppendLine($@"		    			<coordinates>");
                    using (CSSPDBEntities db = new CSSPDBEntities())
                    {
                        List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModel.TVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);

                        foreach (MapInfoPointModel mapInfoPointModel in mapInfoPointModelList.Take(1))
                        {
                            sb.AppendLine($@"{mapInfoPointModel.Lng.ToString("F6")},{mapInfoPointModel.Lat.ToString("F6")},0 ");
                        }
                    }

                    sb.AppendLine($@"						</coordinates>");
                    sb.AppendLine($@"		    		</Point>");
                    sb.AppendLine($@"		    	</Placemark>");
                }
                sb.AppendLine($@"		        </Folder>");

                // ---------------------------------------------------------------
                // doing Inactive Pollution Source Sites
                // ---------------------------------------------------------------
                sb.AppendLine($@"		        <Folder>");
                sb.AppendLine($@"			        <name>Inactive</name> ");
                sb.AppendLine($@"	                <visibility>0</visibility>");
                foreach (TVItemModel tvItemModel in tvitemModelPSSList.Where(c => c.IsActive == false))
                {
                    PolSourceSiteModel polSourceSiteModel = polSourceSiteModelList.Where(c => c.PolSourceSiteTVItemID == tvItemModel.TVItemID).FirstOrDefault();

                    if (polSourceSiteModel == null)
                    {
                        continue;
                    }

                    sb.AppendLine($@"			    <Placemark>");
                    sb.AppendLine($@"			    <name>P{ polSourceSiteModel.Site }</name>");
                    sb.AppendLine($@"	            <visibility>0</visibility>");
                    sb.AppendLine($@"               <description><![CDATA[");

                    if (polSourceObservationModelList.Count > 0)
                    {
                        PolSourceObservationModel polSourceObservationModel = polSourceObservationModelList
                            .Where(c => c.PolSourceSiteID == polSourceSiteModel.PolSourceSiteID)
                            .OrderByDescending(c => c.ObservationDate_Local).FirstOrDefault();

                        if (polSourceObservationModel != null)
                        {

                            sb.AppendLine($@"                <h3>Last Observation</h3>");
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

                        ShowPictures(sb, tvItemModel.TVItemID, false);

                    }
                    sb.AppendLine($@"                   ]]></description>");
                    sb.AppendLine($@"			    	<styleUrl>#s_ylwI-pushpin</styleUrl>");
                    sb.AppendLine($@"			    	<Point>");
                    sb.AppendLine($@"		    			<coordinates>");
                    using (CSSPDBEntities db = new CSSPDBEntities())
                    {
                        List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModel.TVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);

                        foreach (MapInfoPointModel mapInfoPointModel in mapInfoPointModelList.Take(1))
                        {
                            sb.AppendLine($@"{mapInfoPointModel.Lng.ToString("F6")},{mapInfoPointModel.Lat.ToString("F6")},0 ");
                        }
                    }

                    sb.AppendLine($@"						</coordinates>");
                    sb.AppendLine($@"		    		</Point>");
                    sb.AppendLine($@"		    	</Placemark>");
                }
                sb.AppendLine($@"		        </Folder>");

                sb.AppendLine($@"		    </Folder>");

                // ---------------------------------------------------------------
                // doing MWQM Sites
                // ---------------------------------------------------------------
                sb.AppendLine($@"		    <Folder>");
                sb.AppendLine($@"			    <name>MWQM Sites</name> ");
                sb.AppendLine($@"	            <visibility>0</visibility>");

                List<TVItemModel> tvitemModelMWQMSiteList = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSS.TVItemID, TVTypeEnum.MWQMSite).ToList();

                // ---------------------------------------------------------------
                // doing Active MWQM Sites
                // ---------------------------------------------------------------
                sb.AppendLine($@"		        <Folder>");
                sb.AppendLine($@"			        <name>Active</name> ");
                sb.AppendLine($@"	                <visibility>0</visibility>");
                foreach (TVItemModel tvItemModel in tvitemModelMWQMSiteList.Where(c => c.IsActive == true))
                {
                    TVLocation tvlNew = new TVLocation();
                    tvlNew.TVItemID = tvItemModel.TVItemID;
                    tvlNew.TVText = tvItemModel.TVText;
                    tvlNew.TVType = TVTypeEnum.PolSourceSite;
                    tvlNew.SubTVType = TVTypeEnum.MWQMSite;

                    _MapInfoService.GetMWQMSiteMapInfoStatDB(tvItemModel.TVItemID, tvlNew, 30);

                    sb.AppendLine($@"			    <Placemark>");
                    sb.AppendLine($@"			    <name>{ tvlNew.TVText.Substring(0, 8).Replace(" ", "") }</name>");
                    sb.AppendLine($@"	            <visibility>0</visibility>");
                    sb.AppendLine($@"               <description><![CDATA[");

                    ShowPictures(sb, tvItemModel.TVItemID, true);

                    sb.AppendLine($@"                   ]]></description>");
                    if (tvlNew.SubTVType == TVTypeEnum.NoDepuration)
                    {
                        sb.AppendLine($@"			    	<styleUrl>#msn_placemark_square_purple</styleUrl>");
                    }
                    else if (tvlNew.SubTVType == TVTypeEnum.Failed)
                    {
                        sb.AppendLine($@"			    	<styleUrl>#msn_placemark_square_red</styleUrl>");
                    }
                    else // passed
                    {
                        sb.AppendLine($@"			    	<styleUrl>#msn_placemark_square_green</styleUrl>");
                    }
                    sb.AppendLine($@"			    	<Point>");
                    sb.AppendLine($@"		    			<coordinates>");
                    using (CSSPDBEntities db = new CSSPDBEntities())
                    {
                        List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModel.TVItemID, TVTypeEnum.MWQMSite, MapInfoDrawTypeEnum.Point);

                        foreach (MapInfoPointModel mapInfoPointModel in mapInfoPointModelList.Take(1))
                        {
                            sb.AppendLine($@"{mapInfoPointModel.Lng.ToString("F6")},{mapInfoPointModel.Lat.ToString("F6")},0 ");
                        }
                    }

                    sb.AppendLine($@"						</coordinates>");
                    sb.AppendLine($@"		    		</Point>");
                    sb.AppendLine($@"		    	</Placemark>");
                }
                sb.AppendLine($@"		        </Folder>");

                // ---------------------------------------------------------------
                // doing Inactive MWQM Sites
                // ---------------------------------------------------------------
                sb.AppendLine($@"		        <Folder>");
                sb.AppendLine($@"			        <name>Inactive</name> ");
                sb.AppendLine($@"	                <visibility>0</visibility>");
                foreach (TVItemModel tvItemModel in tvitemModelMWQMSiteList.Where(c => c.IsActive == false))
                {
                    TVLocation tvlNew = new TVLocation();
                    tvlNew.TVItemID = tvItemModel.TVItemID;
                    tvlNew.TVText = tvItemModel.TVText;
                    tvlNew.TVType = TVTypeEnum.PolSourceSite;
                    tvlNew.SubTVType = TVTypeEnum.MWQMSite;

                    _MapInfoService.GetMWQMSiteMapInfoStatDB(tvItemModel.TVItemID, tvlNew, 30);

                    sb.AppendLine($@"			    <Placemark>");
                    sb.AppendLine($@"			    <name>{ tvlNew.TVText.Substring(0, 8).Replace(" ", "") }</name>");
                    sb.AppendLine($@"	            <visibility>0</visibility>");
                    sb.AppendLine($@"               <description><![CDATA[");

                    ShowPictures(sb, tvItemModel.TVItemID, true);

                    sb.AppendLine($@"                   ]]></description>");
                    sb.AppendLine($@"			    	<styleUrl>#msn_placemark_square_white</styleUrl>");
                    sb.AppendLine($@"			    	<Point>");
                    sb.AppendLine($@"		    			<coordinates>");
                    using (CSSPDBEntities db = new CSSPDBEntities())
                    {
                        List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(tvItemModel.TVItemID, TVTypeEnum.MWQMSite, MapInfoDrawTypeEnum.Point);

                        foreach (MapInfoPointModel mapInfoPointModel in mapInfoPointModelList.Take(1))
                        {
                            sb.AppendLine($@"{mapInfoPointModel.Lng.ToString("F6")},{mapInfoPointModel.Lat.ToString("F6")},0 ");
                        }
                    }

                    sb.AppendLine($@"						</coordinates>");
                    sb.AppendLine($@"		    		</Point>");
                    sb.AppendLine($@"		    	</Placemark>");
                }
                sb.AppendLine($@"		        </Folder>");

                sb.AppendLine($@"		    </Folder>");

                // ---------------------------------------------------------------
                // doing Municipality
                // ---------------------------------------------------------------
                sb.AppendLine($@"		    <Folder>");
                sb.AppendLine($@"			    <name>Municipalities</name> ");
                sb.AppendLine($@"	            <visibility>0</visibility>");

                List<UseOfSiteModel> useOfSiteModelList = _UseOfSiteService.GetUseOfSiteModelListWithSubsectorTVItemIDDB(tvItemModelSS.TVItemID);

                List<int> MunicipalityTVItemIDList = new List<int>();
                List<TVItemModel> tvItemModelMuniList = new List<TVItemModel>();

                foreach (UseOfSiteModel useOfSiteModel in useOfSiteModelList)
                {
                    if (useOfSiteModel.TVType == TVTypeEnum.Municipality)
                    {
                        MunicipalityTVItemIDList.Add(useOfSiteModel.SiteTVItemID);
                        TVItemModel tvItemModelMuni = _TVItemService.GetTVItemModelWithTVItemIDDB(useOfSiteModel.SiteTVItemID);
                        tvItemModelMuniList.Add(tvItemModelMuni);
                    }
                }

                foreach (TVItemModel tvItemModelMuni in tvItemModelMuniList.OrderBy(c => c.TVText))
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

                                InfrastructureModel infrastructureModel = _InfrastructureService.GetInfrastructureModelWithInfrastructureTVItemIDDB(tvItemModel.TVItemID);

                                sb.AppendLine($@"<table>");
                                sb.AppendLine($@"<thead>");
                                sb.AppendLine($@"<tr><th>Property</th><th>Value</th><tr>");
                                sb.AppendLine($@"</thead>");
                                sb.AppendLine($@"<tbody>");
                                sb.AppendLine($@"<tr><td style=""text-align: right; width: 180px"">Infrastructure Type&nbsp;&nbsp;&nbsp;</td><td style=""text-align: left"">{ _BaseEnumService.GetEnumText_InfrastructureTypeEnum(infrastructureModel.InfrastructureType)}</td><tr>");
                                if (infrastructureModel.InfrastructureType == InfrastructureTypeEnum.WWTP)
                                {
                                    sb.AppendLine($@"<tr><td style=""text-align: right; width: 180px"">Facility Type&nbsp;&nbsp;&nbsp;</td><td style=""text-align: left; min-width: 300px"">{ _BaseEnumService.GetEnumText_FacilityTypeEnum(infrastructureModel.FacilityType)}</td><tr>");

                                    if (infrastructureModel.FacilityType == FacilityTypeEnum.Plant)
                                    {
                                        sb.AppendLine($@"<tr><td style=""text-align: right; width: 180px"">Preliminary Treatment Type&nbsp;&nbsp;&nbsp;</td><td style=""text-align: left; min-width: 300px"">{ _BaseEnumService.GetEnumText_PreliminaryTreatmentTypeEnum(infrastructureModel.PreliminaryTreatmentType)}</td><tr>");
                                        sb.AppendLine($@"<tr><td style=""text-align: right; width: 180px"">Primary Treatment Type&nbsp;&nbsp;&nbsp;</td><td style=""text-align: left; min-width: 300px"">{ _BaseEnumService.GetEnumText_PrimaryTreatmentTypeEnum(infrastructureModel.PrimaryTreatmentType)}</td><tr>");
                                        sb.AppendLine($@"<tr><td style=""text-align: right; width: 180px"">Secondary Treatment Type&nbsp;&nbsp;&nbsp;</td><td style=""text-align: left; min-width: 300px"">{ _BaseEnumService.GetEnumText_SecondaryTreatmentTypeEnum(infrastructureModel.SecondaryTreatmentType)}</td><tr>");
                                        sb.AppendLine($@"<tr><td style=""text-align: right; width: 180px"">Tertiary Treatment Type&nbsp;&nbsp;&nbsp;</td><td style=""text-align: left; min-width: 300px"">{ _BaseEnumService.GetEnumText_TertiaryTreatmentTypeEnum(infrastructureModel.TertiaryTreatmentType)}</td><tr>");
                                    }

                                    sb.AppendLine($@"<tr><td style=""text-align: right; width: 180px"">Is Mechanically Aerated&nbsp;&nbsp;&nbsp;</td><td style=""text-align: left; min-width: 300px"">{ infrastructureModel.IsMechanicallyAerated.ToString()}</td><tr>");
                                    sb.AppendLine($@"<tr><td style=""text-align: right; width: 180px"">Number Of Cells&nbsp;&nbsp;&nbsp;</td><td style=""text-align: left; min-width: 300px"">{ infrastructureModel.NumberOfCells.ToString()}</td><tr>");
                                    sb.AppendLine($@"<tr><td style=""text-align: right; width: 180px"">Number Of Aerated Cells&nbsp;&nbsp;&nbsp;</td><td style=""text-align: left; min-width: 300px"">{ infrastructureModel.NumberOfAeratedCells.ToString()}</td><tr>");
                                    sb.AppendLine($@"<tr><td style=""text-align: right; width: 180px"">Aeration Type&nbsp;&nbsp;&nbsp;</td><td style=""text-align: left; min-width: 300px"">{ _BaseEnumService.GetEnumText_AerationTypeEnum(infrastructureModel.AerationType)}</td><tr>");
                                    sb.AppendLine($@"<tr><td style=""text-align: right; width: 180px"">Disinfection Type&nbsp;&nbsp;&nbsp;</td><td style=""text-align: left; min-width: 300px"">{ _BaseEnumService.GetEnumText_DisinfectionTypeEnum(infrastructureModel.DisinfectionType)}</td><tr>");
                                    sb.AppendLine($@"<tr><td style=""text-align: right; width: 180px"">Collection System Type&nbsp;&nbsp;&nbsp;</td><td style=""text-align: left; min-width: 300px"">{ _BaseEnumService.GetEnumText_CollectionSystemTypeEnum(infrastructureModel.CollectionSystemType)}</td><tr>");
                                    sb.AppendLine($@"<tr><td style=""text-align: right; width: 180px"">Design Flow (m3/day)&nbsp;&nbsp;&nbsp;</td><td style=""text-align: left; min-width: 300px"">{ infrastructureModel.DesignFlow_m3_day.ToString()}</td><tr>");
                                    sb.AppendLine($@"<tr><td style=""text-align: right; width: 180px"">Average Flow (m3/day)&nbsp;&nbsp;&nbsp;</td><td style=""text-align: left; min-width: 300px"">{ infrastructureModel.AverageFlow_m3_day.ToString()}</td><tr>");
                                    sb.AppendLine($@"<tr><td style=""text-align: right; width: 180px"">Peak Flow (m3/day)&nbsp;&nbsp;&nbsp;</td><td style=""text-align: left; min-width: 300px"">{ infrastructureModel.PeakFlow_m3_day.ToString()}</td><tr>");
                                    sb.AppendLine($@"<tr><td style=""text-align: right; width: 180px"">Percent Flow Of Total&nbsp;&nbsp;&nbsp;</td><td style=""text-align: left; min-width: 300px"">{ infrastructureModel.PercFlowOfTotal.ToString()}</td><tr>");
                                    sb.AppendLine($@"<tr><td style=""text-align: right; width: 180px"">Population Served&nbsp;&nbsp;&nbsp;</td><td style=""text-align: left; min-width: 300px"">{ infrastructureModel.PopServed.ToString()}</td><tr>");
                                }
                                sb.AppendLine($@"<tr><td style=""text-align: right; width: 180px"">Alarm System Type&nbsp;&nbsp;&nbsp;</td><td style=""text-align: left; min-width: 300px"">{ _BaseEnumService.GetEnumText_AlarmSystemTypeEnum(infrastructureModel.AlarmSystemType)}</td><tr>");
                                sb.AppendLine($@"<tr><td style=""text-align: right; width: 180px"">Can Overflow&nbsp;&nbsp;&nbsp;</td><td style=""text-align: left; min-width: 300px"">{ infrastructureModel.CanOverflow.ToString()}</td><tr>");
                                sb.AppendLine($@"<tr><td style=""text-align: right; width: 180px"">Comment&nbsp;&nbsp;&nbsp;</td><td style=""text-align: left; min-width: 300px"">{ infrastructureModel.Comment.Replace("\r\n", "<br />")}</td><tr>");
                                sb.AppendLine($@"</tbody>");
                                sb.AppendLine($@"</table>");

                                ShowPictures(sb, tvItemModel.TVItemID, false);

                                sb.AppendLine($@"                           ]]></description>");
                                sb.AppendLine($@"			                	<styleUrl>#sn_shaded_dot</styleUrl>");
                                sb.AppendLine($@"			                	<Point>");
                                sb.AppendLine($@"		                			<coordinates>");
                                using (CSSPDBEntities db = new CSSPDBEntities())
                                {
                                    foreach (MapInfoPointModel mapInfoPointModel in mapInfoPointModelList.Take(1))
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

            try
            {
                File.WriteAllText(fi.FullName, sb.ToString(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateFile_Error_, fi.FullName, ex.Message);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateFile_Error_", fi.FullName, ex.Message);
                return;
            }

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

        private void ShowPictures(StringBuilder sb, int TVItemID, bool IsMWQMSite)
        {
            sb.AppendLine($@"                <span data-tvitemid=""{TVItemID}"">&nbsp;</span>");

            List<TVFileModel> tvFileModelList = _TVFileService.GetTVFileModelListWithParentTVItemIDDB(TVItemID);

            if (tvFileModelList.Where(c => c.FileType == FileTypeEnum.JPEG || c.FileType == FileTypeEnum.JPG).Any())
            {
                sb.AppendLine(@"<script src=""http://ajax.googleapis.com/ajax/libs/jquery/1.8.3/jquery.min.js"" type=""text/javascript""></script>");

                if (IsMWQMSite)
                {
                    sb.AppendLine(@"<h3 style=""text-align:center""><a href=""" + @"http://wmon01dtchlebl2/csspwebtools/en-CA/MWQM/_mwqmSiteSampleMovingAverageStat?MWQMSiteTVItemID=" + TVItemID + @""" target=""_blank"">Open Data Table</a></h3>");
                }

                foreach (TVFileModel tvFileModel in tvFileModelList)
                {
                    string src = @"http://wmon01dtchlebl2/csspwebtools/en-CA/File/ShowImage?TVFileTVItemID=" + tvFileModel.TVFileTVItemID;
                    sb.AppendLine($@"                <p><img src="""" data-src=""{src}"" width=""500"" height=""500""></p>");
                }

                sb.AppendLine(@" <script type=""text/javascript"">");
                sb.AppendLine(@"    $(document).ready(function () {");
                sb.AppendLine(@"        $('img').each(function() {");
                sb.AppendLine(@"            $(this).attr('src', $(this).data('src'));");
                sb.AppendLine(@"        });");
                sb.AppendLine(@"    }); ");
                sb.AppendLine(@"</script>");
            }
        }
        public string ParseKml(FileInfo fi)
        {

            return "Need to finalize ParseKML";
        }
        public void ProvinceToolsGenerateStats()
        {
            int NumberOfSamples = 30;
            StringBuilder sb = new StringBuilder();
            string NotUsed = "";
            List<CSVValues> csvValuesList = new List<CSVValues>();

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

            string FileName = $"StatsAndStatsWithRain_{ Init }.kml";

            FileInfo fi = new FileInfo(ServerPath + FileName);

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

            List<TVItemModel> tvItemModelSSList = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelProv.TVItemID, TVTypeEnum.Subsector);

            sb.AppendLine($@"<?xml version=""1.0"" encoding=""UTF-8""?>");
            sb.AppendLine($@"<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom"">");
            sb.AppendLine($@"<Document>");
            sb.AppendLine($@"	<name>StatsAndStatsWithRain_{ Init }.kml</name>");

            List<LetterColorName> LetterColorNameList = new List<LetterColorName>()
            {
               new LetterColorName() { Letter = "F", Color = "ff7777", Name = "NoDepuration" },
               new LetterColorName() { Letter = "E", Color = "ff7777", Name = "NoDepuration" },
               new LetterColorName() { Letter = "D", Color = "ff7777", Name = "NoDepuration" },
               new LetterColorName() { Letter = "C", Color = "ff7777", Name = "NoDepuration" },
               new LetterColorName() { Letter = "B", Color = "ff7777", Name = "NoDepuration" },
               new LetterColorName() { Letter = "A", Color = "ff7777", Name = "NoDepuration" },
               new LetterColorName() { Letter = "F", Color = "0000ff", Name = "Fail" },
               new LetterColorName() { Letter = "E", Color = "0000ff", Name = "Fail" },
               new LetterColorName() { Letter = "D", Color = "0000ff", Name = "Fail" },
               new LetterColorName() { Letter = "C", Color = "0000ff", Name = "Fail" },
               new LetterColorName() { Letter = "B", Color = "0000ff", Name = "Fail" },
               new LetterColorName() { Letter = "A", Color = "0000ff", Name = "Fail" },
               new LetterColorName() { Letter = "F", Color = "00ff00", Name = "Pass" },
               new LetterColorName() { Letter = "E", Color = "00ff00", Name = "Pass" },
               new LetterColorName() { Letter = "D", Color = "00ff00", Name = "Pass" },
               new LetterColorName() { Letter = "C", Color = "00ff00", Name = "Pass" },
               new LetterColorName() { Letter = "B", Color = "00ff00", Name = "Pass" },
               new LetterColorName() { Letter = "A", Color = "00ff00", Name = "Pass" },
            };

            foreach (LetterColorName letterColorName in LetterColorNameList)
            {
                sb.AppendLine($@"	<Style id=""{ letterColorName.Name }_{ letterColorName.Letter }"">");
                sb.AppendLine($@"		<IconStyle>");
                sb.AppendLine($@"			<color>ff{ letterColorName.Color }</color>");
                sb.AppendLine($@"			<scale>1.0</scale>");
                sb.AppendLine($@"			<Icon>");
                sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/paddle/" + $"{letterColorName.Letter}.png</href>");
                sb.AppendLine($@"			</Icon>");
                sb.AppendLine($@"			<hotSpot x=""32"" y=""1"" xunits=""pixels"" yunits=""pixels""/>");
                sb.AppendLine($@"		</IconStyle>");
                sb.AppendLine($@"   </Style>");
            }
            List<double> DryList = new List<double>() { 4, 8, 12, 16 };
            List<double> WetList = new List<double>() { 12, 25, 37, 50 };

            #region Styles
            sb.AppendLine($@"	<Style id=""SS_Point"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ffffff</color>");
            sb.AppendLine($@"			<scale>1.0</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"			<hotSpot x=""32"" y=""1"" xunits=""pixels"" yunits=""pixels""/>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"   </Style>");

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
            sb.AppendLine($@"			<color>0000ff00</color>");
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
            sb.AppendLine($@"			<color>ff00ff00</color>");
            sb.AppendLine($@"			<width>2</width>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@"		<PolyStyle>");
            sb.AppendLine($@"			<color>0000ff00</color>");
            sb.AppendLine($@"		</PolyStyle>");
            sb.AppendLine($@" </Style>");
            #endregion Styles

            #region Subsectors polygon
            // -------------------------------------------------------------------------------------------------------------
            // Start Subsectors polygon
            // -------------------------------------------------------------------------------------------------------------
            sb.AppendLine($@"	<Folder>");
            sb.AppendLine($@"	<name>Subsectors polygon</name>");

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 5);
            Application.DoEvents();
            foreach (TVItemModel tvItemModelSS in tvItemModelSSList)
            {
                //if (tvItemModelSS.TVItemID != 635)
                //{
                //    continue; // just doing Bouctouche for now
                //}

                List<MapInfo> mapInfoList = new List<MapInfo>();
                List<MapInfoPoint> mapInfoPointList = new List<MapInfoPoint>();

                using (CSSPDBEntities db2 = new CSSPDBEntities())
                {
                    mapInfoList = (from c in db2.MapInfos
                                   where c.TVItemID == tvItemModelSS.TVItemID
                                   select c).ToList();

                    List<int> mapInfoIDList = mapInfoList.Select(c => c.MapInfoID).Distinct().ToList();

                    mapInfoPointList = (from c in db2.MapInfoPoints
                                        from mid in mapInfoIDList
                                        where c.MapInfoID == mid
                                        select c).ToList();
                }

                sb.AppendLine($@"	    <Folder>");
                sb.AppendLine($@"	    <name>{ tvItemModelSS.TVText }</name>");

                sb.AppendLine($@"	<Placemark>");
                sb.AppendLine($@"		<name>{ tvItemModelSS.TVText }</name>");
                sb.AppendLine($@"		<styleUrl>#m_ylw-pushpin</styleUrl>");
                sb.AppendLine($@"		<Polygon>");
                sb.AppendLine($@"			<tessellate>1</tessellate>");
                sb.AppendLine($@"			<outerBoundaryIs>");
                sb.AppendLine($@"				<LinearRing>");
                sb.AppendLine($@"					<coordinates>");
                sb.Append($@"						");
                List<MapInfoPoint> mapInfoPointListPoly = (from mi in mapInfoList
                                                           from mip in mapInfoPointList
                                                           where mi.MapInfoID == mip.MapInfoID
                                                           && mi.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Polygon
                                                           orderby mip.Ordinal
                                                           select mip).ToList();


                foreach (MapInfoPoint mapInfoPoint in mapInfoPointListPoly)
                {
                    sb.Append($@"{ mapInfoPoint.Lng },{ mapInfoPoint.Lat },0 ");
                }
                sb.AppendLine($@"");
                sb.AppendLine($@"					</coordinates>");
                sb.AppendLine($@"				</LinearRing>");
                sb.AppendLine($@"			</outerBoundaryIs>");
                sb.AppendLine($@"		</Polygon>");
                sb.AppendLine($@"   </Placemark>");


                List<MapInfoPoint> mapInfoPointListPoint = (from mi in mapInfoList
                                                            from mip in mapInfoPointList
                                                            where mi.MapInfoID == mip.MapInfoID
                                                            && mi.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Point
                                                            orderby mip.Ordinal
                                                            select mip).ToList();

                if (mapInfoPointListPoint.Count > 0)
                {
                    sb.AppendLine($@"           <Placemark>");
                    sb.AppendLine($@"	        	<name>{ tvItemModelSS.TVText }</name>");
                    sb.AppendLine($@"	        	<styleUrl>#m_ylw-pushpin</styleUrl>");
                    sb.AppendLine($@"	        	<Point>");
                    sb.AppendLine($@"	        		<coordinates>{ mapInfoPointListPoint[0].Lng },{ mapInfoPointListPoint[0].Lat },0</coordinates>");
                    sb.AppendLine($@"	        	</Point>");
                    sb.AppendLine($@"	        </Placemark>");

                }
                sb.AppendLine($@"	    </Folder>");
            }

            sb.AppendLine($@"	</Folder>");
            // -------------------------------------------------------------------------------------------------------------
            // End Subsectors polygon
            // -------------------------------------------------------------------------------------------------------------
            #endregion Subsectors polygon

            #region All-All-All 30 runs
            // -------------------------------------------------------------------------------------------------------------
            // Start All-All-All 30 samples
            // -------------------------------------------------------------------------------------------------------------
            NumberOfSamples = 30;
            sb.AppendLine($@"	<Folder>");
            sb.AppendLine($@"	<name>All-All-All ({ NumberOfSamples })</name>");

            DoAllSubsectorStats(StatType.Run30, tvItemModelSSList, _TVItemService, NumberOfSamples, DryList, WetList, csvValuesList, sb);

            sb.AppendLine($@"	</Folder>");

            // -------------------------------------------------------------------------------------------------------------
            // End All-All-All 30 samples
            // -------------------------------------------------------------------------------------------------------------
            #endregion All-All-All 30 runs

            #region Dry-All-All (4,8,12,16)mm
            // -------------------------------------------------------------------------------------------------------------
            // Start Dry-All-All (4,8,12,16)mm
            // -------------------------------------------------------------------------------------------------------------

            sb.AppendLine($@"	<Folder>");
            sb.AppendLine($@"	<name>Dry-All-All (4,8,12,16)mm</name>");

            DoAllSubsectorStats(StatType.Dry, tvItemModelSSList, _TVItemService, NumberOfSamples, DryList, WetList, csvValuesList, sb);

            sb.AppendLine($@"	</Folder>");

            // -------------------------------------------------------------------------------------------------------------
            // End Dry-All-All (4,8,12,16)mm
            // -------------------------------------------------------------------------------------------------------------
            #endregion Dry-All-All (4,8,12,16)mm

            #region Wet-All-All (12,25,37,50)mm
            // -------------------------------------------------------------------------------------------------------------
            // Start Wet-All-All (12,25,37,50)mm
            // -------------------------------------------------------------------------------------------------------------


            sb.AppendLine($@"	<Folder>");
            sb.AppendLine($@"	<name>Wet-All-All (12,25,37,50)mm</name>");

            DoAllSubsectorStats(StatType.Wet, tvItemModelSSList, _TVItemService, NumberOfSamples, DryList, WetList, csvValuesList, sb);

            sb.AppendLine($@"	</Folder>");

            // -------------------------------------------------------------------------------------------------------------
            // End Wet-All-All (12,25,37,50)mm
            // -------------------------------------------------------------------------------------------------------------
            #endregion Wet-All-All (12,25,37,50)mm

            sb.AppendLine($@"</Document>");
            sb.AppendLine($@"</kml>");

            StreamWriter sw = fi.CreateText();
            sw.Write(sb.ToString());
            sw.Close();

            fi = new FileInfo(ServerPath + FileName);

            TVFileModel tvFileModelNew = new TVFileModel();
            tvFileModelNew.FileCreatedDate_UTC = fi.CreationTimeUtc;
            tvFileModelNew.FileDescription = "Statistics KML file";
            tvFileModelNew.FileInfo = "Statistics KML file";
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


            TVFileModel tvFileModelExist = _TVFileService.GetTVFileModelWithServerFilePathAndServerFileNameDB(ServerPath, FileName);
            if (!string.IsNullOrWhiteSpace(tvFileModelExist.Error))
            {
                TVFileModel tvFileModelRet = _TVFileService.PostAddTVFileDB(tvFileModelNew);
                if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                    return;
                }
            }
            else
            {
                tvFileModelNew.TVFileID = tvFileModelExist.TVFileID;
                TVFileModel tvFileModelRet = _TVFileService.PostUpdateTVFileDB(tvFileModelNew);
                if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                    return;
                }
            }

            WriteCSVFile(Init, csvValuesList, fi, FileName.Replace(".kml", ".csv"), ServerPath, tvItemModelProv);

        }
        #endregion Functions public

        #region Functions private
        private void WriteCSVFile(string Init, List<CSVValues> csvValuesList, FileInfo fi, string FileName, string ServerPath, TVItemModel tvItemModelProv)
        {
            string NotUsed = "";
            StringBuilder sbCSV = new StringBuilder();

            sbCSV.AppendLine("Subsector,Site,StartYear,EndYear,StatType,Class,Letter,NumbSamples,P90,GM,Med,PercOver43,PercOver260,ValueList Dry(4_8_12_16) Wet(12_25_37_50)");

            foreach (CSVValues csvValues in csvValuesList.OrderBy(c => c.Subsector).ThenBy(c => c.Site))
            {
                string P90Str = (csvValues.P90 < 0 ? "" : csvValues.P90.ToString().Replace(",", "."));
                string GMStr = (csvValues.GM < 0 ? "" : csvValues.GM.ToString().Replace(",", "."));
                string MedStr = (csvValues.Med < 0 ? "" : csvValues.Med.ToString().Replace(",", "."));
                string PercOver43Str = (csvValues.PercOver43 < 0 ? "" : csvValues.PercOver43.ToString().Replace(",", "."));
                string PercOver260Str = (csvValues.PercOver260 < 0 ? "" : csvValues.PercOver260.ToString().Replace(",", "."));

                sbCSV.AppendLine($"{ csvValues.Subsector.Replace(",", "_") },{ csvValues.Site.Replace(",", "_") },{ csvValues.StartYear },{ csvValues.EndYear },{ csvValues.statType.ToString() },{ csvValues.ClassStr },{ csvValues.Letter },{ csvValues.NumbSamples },{ P90Str },{ GMStr },{ MedStr },{ PercOver43Str },{ PercOver260Str },{ csvValues.ValueList }");
            }

            fi = new FileInfo(fi.FullName.Replace(".kml", ".csv"));

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

            StreamWriter sw = fi.CreateText();
            sw.Write(sbCSV.ToString());
            sw.Close();

            fi = new FileInfo(fi.FullName);

            TVFileModel tvFileModelNew = new TVFileModel();
            tvFileModelNew.FileCreatedDate_UTC = fi.CreationTimeUtc;
            tvFileModelNew.FileDescription = "Statistics CSV file";
            tvFileModelNew.FileInfo = "Statistics CSV file";
            tvFileModelNew.FilePurpose = FilePurposeEnum.Information;
            tvFileModelNew.FileSize_kb = (int)(fi.Length / 1024);
            tvFileModelNew.FileType = FileTypeEnum.CSV;
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

            TVFileModel tvFileModelExist = _TVFileService.GetTVFileModelWithServerFilePathAndServerFileNameDB(ServerPath, FileName);
            if (!string.IsNullOrWhiteSpace(tvFileModelExist.Error))
            {
                TVFileModel tvFileModelRet = _TVFileService.PostAddTVFileDB(tvFileModelNew);
                if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                    return;
                }
            }
            else
            {
                tvFileModelNew.TVFileID = tvFileModelExist.TVFileID;
                TVFileModel tvFileModelRet = _TVFileService.PostUpdateTVFileDB(tvFileModelNew);
                if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                    return;
                }
            }
        }
        public class LetterColorName
        {
            public string Letter { get; set; }
            public string Color { get; set; }
            public string Name { get; set; }
        }
        private enum StatType
        {
            Run30 = 1,
            Dry = 2,
            Wet = 3,
        }
        private class CSVValues
        {
            public string Subsector { get; set; }
            public string Site { get; set; }
            public int StartYear { get; set; }
            public int EndYear { get; set; }
            public StatType statType { get; set; }
            public string ClassStr { get; set; }
            public string Letter { get; set; }
            public int NumbSamples { get; set; }
            public int P90 { get; set; }
            public int GM { get; set; }
            public int Med { get; set; }
            public int PercOver43 { get; set; }
            public int PercOver260 { get; set; }
            public string ValueList { get; set; }
        }
        private class RainDays
        {
            public DateTime RunDate { get; set; }
            public double R1 { get; set; }
            public double R2 { get; set; }
            public double R3 { get; set; }
            public double R4 { get; set; }
        }
        private void DoAllSubsectorStats(StatType statType, List<TVItemModel> tvItemModelSSList, TVItemService tvItemService, int NumberOfSamples, List<double> DryList, List<double> WetList, List<CSVValues> csvValuesList, StringBuilder sb)
        {

            int TotalCount = tvItemModelSSList.Count;
            int Count = 0;
            if (statType == StatType.Run30)
            {
                TotalCount = 3 * TotalCount;
                Count = 0;
            }
            else if (statType == StatType.Dry)
            {
                Count = TotalCount;
                TotalCount = 3 * TotalCount;
            }
            else
            {
                Count = TotalCount * 2;
                TotalCount = 3 * TotalCount;
            }
            foreach (TVItemModel tvItemModelSS in tvItemModelSSList)
            {
                Count += 1;

                if (Count % 5 == 0)
                {
                    int Perc = (int)(100.0f * ((float)Count / (float)TotalCount));
                    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Perc);
                }
                Application.DoEvents();

                string TVText = tvItemModelSS.TVText;
                if (TVText.Contains(" "))
                {
                    TVText = TVText.Substring(0, TVText.IndexOf(" "));
                }

                //if (tvItemModelSS.TVItemID != 635)
                //{
                //    continue; // just doing Bouctouche for now
                //}

                List<TVItemModel> tvItemModelMWQMSiteList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSS.TVItemID, TVTypeEnum.MWQMSite).Where(c => c.IsActive == true).ToList();
                List<MapInfo> mapInfoList = new List<MapInfo>();
                List<MapInfoPoint> mapInfoPointList = new List<MapInfoPoint>();
                List<MWQMSample> mwqmSampleListAll = new List<MWQMSample>();
                List<MWQMSample> mwqmSampleListStat = new List<MWQMSample>();
                List<TVItemModel> tvItemModelMWQMRunList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSS.TVItemID, TVTypeEnum.MWQMRun).Where(c => c.IsActive == true).ToList();
                List<RainDays> RainList = new List<RainDays>();

                using (CSSPDBEntities db2 = new CSSPDBEntities())
                {
                    List<int> TVItemMWQMSiteList = tvItemModelMWQMSiteList.Select(c => c.TVItemID).Distinct().ToList();
                    List<int> TVItemMWQMRunList = tvItemModelMWQMRunList.Select(c => c.TVItemID).Distinct().ToList();

                    List<MWQMRun> mwqmRunList = (from c in db2.MWQMRuns
                                                 from rid in TVItemMWQMRunList
                                                 where c.MWQMRunTVItemID == rid
                                                 && c.RunSampleType == (int)SampleTypeEnum.Routine
                                                 select c).ToList();

                    List<int> TVItemMWQMRunRoutineList = mwqmRunList.Select(c => c.MWQMRunTVItemID).Distinct().ToList();

                    mwqmSampleListAll = (from c in db2.MWQMSamples
                                         from tid in TVItemMWQMSiteList
                                         from rid in TVItemMWQMRunRoutineList
                                         where c.MWQMSiteTVItemID == tid
                                         && c.MWQMRunTVItemID == rid
                                         select c).ToList();

                    mapInfoList = (from c in db2.MapInfos
                                   from tid in TVItemMWQMSiteList
                                   where c.TVItemID == tid
                                   && c.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Point
                                   select c).ToList();

                    List<int> mapInfoIDList = mapInfoList.Select(c => c.MapInfoID).Distinct().ToList();

                    mapInfoPointList = (from c in db2.MapInfoPoints
                                        from mid in mapInfoIDList
                                        where c.MapInfoID == mid
                                        select c).ToList();

                    if (statType == StatType.Run30)
                    {
                        mwqmSampleListStat = (from c in mwqmSampleListAll
                                              from r in mwqmRunList
                                              where c.MWQMRunTVItemID == r.MWQMRunTVItemID
                                              select c).ToList();
                    }
                    else if (statType == StatType.Dry)
                    {
                        mwqmSampleListStat = (from c in mwqmSampleListAll
                                              from r in mwqmRunList
                                              let R1 = r.RainDay1_mm
                                              let R2 = r.RainDay1_mm + r.RainDay2_mm
                                              let R3 = r.RainDay1_mm + r.RainDay2_mm + r.RainDay3_mm
                                              let R4 = r.RainDay1_mm + r.RainDay2_mm + r.RainDay3_mm + r.RainDay4_mm
                                              where c.MWQMRunTVItemID == r.MWQMRunTVItemID
                                              && (R1 <= DryList[0]
                                              || R2 <= DryList[1]
                                              || R3 <= DryList[2]
                                              || R4 <= DryList[3])
                                              select c).ToList();
                    }
                    else if (statType == StatType.Wet)
                    {
                        var mwqmSampleListStatAndRain = (from c in mwqmSampleListAll
                                                         from r in mwqmRunList
                                                         let R1 = r.RainDay1_mm
                                                         let R2 = r.RainDay1_mm + r.RainDay2_mm
                                                         let R3 = r.RainDay1_mm + r.RainDay2_mm + r.RainDay3_mm
                                                         let R4 = r.RainDay1_mm + r.RainDay2_mm + r.RainDay3_mm + r.RainDay4_mm
                                                         where c.MWQMRunTVItemID == r.MWQMRunTVItemID
                                                         && (R1 >= WetList[0]
                                                         || R2 >= WetList[1]
                                                         || R3 >= WetList[2]
                                                         || R4 >= WetList[3])
                                                         select new { c, R1, R2, R3, R4 }).ToList();

                        foreach (var sampleStatAndRain in mwqmSampleListStatAndRain)
                        {
                            mwqmSampleListStat.Add(sampleStatAndRain.c);
                            RainList.Add(new RainDays() { RunDate = sampleStatAndRain.c.SampleDateTime_Local, R1 = (double)sampleStatAndRain.R1, R2 = (double)sampleStatAndRain.R2, R3 = (double)sampleStatAndRain.R3, R4 = (double)sampleStatAndRain.R4 });
                        }
                    }
                    else
                    {
                        mwqmSampleListStat = new List<MWQMSample>();
                    }

                }

                sb.AppendLine($@"	    <Folder>");
                sb.AppendLine($@"	    <name>{ tvItemModelSS.TVText }</name>");

                foreach (TVItemModel tvItemModelMWQMSite in tvItemModelMWQMSiteList)
                {

                    if (tvItemModelMWQMSite != null)
                    {
                        List<double> mwqmSampleFCList = (from c in mwqmSampleListStat
                                                         where c.MWQMSiteTVItemID == tvItemModelMWQMSite.TVItemID
                                                         orderby c.SampleDateTime_Local descending
                                                         select (c.FecCol_MPN_100ml < 2 ? 1.9D : (double)c.FecCol_MPN_100ml)
                                                        ).Take(NumberOfSamples).ToList<double>();

                        List<MWQMSample> mwqmSampleList = (from c in mwqmSampleListStat
                                                           where c.MWQMSiteTVItemID == tvItemModelMWQMSite.TVItemID
                                                           orderby c.SampleDateTime_Local descending
                                                           select c).Take(NumberOfSamples).ToList<MWQMSample>();


                        if (mwqmSampleList.Count >= 3)
                        {

                            double P90 = tvItemService.GetP90(mwqmSampleFCList);
                            double GeoMean = tvItemService.GeometricMean(mwqmSampleFCList);
                            double Median = tvItemService.GetMedian(mwqmSampleFCList);
                            double PercOver43 = ((((double)mwqmSampleList.Where(c => c.FecCol_MPN_100ml > 43).Count()) / (double)mwqmSampleList.Count()) * 100.0D);
                            double PercOver260 = ((((double)mwqmSampleList.Where(c => c.FecCol_MPN_100ml > 260).Count()) / (double)mwqmSampleList.Count()) * 100.0D);
                            int MinYear = mwqmSampleList.Select(c => c.SampleDateTime_Local).Min().Year;
                            int MaxYear = mwqmSampleList.Select(c => c.SampleDateTime_Local).Max().Year;

                            int P90Int = (int)Math.Round((double)P90, 0);
                            int GeoMeanInt = (int)Math.Round((double)GeoMean, 0);
                            int MedianInt = (int)Math.Round((double)Median, 0);
                            int PercOver43Int = (int)Math.Round((double)PercOver43, 0);
                            int PercOver260Int = (int)Math.Round((double)PercOver260, 0);

                            LetterColorName letterColorName = new LetterColorName();
                            string ClassStr = "";
                            if ((GeoMeanInt > 88) || (MedianInt > 88) || (P90Int > 260) || (PercOver260Int > 10))
                            {
                                ClassStr = "NoDep";
                                if ((GeoMeanInt > 181) || (MedianInt > 181) || (P90Int > 460) || (PercOver260Int > 18))
                                {
                                    letterColorName = new LetterColorName() { Letter = "F", Color = "8888ff", Name = "NoDepuration" };
                                }
                                else if ((GeoMeanInt > 163) || (MedianInt > 163) || (P90Int > 420) || (PercOver260Int > 17))
                                {
                                    letterColorName = new LetterColorName() { Letter = "E", Color = "9999ff", Name = "NoDepuration" };
                                }
                                else if ((GeoMeanInt > 144) || (MedianInt > 144) || (P90Int > 380) || (PercOver260Int > 15))
                                {
                                    letterColorName = new LetterColorName() { Letter = "D", Color = "aaaaff", Name = "NoDepuration" };
                                }
                                else if ((GeoMeanInt > 125) || (MedianInt > 125) || (P90Int > 340) || (PercOver260Int > 13))
                                {
                                    letterColorName = new LetterColorName() { Letter = "C", Color = "bbbbff", Name = "NoDepuration" };
                                }
                                else if ((GeoMeanInt > 107) || (MedianInt > 107) || (P90Int > 300) || (PercOver260Int > 12))
                                {
                                    letterColorName = new LetterColorName() { Letter = "B", Color = "ccccff", Name = "NoDepuration" };
                                }
                                else
                                {
                                    letterColorName = new LetterColorName() { Letter = "A", Color = "ddddff", Name = "NoDepuration" };
                                }
                            }
                            else if ((GeoMeanInt > 14) || (MedianInt > 14) || (P90Int > 43) || (PercOver43Int > 10))
                            {
                                ClassStr = "Fail";
                                if ((GeoMeanInt > 76) || (MedianInt > 76) || (P90Int > 224) || (PercOver43Int > 27))
                                {
                                    letterColorName = new LetterColorName() { Letter = "F", Color = "aa0000", Name = "Fail" };
                                }
                                else if ((GeoMeanInt > 63) || (MedianInt > 63) || (P90Int > 188) || (PercOver43Int > 23))
                                {
                                    letterColorName = new LetterColorName() { Letter = "E", Color = "cc0000", Name = "Fail" };
                                }
                                else if ((GeoMeanInt > 51) || (MedianInt > 51) || (P90Int > 152) || (PercOver43Int > 20))
                                {
                                    letterColorName = new LetterColorName() { Letter = "D", Color = "ff1111", Name = "Fail" };
                                }
                                else if ((GeoMeanInt > 39) || (MedianInt > 39) || (P90Int > 115) || (PercOver43Int > 17))
                                {
                                    letterColorName = new LetterColorName() { Letter = "C", Color = "ff4444", Name = "Fail" };
                                }
                                else if ((GeoMeanInt > 26) || (MedianInt > 26) || (P90Int > 79) || (PercOver43Int > 13))
                                {
                                    letterColorName = new LetterColorName() { Letter = "B", Color = "ff9999", Name = "Fail" };
                                }
                                else
                                {
                                    letterColorName = new LetterColorName() { Letter = "A", Color = "ffcccc", Name = "Fail" };
                                }
                            }
                            else
                            {
                                ClassStr = "Pass";
                                if ((GeoMeanInt > 12) || (MedianInt > 12) || (P90Int > 36) || (PercOver43Int > 8))
                                {
                                    letterColorName = new LetterColorName() { Letter = "F", Color = "ccffcc", Name = "Pass" };
                                }
                                else if ((GeoMeanInt > 9) || (MedianInt > 9) || (P90Int > 29) || (PercOver43Int > 7))
                                {
                                    letterColorName = new LetterColorName() { Letter = "E", Color = "99ff99", Name = "Pass" };
                                }
                                else if ((GeoMeanInt > 7) || (MedianInt > 7) || (P90Int > 22) || (PercOver43Int > 5))
                                {
                                    letterColorName = new LetterColorName() { Letter = "D", Color = "44ff44", Name = "Pass" };
                                }
                                else if ((GeoMeanInt > 5) || (MedianInt > 5) || (P90Int > 14) || (PercOver43Int > 3))
                                {
                                    letterColorName = new LetterColorName() { Letter = "C", Color = "11ff11", Name = "Pass" };
                                }
                                else if ((GeoMeanInt > 2) || (MedianInt > 2) || (P90Int > 7) || (PercOver43Int > 2))
                                {
                                    letterColorName = new LetterColorName() { Letter = "B", Color = "00bb00", Name = "Pass" };
                                }
                                else
                                {
                                    letterColorName = new LetterColorName() { Letter = "A", Color = "009900", Name = "Pass" };
                                }
                            }

                            MapInfoPoint mapInfoPoint = (from mi in mapInfoList
                                                         from mip in mapInfoPointList
                                                         where mi.MapInfoID == mip.MapInfoID
                                                         && mi.TVItemID == tvItemModelMWQMSite.TVItemID
                                                         select mip).FirstOrDefault();

                            if (mapInfoPoint != null)
                            {
                                sb.AppendLine($@"           <Placemark>");
                                sb.AppendLine($@"	        	<name></name>");
                                sb.AppendLine($@"	        	<styleUrl>#{ letterColorName.Name }_{ letterColorName.Letter }</styleUrl>");
                                sb.AppendLine($@"	        	<Point>");
                                if (statType == StatType.Dry)
                                {
                                    sb.AppendLine($@"	        		<coordinates>{ mapInfoPoint.Lng - 0.001D },{ mapInfoPoint.Lat - 0.001D },0</coordinates>");
                                }
                                else if (statType == StatType.Wet)
                                {
                                    sb.AppendLine($@"	        		<coordinates>{ mapInfoPoint.Lng + 0.001D },{ mapInfoPoint.Lat + 0.001D },0</coordinates>");
                                }
                                else
                                {
                                    sb.AppendLine($@"	        		<coordinates>{ mapInfoPoint.Lng },{ mapInfoPoint.Lat },0</coordinates>");
                                }
                                sb.AppendLine($@"	        	</Point>");
                                sb.AppendLine($@"	        </Placemark>");
                            }

                            StringBuilder sbValueList = new StringBuilder();
                            foreach (MWQMSample mwqmSample in mwqmSampleList)
                            {
                                if (statType == StatType.Wet)
                                {
                                    StringBuilder sbRain = new StringBuilder();
                                    RainDays rainDays = (from c in RainList
                                                         where c.RunDate == mwqmSample.SampleDateTime_Local
                                                         select c).FirstOrDefault();

                                    if (rainDays != null)
                                    {
                                        sbRain.Append($"({ Math.Round(rainDays.R1, 0).ToString("F0") }_{ Math.Round(rainDays.R2, 0).ToString("F0") }_{ Math.Round(rainDays.R3, 0).ToString("F0") }_{ Math.Round(rainDays.R4, 0).ToString("F0") })");
                                        sbValueList.Append($"{ mwqmSample.FecCol_MPN_100ml }{ sbRain.ToString() }|");
                                    }
                                    else
                                    {
                                        sbValueList.Append($"{ mwqmSample.FecCol_MPN_100ml }|");
                                    }
                                }
                                else
                                {
                                    sbValueList.Append($"{ mwqmSample.FecCol_MPN_100ml }|");
                                }
                            }
                            CSVValues csvValues = new CSVValues()
                            {
                                Subsector = TVText,
                                Site = tvItemModelMWQMSite.TVText,
                                StartYear = mwqmSampleList[0].SampleDateTime_Local.Year,
                                EndYear = mwqmSampleList[mwqmSampleList.Count - 1].SampleDateTime_Local.Year,
                                statType = statType,
                                ClassStr = ClassStr,
                                Letter = letterColorName.Letter,
                                NumbSamples = mwqmSampleList.Count,
                                P90 = P90Int,
                                GM = GeoMeanInt,
                                Med = MedianInt,
                                PercOver43 = PercOver43Int,
                                PercOver260 = PercOver260Int,
                                ValueList = sbValueList.ToString(),
                            };

                            csvValuesList.Add(csvValues);
                        }
                        else
                        {
                            StringBuilder sbValueList = new StringBuilder();
                            foreach (MWQMSample mwqmSample in mwqmSampleList)
                            {
                                if (statType == StatType.Wet)
                                {
                                    StringBuilder sbRain = new StringBuilder();
                                    RainDays rainDays = (from c in RainList
                                                         where c.RunDate == mwqmSample.SampleDateTime_Local
                                                         select c).FirstOrDefault();

                                    if (rainDays != null)
                                    {
                                        sbRain.Append($"({ Math.Round(rainDays.R1, 0).ToString("F0") }_{ Math.Round(rainDays.R2, 0).ToString("F0") }_{ Math.Round(rainDays.R3, 0).ToString("F0") }_{ Math.Round(rainDays.R4, 0).ToString("F0") })");
                                        sbValueList.Append($"{ mwqmSample.FecCol_MPN_100ml }{ sbRain.ToString() }|");
                                    }
                                    else
                                    {
                                        sbValueList.Append($"{ mwqmSample.FecCol_MPN_100ml }|");
                                    }
                                }
                                else
                                {
                                    sbValueList.Append($"{ mwqmSample.FecCol_MPN_100ml }|");
                                }
                            }
                            CSVValues csvValues = new CSVValues()
                            {
                                Subsector = TVText,
                                Site = tvItemModelMWQMSite.TVText,
                                ClassStr = "",
                                Letter = "",
                                NumbSamples = mwqmSampleList.Count,
                                P90 = -1,
                                GM = -1,
                                Med = -1,
                                PercOver43 = -1,
                                PercOver260 = -1,
                                ValueList = sbValueList.ToString(),
                            };

                            csvValuesList.Add(csvValues);
                        }
                    }
                }

                sb.AppendLine($@"	    </Folder>");
            }
        }
        #endregion Functions private

    }

    public class SensitiveSite
    {
        public string SiteName { get; set; }
        public int Height_m { get; set; }
        public int Width_m { get; set; }
    }
}
