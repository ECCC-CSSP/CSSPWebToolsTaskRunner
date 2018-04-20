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

            FileInfo fi = new FileInfo(_TVFileService.ChoseEDriveOrCDrive(ServerFilePath) + $"MWQMSites_{ProvInit}.kml");

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
                                          from mi in db.MapInfos
                                          from mip in db.MapInfoPoints
                                          let hasSample = (from c in db.MWQMSamples
                                                           where c.MWQMSiteTVItemID == t.TVItemID
                                                           && c.UseForOpenData == true
                                                           select c).Any()
                                          where mi.TVItemID == t.TVItemID
                                          && mip.MapInfoID == mi.MapInfoID
                                          && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                          && t.TVType == (int)TVTypeEnum.MWQMSite
                                          && mi.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Point
                                          && mi.TVType == (int)TVTypeEnum.MWQMSite
                                          && hasSample == true
                                          select new { t, mip, hasSample }).ToList();


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
                        string MS = mwqmSite.t.TVItemID.ToString();
                        string Lat = (mwqmSite.mip != null ? mwqmSite.mip.Lat.ToString("F5") : "");
                        string Lng = (mwqmSite.mip != null ? mwqmSite.mip.Lng.ToString("F5") : "");
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
}
