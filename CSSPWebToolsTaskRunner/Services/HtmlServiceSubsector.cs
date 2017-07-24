using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
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
    public class HtmlServiceSubsector
    {
        #region Variables
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        #endregion Variables

        #region Constructors
        public HtmlServiceSubsector(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
        }
        #endregion Constructors

        #region Functions public
        public void Generate(FileInfo fi)
        {
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

            StringBuilder sbHTML = new StringBuilder();

            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemModel tvItemModelSubsector = tvItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);

            sbHTML.AppendLine(@"<!DOCTYPE html>");
            sbHTML.AppendLine(@"<html xmlns=""http://www.w3.org/1999/xhtml"">");
            sbHTML.AppendLine(@"<head>");
            sbHTML.AppendLine(@"<meta charset=""utf-8"">");
            sbHTML.AppendLine(@"<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"">");
            sbHTML.AppendLine(@"<title>");
            sbHTML.AppendLine(tvItemModelSubsector.TVText);
            sbHTML.AppendLine(@"</title>");
            sbHTML.AppendLine(@"</head>");
            sbHTML.AppendLine(@"<body>");


            // Doing Municipality
            sbHTML.AppendLine(@"<h3>Municipality</h3>");
            List<TVItemModel> tvItemModelMunicipalityList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.Municipality);

            sbHTML.AppendLine(@"<ul>");
            foreach (TVItemModel tvItemModelMunicipality in tvItemModelMunicipalityList)
            {
                sbHTML.AppendLine(@"<li>");
                sbHTML.AppendLine(@"" + tvItemModelMunicipality.TVText + "");
                sbHTML.AppendLine(@"</li>");
            }
            sbHTML.AppendLine(@"</ul>");

            // Doing Pollution Source Site
            sbHTML.AppendLine(@"<h3>Pollution Source Site</h3>");
            List<TVItemModel> tvItemModelPolSourceSiteList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.PolSourceSite);

            sbHTML.AppendLine(@"<ul>");
            foreach (TVItemModel tvItemModelPolSourceSite in tvItemModelPolSourceSiteList)
            {
                sbHTML.AppendLine(@"<li>");
                sbHTML.AppendLine(@"" + tvItemModelPolSourceSite.TVText + "");
                sbHTML.AppendLine(@"</li>");
            }
            sbHTML.AppendLine(@"</ul>");

            // Doing MWQM Site
            sbHTML.AppendLine(@"<h3>MWQM Site</h3>");
            List<TVItemModel> tvItemModelMWQMSiteList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.MWQMSite);

            sbHTML.AppendLine(@"<ul>");
            foreach (TVItemModel tvItemModelMWQMSite in tvItemModelMWQMSiteList)
            {
                sbHTML.AppendLine(@"<li>");
                sbHTML.AppendLine(@"" + tvItemModelMWQMSite.TVText + "");
                sbHTML.AppendLine(@"</li>");
            }
            sbHTML.AppendLine(@"</ul>");


            sbHTML.AppendLine(@"</body>");
            sbHTML.AppendLine(@"</html>");

            StreamWriter sw = fi.CreateText();

            sw.Write(sbHTML.ToString());

            sw.Close();

        }
        #endregion Functions public
    }
}
