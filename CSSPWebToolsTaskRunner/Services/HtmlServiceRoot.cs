﻿using System;
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
    public class HtmlServiceRoot
    {
         #region Variables
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        #endregion Variables

        #region Constructors
        public HtmlServiceRoot(TaskRunnerBaseService taskRunnerBaseService)
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
            TVItemModel tvItemModelRoot = tvItemService.GetRootTVItemModelDB();

            sbHTML.AppendLine(@"<!DOCTYPE html>");
            sbHTML.AppendLine(@"<html xmlns=""http://www.w3.org/1999/xhtml"">");
            sbHTML.AppendLine(@"<head>");
            sbHTML.AppendLine(@"<meta charset=""utf-8"">");
            sbHTML.AppendLine(@"<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"">");
            sbHTML.AppendLine(@"<title>");
            sbHTML.AppendLine(tvItemModelRoot.TVText);
            sbHTML.AppendLine(@"</title>");
            sbHTML.AppendLine(@"</head>");
            sbHTML.AppendLine(@"<body>");

            List<TVItemModel> tvItemModelList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelRoot.TVItemID, TVTypeEnum.Country);

            sbHTML.AppendLine(@"<ul>");
            foreach (TVItemModel tvItemModelCountry in tvItemModelList)
            {
                sbHTML.AppendLine(@"<li>");
                sbHTML.AppendLine(@"" + tvItemModelCountry.TVText + "");
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
