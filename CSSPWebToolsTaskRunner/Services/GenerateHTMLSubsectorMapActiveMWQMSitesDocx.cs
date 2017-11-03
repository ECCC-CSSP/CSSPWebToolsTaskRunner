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
using CSSPWebToolsDBDLL.Models;
using CSSPWebToolsDBDLL.Services;
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;
using System.Windows.Forms;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        private bool GenerateHTMLSubsectorMapMWQMSitesDocx(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            string NotUsed = "";
            int TVItemID = 0;

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 3);

            if (!GetTopHTML(sbHTML))
            {
                return false;
            }

            List<string> ParamValueList = parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            if (!int.TryParse(GetParameters("TVItemID", ParamValueList), out TVItemID))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind__, TaskRunnerServiceRes.Parameter, TaskRunnerServiceRes.TVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind__", TaskRunnerServiceRes.Parameter, TaskRunnerServiceRes.TVItemID);
                return false;
            }

            string SubsectorTVText = _MWQMSubsectorService.GetMWQMSubsectorModelWithMWQMSubsectorTVItemIDDB(TVItemID).MWQMSubsectorTVText;

            sbHTML.AppendLine($@"<h1>{ SubsectorTVText }</h1>");
            sbHTML.AppendLine($@"<br />");

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 5);

            GoogleMapToPNG googleMapToPNG = new GoogleMapToPNG(_TaskRunnerBaseService);

            if (!googleMapToPNG.CreateSubsectorGoogleMapPNGForMWQMSites(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, "hybrid"))
            {
                string Error = _TaskRunnerBaseService._BWObj.TextLanguageList[(_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr ? 1 : 0)].Text;

                sbHTML.AppendLine($@"<h1>{ Error }</h1>");
            }

            sbHTML.AppendLine($@"|||Image|FileName,{ googleMapToPNG.DirName }{ googleMapToPNG.FileNameFullAnnotated }|width,640|height,630|||");

            sbHTML.AppendLine($@"|||FileNameExtra|Random,{ googleMapToPNG.FileNameExtra }|||");
            if (!GetBottomHTML(sbHTML, fi, parameters))
            {
                return false;
            }

            return true;
        }

        // for testing only can comment out when test is completed
        public bool PublicGenerateHTMLSubsectorMapMWQMSitesDocx(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            bool retBool = GenerateHTMLSubsectorMapMWQMSitesDocx(fi, sbHTML, parameters, reportTypeModel);

            StreamWriter sw = fi.CreateText();
            sw.Write(sbHTML.ToString());
            sw.Flush();
            sw.Close();

            return retBool;
        }
    }
}
