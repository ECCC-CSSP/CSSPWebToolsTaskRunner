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
        private bool GenerateHTMLSubsectorMapActivePolSourceSitesDocx(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            string NotUsed = "";
            int TVItemID = 0;
            string HideVerticalScale = "";
            string HideHorizontalScale = "";
            string HideNorthArrow = "";
            string HideSubsectorName = "";

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

            HideVerticalScale = GetParameters("HideVerticalScale", ParamValueList);
            HideHorizontalScale = GetParameters("HideHorizontalScale", ParamValueList);
            HideNorthArrow = GetParameters("HideNorthArrow", ParamValueList);
            HideSubsectorName = GetParameters("HideSubsectorName", ParamValueList);

            string SubsectorTVText = _MWQMSubsectorService.GetMWQMSubsectorModelWithMWQMSubsectorTVItemIDDB(TVItemID).MWQMSubsectorTVText;

            sbHTML.AppendLine($@"<h2>{ SubsectorTVText }</h2>");
            sbHTML.AppendLine($@"<br />");

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 5);

            GoogleMapToPNG googleMapToPNG = new GoogleMapToPNG(_TaskRunnerBaseService, HideVerticalScale, HideHorizontalScale, HideNorthArrow, HideSubsectorName);

            DirectoryInfo di = new DirectoryInfo(googleMapToPNG.DirName);

            if (!di.Exists)
            {
                try
                {
                    di.Create();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateDirectory__, di.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateDirectory__", di.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return false;
                }
            }

            if (!googleMapToPNG.CreateSubsectorGoogleMapPNGForPolSourceSites(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, "hybrid"))
            {
                string Error = _TaskRunnerBaseService._BWObj.TextLanguageList[(_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr ? 1 : 0)].Text;

                sbHTML.AppendLine($@"<h1>{ Error }</h1>");
            }

            sbHTML.AppendLine($@"|||Image|FileName,{ googleMapToPNG.DirName }{ googleMapToPNG.FileNameFullAnnotated }|width,490|height,460|||");

            sbHTML.AppendLine($@"|||FileNameExtra|Random,{ googleMapToPNG.FileNameExtra }|||");
            if (!GetBottomHTML(sbHTML, fi, parameters))
            {
                return false;
            }

            return true;
        }

        // for testing only can comment out when test is completed
        public bool PublicGenerateHTMLSubsectorMapActivePolSourceSitesDocx(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            bool retBool = GenerateHTMLSubsectorMapActivePolSourceSitesDocx(fi, sbHTML, parameters, reportTypeModel);

            StreamWriter sw = fi.CreateText();
            sw.Write(sbHTML.ToString());
            sw.Flush();
            sw.Close();

            return retBool;
        }
    }
}
