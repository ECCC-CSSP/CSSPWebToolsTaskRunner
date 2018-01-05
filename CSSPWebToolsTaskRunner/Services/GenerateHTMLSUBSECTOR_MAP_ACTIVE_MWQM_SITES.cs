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
        private bool GenerateHTMLSUBSECTOR_MAP_ACTIVE_MWQM_SITES(StringBuilder sbTemp)
        {
            int Percent = 10;
            string NotUsed = "";
            string HideVerticalScale = "";
            string HideHorizontalScale = "";
            string HideNorthArrow = "";
            string HideSubsectorName = "";

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", ReportGenerateObjectsKeywordEnum.SUBSECTOR_MAP_ACTIVE_MWQM_SITES.ToString()));

            if (!GetTopHTML())
            {
                return false;
            }
            List<string> ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            // TVItemID and Year alrady loaded

            HideVerticalScale = GetParameters("HideVerticalScale", ParamValueList);
            HideHorizontalScale = GetParameters("HideHorizontalScale", ParamValueList);
            HideNorthArrow = GetParameters("HideNorthArrow", ParamValueList);
            HideSubsectorName = GetParameters("HideSubsectorName", ParamValueList);

            string SubsectorTVText = _MWQMSubsectorService.GetMWQMSubsectorModelWithMWQMSubsectorTVItemIDDB(TVItemID).MWQMSubsectorTVText;
         
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

            Percent = 40;
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);

            if (!googleMapToPNG.CreateSubsectorGoogleMapPNGForMWQMSites(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, "hybrid"))
            {
                string Error = _TaskRunnerBaseService._BWObj.TextLanguageList[(_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr ? 1 : 0)].Text;

                sbTemp.AppendLine($@"<h1>{ Error }</h1>");
            }

            sbTemp.AppendLine($@"|||Image|FileName,{ googleMapToPNG.DirName }{ googleMapToPNG.FileNameFullAnnotated }|width,490|height,460|||");
            sbTemp.AppendLine($@"|||FigureCaption|: { TaskRunnerServiceRes.MWQMSitesApproximateLocation }|||");

            Percent = 70;
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);

            if (!GetBottomHTML())
            {
                return false;
            }

            return true;
        }

        // for testing only can comment out when test is completed
        public bool PublicGenerateHTMLSUBSECTOR_MAP_ACTIVE_MWQM_SITES(StringBuilder sbTemp)
        {
            bool retBool = GenerateHTMLSUBSECTOR_MAP_ACTIVE_MWQM_SITES(sbTemp);

            StreamWriter sw = fi.CreateText();
            sw.Write(sbTemp.ToString());
            sw.Flush();
            sw.Close();

            return retBool;
        }
    }
}
