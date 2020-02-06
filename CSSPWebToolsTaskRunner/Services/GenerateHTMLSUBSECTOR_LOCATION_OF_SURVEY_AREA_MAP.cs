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
using CSSPDBDLL.Models;
using CSSPDBDLL.Services;
using CSSPEnumsDLL.Enums;
using CSSPEnumsDLL.Services;
using CSSPModelsDLL.Models;
using System.Windows.Forms;
//using System.Web.Helpers;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        private bool GenerateHTMLSUBSECTOR_LOCATION_OF_SURVEY_AREA_MAP(StringBuilder sbTemp)
        {
            int Percent = 10;
            string NotUsed = "";
            string HideVerticalScale = "";
            string HideHorizontalScale = "";
            string HideNorthArrow = "";
            string HideSubsectorName = "";

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", ReportGenerateObjectsKeywordEnum.SUBSECTOR_LOCATION_OF_SURVEY_AREA_MAP.ToString()));

            List<string> ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            // TVItemID and Year alrady loaded

            HideVerticalScale = GetParameters("HideVerticalScale", ParamValueList);
            HideHorizontalScale = GetParameters("HideHorizontalScale", ParamValueList);
            HideNorthArrow = GetParameters("HideNorthArrow", ParamValueList);
            HideSubsectorName = GetParameters("HideSubsectorName", ParamValueList);

            string SubsectorTVText = _MWQMSubsectorService.GetMWQMSubsectorModelWithMWQMSubsectorTVItemIDDB(TVItemID).MWQMSubsectorTVText;
            if (!string.IsNullOrEmpty(SubsectorTVText))
            {
                if (SubsectorTVText.Contains(" "))
                {
                    SubsectorTVText = SubsectorTVText.Substring(0, SubsectorTVText.IndexOf(" "));
                }
            }

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 5);

            List<TVFileModel> tvFileModelList = _TVFileService.GetTVFileModelListWithParentTVItemIDDB(TVItemID).OrderByDescending(c => c.ServerFileName).ToList();
            string FileFoundName = "";
            string StartText = $"{SubsectorTVText}_Location_of_Survey_Area_Map_";
            bool FileExist = false;
            int FileYear = 0;
            foreach (TVFileModel tvFileModel in tvFileModelList)
            {
                if (tvFileModel.ServerFileName.StartsWith(StartText) && tvFileModel.ServerFileName.EndsWith(".png"))
                {
                    string YearTxt = tvFileModel.ServerFileName;
                    YearTxt = YearTxt.Replace(StartText, "");
                    YearTxt = YearTxt.Replace(".png", "");

                    if (int.TryParse(YearTxt, out FileYear))
                    {
                        if (Year >= FileYear)
                        {
                            FileFoundName = tvFileModel.ServerFilePath + tvFileModel.ServerFileName;
                            FileExist = true;
                            break;
                        }
                    }
                }
            }

            if (FileExist)
            {
                if (Year == FileYear)
                {
                    sbTemp.AppendLine($@"<div>|||Image|FileName,{ FileFoundName }|width,600|height,440|||</div>");
                    sbTemp.AppendLine($@"<div>|||FigureCaption|Figure 1.0: { TaskRunnerServiceRes.LocationOfSurveyArea }|||</div>");
                }
                else
                {
                    sbTemp.Append($@"<p class=""bgyellow"">{ TaskRunnerServiceRes.From } { FileYear } { TaskRunnerServiceRes.Report }. ({ TaskRunnerServiceRes.Below })</p>");
                    sbTemp.AppendLine($@"<div>|||Image|FileName,{ FileFoundName }|width,600|height,440|||</div>");
                    sbTemp.AppendLine($@"<div>|||FigureCaption|Figure 1.0: { TaskRunnerServiceRes.LocationOfSurveyArea }|||</div>");
                    sbTemp.Append($@"<p class=""bgyellow"">{ TaskRunnerServiceRes.From } { FileYear } { TaskRunnerServiceRes.Report }. ({ TaskRunnerServiceRes.Above })</p>");
                }
            }
            else
            {
                sbTemp.AppendLine($@"<p>{ string.Format(TaskRunnerServiceRes.UploadYourFileNamed_ToReplaceTheImageBelow, StartText + Year.ToString() + ".png") }</p>");

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

                if (!googleMapToPNG.CreateSubsectorGoogleMapPNGForSubsectorOnly(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, "hybrid"))
                {
                    string Error = _TaskRunnerBaseService._BWObj.TextLanguageList[(_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr ? 1 : 0)].Text;

                    sbTemp.AppendLine($@"<h1>{ Error }</h1>");
                }

                sbTemp.AppendLine($@"<div>|||Image|FileName,{ googleMapToPNG.DirName }{ googleMapToPNG.FileNameFullAnnotated }|width,400|height,400|||</div>");
                sbTemp.AppendLine($@"<div>|||FigureCaption|Figure 1.0: { TaskRunnerServiceRes.LocationOfSurveyArea }|||</div>");
            }

            Percent = 70;
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);

            return true;
        }

        // for testing only can comment out when test is completed
        public bool PublicGenerateHTMLSUBSECTOR_LOCATION_OF_SURVEY_AREA_MAP(StringBuilder sbTemp)
        {
            bool retBool = GenerateHTMLSUBSECTOR_LOCATION_OF_SURVEY_AREA_MAP(sbTemp);

            StreamWriter sw = fi.CreateText();
            sw.Write(sbTemp.ToString());
            sw.Flush();
            sw.Close();

            return retBool;
        }
    }
}
