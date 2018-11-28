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
using CSSPModelsDLL.Models;
using System.Windows.Forms;
//using System.Web.Helpers;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        private bool GenerateHTMLSUBSECTOR_ECCC_AND_SWCP_LOGO(StringBuilder sbTemp)
        {
            int Percent = 12;
            string NotUsed = "";

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", ReportGenerateObjectsKeywordEnum.SUBSECTOR_ECCC_AND_SWCP_LOGO.ToString()));

            List<string> ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            // TVItemID and Year already loaded

            TVItemModel tvItemModelSubsector = _TVItemService.GetTVItemModelWithTVItemIDDB(TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelSubsector.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(tvItemModelSubsector.Error);
                return false;
            }

            TVItemModel tvItemModelRoot = _TVItemService.GetRootTVItemModelDB();
            if (!string.IsNullOrWhiteSpace(tvItemModelRoot.Error))
            {
                NotUsed = TaskRunnerServiceRes.CouldNotFindTVItemRoot;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("CouldNotFindTVItemRoot");
                return false;
            }

            string ServerPath = _TVFileService.GetServerFilePath(tvItemModelRoot.TVItemID);

            TVFileModel tvFileModelFullReportECCCAndSWCPLogo = _TVFileService.GetTVFileModelWithServerFilePathAndServerFileNameDB(ServerPath, "ECCCAndSWCPLogo.png");
            if (!string.IsNullOrWhiteSpace(tvFileModelFullReportECCCAndSWCPLogo.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, ServerPath + "ECCCAndSWCPLogo.png");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", ServerPath + "ECCCAndSWCPLogo.png");
                return false;
            }

            FileInfo fiFullReportECCCAndSWCPLogo = new FileInfo(tvFileModelFullReportECCCAndSWCPLogo.ServerFilePath + tvFileModelFullReportECCCAndSWCPLogo.ServerFileName);

            if (!fiFullReportECCCAndSWCPLogo.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiFullReportECCCAndSWCPLogo.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiFullReportECCCAndSWCPLogo.FullName);
                return false;
            }

            sbTemp.AppendLine($@"        |||Image|FileName,{ fiFullReportECCCAndSWCPLogo.FullName }|width,245|height,62|||");
            sbTemp.AppendLine($@" <p>&nbsp;</p>");

            return true;
        }

        // for testing only can comment out when test is completed
        public bool PublicGenerateHTMLSUBSECTOR_ECCC_AND_SWCP_LOGO(StringBuilder sbTemp)
        {
            bool retBool = GenerateHTMLSUBSECTOR_ECCC_AND_SWCP_LOGO(sbTemp);

            StreamWriter sw = fi.CreateText();
            sw.Write(sbTemp.ToString());
            sw.Flush();
            sw.Close();

            return retBool;
        }
    }
}
