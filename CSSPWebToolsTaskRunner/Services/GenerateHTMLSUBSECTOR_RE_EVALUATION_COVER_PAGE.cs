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
        private bool GenerateHTMLSUBSECTOR_RE_EVALUATION_COVER_PAGE(StringBuilder sbTemp)
        {
            int Percent = 10;
            string NotUsed = "";

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", ReportGenerateObjectsKeywordEnum.SUBSECTOR_RE_EVALUATION_COVER_PAGE.ToString()));

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

            TVFileModel tvFileModelFullReportCoverPageCanadaFlag = _TVFileService.GetTVFileModelWithServerFilePathAndServerFileNameDB(ServerPath, "CanadaFlag.png");
            if (!string.IsNullOrWhiteSpace(tvFileModelFullReportCoverPageCanadaFlag.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, ServerPath + "CanadaFlag.png");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", ServerPath + "CanadaFlag.png");
                return false;
            }

            FileInfo fiFullReportCoverPageImageCanadaFlag = new FileInfo(tvFileModelFullReportCoverPageCanadaFlag.ServerFilePath + tvFileModelFullReportCoverPageCanadaFlag.ServerFileName);

            if (!fiFullReportCoverPageImageCanadaFlag.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiFullReportCoverPageImageCanadaFlag.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiFullReportCoverPageImageCanadaFlag.FullName);
                return false;
            }

            TVFileModel tvFileModelFullReportCoverPageCanadaBanner = _TVFileService.GetTVFileModelWithServerFilePathAndServerFileNameDB(ServerPath, "CanadaBanner.png");
            if (!string.IsNullOrWhiteSpace(tvFileModelFullReportCoverPageCanadaBanner.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, ServerPath + "CanadaBanner.png");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", ServerPath + "CanadaBanner.png");
                return false;
            }

            FileInfo fiFullReportCoverPageImageCanadaBanner = new FileInfo(tvFileModelFullReportCoverPageCanadaBanner.ServerFilePath + tvFileModelFullReportCoverPageCanadaBanner.ServerFileName);

            if (!fiFullReportCoverPageImageCanadaBanner.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiFullReportCoverPageImageCanadaBanner.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiFullReportCoverPageImageCanadaBanner.FullName);
                return false;
            }

            TVFileModel tvFileModelFullReportCoverPageCanadaWithFlag = _TVFileService.GetTVFileModelWithServerFilePathAndServerFileNameDB(ServerPath, "CanadaWithFlag.png");
            if (!string.IsNullOrWhiteSpace(tvFileModelFullReportCoverPageCanadaWithFlag.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, ServerPath + "CanadaWithFlag.png");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", ServerPath + "CanadaWithFlag.png");
                return false;
            }

            FileInfo fiFullReportCoverPageImageCanadaWithFlag = new FileInfo(tvFileModelFullReportCoverPageCanadaWithFlag.ServerFilePath + tvFileModelFullReportCoverPageCanadaWithFlag.ServerFileName);

            if (!fiFullReportCoverPageImageCanadaWithFlag.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiFullReportCoverPageImageCanadaBanner.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiFullReportCoverPageImageCanadaWithFlag.FullName);
                return false;
            }

            List<TVItemModel> tvItemModelListParent = _TVItemService.GetParentsTVItemModelList(tvItemModelSubsector.TVPath);
            TVItemModel tvItemModelProvince = new TVItemModel();
            foreach (TVItemModel tvItemModel in tvItemModelListParent)
            {
                if (tvItemModel.TVType == TVTypeEnum.Province)
                {
                    tvItemModelProvince = tvItemModel;
                    break;
                }
            }

            int Pos = tvItemModelSubsector.TVText.IndexOf(" ");
            string SubsectorShort = "Error";
            string SubsectorEndPart = "Error";
            if (Pos > 0)
            {
                SubsectorShort = tvItemModelSubsector.TVText.Substring(0, Pos).Trim();
                SubsectorEndPart = tvItemModelSubsector.TVText.Substring(Pos).Trim();
                if (SubsectorEndPart.StartsWith("("))
                {
                    SubsectorEndPart = SubsectorEndPart.Substring(1);
                }
                if (SubsectorEndPart.EndsWith(")"))
                {
                    SubsectorEndPart = SubsectorEndPart.Substring(0, SubsectorEndPart.Length - 1);
                }
            }

            Percent = 30;
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);

            sbTemp.AppendLine($@" <p>&nbsp;</p>");
            sbTemp.AppendLine($@" <table>");
            sbTemp.AppendLine($@"    <tr> ");
            sbTemp.AppendLine($@"        <td>|||Image|FileName,{ fiFullReportCoverPageImageCanadaFlag.FullName }|width,45|height,20|||</td> ");
            sbTemp.AppendLine($@"        <td>&nbsp;&nbsp;&nbsp;</td>");
            sbTemp.AppendLine($@"        <td class=""textAlignLeft""><p style=""font-size: 0.8em"">Environment and <br />Climate Change Canada</p></td>");
            sbTemp.AppendLine($@"        <td>&nbsp;&nbsp;&nbsp;</td>");
            sbTemp.AppendLine($@"        <td class=""textAlignLeft""><p style=""font-size: 0.8em"">Environnement et <br />Changement climatique Canada</p></td>");
            sbTemp.AppendLine($@"    </tr>");
            sbTemp.AppendLine($@" </table>");
            sbTemp.AppendLine($@" <div class=""textAlignLeft"">");
            sbTemp.AppendLine($@"        |||Image|FileName,{ fiFullReportCoverPageImageCanadaBanner.FullName }|width,480|height,48|||");
            sbTemp.AppendLine($@" </div>");
            sbTemp.AppendLine($@" <div>");
            sbTemp.AppendLine($@"   <br />");
            sbTemp.AppendLine($@"   <br />");
            sbTemp.AppendLine($@"   <blockquote>");
            sbTemp.AppendLine($@"       <hr />");
            sbTemp.AppendLine($@"       <p class=""textAlignLeft"" style=""font-size: 1.2em;"">");
            sbTemp.AppendLine($@"           <strong>{ TaskRunnerServiceRes.ShellfishWaterClassificationProgram } ({ TaskRunnerServiceRes.SWCP }) </strong>");
            sbTemp.AppendLine($@"       </p>");
            sbTemp.AppendLine($@"       <p class=""textAlignLeft"" style=""font-size: 1.2em;"">");
            sbTemp.AppendLine($@"           <strong>{ TaskRunnerServiceRes.ReEvaluationReport }</strong>");
            sbTemp.AppendLine($@"       </p>");
            sbTemp.AppendLine($@"       <hr />");
            sbTemp.AppendLine($@"       <p class=""textAlignLeft"" style=""font-size: 1.2em;""><strong>{ tvItemModelProvince.TVText } { TaskRunnerServiceRes.ShellfishGrowingArea}</strong></p>");
            sbTemp.AppendLine($@"       <p class=""textAlignLeft"" style=""font-size: 1.2em;""><strong>{ SubsectorShort }</strong></p>");
            sbTemp.AppendLine($@"       <p class=""textAlignLeft"" style=""font-size: 1.2em;""><strong>{ SubsectorEndPart }</strong></p>");
            sbTemp.AppendLine($@"       <hr />");
            sbTemp.AppendLine($@"       <p class=""textAlignLeft"" style=""font-size: 1.2em;""><strong>{ TaskRunnerServiceRes.Authors }:</strong> { tvItemModelProvince.TVText } { TaskRunnerServiceRes.SWCP } { TaskRunnerServiceRes.Staff }</p>");
            sbTemp.AppendLine($@"       <hr />");
            sbTemp.AppendLine($@"   </blockquote>");
            for (int i = 0; i < 5; i++)
            {
                sbTemp.AppendLine($@"   <p>&nbsp;</p>");
            }
            sbTemp.AppendLine($@"   <p class=""textAlignRight"" style=""font-size: 1.2em;"">");
            sbTemp.AppendLine($@"       <strong>{ TaskRunnerServiceRes.ReportID }:</strong> { TaskRunnerServiceRes.ReEvaluation } ({ SubsectorShort }) { DateTime.Now.ToString("MMMM") } { Year }");
            sbTemp.AppendLine($@"   </p>");
            sbTemp.AppendLine($@" </div>");
            sbTemp.AppendLine($@" <div class=""textAlignRight"">|||Image|FileName,{ fiFullReportCoverPageImageCanadaWithFlag.FullName }|width,76|height,22|||</div>");

            return true;
        }

        // for testing only can comment out when test is completed
        public bool PublicGenerateHTMLSUBSECTOR_RE_EVALUATION_COVER_PAGE(StringBuilder sbTemp)
        {
            bool retBool = GenerateHTMLSUBSECTOR_RE_EVALUATION_COVER_PAGE(sbTemp);

            StreamWriter sw = fi.CreateText();
            sw.Write(sbTemp.ToString());
            sw.Flush();
            sw.Close();

            return retBool;
        }
    }
}
