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
using System.Web.Helpers;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        private bool GenerateHTMLSubsectorFullReportCoverPage(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            string NotUsed = "";
            int TVItemID = 0;
            int Year = 0;

            Random random = new Random();
            string FileNameExtra = "";
            for (int i = 0; i < 10; i++)
            {
                FileNameExtra = FileNameExtra + (char)random.Next(97, 122);
            }

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

            TVItemModel tvItemModelSubsector = _TVItemService.GetTVItemModelWithTVItemIDDB(TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelSubsector.Error))
            {
                NotUsed = tvItemModelSubsector.Error;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(tvItemModelSubsector.Error);
                return false;
            }

            if (!int.TryParse(GetParameters("Year", ParamValueList), out Year))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind__, TaskRunnerServiceRes.Parameter, TaskRunnerServiceRes.Year);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind__", TaskRunnerServiceRes.Parameter, TaskRunnerServiceRes.Year);
                return false;
            }

            TVItemModel tvItemModelRoot = _TVItemService.GetRootTVItemModelDB();
            if (!string.IsNullOrWhiteSpace(tvItemModelRoot.Error))
            {
                NotUsed = tvItemModelRoot.Error;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(tvItemModelRoot.Error);
                return false;
            }

            string ServerPath = _TVFileService.GetServerFilePath(tvItemModelRoot.TVItemID);

            TVFileModel tvFileModelFullReportCoverPageCanadaFlag = _TVFileService.GetTVFileModelWithServerFilePathAndServerFileNameDB(ServerPath, "CanadaFlag.png");
            if (!string.IsNullOrWhiteSpace(tvFileModelFullReportCoverPageCanadaFlag.Error))
            {
                NotUsed = tvFileModelFullReportCoverPageCanadaFlag.Error;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(tvFileModelFullReportCoverPageCanadaFlag.Error);
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
                NotUsed = tvFileModelFullReportCoverPageCanadaBanner.Error;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(tvFileModelFullReportCoverPageCanadaBanner.Error);
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
                NotUsed = tvFileModelFullReportCoverPageCanadaWithFlag.Error;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(tvFileModelFullReportCoverPageCanadaWithFlag.Error);
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
            }

            string NamesOfAuthors = "Name of authors";
            string ReportDateText = DateTime.Now.ToString("MMMM yyyy");

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 5);

            sbHTML.AppendLine($@" <p>&nbsp;</p>");
            sbHTML.AppendLine($@" <table>");
            sbHTML.AppendLine($@"    <tr> ");
            sbHTML.AppendLine($@"        <td>|||Image|FileName,{ fiFullReportCoverPageImageCanadaFlag.FullName }|width,66|height,28|||</td> ");
            sbHTML.AppendLine($@"        <td>&nbsp;&nbsp;&nbsp;</td>");
            sbHTML.AppendLine($@"        <td class=""textAlignLeft""><h5>Environment and <br />Climate Change Canada</h5></td>");
            sbHTML.AppendLine($@"        <td>&nbsp;&nbsp;&nbsp;</td>");
            sbHTML.AppendLine($@"        <td class=""textAlignLeft""><h5>Environnement et <br />Changement climatique Canada</h5></td>");
            sbHTML.AppendLine($@"    </tr>");
            sbHTML.AppendLine($@" </table>");
            sbHTML.AppendLine($@" <div class=""textAlignLeft"">");
            sbHTML.AppendLine($@"        |||Image|FileName,{ fiFullReportCoverPageImageCanadaBanner.FullName }|width,480|height,48|||");
            sbHTML.AppendLine($@" </div>");
            sbHTML.AppendLine($@" <div>");
            sbHTML.AppendLine($@"   <br />");
            sbHTML.AppendLine($@"   <br />");
            sbHTML.AppendLine($@"   <blockquote>");
            sbHTML.AppendLine($@"       <hr />");
            sbHTML.AppendLine($@"       <h5 class=""textAlignLeft"">{ TaskRunnerServiceRes.MarineWaterQualityReEvaluationReport }</h5>");
            sbHTML.AppendLine($@"       <hr />");
            sbHTML.AppendLine($@"       <h5 class=""textAlignLeft"">{ tvItemModelProvince.TVText } { TaskRunnerServiceRes.ShellfishGrowingArea}</h5>");
            sbHTML.AppendLine($@"       <h5 class=""textAlignLeft"">{ SubsectorShort }</h5>");
            sbHTML.AppendLine($@"       <h5 class=""textAlignLeft"">{ SubsectorEndPart }</h5>");
            sbHTML.AppendLine($@"       <hr />");
            sbHTML.AppendLine($@"       <h5 class=""textAlignLeft"">{ NamesOfAuthors }</h5>");
            sbHTML.AppendLine($@"       <hr />");
            sbHTML.AppendLine($@"   </blockquote>");
            sbHTML.AppendLine($@"   <h5>&nbsp;</h5>");
            sbHTML.AppendLine($@"   <h5>&nbsp;</h5>");
            sbHTML.AppendLine($@"   <h5>&nbsp;</h5>");
            sbHTML.AppendLine($@"   <h5>&nbsp;</h5>");
            sbHTML.AppendLine($@"   <h5>&nbsp;</h5>");
            sbHTML.AppendLine($@"   <h5>&nbsp;</h5>");
            sbHTML.AppendLine($@"   <h5 class=""textAlignRight"">{ TaskRunnerServiceRes.AtlanticMarineWaterQualityMonitoring }</h5>");
            sbHTML.AppendLine($@"   <h5 class=""textAlignRight"">{ TaskRunnerServiceRes.Report } <span>_______________________</span></h5>");
            sbHTML.AppendLine($@"   <h5 class=""textAlignRight"">{ ReportDateText }</h5>");
            sbHTML.AppendLine($@" </div>");
            sbHTML.AppendLine($@" <div class=""textAlignRight"">|||Image|FileName,{ fiFullReportCoverPageImageCanadaWithFlag.FullName }|width,76|height,25|||</div>");

            sbHTML.AppendLine(@"<span>|||PageBreak|||</span>");

            if (!GetBottomHTML(sbHTML, fi, parameters))
            {
                return false;
            }

            return true;
        }

        // for testing only can comment out when test is completed
        public bool PublicGenerateHTMLSubsectorFullReportCoverPage(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            bool retBool = GenerateHTMLSubsectorFullReportCoverPage(fi, sbHTML, parameters, reportTypeModel);

            StreamWriter sw = fi.CreateText();
            sw.Write(sbHTML.ToString());
            sw.Flush();
            sw.Close();

            return retBool;
        }
    }
}
