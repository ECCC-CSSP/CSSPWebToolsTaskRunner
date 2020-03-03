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

            List<TVItemModel> tvItemModelParentList = _TVItemService.GetParentsTVItemModelList(tvItemModelSubsector.TVPath);
            TVItemModel tvItemModelProv = new TVItemModel();
            foreach (TVItemModel tvItemModel in tvItemModelParentList)
            {
                if (tvItemModel.TVType == TVTypeEnum.Province)
                {
                    tvItemModelProv = tvItemModel;
                }
            }

            string ProvInitCap = "";
            switch (tvItemModelProv.TVText)
            {
                case "New Brunswick":
                case "Nouveau-Brunswick":
                    {
                        ProvInitCap = "NB";
                    }
                    break;
                case "Newfoundland and Labrador":
                case "Terre-Neuve_et-Labrador":
                    {
                        ProvInitCap = "NL";
                    }
                    break;
                case "Nova Scotia":
                case "Nouvelle-Écosse":
                    {
                        ProvInitCap = "NS";
                    }
                    break;
                case "Prince Edward Island":
                case "Île-du-Prince-Édouard":
                    {
                        ProvInitCap = "PE";
                    }
                    break;
                case "British Columbia":
                case "Colombie-Britannique":
                    {
                        ProvInitCap = "BC";
                    }
                    break;
                case "Quebec":
                case "Québec":
                    {
                        ProvInitCap = "QC";
                    }
                    break;
                default:
                    break;
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

            TVFileModel tvFileModelFullReportCoverPageLeafRight = _TVFileService.GetTVFileModelWithServerFilePathAndServerFileNameDB(ServerPath, "LeafRight.png");
            if (!string.IsNullOrWhiteSpace(tvFileModelFullReportCoverPageLeafRight.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, ServerPath + "LeafRight.png");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", ServerPath + "LeafRight.png");
                return false;
            }

            FileInfo fiFullReportCoverPageImageLeafRight = new FileInfo(tvFileModelFullReportCoverPageLeafRight.ServerFilePath + tvFileModelFullReportCoverPageLeafRight.ServerFileName);

            if (!fiFullReportCoverPageImageLeafRight.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiFullReportCoverPageImageLeafRight.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiFullReportCoverPageImageLeafRight.FullName);
                return false;
            }

            TVFileModel tvFileModelFullReportCoverPageBarTopBottom = _TVFileService.GetTVFileModelWithServerFilePathAndServerFileNameDB(ServerPath, "BarTopBottom.png");
            if (!string.IsNullOrWhiteSpace(tvFileModelFullReportCoverPageBarTopBottom.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, ServerPath + "BarTopBottom.png");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", ServerPath + "BarTopBottom.png");
                return false;
            }

            FileInfo fiFullReportCoverPageImageBarTopBottom = new FileInfo(tvFileModelFullReportCoverPageBarTopBottom.ServerFilePath + tvFileModelFullReportCoverPageBarTopBottom.ServerFileName);

            if (!fiFullReportCoverPageImageBarTopBottom.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiFullReportCoverPageImageBarTopBottom.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiFullReportCoverPageImageBarTopBottom.FullName);
                return false;
            }

            TVFileModel tvFileModelFullReportCoverPageThreeImagesBottom = _TVFileService.GetTVFileModelWithServerFilePathAndServerFileNameDB(ServerPath, "ThreeImagesBottom.png");
            if (!string.IsNullOrWhiteSpace(tvFileModelFullReportCoverPageThreeImagesBottom.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, ServerPath + "ThreeImagesBottom.png");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", ServerPath + "ThreeImagesBottom.png");
                return false;
            }

            FileInfo fiFullReportCoverPageImageThreeImagesBottom = new FileInfo(tvFileModelFullReportCoverPageThreeImagesBottom.ServerFilePath + tvFileModelFullReportCoverPageThreeImagesBottom.ServerFileName);

            if (!fiFullReportCoverPageImageThreeImagesBottom.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiFullReportCoverPageImageThreeImagesBottom.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiFullReportCoverPageImageThreeImagesBottom.FullName);
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
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiFullReportCoverPageImageCanadaWithFlag.FullName);
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

            sbTemp.AppendLine($@" <table>");
            sbTemp.AppendLine($@"    <tr>");
            sbTemp.AppendLine($@"        <td>|||Image|FileName,{ fiFullReportCoverPageImageCanadaFlag.FullName }|width,45|height,20|||<br /><br /><br /><br /><br /><br /></td> ");
            sbTemp.AppendLine($@"        <td>&nbsp;&nbsp;&nbsp;<br /><br /><br /><br /><br /></td>");
            sbTemp.AppendLine($@"        <td class=""textAlignLeft""><p>Environment and <br />Climate Change Canada</p><br /><br /><br /><br /></td>");
            sbTemp.AppendLine($@"        <td>&nbsp;&nbsp;&nbsp;<br /><br /><br /><br /><br /></td>");
            sbTemp.AppendLine($@"        <td class=""textAlignLeft""><p>Environnement et <br />Changement climatique Canada</p><br /><br /><br /><br /></td>");
            sbTemp.AppendLine($@"        <td>|||Image|FileName,{ fiFullReportCoverPageImageLeafRight.FullName }|width,180|height,100|||</td> ");
            sbTemp.AppendLine($@"    </tr>");
            sbTemp.AppendLine($@"    <tr>");
            sbTemp.AppendLine($@"        <td colspan=""6"">|||Image|FileName,{ fiFullReportCoverPageImageBarTopBottom.FullName }|width,467|height,10|||</td> ");
            sbTemp.AppendLine($@"    </tr>");
            sbTemp.AppendLine($@" </table>");
            sbTemp.AppendLine($@" <div>");
            sbTemp.AppendLine($@"   <p>&nbsp;</p>");
            sbTemp.AppendLine($@"   <hr style=""color: blue"" />");
            sbTemp.AppendLine($@"   <p class=""textAlignLeft"" style=""font-size: 1.2em;"">");
            sbTemp.AppendLine($@"       <strong>{ TaskRunnerServiceRes.ShellfishWaterClassificationProgram } ({ TaskRunnerServiceRes.SWCP }) </strong>");
            sbTemp.AppendLine($@"   </p>");
            sbTemp.AppendLine($@"   <p class=""textAlignLeft"" style=""font-size: 1.2em;"">");
            sbTemp.AppendLine($@"       <strong>{ TaskRunnerServiceRes.ReEvaluationReport }</strong>&nbsp;|||STATISTICS_LAST_YEAR|||");
            sbTemp.AppendLine($@"   </p>");
            sbTemp.AppendLine($@"   <hr style=""color: blue"" />");
            sbTemp.AppendLine($@"   <p class=""textAlignLeft"" style=""font-size: 1.2em;""><strong>|||PROVINCE_INITIAL||| { TaskRunnerServiceRes.ShellfishGrowingArea}</strong></p>");
            sbTemp.AppendLine($@"   <p class=""textAlignLeft"" style=""font-size: 1.2em;"">|||SUBSECTOR_NAME_SHORT|||</p>");
            sbTemp.AppendLine($@"   <p class=""textAlignLeft"" style=""font-size: 1.2em;"">|||SUBSECTOR_NAME_TEXT|||</p>");
            sbTemp.AppendLine($@"   <hr style=""color: blue"" />");
            sbTemp.AppendLine($@"   <p class=""textAlignLeft"" style=""font-size: 1.2em;""><strong>{ TaskRunnerServiceRes.Authors }:</strong>&nbsp;|||{ProvInitCap}_AUTHORS|||</p>");
            sbTemp.AppendLine($@"   <p class=""textAlignLeft"" style=""font-size: 1.2em;""><strong>{ TaskRunnerServiceRes.ReportID }:</strong>&nbsp;R-|||STATISTICS_LAST_YEAR|||&nbsp;({ SubsectorShort })&nbsp;Created&nbsp;{ DateTime.Now.ToString("MMMM dd, yyyy") }</p>");
            sbTemp.AppendLine($@"   <hr style=""color: blue"" />");
            sbTemp.AppendLine($@"   <p>&nbsp;</p>");
            sbTemp.AppendLine($@" </div>");
            sbTemp.AppendLine($@" <div>");
            sbTemp.AppendLine($@" |||Image|FileName,{ fiFullReportCoverPageImageThreeImagesBottom.FullName }|width,467|height,110|||");
            sbTemp.AppendLine($@" </div>");
            sbTemp.AppendLine($@" <div class=""textAlignRight"">");
            sbTemp.AppendLine($@" |||Image|FileName,{ fiFullReportCoverPageImageCanadaWithFlag.FullName }|width,76|height,22|||");
            sbTemp.AppendLine($@" </div>");

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
