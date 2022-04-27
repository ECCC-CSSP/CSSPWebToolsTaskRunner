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
using CSSPDBDLL;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        private bool GenerateHTMLSubsector()
        {
            switch (reportTypeModel.UniqueCode)
            {
                case "AllWQDataFRXLSX":
                case "AllWQDataENXLSX":
                    {
                        if (!GenerateHTMLAllWQDataXlsx())
                        {
                            return false;
                        }
                    }
                    break;
                case "TestObjectsFRDOCX":
                case "TestObjectsENDOCX":
                case "FCSummaryStatFRDOCX":
                case "FCSummaryStatENDOCX":
                case "ReEvaluationFRDOCX":
                case "ReEvaluationENDOCX":
                case "AnnualReviewFRDOCX":
                case "AnnualReviewENDOCX":
                case "PollutionAndSanitaryEvaluationFRDOCX":
                case "PollutionAndSanitaryEvaluationENDOCX":
                    {
                        if (!GenerateHTMLDocx())
                        {
                            return false;
                        }
                    }
                    break;
                case "WQMSitesPhotoAlbumENDOCX":
                case "WQMSitesPhotoAlbumFRDOCX":
                    {
                        if (!GenerateHTMLWQMSitesPhotoAlbumDocx())
                        {
                            return false;
                        }

                        if (!GenerateObjects())
                        {
                            return false;
                        }
                    }
                    break;
                case "FCSummaryStatFRXLSX":
                case "FCSummaryStatENXLSX":
                    {
                        if (!GenerateHTMLSubsector_SubsectorTestXlsx())
                        {
                            return false;
                        }
                    }
                    break;
                case "SomethingElseAsUniqueCode":
                    {
                    }
                    break;
                default:
                    if (!GenerateHTMLSubsector_NotImplemented())
                    {
                        return false;
                    }
                    break;
            }
            return true;
        }
        private bool GenerateHTMLAllWQDataXlsx()
        {
            string NotUsed = "";

            int SubsectorTVItemID = 0;

            List<string> ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            string tempVal = GetParameters("TVItemID", ParamValueList);
            if (string.IsNullOrWhiteSpace(tempVal))
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, "TVItemID");
                List<TextLanguage> TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", "TVItemID");
            }
            SubsectorTVItemID = int.Parse(tempVal);

            using (CSSPDBEntities db2 = new CSSPDBEntities())
            {
                var subsectorItem = (from c in db2.TVItems
                                     from cl in db2.TVItemLanguages
                                     where c.TVItemID == cl.TVItemID
                                     && c.TVItemID == SubsectorTVItemID
                                     && cl.Language == (int)LanguageEnum.en
                                     select new { c, cl }).FirstOrDefault();


                string tvText = subsectorItem.cl.TVText;
                string locator = tvText;
                string name = tvText;

                if (tvText.Contains(" "))
                {
                    locator = tvText.Substring(0, tvText.IndexOf(" "));
                    name = tvText.Substring(tvText.IndexOf(" ") + 1);
                }

                name = name.Trim();

                if (!string.IsNullOrWhiteSpace(name) && name.StartsWith("("))
                {
                    name = name.Substring(1);
                }
                if (!string.IsNullOrWhiteSpace(name) && name.StartsWith("("))
                {
                    name = name.Substring(1);
                }

                if (!string.IsNullOrWhiteSpace(name) && name.EndsWith(")"))
                {
                    name = name.Substring(0, name.Length - 1);
                }
                if (!string.IsNullOrWhiteSpace(name) && name.EndsWith(")"))
                {
                    name = name.Substring(0, name.Length - 1);
                }

                NotUsed = string.Format(TaskRunnerServiceRes.Creating_, $"{ locator } AllWQData.html");
                List<TextLanguage> TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", $"{ locator } AllWQData.html");

                _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);

                _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 10);

                sb.AppendLine("<!DOCTYPE html>");
                sb.AppendLine("");
                sb.AppendLine(@"<html lang=""en"" xmlns=""http://www.w3.org/1999/xhtml"">");
                sb.AppendLine("<head>");
                sb.AppendLine(@"    <meta charset=""utf-8"" />");
                sb.AppendLine("    <title></title>");
                //sb.AppendLine(@"    <style type=""text/css"">");
                //sb.AppendLine("        th,td {");
                //sb.AppendLine("            border: 1px solid black;");
                //sb.AppendLine("        }");
                //sb.AppendLine("    </style>");
                sb.AppendLine("</head>");
                sb.AppendLine("<body>");
                sb.AppendLine("<table>");
                sb.AppendLine("<tr>");
                sb.AppendLine("<th colspan=7>");
                sb.AppendLine($"{ tvText } Water Quality Data");
                sb.AppendLine("</th>");
                sb.AppendLine("</tr>");
                sb.AppendLine("<tr>");
                sb.AppendLine("<th>");
                sb.AppendLine("Site");
                sb.AppendLine("</th>");
                sb.AppendLine("<th>");
                sb.AppendLine("Date");
                sb.AppendLine("</th>");
                sb.AppendLine("<th>");
                sb.AppendLine("Time");
                sb.AppendLine("</th>");
                sb.AppendLine("<th>");
                sb.AppendLine("Type");
                sb.AppendLine("</th>");
                sb.AppendLine("<th>");
                sb.AppendLine("FC");
                sb.AppendLine("</th>");
                sb.AppendLine("<th>");
                sb.AppendLine("Sal");
                sb.AppendLine("</th>");
                sb.AppendLine("<th>");
                sb.AppendLine("Temp");
                sb.AppendLine("</th>");
                sb.AppendLine("</tr>");

                List<MWQMSite> mwqmSiteList = (from c in db2.TVItems
                                               from s in db2.MWQMSites
                                               where c.TVItemID == s.MWQMSiteTVItemID
                                               && c.TVType == (int)TVTypeEnum.MWQMSite
                                               && c.TVPath.Contains(subsectorItem.c.TVPath + "p")
                                               orderby s.MWQMSiteNumber ascending
                                               select s).ToList();

                List<TVItemLanguage> tvItemLanguageList = (from c in db2.TVItems
                                                           from cl in db2.TVItemLanguages
                                                           where c.TVItemID == cl.TVItemID
                                                           && c.TVType == (int)TVTypeEnum.MWQMSite
                                                           && c.TVPath.Contains(subsectorItem.c.TVPath + "p")
                                                           && c.ParentID == SubsectorTVItemID
                                                           && cl.Language == (int)LanguageEnum.en
                                                           select cl).ToList();

                foreach (MWQMSite mwqmSite in mwqmSiteList)
                {
                    string mwqmSiteTVText = (from c in tvItemLanguageList
                                             where c.TVItemID == mwqmSite.MWQMSiteTVItemID
                                             select c.TVText).FirstOrDefault();

                    List<MWQMSample> mwqmSampleList = (from s in db2.MWQMSamples
                                                       where s.MWQMSiteTVItemID == mwqmSite.MWQMSiteTVItemID
                                                       orderby s.SampleDateTime_Local descending
                                                       select s).ToList();

                    foreach (MWQMSample mwqmSample in mwqmSampleList)
                    {
                        sb.AppendLine("<tr>");
                        sb.AppendLine("<th>");
                        sb.AppendLine($"{ mwqmSiteTVText }");
                        sb.AppendLine("</th>");
                        sb.AppendLine("<th>");
                        sb.AppendLine($"{ mwqmSample.SampleDateTime_Local.ToString("yyyy-MM-dd") }");
                        sb.AppendLine("</th>");
                        sb.AppendLine("<th>");
                        sb.AppendLine($"{ mwqmSample.SampleDateTime_Local.ToString("hh:mm") }");
                        sb.AppendLine("</th>");
                        sb.AppendLine("<th>");
                        string type = "";
                        if (mwqmSample.SampleTypesText.Contains($"{ ((int)SampleTypeEnum.Routine).ToString() },"))
                        {
                            type = "Routine";
                        }
                        else if (mwqmSample.SampleTypesText.Contains($"{ ((int)SampleTypeEnum.RainCMP).ToString() },"))
                        {
                            type = "RainCMP";
                        }
                        else if (mwqmSample.SampleTypesText.Contains($"{ ((int)SampleTypeEnum.RainRun).ToString() },"))
                        {
                            type = "RainRun";
                        }
                        else if (mwqmSample.SampleTypesText.Contains($"{ ((int)SampleTypeEnum.DailyDuplicate).ToString() },"))
                        {
                            type = "DailyDuplicate";
                        }
                        else if (mwqmSample.SampleTypesText.Contains($"{ ((int)SampleTypeEnum.Infrastructure).ToString() },"))
                        {
                            type = "Infrastructure";
                        }
                        else if (mwqmSample.SampleTypesText.Contains($"{ ((int)SampleTypeEnum.IntertechDuplicate).ToString() },"))
                        {
                            type = "IntertechDuplicate";
                        }
                        else if (mwqmSample.SampleTypesText.Contains($"{ ((int)SampleTypeEnum.IntertechRead).ToString() },"))
                        {
                            type = "IntertechRead";
                        }
                        else if (mwqmSample.SampleTypesText.Contains($"{ ((int)SampleTypeEnum.ReopeningEmergencyRain).ToString() },"))
                        {
                            type = "ReopeningEmergencyRain";
                        }
                        else if (mwqmSample.SampleTypesText.Contains($"{ ((int)SampleTypeEnum.ReopeningSpill).ToString() },"))
                        {
                            type = "ReopeningSpill";
                        }
                        else if (mwqmSample.SampleTypesText.Contains($"{ ((int)SampleTypeEnum.Sanitary).ToString() },"))
                        {
                            type = "Sanitary";
                        }
                        else if (mwqmSample.SampleTypesText.Contains($"{ ((int)SampleTypeEnum.Study).ToString() },"))
                        {
                            type = "Study";
                        }
                        else if (mwqmSample.SampleTypesText.Contains($"{ ((int)SampleTypeEnum.Sediment).ToString() },"))
                        {
                            type = "Sediment";
                        }
                        else if (mwqmSample.SampleTypesText.Contains($"{ ((int)SampleTypeEnum.Bivalve).ToString() },"))
                        {
                            type = "Bivalve";
                        }
                        else
                        {
                            type = "Error";
                        }
                        sb.AppendLine($"{ type }");
                        sb.AppendLine("</th>");
                        sb.AppendLine("<th>");
                        sb.AppendLine($"{ mwqmSample.FecCol_MPN_100ml }");
                        sb.AppendLine("</th>");
                        sb.AppendLine("<th>");
                        sb.AppendLine(mwqmSample.Salinity_PPT == null ? "" : mwqmSample.Salinity_PPT.Value.ToString("F1"));
                        sb.AppendLine("</th>");
                        sb.AppendLine("<th>");
                        sb.AppendLine(mwqmSample.WaterTemp_C == null ? "" : mwqmSample.WaterTemp_C.Value.ToString("F1"));
                        sb.AppendLine("</th>");
                        sb.AppendLine("</tr>");
                    }
                }

                sb.AppendLine("</table>");
                sb.AppendLine("</body>");
                sb.AppendLine("</html>");

                FileInfo fi = new FileInfo($@"C:\CSSP\{ locator } AllWQData.html");

                StreamWriter sw = fi.CreateText();
                sw.WriteLine(sb.ToString());
                sw.Close();

                return true;
            }
        }

        private bool GenerateHTMLWQMSitesPhotoAlbumDocx()
        {
            string NotUsed = "";

            int SubsectorTVItemID = 0;

            List<string> ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            string tempVal = GetParameters("TVItemID", ParamValueList);
            if (string.IsNullOrWhiteSpace(tempVal))
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, "TVItemID");
                List<TextLanguage> TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", "TVItemID");
            }
            SubsectorTVItemID = int.Parse(tempVal);

            using (CSSPDBEntities db2 = new CSSPDBEntities())
            {
                var subsectorItem = (from c in db2.TVItems
                                     from cl in db2.TVItemLanguages
                                     where c.TVItemID == cl.TVItemID
                                     && c.TVItemID == SubsectorTVItemID
                                     && cl.Language == (int)LanguageEnum.en
                                     select new { c, cl }).FirstOrDefault();


                string tvText = subsectorItem.cl.TVText;
                string locator = tvText;
                string name = tvText;

                if (tvText.Contains(" "))
                {
                    locator = tvText.Substring(0, tvText.IndexOf(" "));
                    name = tvText.Substring(tvText.IndexOf(" ") + 1);
                }

                name = name.Trim();

                if (!string.IsNullOrWhiteSpace(name) && name.StartsWith("("))
                {
                    name = name.Substring(1);
                }
                if (!string.IsNullOrWhiteSpace(name) && name.StartsWith("("))
                {
                    name = name.Substring(1);
                }

                if (!string.IsNullOrWhiteSpace(name) && name.EndsWith(")"))
                {
                    name = name.Substring(0, name.Length - 1);
                }
                if (!string.IsNullOrWhiteSpace(name) && name.EndsWith(")"))
                {
                    name = name.Substring(0, name.Length - 1);
                }

                NotUsed = string.Format(TaskRunnerServiceRes.Creating_, $"{ locator } SitePhotoAlbum.html");
                List<TextLanguage> TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", $"{ locator } SitePhotoAlbum.html");

                _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);

                _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 10);

                sb.AppendLine("<!DOCTYPE html>");
                sb.AppendLine("");
                sb.AppendLine(@"<html lang=""en"" xmlns=""http://www.w3.org/1999/xhtml"">");
                sb.AppendLine("<head>");
                sb.AppendLine(@"    <meta charset=""utf-8"" />");
                sb.AppendLine($"    <title>{ tvText }</title>");
                sb.AppendLine("</head>");
                sb.AppendLine("<body>");

                var tvItemList = (from c in db2.TVItems
                                  from cl in db2.TVItemLanguages
                                  from ms in db2.MWQMSites
                                  where c.TVItemID == cl.TVItemID
                                  && c.TVItemID == ms.MWQMSiteTVItemID
                                  && c.TVType == (int)TVTypeEnum.MWQMSite
                                  && c.TVPath.Contains(subsectorItem.c.TVPath + "p")
                                  && cl.Language == (int)LanguageEnum.en
                                  && c.IsActive == true
                                  orderby cl.TVText ascending
                                  select new { c, cl, ms }).ToList();

                var tvFileList = (from c in db2.TVItems
                                  from cf in db2.TVFiles
                                  where c.TVItemID == cf.TVFileTVItemID
                                  && c.TVType == (int)TVTypeEnum.File
                                  && c.TVPath.Contains(subsectorItem.c.TVPath + "p")
                                  select new { c, cf }).ToList();

                int TotalPageCount = (int)Math.Ceiling((tvItemList.Count / 6.0D));
                int PageCount = 0;
                for (int i = 0, count = tvItemList.Count; i < count; i += 6)
                {
                    PageCount++;

                    sb.AppendLine($"<h3>{ tvText }&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;({ PageCount }/{ TotalPageCount })</h3>");
                    sb.AppendLine(@"<table style=""width: 95%; height: 95%;"">");
                    sb.AppendLine(@"<tr style=""height: 32%"">");

                    for (int j = 0; j < 6; j++)
                    {
                        if ((i + j) < tvItemList.Count)
                        {
                            var tvFileItem = (from c in tvFileList
                                              where c.c.TVPath.Contains(tvItemList[i + j].c.TVPath + "p")
                                              && c.cf.ServerFileName.ToLower().Contains("monitoring.jpg")
                                              select c).FirstOrDefault();

                            if (tvFileItem == null)
                            {
                                sb.AppendLine(@"<td style=""width: 45%;"">");
                                sb.AppendLine("<br><br>");
                                sb.AppendLine($@"<span style=""font-size: 10px"">{ tvItemList[i + j].cl.TVText } No Image</span>");
                                sb.AppendLine("<br><br>");
                                sb.AppendLine($@"<span style=""font-size: 10px"">with name ending with monitoring.jpg</span>");
                                sb.AppendLine("</td>");
                            }
                            else
                            {
                                sb.AppendLine(@"<td style=""width: 45%;"">");
                                sb.AppendLine($@"|||Image|FileName,{ tvFileItem.cf.ServerFilePath }{ tvFileItem.cf.ServerFileName }|width,224|height,190|||");
                                sb.AppendLine("<br>");
                                if (tvItemList[i + j].ms.MWQMSiteDescription.Length > 50)
                                {
                                    sb.AppendLine($@"<span style=""font-size: 10px"">{ tvItemList[i + j].cl.TVText } --- { tvItemList[i + j].ms.MWQMSiteDescription.Substring(0, 50) } ...</span>");
                                }
                                else
                                {
                                    sb.AppendLine($@"<span style=""font-size: 10px"">{ tvItemList[i + j].cl.TVText } --- { tvItemList[i + j].ms.MWQMSiteDescription }</span>");
                                }
                                sb.AppendLine("<br>");
                                sb.AppendLine("</td>");
                            }
                        }
                        else
                        {
                            sb.AppendLine(@"<td style=""width: 45%;"">");
                            sb.AppendLine("<br><br><br>");
                            sb.AppendLine($@"&nbsp;");
                            sb.AppendLine("</td>");
                        }
                        if (j % 2 != 0)
                        {
                            sb.AppendLine("</tr>");
                            if (j != 5)
                            {
                                sb.AppendLine(@"<tr style=""height: 32%"">");
                            }
                        }
                        if (j == 5)
                        {
                            sb.AppendLine("</tr>");
                            sb.AppendLine("</table>");

                        }
                    }
                }

                sb.AppendLine("</body>");
                sb.AppendLine("</html>");

                FileInfo fi = new FileInfo($@"C:\CSSP\{ locator } SitePhotoAlbum.html");

                StreamWriter sw = fi.CreateText();
                sw.WriteLine(sb.ToString());
                sw.Close();

                return true;
            }
        }

        private bool GenerateHTMLSubsector_SubsectorTestXlsx()
        {
            if (!GetTopHTML())
            {
                return false;
            }

            sb.AppendLine(@"<h2>This will contain the Summary statistics FC densities (MPN/100mL)</h2>");

            if (!GetBottomHTML())
            {
                return false;
            }

            return true;
        }
        private bool GenerateHTMLSubsector_NotImplemented()
        {
            if (!GetTopHTML())
            {
                return false;
            }

            sb.AppendLine(@"<h2>UniqueCode [" + reportTypeModel.UniqueCode + " is not implemented.</h2>");

            if (!GetBottomHTML())
            {
                return false;
            }

            return true;
        }
        private string GetLastClassificationColor(MWQMSiteLatestClassificationEnum? mwqmSiteLatestClassification)
        {
            if (mwqmSiteLatestClassification == null)
            {
                return "";
            }

            switch (mwqmSiteLatestClassification)
            {
                case MWQMSiteLatestClassificationEnum.Approved:
                    return "bggreena";
                case MWQMSiteLatestClassificationEnum.ConditionallyApproved:
                    return "bggreenf";
                case MWQMSiteLatestClassificationEnum.ConditionallyRestricted:
                    return "bgredf";
                case MWQMSiteLatestClassificationEnum.Prohibited:
                    return "bgblack";
                case MWQMSiteLatestClassificationEnum.Restricted:
                    return "bgredf";
                case MWQMSiteLatestClassificationEnum.Unclassified:
                    return "";
                default:
                    return "";
            }
        }
        private string GetLastClassificationInitial(MWQMSiteLatestClassificationEnum? mwqmSiteLatestClassification)
        {
            if (mwqmSiteLatestClassification == null)
            {
                return "";
            }

            switch (mwqmSiteLatestClassification)
            {
                case MWQMSiteLatestClassificationEnum.Approved:
                    return (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr ? "A" : "A");
                case MWQMSiteLatestClassificationEnum.ConditionallyApproved:
                    return (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr ? "AC" : "CA");
                case MWQMSiteLatestClassificationEnum.ConditionallyRestricted:
                    return (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr ? "RC" : "CR");
                case MWQMSiteLatestClassificationEnum.Prohibited:
                    return (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr ? "P" : "P");
                case MWQMSiteLatestClassificationEnum.Restricted:
                    return (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr ? "R" : "R");
                case MWQMSiteLatestClassificationEnum.Unclassified:
                    return (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr ? "" : "");
                default:
                    return "";
            }
        }

        // for testing only can comment out when test is completed
    }
}
