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
using CSSPDBDLL;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        private bool GenerateHTMLProvince()
        {
            switch (reportTypeModel.UniqueCode)
            {
                case "ProvincialSubsectorsReviewFRXLSX":
                case "ProvincialSubsectorsReviewENXLSX":
                    {
                        if (!GenerateHTMLProvincialSubsectorsReviewXlsx())
                        {
                            return false;
                        }
                    }
                    break;
                case "ProvinceTestFRDOCX":
                case "ProvinceTestENDOCX":
                    {
                        if (!GenerateHTMLProvince_ProvinceTestDocx())
                        {
                            return false;
                        }
                    }
                    break;
                case "ProvinceTestExcelFRXLSX":
                case "ProvinceTestExcelENXLSX":
                    {
                        if (!GenerateHTMLProvince_ProvinceTestXlsx())
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
                    if (!GenerateHTMLProvince_NotImplemented())
                    {
                        return false;
                    }
                    break;
            }
            return true;
        }
        private bool GenerateHTMLProvince_ProvinceTestDocx()
        {
            if (!GetTopHTML())
            {
                return false;
            }

            sb.AppendLine(@"<h2>Bonjour</h2>");

            if (!GetBottomHTML())
            {
                return false;
            }

            return true;
        }
        private bool GenerateHTMLProvince_ProvinceTestXlsx()
        {
            if (!GetTopHTML())
            {
                return false;
            }

            sb.AppendLine(@"<h2>Bonjour 2 for xlsx</h2>");

            if (!GetBottomHTML())
            {
                return false;
            }

            return true;
        }
        private bool GenerateHTMLProvincialSubsectorsReviewXlsx()
        {
            string NotUsed = "";
            bool ErrorInDoc = false;

            int ProvinceTVItemID = 7;
            int AfterYear = 1989;
            int NumberOfRuns = 30;
            bool OnlyActiveSubsectors = true;
            bool OnlyActiveMWQMSites = true;
            bool FullYear = true;

            List<string> ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            string tempVal = GetParameters("TVItemID", ParamValueList);
            if (string.IsNullOrWhiteSpace(tempVal))
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, "TVItemID");
                List<TextLanguage> TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", "TVItemID");
            }
            ProvinceTVItemID = int.Parse(tempVal);
            tempVal = GetParameters("AfterYear", ParamValueList);
            if (string.IsNullOrWhiteSpace(tempVal))
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, "AfterYear");
                List<TextLanguage> TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", "AfterYear");
            }
            tempVal = GetParameters("NumberOfRuns", ParamValueList);
            if (string.IsNullOrWhiteSpace(tempVal))
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, "NumberOfRuns");
                List<TextLanguage> TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", "NumberOfRuns");
            }
            NumberOfRuns = int.Parse(tempVal);
            OnlyActiveSubsectors = GetParameters("OnlyActiveSubsectors", ParamValueList) == "" ? false : true;
            OnlyActiveMWQMSites = GetParameters("OnlyActiveMWQMSites", ParamValueList) == "" ? false : true;
            FullYear = GetParameters("FullYear", ParamValueList) == "" ? false : true;

            string provInit = "";
            string prov = "";

            List<string> ProvInitList = new List<string>()
            {
                "BC", "ME", "NB", "NL", "NS", "PE", "QC",
            };
            List<string> ProvList = new List<string>()
            {
                "British Columbia", "Maine", "New Brunswick", "Newfoundland and Labrador", "Nova Scotia", "Prince Edward Island", "Québec",
            };

            using (CSSPDBEntities db2 = new CSSPDBEntities())
            {
                var provItem = (from c in db2.TVItems
                                from cl in db2.TVItemLanguages
                                where c.TVItemID == cl.TVItemID
                                && c.TVItemID == ProvinceTVItemID
                                && cl.Language == (int)LanguageEnum.en
                                select new { c, cl }).FirstOrDefault();

                for (int i = 0; i < ProvList.Count; i++)
                {
                    if (ProvList[i] == provItem.cl.TVText)
                    {
                        provInit = ProvInitList[i];
                        prov = ProvList[i];
                        break;
                    }
                }

                if (string.IsNullOrEmpty(provInit))
                {
                    return false;
                }

                NotUsed = string.Format(TaskRunnerServiceRes.Creating_, $"{ provInit } Subsector Review.html");
                List<TextLanguage> TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", $"{ provInit } Subsector Review.html");

                _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);

                _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 10);

                sb.AppendLine("<!DOCTYPE html>");
                sb.AppendLine("");
                sb.AppendLine(@"<html lang=""en"" xmlns=""http://www.w3.org/1999/xhtml"">");
                sb.AppendLine("<head>");
                sb.AppendLine(@"    <meta charset=""utf-8"" />");
                sb.AppendLine("    <title></title>");
                sb.AppendLine(@"    <style type=""text/css"">");
                sb.AppendLine("        th,td {");
                sb.AppendLine("            border: 1px solid black;");
                sb.AppendLine("        }");
                sb.AppendLine("    </style>");
                sb.AppendLine("</head>");
                sb.AppendLine("<body>");
                sb.AppendLine("<table>");
                sb.AppendLine("<tr>");
                sb.AppendLine("<th colspan=19>");
                sb.AppendLine("Data extrated from Webtools on DATE  (All-All-All, N= 30, Full year)");
                sb.AppendLine("</th>");
                sb.AppendLine("</tr>");
                sb.AppendLine("<tr>");
                sb.AppendLine("<th>");
                sb.AppendLine("Locator");
                sb.AppendLine("</th>");
                sb.AppendLine("<th>");
                sb.AppendLine("Name");
                sb.AppendLine("</th>");
                sb.AppendLine("<th>");
                sb.AppendLine("Max <br/># Samples");
                sb.AppendLine("</th>");
                sb.AppendLine("<th>");
                sb.AppendLine("Max-Min<br/>Years");
                sb.AppendLine("</th>");
                sb.AppendLine("<th>");
                sb.AppendLine("Sites Classified as <br/>Approved <br/>with Red A to F rating");
                sb.AppendLine("</th>");
                sb.AppendLine("<th>");
                sb.AppendLine("Sites Classified as <br/>Approved <br/>with Green F rating");
                sb.AppendLine("</th>");
                sb.AppendLine("<th>");
                sb.AppendLine("Sites Classified as <br/>Restricted or Conditional Restricted <br/>with red F rating");
                sb.AppendLine("</th>");
                sb.AppendLine("<th>");
                sb.AppendLine("Sites Classified as <br/>Restricted or Conditional Restricted <br/>with purple A-F rating");
                sb.AppendLine("</th>");
                sb.AppendLine("<th>");
                sb.AppendLine("Sites Classified as <br/>Restricted or Conditional Restricted <br/>with green A to D rating");
                sb.AppendLine("</th>");
                sb.AppendLine("<th>");
                sb.AppendLine("Sites Classified as <br/>Unclassified");
                sb.AppendLine("</th>");
                sb.AppendLine("<th>");
                sb.AppendLine("# Rain runs<br/> &ge; 12mm <br/>on Run day");
                sb.AppendLine("</th>");
                sb.AppendLine("<th>");
                sb.AppendLine("# Rain runs<br/> &ge; 12mm <br/>0-24 hrs");
                sb.AppendLine("</th>");
                sb.AppendLine("<th>");
                sb.AppendLine("# Rain runs<br/> &ge; 12mm <br/>24-48 hrs");
                sb.AppendLine("</th>");
                sb.AppendLine("<th>");
                sb.AppendLine("# Rain runs<br/> &ge; 25mm <br/>0-24 hrs");
                sb.AppendLine("</th>");
                sb.AppendLine("<th>");
                sb.AppendLine("Missing Rain data <br/>in Webtools <br/>(Yes/No)");
                sb.AppendLine("</th>");
                sb.AppendLine("<th>");
                sb.AppendLine("Last Relay<br/>request year");
                sb.AppendLine("</th>");
                sb.AppendLine("<th>");
                sb.AppendLine("Last depuration<br/>request year");
                sb.AppendLine("</th>");
                sb.AppendLine("<th>");
                sb.AppendLine("More Rain Runs Needed");
                sb.AppendLine("</th>");
                sb.AppendLine("<th>");
                sb.AppendLine("Action suggested");
                sb.AppendLine("</th>");
                sb.AppendLine("<th>");
                sb.AppendLine("Initials/Date reviewed");
                sb.AppendLine("</th>");
                sb.AppendLine("</tr>");

                sb2.AppendLine("<!DOCTYPE html>");
                sb2.AppendLine("");
                sb2.AppendLine(@"<html lang=""en"" xmlns=""http://www.w3.org/1999/xhtml"">");
                sb2.AppendLine("<head>");
                sb2.AppendLine(@"    <meta charset=""utf-8"" />");
                sb2.AppendLine("    <title></title>");
                sb2.AppendLine(@"    <style type=""text/css"">");
                sb2.AppendLine("        th,td {");
                sb2.AppendLine("            border: 1px solid black;");
                sb2.AppendLine("        }");
                sb2.AppendLine("    </style>");
                sb2.AppendLine("</head>");
                sb2.AppendLine("<body>");
                sb2.AppendLine("<table>");
                sb2.AppendLine("<tr>");
                sb2.AppendLine("<th colspan=8>");
                sb2.AppendLine("Data extrated from Webtools on DATE  (All-All-All, N= 30, Full year) by site");
                sb2.AppendLine("</th>");
                sb2.AppendLine("</tr>");
                sb2.AppendLine("<tr>");
                sb2.AppendLine("<th>");
                sb2.AppendLine("Locator");
                sb2.AppendLine("</th>");
                sb2.AppendLine("<th>");
                sb2.AppendLine("Name");
                sb2.AppendLine("</th>");
                sb2.AppendLine("<th>");
                sb2.AppendLine("# Samples");
                sb2.AppendLine("</th>");
                sb2.AppendLine("<th>");
                sb2.AppendLine("Years");
                sb2.AppendLine("</th>");
                sb2.AppendLine("<th>");
                sb2.AppendLine("Site");
                sb2.AppendLine("</th>");
                sb2.AppendLine("<th>");
                sb2.AppendLine("Currently <br/>Classified");
                sb2.AppendLine("</th>");
                sb2.AppendLine("<th>");
                sb2.AppendLine("Stat <br/>Color");
                sb2.AppendLine("</th>");
                sb2.AppendLine("<th>");
                sb2.AppendLine("Stat <br/>Letter");
                sb2.AppendLine("</th>");
                sb2.AppendLine("</tr>");

                TVItem tvItemProv = (from c in db2.TVItems
                                     from cl in db2.TVItemLanguages
                                     where c.TVItemID == cl.TVItemID
                                     && cl.Language == (int)LanguageEnum.en
                                     && c.TVType == (int)TVTypeEnum.Province
                                     && cl.TVText == prov
                                     select c).FirstOrDefault();

                var subsectorList = (from c in db2.TVItems
                                     from cl in db2.TVItemLanguages
                                     where c.TVItemID == cl.TVItemID
                                     && cl.Language == (int)LanguageEnum.en
                                     && c.TVType == (int)TVTypeEnum.Subsector
                                     && c.TVPath.Contains(tvItemProv.TVPath + "p")
                                     && (c.IsActive == true
                                     || c.IsActive == OnlyActiveSubsectors)
                                     orderby cl.TVText ascending
                                     select new { c, cl }).ToList();

                int SSCount = 0;
                foreach (var subsector in subsectorList)
                {
                    SSCount += 1;

                    string tvText = subsector.cl.TVText;
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

                    if (SSCount % 10 == 0)
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.Doing_, tvText);
                        TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Doing_", tvText);
                        _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);
                        _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)(10 + (80.0D*((double)SSCount/ (double)subsectorList.Count))));
                    }

                    List<MWQMSite> mwqmSiteList = (from c in db2.TVItems
                                                   from s in db2.MWQMSites
                                                   where c.TVItemID == s.MWQMSiteTVItemID
                                                   && c.TVType == (int)TVTypeEnum.MWQMSite
                                                   && c.TVPath.Contains(subsector.c.TVPath + "p")
                                                   && (c.IsActive == true
                                                   || c.IsActive == OnlyActiveMWQMSites)
                                                   orderby s.MWQMSiteNumber ascending
                                                   select s).ToList();

                    List<int> mwqmSiteTVItemIDList = (from c in mwqmSiteList
                                                      select c.MWQMSiteTVItemID).Distinct().ToList();

                    string routine = $"{ (int)SampleTypeEnum.Routine },";
                    List<MWQMSample> mwqmSampleList = (from s in db2.MWQMSamples
                                                       from c in mwqmSiteTVItemIDList
                                                       where s.MWQMSiteTVItemID == c
                                                       && s.SampleTypesText.Contains(routine)
                                                       && s.SampleDateTime_Local.Year > AfterYear
                                                       select s).ToList();

                    List<int> mwqmRunTVItemIDList = (from c in mwqmSampleList
                                                     select c.MWQMRunTVItemID).Distinct().ToList();

                    List<MWQMRun> mwqmRunList = (from r in db2.MWQMRuns
                                                 from c in mwqmRunTVItemIDList
                                                 where r.MWQMRunTVItemID == c
                                                 && r.DateTime_Local.Year > AfterYear
                                                 && r.RunSampleType == (int)SampleTypeEnum.Routine
                                                 select r).Distinct().ToList();


                    List<MWQMSiteClassStat> MWQMSiteClassStatList = new List<MWQMSiteClassStat>();

                    int MaxNumberOfSample = 0;
                    int SampleMaxYear = 0;
                    int SampleMinYear = 100000;

                    foreach (MWQMSite mwqmSite in mwqmSiteList)
                    {
                        MWQMSiteClassStat mwqmSiteClassStat = new MWQMSiteClassStat();
                        mwqmSiteClassStat.MWQMSite = mwqmSite;

                        var mwqmSampleListForMWQMSite = (from c in mwqmSampleList
                                                         where c.MWQMSiteTVItemID == mwqmSite.MWQMSiteTVItemID
                                                         orderby c.SampleDateTime_Local descending
                                                         select c).ToList();

                        int count = 0;
                        int lastYear = 0;
                        foreach (MWQMSample mwqmSample in mwqmSampleListForMWQMSite)
                        {
                            count += 1;

                            if (count > NumberOfRuns)
                            {
                                if (!FullYear)
                                {
                                    break;
                                }

                                if (lastYear != mwqmSample.SampleDateTime_Local.Year)
                                {
                                    break;
                                }
                            }

                            mwqmSiteClassStat.MWQMSampleList.Add(mwqmSample);

                            lastYear = mwqmSample.SampleDateTime_Local.Year;
                        }

                        mwqmSiteClassStat.MWQMRunList = (from c in mwqmSiteClassStat.MWQMSampleList
                                                         from r in mwqmRunList
                                                         where c.MWQMRunTVItemID == r.MWQMRunTVItemID
                                                         && c.MWQMSiteTVItemID == mwqmSite.MWQMSiteTVItemID
                                                         select r).Distinct().ToList();

                        if (mwqmSampleList.Count > 3)
                        {
                            List<double> mwqmSampleFCList = (from c in mwqmSiteClassStat.MWQMSampleList
                                                             where c.MWQMSiteTVItemID == mwqmSite.MWQMSiteTVItemID
                                                             orderby c.SampleDateTime_Local descending
                                                             select (c.FecCol_MPN_100ml < 2 ? 1.9D : (double)c.FecCol_MPN_100ml)).ToList<double>();

                            if (mwqmSampleFCList.Count > 3)
                            {
                                if (MaxNumberOfSample < mwqmSampleFCList.Count)
                                {
                                    MaxNumberOfSample = mwqmSampleFCList.Count;
                                }

                                double P90 = _TVItemService.GetP90(mwqmSampleFCList);
                                double GeoMean = _TVItemService.GeometricMean(mwqmSampleFCList);
                                double Median = _TVItemService.GetMedian(mwqmSampleFCList);
                                double PercOver43 = ((((double)mwqmSiteClassStat.MWQMSampleList.Where(c => c.FecCol_MPN_100ml > 43).Count()) / (double)mwqmSampleFCList.Count()) * 100.0D);
                                double PercOver260 = ((((double)mwqmSiteClassStat.MWQMSampleList.Where(c => c.FecCol_MPN_100ml > 260).Count()) / (double)mwqmSampleFCList.Count()) * 100.0D);
                                int MinYear = mwqmSiteClassStat.MWQMSampleList.Select(c => c.SampleDateTime_Local).Min().Year;
                                int MaxYear = mwqmSiteClassStat.MWQMSampleList.Select(c => c.SampleDateTime_Local).Max().Year;

                                if (SampleMaxYear < MaxYear)
                                {
                                    SampleMaxYear = MaxYear;
                                }

                                if (SampleMinYear > MinYear)
                                {
                                    SampleMinYear = MinYear;
                                }

                                LetterColorName letterColorName = new LetterColorName();

                                if ((GeoMean > 88) || (Median > 88) || (P90 > 260) || (PercOver260 > 10))
                                {
                                    if ((GeoMean > 181) || (Median > 181) || (P90 > 460) || (PercOver260 > 18))
                                    {
                                        letterColorName = new LetterColorName() { Letter = "F", Color = "Purple", Name = "NoDepuration" };
                                    }
                                    else if ((GeoMean > 163) || (Median > 163) || (P90 > 420) || (PercOver260 > 17))
                                    {
                                        letterColorName = new LetterColorName() { Letter = "E", Color = "Purple", Name = "NoDepuration" };
                                    }
                                    else if ((GeoMean > 144) || (Median > 144) || (P90 > 380) || (PercOver260 > 15))
                                    {
                                        letterColorName = new LetterColorName() { Letter = "D", Color = "Purple", Name = "NoDepuration" };
                                    }
                                    else if ((GeoMean > 125) || (Median > 125) || (P90 > 340) || (PercOver260 > 13))
                                    {
                                        letterColorName = new LetterColorName() { Letter = "C", Color = "Purple", Name = "NoDepuration" };
                                    }
                                    else if ((GeoMean > 107) || (Median > 107) || (P90 > 300) || (PercOver260 > 12))
                                    {
                                        letterColorName = new LetterColorName() { Letter = "B", Color = "Purple", Name = "NoDepuration" };
                                    }
                                    else
                                    {
                                        letterColorName = new LetterColorName() { Letter = "A", Color = "Purple", Name = "NoDepuration" };
                                    }
                                }
                                else if ((GeoMean > 14) || (Median > 14) || (P90 > 43) || (PercOver43 > 10))
                                {
                                    if ((GeoMean > 76) || (Median > 76) || (P90 > 224) || (PercOver43 > 27))
                                    {
                                        letterColorName = new LetterColorName() { Letter = "F", Color = "Red", Name = "Fail" };
                                    }
                                    else if ((GeoMean > 63) || (Median > 63) || (P90 > 188) || (PercOver43 > 23))
                                    {
                                        letterColorName = new LetterColorName() { Letter = "E", Color = "Red", Name = "Fail" };
                                    }
                                    else if ((GeoMean > 51) || (Median > 51) || (P90 > 152) || (PercOver43 > 20))
                                    {
                                        letterColorName = new LetterColorName() { Letter = "D", Color = "Red", Name = "Fail" };
                                    }
                                    else if ((GeoMean > 39) || (Median > 39) || (P90 > 115) || (PercOver43 > 17))
                                    {
                                        letterColorName = new LetterColorName() { Letter = "C", Color = "Red", Name = "Fail" };
                                    }
                                    else if ((GeoMean > 26) || (Median > 26) || (P90 > 79) || (PercOver43 > 13))
                                    {
                                        letterColorName = new LetterColorName() { Letter = "B", Color = "Red", Name = "Fail" };
                                    }
                                    else
                                    {
                                        letterColorName = new LetterColorName() { Letter = "A", Color = "Red", Name = "Fail" };
                                    }
                                }
                                else
                                {
                                    if ((GeoMean > 12) || (Median > 12) || (P90 > 36) || (PercOver43 > 8))
                                    {
                                        letterColorName = new LetterColorName() { Letter = "F", Color = "Green", Name = "Pass" };
                                    }
                                    else if ((GeoMean > 9) || (Median > 9) || (P90 > 29) || (PercOver43 > 7))
                                    {
                                        letterColorName = new LetterColorName() { Letter = "E", Color = "Green", Name = "Pass" };
                                    }
                                    else if ((GeoMean > 7) || (Median > 7) || (P90 > 22) || (PercOver43 > 5))
                                    {
                                        letterColorName = new LetterColorName() { Letter = "D", Color = "Green", Name = "Pass" };
                                    }
                                    else if ((GeoMean > 5) || (Median > 5) || (P90 > 14) || (PercOver43 > 3))
                                    {
                                        letterColorName = new LetterColorName() { Letter = "C", Color = "Green", Name = "Pass" };
                                    }
                                    else if ((GeoMean > 2) || (Median > 2) || (P90 > 7) || (PercOver43 > 2))
                                    {
                                        letterColorName = new LetterColorName() { Letter = "B", Color = "Green", Name = "Pass" };
                                    }
                                    else
                                    {
                                        letterColorName = new LetterColorName() { Letter = "A", Color = "Green", Name = "Pass" };
                                    }
                                }

                                mwqmSiteClassStat.LetterColorName = letterColorName;

                                sb2.AppendLine("<tr>");
                                sb2.AppendLine("<th>");
                                sb2.AppendLine($"{ locator }");
                                sb2.AppendLine("</th>");
                                sb2.AppendLine("<th>");
                                sb2.AppendLine($"{ name }");
                                sb2.AppendLine("</th>");
                                sb2.AppendLine("<th>");
                                sb2.AppendLine($"{ mwqmSampleFCList.Count }");
                                sb2.AppendLine("</th>");
                                sb2.AppendLine("<th>");
                                sb2.AppendLine($"{ MaxYear }-{ MinYear }");
                                sb2.AppendLine("</th>");
                                sb2.AppendLine("<th>");
                                sb2.AppendLine($"{ RemoveStart0(mwqmSiteClassStat.MWQMSite.MWQMSiteNumber) }");
                                sb2.AppendLine("</th>");
                                sb2.AppendLine("<th>");
                                sb2.AppendLine($"{ ((MWQMSiteLatestClassificationEnum)mwqmSiteClassStat.MWQMSite.MWQMSiteLatestClassification).ToString() }");
                                sb2.AppendLine("</th>");
                                sb2.AppendLine("<th>");
                                sb2.AppendLine($"{ mwqmSiteClassStat.LetterColorName.Color }");
                                sb2.AppendLine("</th>");
                                sb2.AppendLine("<th>");
                                sb2.AppendLine($"{ mwqmSiteClassStat.LetterColorName.Letter }");
                                sb2.AppendLine("</th>");
                                sb2.AppendLine("</tr>");

                            }
                            else
                            {
                                mwqmSiteClassStat.LetterColorName = new LetterColorName() { Color = "ND", Letter = "ND", Name = "ND" };

                                sb2.AppendLine("<tr>");
                                sb2.AppendLine("<th>");
                                sb2.AppendLine($"{ locator }");
                                sb2.AppendLine("</th>");
                                sb2.AppendLine("<th>");
                                sb2.AppendLine($"{ name }");
                                sb2.AppendLine("</th>");
                                sb2.AppendLine("<th>");
                                sb2.AppendLine("ND");
                                sb2.AppendLine("</th>");
                                sb2.AppendLine("<th>");
                                sb2.AppendLine("ND");
                                sb2.AppendLine("</th>");
                                sb2.AppendLine("<th>");
                                sb2.AppendLine("ND");
                                sb2.AppendLine("</th>");
                                sb2.AppendLine("<th>");
                                sb2.AppendLine("ND");
                                sb2.AppendLine("</th>");
                                sb2.AppendLine("<th>");
                                sb2.AppendLine("ND");
                                sb2.AppendLine("</th>");
                                sb2.AppendLine("<th>");
                                sb2.AppendLine("ND");
                                sb2.AppendLine("</th>");
                                sb2.AppendLine("</tr>");
                            }
                        }
                        else
                        {
                            mwqmSiteClassStat.LetterColorName = new LetterColorName() { Color = "ND", Letter = "ND", Name = "ND" };

                            sb2.AppendLine("<tr>");
                            sb2.AppendLine("<th>");
                            sb2.AppendLine($"{ locator }");
                            sb2.AppendLine("</th>");
                            sb2.AppendLine("<th>");
                            sb2.AppendLine($"{ name }");
                            sb2.AppendLine("</th>");
                            sb2.AppendLine("<th>");
                            sb2.AppendLine("ND");
                            sb2.AppendLine("</th>");
                            sb2.AppendLine("<th>");
                            sb2.AppendLine("ND");
                            sb2.AppendLine("</th>");
                            sb2.AppendLine("<th>");
                            sb2.AppendLine("ND");
                            sb2.AppendLine("</th>");
                            sb2.AppendLine("<th>");
                            sb2.AppendLine("ND");
                            sb2.AppendLine("</th>");
                            sb2.AppendLine("<th>");
                            sb2.AppendLine("ND");
                            sb2.AppendLine("</th>");
                            sb2.AppendLine("<th>");
                            sb2.AppendLine("ND");
                            sb2.AppendLine("</th>");
                            sb2.AppendLine("</tr>");
                        }

                        MWQMSiteClassStatList.Add(mwqmSiteClassStat);
                    }

                    sb.AppendLine("<tr>");
                    sb.AppendLine("<td>");
                    sb.Append($@"<a href=""http://wmon01dtchlebl2/csspwebtools/en-CA/#!View/a|||");
                    sb.Append($@"{ subsector.c.TVItemID }|||00000000001000000000000000000000"">{ locator }</a>");
                    sb.AppendLine("</td>");
                    sb.AppendLine("<td>");
                    sb.AppendLine($"{ name }");
                    sb.AppendLine("</td>");
                    sb.AppendLine("<td>");
                    sb.AppendLine($"{ MaxNumberOfSample }");
                    sb.AppendLine("</td>");
                    sb.AppendLine("<td>");
                    sb.AppendLine($"{ SampleMaxYear }-{ SampleMinYear }");
                    sb.AppendLine("</td>");
                    sb.AppendLine("<td>");
                    if (mwqmSampleList.Count > 3)
                    {
                        foreach (MWQMSiteClassStat mwqmSiteClassStat in MWQMSiteClassStatList)
                        {
                            if (mwqmSiteClassStat.MWQMSite.MWQMSiteLatestClassification == (int)MWQMSiteLatestClassificationEnum.Approved)
                            {
                                if (mwqmSiteClassStat.LetterColorName.Color == "Red")
                                {
                                    sb.Append($"{ RemoveStart0(mwqmSiteClassStat.MWQMSite.MWQMSiteNumber)},");
                                }
                            }
                        }
                    }
                    else
                    {
                        sb.AppendLine("ND");
                    }
                    sb.AppendLine("</td>");
                    sb.AppendLine("<td>");
                    if (mwqmSampleList.Count > 3)
                    {
                        foreach (MWQMSiteClassStat mwqmSiteClassStat in MWQMSiteClassStatList)
                        {
                            if (mwqmSiteClassStat.MWQMSite.MWQMSiteLatestClassification == (int)MWQMSiteLatestClassificationEnum.Approved)
                            {
                                if (mwqmSiteClassStat.LetterColorName.Color == "Green" && mwqmSiteClassStat.LetterColorName.Letter == "F")
                                {
                                    sb.Append($"{ RemoveStart0(mwqmSiteClassStat.MWQMSite.MWQMSiteNumber)},");
                                }
                            }
                        }
                    }
                    else
                    {
                        sb.AppendLine("ND");
                    }
                    sb.AppendLine("</td>");
                    sb.AppendLine("<td>");
                    if (mwqmSampleList.Count > 3)
                    {
                        foreach (MWQMSiteClassStat mwqmSiteClassStat in MWQMSiteClassStatList)
                        {
                            if (mwqmSiteClassStat.MWQMSite.MWQMSiteLatestClassification == (int)MWQMSiteLatestClassificationEnum.Restricted
                                || mwqmSiteClassStat.MWQMSite.MWQMSiteLatestClassification == (int)MWQMSiteLatestClassificationEnum.ConditionallyRestricted)
                            {
                                if (mwqmSiteClassStat.LetterColorName.Color == "Red" && mwqmSiteClassStat.LetterColorName.Letter == "F")
                                {
                                    sb.Append($"{ RemoveStart0(mwqmSiteClassStat.MWQMSite.MWQMSiteNumber)},");
                                }
                            }
                        }
                    }
                    else
                    {
                        sb.AppendLine("ND");
                    }
                    sb.AppendLine("</td>");
                    sb.AppendLine("<td>");
                    if (mwqmSampleList.Count > 3)
                    {
                        foreach (MWQMSiteClassStat mwqmSiteClassStat in MWQMSiteClassStatList)
                        {
                            if (mwqmSiteClassStat.MWQMSite.MWQMSiteLatestClassification == (int)MWQMSiteLatestClassificationEnum.Restricted
                                || mwqmSiteClassStat.MWQMSite.MWQMSiteLatestClassification == (int)MWQMSiteLatestClassificationEnum.ConditionallyRestricted)
                            {
                                if (mwqmSiteClassStat.LetterColorName.Color == "Purple")
                                {
                                    sb.Append($"{ RemoveStart0(mwqmSiteClassStat.MWQMSite.MWQMSiteNumber)},");
                                }
                            }
                        }
                    }
                    else
                    {
                        sb.AppendLine("ND");
                    }
                    sb.AppendLine("</td>");
                    sb.AppendLine("<td>");
                    if (mwqmSampleList.Count > 3)
                    {
                        foreach (MWQMSiteClassStat mwqmSiteClassStat in MWQMSiteClassStatList)
                        {
                            if (mwqmSiteClassStat.MWQMSite.MWQMSiteLatestClassification == (int)MWQMSiteLatestClassificationEnum.Restricted
                                || mwqmSiteClassStat.MWQMSite.MWQMSiteLatestClassification == (int)MWQMSiteLatestClassificationEnum.ConditionallyRestricted)
                            {
                                if (mwqmSiteClassStat.LetterColorName.Color == "Green" && !(mwqmSiteClassStat.LetterColorName.Letter == "F"))
                                {
                                    sb.Append($"{ RemoveStart0(mwqmSiteClassStat.MWQMSite.MWQMSiteNumber)},");
                                }
                            }
                        }
                    }
                    else
                    {
                        sb.AppendLine("ND");
                    }
                    sb.AppendLine("</td>");
                    sb.AppendLine("<td>");
                    foreach (MWQMSiteClassStat mwqmSiteClassStat in MWQMSiteClassStatList)
                    {
                        if (mwqmSiteClassStat.MWQMSite.MWQMSiteLatestClassification == (int)MWQMSiteLatestClassificationEnum.Error)
                        {
                            sb.Append($"{ RemoveStart0(mwqmSiteClassStat.MWQMSite.MWQMSiteNumber)},");
                        }
                    }
                    sb.AppendLine("</td>");
                    int countRainDay = 0;
                    int countRainDay24h = 0;
                    int countRainDay24hOver25mm = 0;
                    int countRainDay48h = 0;
                    bool rainDataMissing = false;

                    List<MWQMRun> MWQMRunList = new List<MWQMRun>();

                    foreach (MWQMSiteClassStat mwqmSiteClassStat in MWQMSiteClassStatList)
                    {
                        foreach (MWQMSample mwqmSample in mwqmSiteClassStat.MWQMSampleList)
                        {
                            MWQMRun mwqmRun = (from c in mwqmSiteClassStat.MWQMRunList
                                               where c.MWQMRunTVItemID == mwqmSample.MWQMRunTVItemID
                                               select c).FirstOrDefault();

                            if (mwqmRun != null)
                            {
                                if (!MWQMRunList.Where(c => c.MWQMRunTVItemID == mwqmRun.MWQMRunTVItemID).Any())
                                {
                                    MWQMRunList.Add(mwqmRun);
                                }
                            }
                        }
                    }

                    foreach (MWQMRun mwqmRun in MWQMRunList)
                    {
                        if (mwqmRun.RainDay0_mm == null || mwqmRun.RainDay1_mm == null || mwqmRun.RainDay2_mm == null)
                        {
                            rainDataMissing = true;
                        }

                        if (mwqmRun.RainDay0_mm >= 12)
                        {
                            countRainDay += 1;
                        }

                        if (mwqmRun.RainDay1_mm >= 12)
                        {
                            countRainDay24h += 1;
                        }

                        if (mwqmRun.RainDay1_mm >= 25)
                        {
                            countRainDay24hOver25mm += 1;
                        }

                        if (mwqmRun.RainDay2_mm >= 12)
                        {
                            countRainDay48h += 1;
                        }
                    }

                    sb.AppendLine("<td>");
                    sb.AppendLine($"{ countRainDay }");
                    sb.AppendLine("</td>");
                    sb.AppendLine("<td>");
                    sb.AppendLine($"{ countRainDay24h }");
                    sb.AppendLine("</td>");
                    sb.AppendLine("<td>");
                    sb.AppendLine($"{ countRainDay48h }");
                    sb.AppendLine("</td>");
                    sb.AppendLine("<td>");
                    sb.AppendLine($"{ countRainDay24hOver25mm }");
                    sb.AppendLine("</td>");
                    sb.AppendLine("<td>");
                    sb.AppendLine(rainDataMissing ? "Yes" : "No");
                    sb.AppendLine("</td>");
                    sb.AppendLine("<td>");
                    sb.AppendLine("");
                    sb.AppendLine("</td>");
                    sb.AppendLine("<td>");
                    sb.AppendLine("");
                    sb.AppendLine("</td>");
                    sb.AppendLine("<td>");
                    sb.AppendLine("");
                    sb.AppendLine("</td>");
                    sb.AppendLine("<td>");
                    sb.AppendLine("");
                    sb.AppendLine("</td>");
                    sb.AppendLine("<td>");
                    sb.AppendLine("");
                    sb.AppendLine("</td>");
                    sb.AppendLine("</tr>");

                }

                sb.AppendLine("</table>");
                sb.AppendLine("</body>");
                sb.AppendLine("</html>");

                //FileInfo fi = new FileInfo($@"C:\CSSP\{ provInit } Subsector Review.html");

                //StreamWriter sw = fi.CreateText();
                //sw.WriteLine(sb.ToString());
                //sw.Close();

                //FileInfo fiSite = new FileInfo($@"C:\CSSP\{ provInit } Subsector Review By Site.html");

                //StreamWriter swSite = fiSite.CreateText();
                //swSite.WriteLine(sb2.ToString());
                //swSite.Close();

                return true;
            }
        }
        private bool GenerateHTMLProvince_NotImplemented()
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
        public class MWQMSiteClassStat
        {
            public MWQMSite MWQMSite { get; set; }
            public List<MWQMSample> MWQMSampleList { get; set; }
            public List<MWQMRun> MWQMRunList { get; set; }
            public LetterColorName LetterColorName { get; set; }

            public MWQMSiteClassStat()
            {
                MWQMSampleList = new List<MWQMSample>();
                MWQMRunList = new List<MWQMRun>();
            }
        }
        public class LetterColorName
        {
            public string Letter { get; set; }
            public string Color { get; set; }
            public string Name { get; set; }
        }
        private string RemoveStart0(string TVText)
        {
            while (TVText.StartsWith("0"))
            {
                TVText = TVText.Substring(1);
            }

            return TVText;
        }

    }
}
