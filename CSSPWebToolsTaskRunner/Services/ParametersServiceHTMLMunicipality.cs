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

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        public enum TideType
        {
            Low,
            High
        }
        public class Peaks
        {
            public DateTime Date { get; set; }
            public float Value { get; set; }
        }
        public enum Direction
        {
            Up,
            Down
        }

        private bool GenerateHTMLMunicipality()
        {

            switch (reportTypeModel.UniqueCode)
            {
                case "MunicipalityTestDocx":
                    {
                        if (!GenerateHTMLMunicipality_MunicipalityTestDocx())
                        {
                            return false;
                        }
                    }
                    break;
                case "MunicipalityHighestAndLowestMonthlyTidesXLSX":
                    {
                        if (!GenerateHTMLMunicipality_MunicipalityHighestAndLowestMonthlyTidesXLSX())
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
                    if (!GenerateHTMLMunicipality_NotImplemented())
                    {
                        return false;
                    }
                    break;
            }
            return true;
        }
        private bool GenerateHTMLMunicipality_MunicipalityTestDocx()
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
        private bool GenerateHTMLMunicipality_MunicipalityHighestAndLowestMonthlyTidesXLSX()
        {
            string NotUsed = "";
            bool ErrorInDoc = false;

            string WebTideDataSetText = "";
            string YearText = "";
            string LatText = "";
            string LngText = "";

            WebTideDataSetEnum webTideDataSet = WebTideDataSetEnum.Error;
            int Year = 0;
            double Lat = 0.0f;
            double Lng = 0.0f;

            List<string> ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            WebTideDataSetText = GetParameters("WebTideDataSet", ParamValueList);
            YearText = GetParameters("Year", ParamValueList);
            LatText = GetParameters("Lat", ParamValueList);
            LngText = GetParameters("Lng", ParamValueList);
            if (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr)
            {
                LatText = LatText.Replace(".", ",");
                LngText = LngText.Replace(".", ",");
            }

            if (string.IsNullOrWhiteSpace(WebTideDataSetText))
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, "Web Tide Data");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", "Web Tide Data");
                return false;
            }

            if (string.IsNullOrWhiteSpace(YearText))
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, "Year");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", "Year");
                return false;
            }

            if (string.IsNullOrWhiteSpace(LatText))
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, "Lat");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", "Lat");
                return false;
            }

            if (string.IsNullOrWhiteSpace(LngText))
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, "Lng");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", "Lng");
                return false;
            }

            int tempInt = 0;
            if (!(int.TryParse(WebTideDataSetText, out tempInt)))
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, "Web Tide Data");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", "Web Tide Data");
                return false;
            }

            webTideDataSet = (WebTideDataSetEnum)tempInt;

            if (!(int.TryParse(YearText, out Year)))
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, "Year");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", "Year");
                return false;
            }

            if (!(double.TryParse(LatText, out Lat)))
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, "Lat");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", "Lat");
                return false;
            }

            if (!(double.TryParse(LngText, out Lng)))
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, "Lng");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", "Lng");
                return false;
            }

            if (!GetTopHTML())
            {
                return false;
            }

            if (!WriteHTMLMunicipalityHighestAndLowestMonthlyTides(sb, webTideDataSet, Year, Lat, Lng))
            {
                ErrorInDoc = true;
            }

            if (!GetBottomHTML())
            {
                return false;
            }

            if (ErrorInDoc)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.ErrorOccuredWhileCreating_Document_, "XLSX", fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("ErrorOccuredWhileCreating_Document_", "XLSX", fi.FullName);
                return false;
            }
            return true;

        }
        private bool GenerateHTMLMunicipality_NotImplemented()
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

        private bool WriteHTMLMunicipalityHighestAndLowestMonthlyTides(StringBuilder sbHTML, WebTideDataSetEnum webTideDataSet, int Year, double Lat, double Lng)
        {
            int Days = 5;

            TideModel tideModel = new TideModel(_TVFileService.ChoseEDriveOrCDrive(_TVFileService.BasePath), webTideDataSet)
            {
                StartDate = new DateTime(Year, 1, 1),
                EndDate = new DateTime(Year + 1, 1, 1),
                Lat = Lat,
                Lng = Lng,
                Steps_min = 15,
                WebTideDataSet = webTideDataSet,
                DoWaterLevels = true,
                //TideDataPath --- already set in the constructor
            };

            List<WaterLevelResult> waterLevelResultList = _TidesAndCurrentsService.GetTides(tideModel);

            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
            {
                if (_TaskRunnerBaseService._BWObj.appTaskModel.Language == LanguageEnum.fr)
                {
                    sbHTML.AppendLine(_TaskRunnerBaseService._BWObj.TextLanguageList[1].Text);
                }
                else
                {
                    sbHTML.AppendLine(_TaskRunnerBaseService._BWObj.TextLanguageList[0].Text);
                }
                return true;
            }

            List<PeakDifference> HighPeakDifferenceList = FindMonthlyHighAndLowTide(Year, waterLevelResultList, TideType.High, Days);
            List<PeakDifference> LowPeakDifferenceList = FindMonthlyHighAndLowTide(Year, waterLevelResultList, TideType.Low, Days);

            List<DataPathOfTide> dataPathOfTideList = _TideSiteService.GetTideDataPathsDB();

            DataPathOfTide dataPathOfTide = dataPathOfTideList.Where(c => c.WebTideDataSet == webTideDataSet).FirstOrDefault();

            sbHTML.AppendLine($"<h2>{ TaskRunnerServiceRes.WebTideDataSet}: { dataPathOfTide.Text }</h2>");
            sbHTML.AppendLine($"<h2>{ TaskRunnerServiceRes.Location }: { Lat } { Lng }</h2>");
            sbHTML.AppendLine($"<h3>{ TaskRunnerServiceRes.AllHoursAreInUTC}</h3>");
            sbHTML.AppendLine("");
            sbHTML.AppendLine("");
            sbHTML.AppendLine($"<h2>{ TaskRunnerServiceRes.MonthlyHighTides }</h2>");
            sbHTML.AppendLine("");
            sbHTML.AppendLine(@"<table>");
            sbHTML.AppendLine(@"<thead>");
            sbHTML.AppendLine(@"<tr>");
            sbHTML.AppendLine($@"<th>{ TaskRunnerServiceRes.Monthly }</th>");
            sbHTML.AppendLine($@"<th>{ TaskRunnerServiceRes.StartTime }</th>");
            sbHTML.AppendLine($@"<th>{ TaskRunnerServiceRes.EndTime }</th>");
            sbHTML.AppendLine($@"<th>{ TaskRunnerServiceRes.Amplitude }</th>");
            sbHTML.AppendLine(@"<th></th>");
            sbHTML.AppendLine(@"<th></th>");
            sbHTML.AppendLine(@"</tr>");
            sbHTML.AppendLine(@"</thead>");
            sbHTML.AppendLine(@"<tbody>");

            foreach (PeakDifference peakDifference in HighPeakDifferenceList)
            {
                sbHTML.AppendLine(@"<tr>");
                sbHTML.AppendLine($@"<td>{ peakDifference.StartDate.ToString("MMMM") }</td>");
                sbHTML.AppendLine($@"<td>{ peakDifference.StartDate.ToString("yyyy-MMM-dd HH:mm:ss") }</td>");
                sbHTML.AppendLine($@"<td>{ peakDifference.EndDate.ToString("yyyy-MMM-dd HH:mm:ss") }</td>");
                sbHTML.AppendLine($@"<td>{ peakDifference.Value }</td>");
                sbHTML.AppendLine($@"<td></td>");
                sbHTML.AppendLine($@"<td></td>");

                sbHTML.AppendLine(@"</tr>");
            }

            sbHTML.AppendLine(@"</tbody>");
            sbHTML.AppendLine(@"</table>");

            sbHTML.AppendLine("");
            sbHTML.AppendLine("");
            sbHTML.AppendLine($"<h2>{ TaskRunnerServiceRes.MonthlyLowTides }</h2>");
            sbHTML.AppendLine("");
            sbHTML.AppendLine("");
            sbHTML.AppendLine(@"<table>");
            sbHTML.AppendLine(@"<thead>");
            sbHTML.AppendLine(@"<tr>");
            sbHTML.AppendLine($@"<th>{ TaskRunnerServiceRes.Monthly }</th>");
            sbHTML.AppendLine($@"<th>{ TaskRunnerServiceRes.StartTime }</th>");
            sbHTML.AppendLine($@"<th>{ TaskRunnerServiceRes.EndTime }</th>");
            sbHTML.AppendLine($@"<th>{ TaskRunnerServiceRes.Amplitude }</th>");
            sbHTML.AppendLine(@"</tr>");
            sbHTML.AppendLine(@"</thead>");
            sbHTML.AppendLine(@"<tbody>");

            foreach (PeakDifference peakDifference in LowPeakDifferenceList)
            {
                sbHTML.AppendLine(@"<tr>");
                sbHTML.AppendLine($@"<td>{ peakDifference.StartDate.ToString("MMMM") }</td>");
                sbHTML.AppendLine($@"<td>{ peakDifference.StartDate.ToString("yyyy-MMM-dd HH:mm:ss") }</td>");
                sbHTML.AppendLine($@"<td>{ peakDifference.EndDate.ToString("yyyy-MMM-dd HH:mm:ss") }</td>");
                sbHTML.AppendLine($@"<td>{ peakDifference.Value }</td>");

                sbHTML.AppendLine(@"</tr>");
            }

            sbHTML.AppendLine(@"</tbody>");
            sbHTML.AppendLine(@"</table>");

            sbHTML.AppendLine("");
            sbHTML.AppendLine("");
            sbHTML.AppendLine($"<h2>{ TaskRunnerServiceRes.WaterLevelsData}</h2>");
            sbHTML.AppendLine("");
            sbHTML.AppendLine("");
            sbHTML.AppendLine(@"<table>");
            sbHTML.AppendLine(@"<thead>");
            sbHTML.AppendLine(@"<tr>");
            sbHTML.AppendLine($@"<th>{ TaskRunnerServiceRes.Time }</th>");
            sbHTML.AppendLine($@"<th>{ TaskRunnerServiceRes.WaterLevel }</th>");
            sbHTML.AppendLine(@"</tr>");
            sbHTML.AppendLine(@"</thead>");
            sbHTML.AppendLine(@"<tbody>");

            foreach (WaterLevelResult waterLevelResult in waterLevelResultList)
            {
                sbHTML.AppendLine(@"<tr>");
                sbHTML.AppendLine($@"<td>{ waterLevelResult.Date.ToString("yyyy-MMM-dd HH:mm:ss") }</td>");
                sbHTML.AppendLine($@"<td>{ waterLevelResult.WaterLevel }</td>");

                sbHTML.AppendLine(@"</tr>");
            }

            sbHTML.AppendLine(@"</tbody>");
            sbHTML.AppendLine(@"</table>");

            return true;
        }

        private List<PeakDifference> FindMonthlyHighAndLowTide(int Year, List<WaterLevelResult> waterLevelResultList, TideType tideType, int Days)
        {

            DateTime StartUpDate = new DateTime(Year, 1, 1);

            List<Peaks> PeakList = new List<Peaks>();

            Direction direction = new Direction();

            if (waterLevelResultList[0].WaterLevel > waterLevelResultList[1].WaterLevel)
            {
                direction = Direction.Down;
            }
            else
            {
                direction = Direction.Up;
            }

            for (int i = 1; i < waterLevelResultList.Count - 2; i++)
            {
                if (waterLevelResultList[i].WaterLevel > waterLevelResultList[i + 1].WaterLevel)
                {
                    if (direction == Direction.Up)
                    {
                        PeakList.Add(new Peaks() { Date = waterLevelResultList[i].Date, Value = (float)waterLevelResultList[i].WaterLevel });
                        direction = Direction.Down;
                    }
                }
                else
                {
                    if (direction == Direction.Down)
                    {
                        PeakList.Add(new Peaks() { Date = waterLevelResultList[i].Date, Value = (float)waterLevelResultList[i].WaterLevel });
                        direction = Direction.Up;
                    }
                }
            }

            List<PeakDifference> PeakDiffList = new List<PeakDifference>();

            for (int i = 0; i < PeakList.Count - 1; i++)
            {
                PeakDiffList.Add(new PeakDifference() { StartDate = PeakList[i].Date, EndDate = PeakList[i + 1].Date, Value = Math.Abs(PeakList[i].Value - PeakList[i + 1].Value) });
            }

            float TempFloat = Days;
            float Hours = TempFloat * 24;
            int PeakDiffNumber = (int)(Hours / 6.2) == 0 ? 1 : (int)(Hours / 6.2);

            List<PeakDifference> MovingAverageOfPeakDiffList = new List<PeakDifference>();

            for (int i = 0; i < PeakDiffList.Count - PeakDiffNumber; i++)
            {
                float Average = 0;
                for (int j = 0; j < PeakDiffNumber; j++)
                {
                    Average += PeakDiffList[i + j].Value;
                }
                Average = Average / PeakDiffNumber;
                MovingAverageOfPeakDiffList.Add(new PeakDifference() { StartDate = PeakDiffList[i].StartDate, EndDate = PeakDiffList[i + PeakDiffNumber - 1].EndDate, Value = Average });
            }

            List<PeakDifference> MonthlyPeaks = new List<PeakDifference>();
            for (int i = 1; i < 13; i++)
            {
                switch (tideType)
                {
                    case TideType.Low:
                        {
                            PeakDifference peakDifference = (from pd in MovingAverageOfPeakDiffList
                                                             where pd.StartDate.Month == i
                                                             orderby pd.Value
                                                             select pd).FirstOrDefault<PeakDifference>();

                            MonthlyPeaks.Add(peakDifference);
                        }
                        break;
                    case TideType.High:
                        {
                            PeakDifference peakDifference = (from pd in MovingAverageOfPeakDiffList
                                                             where pd.StartDate.Month == i
                                                             orderby pd.Value descending
                                                             select pd).FirstOrDefault<PeakDifference>();

                            MonthlyPeaks.Add(peakDifference);
                        }
                        break;
                    default:
                        break;
                }
            }

            return MonthlyPeaks;
        }

    }
}
