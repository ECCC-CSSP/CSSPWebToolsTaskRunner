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
using System.Diagnostics;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        private bool GenerateKMZ()
        {
            string NotUsed = "";
            //StringBuilder sbKMZ = new StringBuilder();
            //DateTime CD = DateTime.Now;

            //string Language = "_" + _TaskRunnerBaseService._BWObj.appTaskModel.Language;

            //string DateText = "_" + CD.Year.ToString() +
            //    "_" + (CD.Month > 9 ? CD.Month.ToString() : "0" + CD.Month.ToString()) +
            //    "_" + (CD.Day > 9 ? CD.Day.ToString() : "0" + CD.Day.ToString()) +
            //    "_" + (CD.Hour > 9 ? CD.Hour.ToString() : "0" + CD.Hour.ToString()) +
            //    "_" + (CD.Minute > 9 ? CD.Minute.ToString() : "0" + CD.Minute.ToString());

            switch (reportTypeModel.TVType)
            {
                case TVTypeEnum.MikeScenario:
                    break;
                case TVTypeEnum.Subsector:
                    break;
                default:
                    break;
            }

            //if (!RenameStartOfFileNameKMZ(fi))
            //{
            //    return false;
            //}

            fi = new FileInfo(fi.FullName.Replace(".html", ".kml"));

            if (fi.Exists)
            {
                try
                {
                    fi.Delete();
                    fi = new FileInfo(fi.FullName);
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDeleteFile_Error_, fi.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDeleteFile_Error_", fi.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return false;
                }
            }

            switch (reportTypeModel.TVType)
            {
                case TVTypeEnum.Root:
                    {
                        if (!GenerateKMZRoot())
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.Area:
                case TVTypeEnum.Country:
                case TVTypeEnum.Infrastructure:
                case TVTypeEnum.MikeScenario:
                    {
                        if (!GenerateKMZMikeScenario())
                        {
                            return false;
                        }
                    }
                    break;
                case TVTypeEnum.MikeSource:
                case TVTypeEnum.Municipality:
                case TVTypeEnum.MWQMSite:
                case TVTypeEnum.PolSourceSite:
                case TVTypeEnum.Province:
                case TVTypeEnum.Sector:
                case TVTypeEnum.Subsector:
                case TVTypeEnum.BoxModel:
                case TVTypeEnum.VisualPlumesScenario:
                    if (!GenerateKMZNotImplemented())
                    {
                        return false;
                    }
                    break;
                default:
                    break;
            }


            DirectoryInfo di = new DirectoryInfo(fi.DirectoryName);

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

            StreamWriter sw = fi.CreateText();
            sw.Write(sb.ToString());
            sw.Flush();
            sw.Close();

            fi = new FileInfo(fi.FullName);

            if (!fi.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fi.FullName);
                return false;
            }


            ProcessStartInfo pZip = new ProcessStartInfo();
            pZip.Arguments = "a -tzip \"" + fi.FullName.Replace(".kml", ".kmz") + "\" \"" + fi.FullName + "\"";
            pZip.RedirectStandardInput = true;
            pZip.UseShellExecute = false;
            pZip.CreateNoWindow = true;
            pZip.WindowStyle = ProcessWindowStyle.Hidden;

            Process processZip = new Process();
            processZip.StartInfo = pZip;
            try
            {
                pZip.FileName = @"C:\Program Files\7-Zip\7z.exe";
                processZip.Start();
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CompressKMLDidNotWorkWith7zError_, ex.Message);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CompressKMLDidNotWorkWith7zError_", ex.Message);
                return false;
            }

            while (!processZip.HasExited)
            {
                // waiting for the processZip to finish then continue
            }

            fi = new FileInfo(fi.FullName);

            if (fi.Exists)
            {
                try
                {
                    fi.Delete();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDeleteFile_Error_, fi.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDeleteFile_Error_", fi.FullName, ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return false;
                }
            }

            fi = new FileInfo(fi.FullName.Replace(".kml", ".kmz"));

            if (!fi.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fi.FullName);
                return false;
            }

            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemModel tvItemModel = tvItemService.PostAddChildTVItemDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi.Name.Replace(fi.Extension, ""), TVTypeEnum.File);
            if (!string.IsNullOrWhiteSpace(tvItemModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVItem, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVItem, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                return false;
            }

            TVFileModel tvFileModelNew = new TVFileModel();
            tvFileModelNew.TVFileTVItemID = tvItemModel.TVItemID;
            tvFileModelNew.TemplateTVType = 0;
            tvFileModelNew.ReportTypeID = reportTypeModel.ReportTypeID;
            tvFileModelNew.Parameters = Parameters;
            tvFileModelNew.ServerFileName = fi.Name;
            tvFileModelNew.FilePurpose = FilePurposeEnum.ReportGenerated;
            tvFileModelNew.Language = _TaskRunnerBaseService._BWObj.appTaskModel.Language;
            tvFileModelNew.Year = DateTime.Now.Year;
            tvFileModelNew.FileDescription = reportTypeModel.Description;
            tvFileModelNew.FileType = tvFileService.GetFileType(fi.Extension);
            tvFileModelNew.FileSize_kb = (((int)fi.Length / 1024) == 0 ? 1 : (int)fi.Length / 1024);
            tvFileModelNew.FileInfo = TaskRunnerServiceRes.FileName + "[" + fi.Name + "]\r\n" + TaskRunnerServiceRes.FileType + "[" + fi.Extension + "]\r\n";
            tvFileModelNew.FileCreatedDate_UTC = fi.LastWriteTimeUtc;
            tvFileModelNew.ServerFilePath = (fi.DirectoryName + @"\").Replace(@"C:\", @"E:\");
            tvFileModelNew.LastUpdateDate_UTC = DateTime.UtcNow;
            tvFileModelNew.LastUpdateContactTVItemID = _TaskRunnerBaseService._BWObj.appTaskModel.LastUpdateContactTVItemID;

            TVFileModel tvFileModelRet = tvFileService.PostAddTVFileDB(tvFileModelNew);
            if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                return false;
            }

            return true;
        }
        private bool GenerateKMZNotImplemented()
        {
            sb.AppendLine(@"<?xml version=""1.0"" encoding=""UTF-8""?>");
            sb.AppendLine(@"<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom"">");
            sb.AppendLine(@"<Document>");
            sb.AppendLine(@"	<name>" + fi.FullName + "</name> ");
            sb.AppendLine(@"	<Placemark>");
            sb.AppendLine(@"		<name>" + reportTypeModel.TVType.ToString() + " is not implemented</name> ");
            sb.AppendLine(@"		<Point>");
            sb.AppendLine(@"			<coordinates>-90,50,0</coordinates>");
            sb.AppendLine(@"		</Point> ");
            sb.AppendLine(@"	</Placemark>");
            sb.AppendLine(@"</Document> ");
            sb.AppendLine(@"</kml>");

            return true;
        }
        //private bool RenameStartOfFileNameKMZ()
        //{
        //    string NotUsed = "";
        //    switch (reportTypeModel.TVType)
        //    {
        //        case TVTypeEnum.MikeScenario:
        //            {
        //                string TVItemIDStr = "";
        //                int TVItemID = 0;
        //                string ContourValues = "";
        //                List<string> ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

        //                TVItemIDStr = GetParameters("TVItemID", ParamValueList);
        //                if (string.IsNullOrWhiteSpace(TVItemIDStr))
        //                {
        //                    NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.TVItemID);
        //                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.TVItemID);
        //                }

        //                int.TryParse(TVItemIDStr, out TVItemID);
        //                if (TVItemID == 0)
        //                {
        //                    NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.TVItemID);
        //                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.TVItemID);
        //                }

        //                ContourValues = GetParameters("ContourValues", ParamValueList);
        //                ContourValues = ContourValues.Trim().Replace(" ", "_");

        //                TVItemModel tvItemModelMikeScenario = _TVItemService.GetTVItemModelWithTVItemIDDB(TVItemID);
        //                if (!string.IsNullOrWhiteSpace(tvItemModelMikeScenario.Error))
        //                {
        //                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
        //                    return false;
        //                }
        //                string MikeScenarioName = tvItemModelMikeScenario.TVText;
        //                int pos = MikeScenarioName.IndexOf(" ");
        //                if (pos > 0)
        //                {
        //                    MikeScenarioName = MikeScenarioName.Trim();
        //                }
        //                reportTypeModel.StartOfFileName = reportTypeModel.StartOfFileName.Replace("{MikeScenarioName}", MikeScenarioName);

                    
        //                // it is possible that ContourValues parameter does not exist or is empty
        //                ContourValues = GetParameters("ContourValues", ParamValueList);
        //                ContourValues = ContourValues.Trim().Replace(" ", "_");

        //                if (!string.IsNullOrWhiteSpace(ContourValues))
        //                {

        //                    reportTypeModel.StartOfFileName = reportTypeModel.StartOfFileName.Replace("{ContourValues}", ContourValues);
        //                }
        //            }
        //            break;
        //        default:
        //            break;
        //    }

        //    return true;
        //}

    }
}
