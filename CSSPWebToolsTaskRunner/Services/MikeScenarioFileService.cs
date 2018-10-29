using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CSSPDBDLL.Models;
using CSSPDBDLL.Services;
using CSSPWebToolsTaskRunner.Services.Resources;
using CSSPWebToolsTaskRunner.Services;
using DHI.Generic.MikeZero.DFS.dfsu;
using System.IO;
using System.Transactions;
using System.Windows.Forms;
using System.Threading;
using DHI.Generic.MikeZero;
using DHI.Generic.MikeZero.DFS;
using System.Diagnostics;
using CSSPDBDLL;
using System.Security.Principal;
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;
using DHI.PFS;
using System.Net;
using System.Collections.Specialized;
using Newtonsoft.Json;
using DHI.Generic.MikeZero.DFS.mesh;

namespace CSSPWebToolsTaskRunner.Services
{
    public class MikeScenarioFileService
    {
        #region Variables
        #endregion Variables

        #region Properties
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        public TidesAndCurrentsService tidesAndCurrentsService { get; private set; }
        public KmzServiceMikeScenario kmzServiceMikeScenario { get; private set; }
        #endregion Properties

        #region Constructors
        public MikeScenarioFileService(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
            tidesAndCurrentsService = new TidesAndCurrentsService(taskRunnerBaseService);
            kmzServiceMikeScenario = new KmzServiceMikeScenario(taskRunnerBaseService);
        }
        #endregion Constructors

        #region Events
        #endregion Events

        #region Functions public
        public void CreateWebTideDataWLAtFirstNode()
        {
            MikeScenarioService mikeScenarioService = new MikeScenarioService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            string NotUsed = "";

            string Parameters = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters;
            string[] ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            int MikeScenarioTVItemID = 0;
            foreach (string s in ParamValueList)
            {
                string[] ParamValue = s.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                if (ParamValue.Length != 2)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, TaskRunnerServiceRes.Parameters);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", TaskRunnerServiceRes.Parameters);
                    return;
                }

                if (ParamValue[0] == "MikeScenarioTVItemID")
                {
                    MikeScenarioTVItemID = int.Parse(ParamValue[1]);
                }
                else
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind__, TaskRunnerServiceRes.Parameter, ParamValue[0]);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind__", TaskRunnerServiceRes.Parameter, ParamValue[0].ToString());
                    return;
                }
            }

            if (MikeScenarioTVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.MikeScenarioTVItemID, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                return;
            }

            NotUsed = TaskRunnerServiceRes.CreatingWebTideBoundaryConditionFiles;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("CreatingWebTideBoundaryConditionFiles"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 2);

            MikeScenarioService MikeScenarioService = new MikeScenarioService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MikeScenarioModel mikeScenarioModel = MikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(MikeScenarioTVItemID);
            if (!string.IsNullOrWhiteSpace(mikeScenarioModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                return;
            }

            MikeBoundaryConditionService mikeBoundaryConditionService = new MikeBoundaryConditionService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            List<MikeBoundaryConditionModel> mbcModelList = mikeBoundaryConditionService.GetMikeBoundaryConditionModelListWithMikeScenarioTVItemIDAndTVTypeDB(mikeScenarioModel.MikeScenarioTVItemID, TVTypeEnum.MikeBoundaryConditionWebTide);
            if (mbcModelList.Count == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_StartingWith_, TaskRunnerServiceRes.MIKEBoundaryCondition, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_StartingWith_", TaskRunnerServiceRes.MIKEBoundaryCondition, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                return;
            }

            MikeBoundaryConditionModel mbcm = mbcModelList.Where(c => c.MikeBoundaryConditionLevelOrVelocity == MikeBoundaryConditionLevelOrVelocityEnum.Level && c.TVType == TVTypeEnum.MikeBoundaryConditionWebTide).FirstOrDefault();
            if (mbcm == null)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MIKEBoundaryCondition, TaskRunnerServiceRes.MikeBoundaryConditionLevelOrVelocity + "," + TaskRunnerServiceRes.TVType, MikeBoundaryConditionLevelOrVelocityEnum.Level.ToString() + "," + TVTypeEnum.MikeBoundaryConditionWebTide.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MIKEBoundaryCondition, TaskRunnerServiceRes.MikeBoundaryConditionLevelOrVelocity + "," + TaskRunnerServiceRes.TVType, MikeBoundaryConditionLevelOrVelocityEnum.Level.ToString() + "," + TVTypeEnum.MikeBoundaryConditionWebTide.ToString());
                return;
            }

            MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            List<MapInfoPointModel> mapInfoPointList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(mbcm.MikeBoundaryConditionTVItemID, TVTypeEnum.MikeBoundaryConditionWebTide, MapInfoDrawTypeEnum.Polyline);
            if (mapInfoPointList.Count > 0)
            {
                List<IEnumerable<WaterLevelResult>> AllWLResults = new List<IEnumerable<WaterLevelResult>>();
                IEnumerable<WaterLevelResult> WLResults = null;

                GoogleTimeZoneJSON googleTimeZoneJSON = new GoogleTimeZoneJSON();
                using (WebClient webClient = new WebClient() { Encoding = Encoding.UTF8 })
                {
                    WebProxy webProxy = new WebProxy();
                    webClient.Proxy = webProxy;

                    webClient.UseDefaultCredentials = true;

                    var json_data = string.Empty;
                    // attempt to download JSON data as a string
                    try
                    {
                        DateTime epoch = new DateTime(1970, 1, 1, 0, 0, 0, 0).ToLocalTime();
                        TimeSpan span = (mikeScenarioModel.MikeScenarioStartDateTime_Local - epoch);
                        string timeStamp = ((long)Convert.ToDouble(span.TotalSeconds)).ToString();
                        string url = "https://maps.googleapis.com/maps/api/timezone/json?location=" +
                            mapInfoPointList[0].Lat.ToString("F5") + "," + mapInfoPointList[0].Lng.ToString("F5") + "&timestamp=" + timeStamp + "&key=AIzaSyAwPGpdSM6z0A7DFdWPbS3vIDTk2mxINaA";
                        byte[] responseBytes = webClient.DownloadData(url);
                        json_data = Encoding.UTF8.GetString(responseBytes);
                    }
                    catch (Exception)
                    {
                        json_data = "Error";
                    }
                    JsonSerializerSettings jsonSerializerSettings = new JsonSerializerSettings
                    {
                        DateTimeZoneHandling = DateTimeZoneHandling.Local
                    };
                    googleTimeZoneJSON = JsonConvert.DeserializeObject<GoogleTimeZoneJSON>(json_data, jsonSerializerSettings);
                }

                float offset = (float)((googleTimeZoneJSON.rawOffset + googleTimeZoneJSON.dstOffset) / 3600.0f);

                string offsetText = offset.ToString("F1");
                if (offsetText.EndsWith("0"))
                {
                    offsetText = offsetText.Substring(0, offsetText.Length - 2);
                }

                TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                TideModel tideModel = new TideModel(tvFileService.ChoseEDriveOrCDrive(tvFileService.BasePath), mbcm.WebTideDataSet)
                {
                    StartDate = mikeScenarioModel.MikeScenarioStartDateTime_Local.AddHours(-offset - 24),
                    EndDate = mikeScenarioModel.MikeScenarioEndDateTime_Local.AddHours(-offset),
                    Lng = (double)mapInfoPointList[0].Lng,
                    Lat = (double)mapInfoPointList[0].Lat,
                    Steps_min = 60,
                    DoWaterLevels = true,
                };
                WLResults = tidesAndCurrentsService.GetTides(tideModel);
                if (WLResults != null)
                {
                    foreach (WaterLevelResult wlr in WLResults)
                    {
                        wlr.Date = wlr.Date.AddHours(offset);
                    }
                    AllWLResults.Add(WLResults);
                }

                StringBuilder sb = new StringBuilder();
                sb.AppendLine(TaskRunnerServiceRes.TimeOffsetFromGMT + ": " + offsetText + " " + TaskRunnerServiceRes.Hours + " --- " + googleTimeZoneJSON.timeZoneName);
                sb.AppendLine(TaskRunnerServiceRes.From + " " + mikeScenarioModel.MikeScenarioStartDateTime_Local.ToString("yyyy MMMM dd HH:mm") +
                    " " + TaskRunnerServiceRes.To + " " + mikeScenarioModel.MikeScenarioEndDateTime_Local.ToString("yyyy MMMM dd HH:mm") + " (" +
                    TaskRunnerServiceRes.LocalTime + ")");
                sb.AppendLine(TaskRunnerServiceRes.AtCoordinates + " " + tideModel.Lat.ToString("F5") + " " + tideModel.Lng.ToString("F5"));
                sb.AppendLine(TaskRunnerServiceRes.Showing24HoursBeforeStartDateJustInCaseYouNeedToSetADifferentStartDate);
                sb.AppendLine("");
                sb.AppendLine(TaskRunnerServiceRes.DateAndTime + "\t\t" + TaskRunnerServiceRes.WaterLevel);
                foreach (List<WaterLevelResult> waterLevelResultList in AllWLResults)
                {
                    foreach (WaterLevelResult waterLevelResult in waterLevelResultList)
                    {
                        sb.AppendLine(waterLevelResult.Date.ToString("yyyy MMMM dd HH:mm") + "\t" + waterLevelResult.WaterLevel.ToString("F4"));
                    }
                }
                mbcm.WebTideDataFromStartToEndDate = sb.ToString();

                MikeBoundaryConditionModel mikeBoundaryConditionModelRet = mikeBoundaryConditionService.PostUpdateMikeBoundaryConditionDB(mbcm);
                if (!string.IsNullOrWhiteSpace(mikeBoundaryConditionModelRet.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.MIKEBoundaryCondition, mikeBoundaryConditionModelRet.Error);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.MIKEBoundaryCondition, mikeBoundaryConditionModelRet.Error);
                    return;
                }
            }
        }
        public void MikeScenarioImportDB()
        {
            string NotUsed = "";

            int TVItemID = 0;
            string UploadClientPath = "";

            string Parameters = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters;
            string[] ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

            if (_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.TVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.TVItemID);
            }

            if (ParamValueList.Count() != 2)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.ParameterCount_NotEqual_, ParamValueList.Count().ToString(), "4");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("ParameterCount_NotEqual_", ParamValueList.Count().ToString(), "4");
            }

            foreach (string s in ParamValueList)
            {
                string[] ParamValue = s.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                if (ParamValue.Length != 2)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, TaskRunnerServiceRes.Parameters);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", TaskRunnerServiceRes.Parameters);
                    return;
                }

                if (ParamValue[0] == "TVItemID")
                {
                    TVItemID = int.Parse(ParamValue[1]);
                }
                else if (ParamValue[0] == "UploadClientPath")
                {
                    UploadClientPath = ParamValue[1];
                }
                else
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, ParamValue[0]);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", ParamValue[0].ToString());
                    return;
                }
            }

            if (TVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBeNullOrEmpty, TaskRunnerServiceRes.TVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_ShouldNotBeNullOrEmpty", TaskRunnerServiceRes.TVItemID);
                return;
            }

            if (string.IsNullOrEmpty(UploadClientPath))
            {
                NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBeNullOrEmpty, TaskRunnerServiceRes.UploadClientPath);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_ShouldNotBeNullOrEmpty", TaskRunnerServiceRes.UploadClientPath);
                return;
            }

            // update percentage
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 10);

            MikeScenarioModel mikeScenarioModel = new MikeScenarioModel();
            List<OtherFileInfo> otherFileInfoList = new List<OtherFileInfo>();

            TVItemModel tvItemModelMikeScenario = new TVItemModel();
            //using (TransactionScope ts = new TransactionScope())
            //{
            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            tvItemModelMikeScenario = tvItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelMikeScenario.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                return;
            }

            int TVLevel = tvItemModelMikeScenario.TVLevel + 1;

            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVFileModel tvFileModelM21_3FM = tvFileService.GetTVFileModelWithTVItemIDAndTVFileTypeM21FMOrM3FMDB(tvItemModelMikeScenario.TVItemID);
            if (tvFileModelM21_3FM == null)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.TVItemID + "," + TaskRunnerServiceRes.FileType, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID + " .m21fm or .m3fm");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.TVItemID + "," + TaskRunnerServiceRes.FileType, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID + " .m21fm or .m3fm");
                return;
            }

            MikeScenarioService mikeScenarioService = new MikeScenarioService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            mikeScenarioModel = mikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(mikeScenarioModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                return;
            }

            string ServerFileName = tvFileModelM21_3FM.ServerFileName;
            string ServerFilePath = tvFileService.GetServerFilePath(mikeScenarioModel.MikeScenarioTVItemID);

            List<string> FileNameList = new List<string>();

            FileInfo fiServer = new FileInfo(ServerFilePath + ServerFileName);

            if (!fiServer.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiServer.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiServer.FullName);
                return;
            }

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 15);

            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotOpenFile_, fiServer.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotOpenFile_", fiServer.FullName);
                return;
            }

            PFSFile pfsFile = new PFSFile(fiServer.FullName);

            if (!IsEachSourceContinuous(pfsFile))
            {
                pfsFile.Close();
                return;
            }

            if (!IsTransportModuleDecayConstant(pfsFile))
            {
                pfsFile.Close();
                return;
            }

            // getting the log file
            string ClientFullFileName = UploadClientPath + tvFileModelM21_3FM.ServerFileName;
            string LogFileName = ClientFullFileName.Substring(0, ClientFullFileName.LastIndexOf(".")) + ".log";
            string ServerLogFileName = ServerFilePath + ServerFileName.Substring(0, ServerFileName.LastIndexOf(".")) + ".log";

            string m21fmServerFileName = ServerFilePath + ServerFileName;


            // updating rest of new MikeScenario
            DateTime? dateTimeTemp = GetParameterStartTime(pfsFile, "FemEngineHD/TIME", "start_time");
            if (dateTimeTemp == null)
            {
                pfsFile.Close();
                return;
            }
            mikeScenarioModel.MikeScenarioStartDateTime_Local = (DateTime)dateTimeTemp;
            int? TimeStepInterval = GetParameterInt(pfsFile, "FemEngineHD/TIME", "time_step_interval");
            if (TimeStepInterval == null)
            {
                pfsFile.Close();
                return;
            }
            int? NumberOfTimeSteps = GetParameterInt(pfsFile, "FemEngineHD/TIME", "number_of_time_steps");
            if (NumberOfTimeSteps == null)
            {
                pfsFile.Close();
                return;
            }
            mikeScenarioModel.MikeScenarioEndDateTime_Local = ((DateTime)dateTimeTemp).AddSeconds((int)TimeStepInterval * (int)NumberOfTimeSteps);
            //mikeScenarioModel.UseWebTide = false;

            TVItemModel tvItemModelMikeScenarioRet = tvItemService.GetTVItemModelWithTVItemIDDB(mikeScenarioModel.MikeScenarioTVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelMikeScenarioRet.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.MikeScenarioTVItemID, mikeScenarioModel.MikeScenarioTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.MikeScenarioTVItemID, mikeScenarioModel.MikeScenarioTVItemID.ToString());
                pfsFile.Close();
                return;
            }

            MikeScenarioModel mikeScenarioModelRet = mikeScenarioService.PostUpdateMikeScenarioDB(mikeScenarioModel);
            if (!string.IsNullOrWhiteSpace(mikeScenarioModelRet.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.MikeScenario, mikeScenarioModelRet.Error);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.MikeScenario, mikeScenarioModelRet.Error);
                pfsFile.Close();
                return;
            }

            otherFileInfoList.Add(new OtherFileInfo() { TVFileID = 0, ClientFullFileName = LogFileName, ServerFullFileName = ServerLogFileName, IsOutput = true });

            // update percentage
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 30);


            //__________________________________________________
            // Filling the FileNameList with files to be downloaded
            //__________________________________________________
            GetAllInputFilesToUpload(ServerFilePath, ClientFullFileName, otherFileInfoList, mikeScenarioModel, pfsFile);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.No_FileFoundInM21Or3fmFile, TaskRunnerServiceRes.Input);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("No_FileFoundInM21Or3fmFile", TaskRunnerServiceRes.Input);
                return;
            }

            // should only have 1 output file for hydrodynamic and 1 output file for transport
            int? NumberOfOutputs = GetParameterInt(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/OUTPUTS", "number_of_outputs");

            if (NumberOfOutputs == null)
            {
                NotUsed = TaskRunnerServiceRes.OnlyOneHydrodynamicAndTransportOutputIsAllowed;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("OnlyOneHydrodynamicAndTransportOutputIsAllowed");
                pfsFile.Close();
                return;
            }
            if (NumberOfOutputs != 1)
            {
                NotUsed = TaskRunnerServiceRes.OnlyOneHydrodynamicAndTransportOutputIsAllowed;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("OnlyOneHydrodynamicAndTransportOutputIsAllowed");
                pfsFile.Close();
                return;
            }

            //__________________________________________________
            // Filling the FileNameList with files to be downloaded
            //__________________________________________________
            //FileNameList.Clear();
            GetAllResultFilesToUpload(ServerFilePath, ClientFullFileName, otherFileInfoList, mikeScenarioModel, pfsFile);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.No_FileFoundInM21Or3fmFile, TaskRunnerServiceRes.Output);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("No_FileFoundInM21Or3fmFile", TaskRunnerServiceRes.Output);
                return;
            }

            // _______________________________________
            // Creating MikeParameters associated with the scenario
            // _______________________________________

            float? WindConstantSpeed = (float?)GetParameterDouble(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/WIND_FORCING", "constant_speed");
            if (WindConstantSpeed == null)
            {
                pfsFile.Close();
                return;
            }

            float? WindConstantDirection = (float?)GetParameterDouble(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/WIND_FORCING", "constant_direction");
            if (WindConstantDirection == null)
            {
                pfsFile.Close();
                return;
            }
            float? DecayValue = (float?)GetParameterDouble(pfsFile, "FemEngineHD/TRANSPORT_MODULE/DECAY/COMPONENT_1", "constant_value");
            if (DecayValue == null)
            {
                pfsFile.Close();
                return;
            }

            mikeScenarioModel.WindSpeed_km_h = (float)WindConstantSpeed * 3.6;
            mikeScenarioModel.WindDirection_deg = (float)WindConstantDirection;
            mikeScenarioModel.DecayFactor_per_day = (float)DecayValue * 3600 * 24;
            mikeScenarioModel.DecayIsConstant = true;

            int? TimeStepFrequency = GetParameterInt(pfsFile, "FemEngineHD/TRANSPORT_MODULE/OUTPUTS/OUTPUT_1", "time_step_frequency");
            if (TimeStepFrequency == null)
            {
                pfsFile.Close();
                return;
            }
            mikeScenarioModel.ResultFrequency_min = (int)TimeStepFrequency;

            float? AmbientTemperature = (float?)GetParameterDouble(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/DENSITY", "temperature_reference");
            if (AmbientTemperature == null)
            {
                pfsFile.Close();
                return;
            }
            mikeScenarioModel.AmbientTemperature_C = (float)AmbientTemperature;

            float? AmbientSalinity = (float?)GetParameterDouble(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/DENSITY", "salinity_reference");
            if (AmbientSalinity == null)
            {
                pfsFile.Close();
                return;
            }
            mikeScenarioModel.AmbientSalinity_PSU = (float)AmbientSalinity;

            float? ManningNumber = (float?)GetParameterDouble(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/BED_RESISTANCE/MANNING_NUMBER", "constant_value");
            if (ManningNumber == null)
            {
                pfsFile.Close();
                return;
            }
            mikeScenarioModel.ManningNumber = (float)ManningNumber;

            TVItemModel tvItemModelMikeScenarioRet3 = tvItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelMikeScenarioRet3.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                pfsFile.Close();
                return;
            }

            TVItemModel tvItemModelMikeScenarioParent = tvItemService.GetTVItemModelWithTVItemIDDB((int)tvItemModelMikeScenario.ParentID);
            if (!string.IsNullOrWhiteSpace(tvItemModelMikeScenarioParent.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, tvItemModelMikeScenario.ParentID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, tvItemModelMikeScenario.ParentID.ToString()));
                pfsFile.Close();
                return;
            }

            // update percentage
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 60);

            mikeScenarioModel.DecayFactorAmplitude = mikeScenarioModel.DecayFactor_per_day - (mikeScenarioModel.DecayFactor_per_day * 0.03);
            MikeScenarioModel mikeScenarioModelRet3 = mikeScenarioService.PostUpdateMikeScenarioDB(mikeScenarioModel);
            if (!string.IsNullOrWhiteSpace(mikeScenarioModelRet3.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.MikeScenario, mikeScenarioModelRet3.Error.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.MikeScenario, mikeScenarioModelRet3.Error.ToString());
                pfsFile.Close();
                return;
            }

            // _______________________________________
            // creating all the MikeSources associated with the scenario
            // _______________________________________
            for (int i = 1; i < 1000; i++)
            {
                PFSSection pfsSectionSource = pfsFile.GetSectionFromHandle("FemEngineHD/HYDRODYNAMIC_MODULE/SOURCES/SOURCE_" + i.ToString());

                if (pfsSectionSource == null)
                {
                    break;
                }

                string SourceName = GetParameterString(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/SOURCES/SOURCE_" + i.ToString(), "Name");
                TVItemModel tvItemMikeSource = tvItemService.PostAddChildTVItemDB(tvItemModelMikeScenario.TVItemID, SourceName, TVTypeEnum.MikeSource);

                MikeSourceModel mikeSourceModelNew = new MikeSourceModel();

                mikeSourceModelNew.SourceNumberString = "SOURCE_" + i.ToString();
                mikeSourceModelNew.MikeSourceTVItemID = tvItemMikeSource.TVItemID;

                int? SourceIncluded = (int?)GetParameterInt(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/SOURCES/SOURCE_" + i.ToString(), "include");
                if (SourceIncluded == null)
                {
                    pfsFile.Close();
                    return;
                }

                mikeSourceModelNew.Include = SourceIncluded == 1 ? true : false;
                if (SourceName.ToUpper().Contains("RIV") || SourceName.ToUpper().Contains("BROOK") || SourceName.ToUpper().Contains("STREAM"))
                {
                    mikeSourceModelNew.IsRiver = true;
                }
                else
                {
                    mikeSourceModelNew.IsRiver = false;
                }
                mikeSourceModelNew.UseHydrometric = false;
                mikeSourceModelNew.HydrometricTVItemID = null;
                mikeSourceModelNew.DrainageArea_km2 = null;
                mikeSourceModelNew.Factor = null;

                mikeSourceModelNew.MikeSourceTVText = SourceName;

                MikeSourceStartEndModel mikeSourceStartEndModelNew = new MikeSourceStartEndModel();

                float? SourceFlow = (float?)GetParameterDouble(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/SOURCES/SOURCE_" + i.ToString(), "constant_value");
                if (SourceFlow == null)
                {
                    pfsFile.Close();
                    return;
                }

                mikeSourceStartEndModelNew.SourceFlowStart_m3_day = (float)SourceFlow * 24 * 3600;
                mikeSourceStartEndModelNew.SourceFlowEnd_m3_day = (float)SourceFlow * 24 * 3600;

                Coord SourceCoord = GetParameterCoord(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/SOURCES/SOURCE_" + i.ToString(), "coordinates");
                if (SourceCoord.Lat == 0.0f || SourceCoord.Lng == 0.0f)
                {
                    pfsFile.Close();
                    return;
                }

                List<Coord> coordList = new List<Coord>()
                    {
                        SourceCoord,
                    };

                PFSSection pfsSectionSourceTransport = pfsFile.GetSectionFromHandle("FemEngineHD/TRANSPORT_MODULE/SOURCES/SOURCE_" + i.ToString() + "/COMPONENT_1");

                if (pfsSectionSourceTransport == null)
                {
                    pfsFile.Close();
                    break;
                }

                float? SourcePollution = (float?)GetParameterDouble(pfsFile, "FemEngineHD/TRANSPORT_MODULE/SOURCES/SOURCE_" + i.ToString() + "/COMPONENT_1", "constant_value");
                if (SourcePollution == null)
                {
                    pfsFile.Close();
                    return;
                }
                mikeSourceStartEndModelNew.SourcePollutionStart_MPN_100ml = (int)SourcePollution;
                mikeSourceStartEndModelNew.SourcePollutionEnd_MPN_100ml = (int)SourcePollution;

                int? SourcePollutionContinuous = GetParameterInt(pfsFile, "FemEngineHD/TRANSPORT_MODULE/SOURCES/SOURCE_" + i.ToString() + "/COMPONENT_1", "format");
                if (SourcePollutionContinuous == null)
                {
                    pfsFile.Close();
                    return;
                }
                mikeSourceModelNew.IsContinuous = SourcePollutionContinuous == 0 ? true : false;

                mikeSourceStartEndModelNew.StartDateAndTime_Local = mikeScenarioModel.MikeScenarioStartDateTime_Local;
                mikeSourceStartEndModelNew.EndDateAndTime_Local = mikeScenarioModel.MikeScenarioEndDateTime_Local;

                MikeSourceService mikeSourceService = new MikeSourceService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                MikeSourceModel mikeSourceModelRet = mikeSourceService.PostAddMikeSourceDB(mikeSourceModelNew);
                if (!string.IsNullOrWhiteSpace(mikeSourceModelRet.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.MikeSource, mikeScenarioModelRet3.Error.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.MikeSource, mikeScenarioModelRet3.Error.ToString());
                    pfsFile.Close();
                    return;
                }

                MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                MapInfoModel mapInfoModel = mapInfoService.CreateMapInfoObjectDB(coordList, MapInfoDrawTypeEnum.Point, TVTypeEnum.MikeSource, tvItemMikeSource.TVItemID);
                if (!string.IsNullOrWhiteSpace(mapInfoModel.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_For_With_Equal_, TaskRunnerServiceRes.MapInfo, TaskRunnerServiceRes.MikeSource, TaskRunnerServiceRes.TVItemID, tvItemMikeSource.TVItemID.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat4List("CouldNotCreate_For_With_Equal_", TaskRunnerServiceRes.MapInfo, TaskRunnerServiceRes.MikeSource, TaskRunnerServiceRes.TVItemID, tvItemMikeSource.TVItemID.ToString());
                    pfsFile.Close();
                    return;
                }

                PFSSection pfsSectionSourceTemperature = pfsFile.GetSectionFromHandle("FemEngineHD/HYDRODYNAMIC_MODULE/TEMPERATURE_SALINITY_MODULE/SOURCES/SOURCE_" + i.ToString() + "/TEMPERATURE");

                if (pfsSectionSourceTemperature == null)
                {
                    break;
                }

                float? SourceTemperature = (float?)GetParameterDouble(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/TEMPERATURE_SALINITY_MODULE/SOURCES/SOURCE_" + i.ToString() + "/TEMPERATURE", "constant_value");
                if (SourceTemperature == null)
                {
                    pfsFile.Close();
                    return;
                }
                mikeSourceStartEndModelNew.SourceTemperatureStart_C = (float)SourceTemperature;
                mikeSourceStartEndModelNew.SourceTemperatureEnd_C = (float)SourceTemperature;

                PFSSection pfsSectionSourceSalinity = pfsFile.GetSectionFromHandle("FemEngineHD/HYDRODYNAMIC_MODULE/TEMPERATURE_SALINITY_MODULE/SOURCES/SOURCE_" + i.ToString() + "/Salinity");

                if (pfsSectionSourceSalinity == null)
                {
                    break;
                }

                float? SourceSalinity = (float?)GetParameterDouble(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/TEMPERATURE_SALINITY_MODULE/SOURCES/SOURCE_" + i.ToString() + "/Salinity", "constant_value");
                if (SourceSalinity == null)
                {
                    pfsFile.Close();
                    return;
                }
                mikeSourceStartEndModelNew.SourceSalinityStart_PSU = (float)SourceSalinity;
                mikeSourceStartEndModelNew.SourceSalinityEnd_PSU = (float)SourceSalinity;

                mikeSourceStartEndModelNew.MikeSourceID = mikeSourceModelRet.MikeSourceID;

                MikeSourceStartEndService mikeSourceStartEndService = new MikeSourceStartEndService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                MikeSourceStartEndModel mikeSourceStartEndModelRet = mikeSourceStartEndService.PostAddMikeSourceStartEndDB(mikeSourceStartEndModelNew);
                if (!string.IsNullOrWhiteSpace(mikeSourceStartEndModelRet.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.MikeSourceStartEnd, mikeSourceStartEndModelRet.Error.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.MikeSourceStartEnd, mikeSourceStartEndModelRet.Error.ToString());
                    pfsFile.Close();
                    return;
                }

            }

            // update percentage
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 80);

            foreach (OtherFileInfo otherFileInfo in otherFileInfoList)
            {
                FileInfo fi = new FileInfo(otherFileInfo.ServerFullFileName);

                TVItemModel tvItemFileModel = tvItemService.PostAddChildTVItemDB(tvItemModelMikeScenario.TVItemID, fi.Name, TVTypeEnum.File);
                if (!string.IsNullOrEmpty(tvItemFileModel.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_For_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.MIKEScenarioDocumentation, TaskRunnerServiceRes.TVText, fi.Name);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat4List("CouldNotCreate_For_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.MIKEScenarioDocumentation, TaskRunnerServiceRes.TVText, fi.Name);
                    pfsFile.Close();
                    return;
                }

                TVFileModel tvFileModelNew = new TVFileModel();
                tvFileModelNew.TVFileTVItemID = tvItemFileModel.TVItemID;

                if (otherFileInfo.IsOutput)
                {
                    tvFileModelNew.FilePurpose = FilePurposeEnum.MikeResultDFSU;
                }
                else
                {
                    tvFileModelNew.FilePurpose = FilePurposeEnum.MikeInput;
                }

                tvFileModelNew.Language = _TaskRunnerBaseService._BWObj.appTaskModel.Language;
                tvFileModelNew.Year = DateTime.Now.Year;
                tvFileModelNew.FileDescription = ""; // nothing for now
                tvFileModelNew.FileType = tvFileService.GetFileType(fi.Extension);
                tvFileModelNew.FileSize_kb = 0;
                tvFileModelNew.FileInfo = "Mike Scenario File";
                tvFileModelNew.FileDescription = "Mike Scenario File";
                tvFileModelNew.FileCreatedDate_UTC = DateTime.UtcNow;
                tvFileModelNew.ClientFilePath = otherFileInfo.ClientFullFileName.Substring(0, otherFileInfo.ClientFullFileName.LastIndexOf(@"\") + 1);
                tvFileModelNew.ServerFileName = fi.Name;
                tvFileModelNew.ServerFilePath = ServerFilePath.Replace(@"C:\", @"E:\");

                TVFileModel tvFileModelRet = tvFileService.PostAddTVFileDB(tvFileModelNew);
                if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVFile, tvFileModelRet.Error.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVFile, tvFileModelRet.Error.ToString());
                    pfsFile.Close();
                    return;
                }

                otherFileInfo.TVFileID = tvFileModelNew.TVFileID;

            }

            try
            {
                pfsFile.Write(fiServer.FullName);
            }
            catch (Exception)
            {
                // nothing
            }
            FixPFSFileSystemPart(fiServer.FullName);

            //   ts.Complete();
            //}

            // update percentage
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 99);

            return;
        }
        public void MikeScenarioOtherFileImportDB()
        {
            string NotUsed = "";

            string Parameters = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters;
            string[] ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            int MikeScenarioTVItemID = 0;
            int TVFileTVItemID = 0;
            foreach (string s in ParamValueList)
            {
                string[] ParamValue = s.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                if (ParamValue.Length != 2)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, TaskRunnerServiceRes.Parameters);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", TaskRunnerServiceRes.Parameters);
                    return;
                }

                if (ParamValue[0] == "MikeScenarioTVItemID")
                {
                    MikeScenarioTVItemID = int.Parse(ParamValue[1]);
                }
                else if (ParamValue[0] == "TVFileTVItemID")
                {
                    TVFileTVItemID = int.Parse(ParamValue[1]);
                }
                else
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, ParamValue[0]);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", ParamValue[0].ToString());
                    return;
                }
            }

            if (MikeScenarioTVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBeNullOrEmpty, TaskRunnerServiceRes.MikeScenarioTVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_ShouldNotBeNullOrEmpty", TaskRunnerServiceRes.MikeScenarioTVItemID);
                return;
            }

            if (TVFileTVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBeNullOrEmpty, TaskRunnerServiceRes.TVFileTVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_ShouldNotBeNullOrEmpty", TaskRunnerServiceRes.TVFileTVItemID);
                return;
            }

            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVFileModel tvFileModel = tvFileService.GetTVFileModelWithTVFileTVItemIDDB(TVFileTVItemID);
            if (!string.IsNullOrWhiteSpace(tvFileModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.TVFileTVItemID, TVFileTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.TVFileTVItemID, TVFileTVItemID.ToString());
                return;
            }

            // ready to save the file uploaded
            string ServerFilePath = tvFileService.GetServerFilePath(MikeScenarioTVItemID);

            FileInfo fi = new FileInfo(ServerFilePath + tvFileModel.ServerFileName);


            if (!fi.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fi.FullName);
                return;
            }

            if (fi.Extension.ToLower() == ".mdf")
            {
                GetMoreFilesToImportFromMDFFile(MikeScenarioTVItemID, TVFileTVItemID, tvFileModel.ClientFilePath + tvFileModel.ServerFileName, tvFileModel.ServerFilePath + tvFileModel.ServerFileName);
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                    return;
            }
            else if (fi.Extension.ToLower() == ".dfs0")
            {
            }
            else if (fi.Extension.ToLower() == ".dfs1")
            {
            }
            else if (fi.Extension.ToLower() == ".dfs2")
            {
            }
            else if (fi.Extension.ToLower() == ".dfsu")
            {
            }
            else
            {
            }

            return;
        }
        public void MikeScenarioAskToRun()
        {
            MikeScenarioService mikeScenarioService = new MikeScenarioService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            _TaskRunnerBaseService.SendStatusToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, AppTaskStatusEnum.Running);

            string NotUsed = "";

            string Parameters = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters;
            string[] ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            int MikeScenarioTVItemID = 0;
            foreach (string s in ParamValueList)
            {
                string[] ParamValue = s.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                if (ParamValue.Length != 2)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, TaskRunnerServiceRes.Parameters);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", TaskRunnerServiceRes.Parameters);
                    return;
                }

                if (ParamValue[0] == "MikeScenarioTVItemID")
                {
                    MikeScenarioTVItemID = int.Parse(ParamValue[1]);
                }
                else
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind__, TaskRunnerServiceRes.Parameter, ParamValue[0]);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind__", TaskRunnerServiceRes.Parameter, ParamValue[0].ToString());
                    return;
                }
            }

            if (MikeScenarioTVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBe0, TaskRunnerServiceRes.MikeScenarioTVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_ShouldNotBe0", TaskRunnerServiceRes.MikeScenarioTVItemID);
                return;
            }

            MikeScenarioModel mikeScenarioModel = mikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(MikeScenarioTVItemID);
            if (!string.IsNullOrWhiteSpace(mikeScenarioModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                return;
            }

            // should save the new start end date to the .m21fm or .m3fm file
            // will also create the initial conditions (sal, temp)
            // will also create file for pollution source (flow, concentration)
            SaveDBInfoToMikeScenarioFile(MikeScenarioTVItemID);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
            {
                return;
            }

            CreateWebTideDfs1Files(MikeScenarioTVItemID);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
            {
                return;
            }

            NotUsed = TaskRunnerServiceRes.WaitingToRun;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("WaitingToRun"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 3);
            _TaskRunnerBaseService.SendTaskCommandToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, AppTaskCommandEnum.MikeScenarioWaitingToRun);
            //_TaskRunnerBaseService.SendStatusToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, AppTaskStatusEnum.Created);

            int CountMikeScenarioRunning = mikeScenarioService.GetMikeScenarioCountWithMikeScenarioStatusDB(ScenarioStatusEnum.Running);
            if (CountMikeScenarioRunning > 0)
            {
                // Can only run 3 mike scenarios at once
                return;
            }

            // starting the Updater.exe application
            ProcessStartInfo pInfoUpdater = new ProcessStartInfo();
            pInfoUpdater.Arguments = "";
            pInfoUpdater.WindowStyle = ProcessWindowStyle.Minimized;
            pInfoUpdater.UseShellExecute = true;

            Process processUpdater = new Process();
            processUpdater.StartInfo = pInfoUpdater;
            try
            {
                pInfoUpdater.FileName = @"C:\CSSP_Execs\Updater\Debug\Updater.exe";
                processUpdater.Start();
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotRun_Error_, pInfoUpdater.FileName, ex.Message);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotRun_Error_", pInfoUpdater.FileName, ex.Message);
                return;
            }

            processUpdater.WaitForInputIdle(2000);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
            {
                return;
            }
        }
        #endregion Functions public

        #region Functions private
        private string AddSourceKeywordAndParameter(PFSSection pfsSectionNew, PFSSection pfsSectionSourcePrev, string keyword, int position, int index,
            string type /* possible values 'string', 'bool', 'int', 'double', 'filename' */)
        {
            PFSKeyword pfsKeywordOld = null;
            PFSKeyword pfsKeywordNew = null;
            PFSParameter pfsParameterOld = null;
            PFSParameter pfsParameter = null;

            try
            {
                pfsKeywordOld = pfsSectionSourcePrev.GetKeyword(keyword);
            }
            catch (Exception ex)
            {
                return string.Format(TaskRunnerServiceRes.PFS_Error_, "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
            }

            try
            {
                if (index == 1)
                {
                    pfsKeywordNew = pfsSectionNew.InsertNewKeyword(keyword, position);
                }
                else
                {
                    pfsKeywordNew = pfsSectionNew.GetKeyword(keyword);
                }
            }
            catch (Exception ex)
            {
                return string.Format(TaskRunnerServiceRes.PFS_Error_, "InsertNewKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
            }

            try
            {
                pfsParameterOld = pfsKeywordOld.GetParameter(index);
            }
            catch (Exception ex)
            {
                return string.Format(TaskRunnerServiceRes.PFS_Error_, "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
            }

            try
            {
                switch (type)
                {
                    case "string":
                        {
                            pfsParameter = pfsKeywordNew.InsertNewParameterString(pfsParameterOld.ToStringValue(), index);
                        }
                        break;
                    case "bool":
                        {
                            pfsParameter = pfsKeywordNew.InsertNewParameterBool(pfsParameterOld.ToBoolean(), index);
                        }
                        break;
                    case "int":
                        {
                            pfsParameter = pfsKeywordNew.InsertNewParameterInt(pfsParameterOld.ToInt(), index);
                        }
                        break;
                    case "double":
                        {
                            pfsParameter = pfsKeywordNew.InsertNewParameterDouble(pfsParameterOld.ToDouble(), index);
                        }
                        break;
                    case "filename":
                        {
                            pfsParameter = pfsKeywordNew.InsertNewParameterFileName(pfsParameterOld.ToFileName(), index);
                        }
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                return string.Format(TaskRunnerServiceRes.PFS_Error_, "InsertNewParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
            }

            return "";
        }
        private string AddHydroSourceKeyAndParam(PFSSection pfsSectionNew, PFSSection pfsSectionSourcePrev)
        {
            string ret = "";

            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "Name", 1, 1, "string");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "include", 2, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "interpolation_type", 3, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "coordinate_type", 4, 1, "string");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "zone", 5, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "coordinates", 6, 1, "double");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "coordinates", 6, 2, "double");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "coordinates", 6, 3, "double");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "layer", 7, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "distribution_type", 8, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "connected_source", 9, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "type", 10, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "format", 11, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "file_name", 12, 1, "filename");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "constant_value", 13, 1, "double");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "item_number", 14, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "item_name", 15, 1, "string");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "type_of_soft_start", 16, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "soft_time_interval", 17, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "reference_value", 18, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "type_of_time_interpolation", 19, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "diameter", 20, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "sigma", 21, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "theta", 22, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "maximum_distance", 23, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "upstream", 24, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "distance_upstream", 25, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;


            return "";
        }
        private string AddHydroTempSalSourceKeyAndParam(PFSSection pfsSectionNew, PFSSection pfsSectionSourcePrev)
        {
            string ret = "";

            // doing SALINITY
            PFSSection pfsSectionNew3 = pfsSectionNew.InsertNewSection("SALINITY", 1);
            if (pfsSectionNew3 == null)
            {
                return string.Format(TaskRunnerServiceRes.CouldNotAddNewSection_, "SALINITY");
            }

            PFSSection pfsSectionSourcePrev3 = pfsSectionSourcePrev.GetSection("SALINITY", 1);
            if (pfsSectionSourcePrev3 == null)
            {
                return string.Format(TaskRunnerServiceRes.CouldNotGetSection_, "SALINITY");
            }

            ret = AddSourceKeywordAndParameter(pfsSectionNew3, pfsSectionSourcePrev3, "Touched", 1, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew3, pfsSectionSourcePrev3, "type", 2, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew3, pfsSectionSourcePrev3, "format", 3, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew3, pfsSectionSourcePrev3, "constant_value", 4, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew3, pfsSectionSourcePrev3, "file_name", 5, 1, "filename");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew3, pfsSectionSourcePrev3, "item_name", 6, 1, "string");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew3, pfsSectionSourcePrev3, "type_of_soft_start", 7, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew3, pfsSectionSourcePrev3, "soft_time_interval", 8, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew3, pfsSectionSourcePrev3, "reference_value", 9, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew3, pfsSectionSourcePrev3, "type_of_time_interpolation", 10, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;

            // doing TEMPERATURE
            PFSSection pfsSectionNew2 = pfsSectionNew.InsertNewSection("TEMPERATURE", 1);
            if (pfsSectionNew2 == null)
            {
                return string.Format(TaskRunnerServiceRes.CouldNotAddNewSection_, "TEMPERATURE");
            }

            PFSSection pfsSectionSourcePrev2 = pfsSectionSourcePrev.GetSection("TEMPERATURE", 1);
            if (pfsSectionSourcePrev2 == null)
            {
                return string.Format(TaskRunnerServiceRes.CouldNotGetSection_, "TEMPERATURE");
            }

            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "Touched", 1, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "type", 2, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "format", 3, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "constant_value", 4, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "file_name", 5, 1, "filename");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "item_name", 6, 1, "string");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "type_of_soft_start", 7, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "soft_time_interval", 8, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "reference_value", 9, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "type_of_time_interpolation", 10, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;


            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "type_of_salinity", 1, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "type_of_temperature", 1, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "name", 1, 1, "string");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;

            return "";
        }
        private string AddHydroTurbSourceKeyAndParam(PFSSection pfsSectionNew, PFSSection pfsSectionSourcePrev)
        {
            string ret = "";

            // doing DISSIPATION_OF_KINETIC_ENERGY
            PFSSection pfsSectionNew2 = pfsSectionNew.InsertNewSection("DISSIPATION_OF_KINETIC_ENERGY", 1);
            if (pfsSectionNew2 == null)
            {
                return string.Format(TaskRunnerServiceRes.CouldNotAddNewSection_, "DISSIPATION_OF_KINETIC_ENERGY");
            }

            PFSSection pfsSectionSourcePrev2 = pfsSectionSourcePrev.GetSection("DISSIPATION_OF_KINETIC_ENERGY", 1);
            if (pfsSectionSourcePrev2 == null)
            {
                return string.Format(TaskRunnerServiceRes.CouldNotGetSection_, "DISSIPATION_OF_KINETIC_ENERGY");
            }

            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "Touched", 1, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "type", 2, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "format", 3, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "constant_value", 4, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "file_name", 5, 1, "filename");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "item_name", 6, 1, "string");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "type_of_soft_start", 7, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "soft_time_interval", 8, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "reference_value", 9, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "type_of_time_interpolation", 10, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;

            // doing KINETIC_ENERGY
            PFSSection pfsSectionNew3 = pfsSectionNew.InsertNewSection("KINETIC_ENERGY", 1);
            if (pfsSectionNew3 == null)
            {
                return string.Format(TaskRunnerServiceRes.CouldNotAddNewSection_, "KINETIC_ENERGY");
            }

            PFSSection pfsSectionSourcePrev3 = pfsSectionSourcePrev.GetSection("KINETIC_ENERGY", 1);
            if (pfsSectionSourcePrev3 == null)
            {
                return string.Format(TaskRunnerServiceRes.CouldNotGetSection_, "KINETIC_ENERGY");
            }

            ret = AddSourceKeywordAndParameter(pfsSectionNew3, pfsSectionSourcePrev3, "Touched", 1, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew3, pfsSectionSourcePrev3, "type", 2, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew3, pfsSectionSourcePrev3, "format", 3, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew3, pfsSectionSourcePrev3, "constant_value", 4, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew3, pfsSectionSourcePrev3, "file_name", 5, 1, "filename");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew3, pfsSectionSourcePrev3, "item_name", 6, 1, "string");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew3, pfsSectionSourcePrev3, "type_of_soft_start", 7, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew3, pfsSectionSourcePrev3, "soft_time_interval", 8, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew3, pfsSectionSourcePrev3, "reference_value", 9, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew3, pfsSectionSourcePrev3, "type_of_time_interpolation", 10, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;

            return "";
        }
        private string AddTransSourceKeyAndParam(PFSSection pfsSectionNew, PFSSection pfsSectionSourcePrev)
        {
            string ret = "";

            // doing COMPONENT_1
            PFSSection pfsSectionNew2 = pfsSectionNew.InsertNewSection("COMPONENT_1", 1);
            if (pfsSectionNew2 == null)
            {
                return string.Format(TaskRunnerServiceRes.CouldNotAddNewSection_, "COMPONENT_1");
            }

            PFSSection pfsSectionSourcePrev2 = pfsSectionSourcePrev.GetSection("COMPONENT_1", 1);
            if (pfsSectionSourcePrev2 == null)
            {
                return string.Format(TaskRunnerServiceRes.CouldNotGetSection_, "COMPONENT_1");
            }

            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "type_of_component", 1, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "type", 2, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "format", 3, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "constant_value", 4, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "file_name", 5, 1, "filename");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "item_name", 6, 1, "string");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "type_of_soft_start", 7, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "soft_time_interval", 8, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "reference_value", 9, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "type_of_time_interpolation", 10, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;

            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "name", 1, 1, "string");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "MzSEPfsListItemCount", 1, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "Touched", 1, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;

            return "";
        }
        private string AddEcolabSourceKeyAndParam(PFSSection pfsSectionNew, PFSSection pfsSectionSourcePrev)
        {
            string ret = "";

            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "name", 1, 1, "string");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "MzSEPfsListItemCount", 1, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "Touched", 1, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;

            return "";
        }
        private string AddMudTransSourceKeyAndParam(PFSSection pfsSectionNew, PFSSection pfsSectionSourcePrev)
        {

            string ret = "";

            // doing SSC_FRACTION_1
            PFSSection pfsSectionNew2 = pfsSectionNew.InsertNewSection("SSC_FRACTION_1", 1);
            if (pfsSectionNew2 == null)
            {
                return string.Format(TaskRunnerServiceRes.CouldNotAddNewSection_, "SSC_FRACTION_1");
            }

            PFSSection pfsSectionSourcePrev2 = pfsSectionSourcePrev.GetSection("SSC_FRACTION_1", 1);
            if (pfsSectionSourcePrev2 == null)
            {
                return string.Format(TaskRunnerServiceRes.CouldNotGetSection_, "SSC_FRACTION_1");
            }

            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "type_of_component", 1, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "type", 2, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "format", 3, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "constant_value", 4, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "file_name", 5, 1, "filename");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "item_name", 6, 1, "string");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "type_of_soft_start", 7, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "soft_time_interval", 8, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "reference_value", 9, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "type_of_time_interpolation", 10, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;

            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "name", 1, 1, "string");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "MzSEPfsListItemCount", 1, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "Touched", 1, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;

            return "";
        }
        private string AddSandTransSourceKeyAndParam(PFSSection pfsSectionNew, PFSSection pfsSectionSourcePrev)
        {
            string ret = "";

            // doing SSC_FRACTION_1
            PFSSection pfsSectionNew2 = pfsSectionNew.InsertNewSection("SSC_FRACTION_1", 1);
            if (pfsSectionNew2 == null)
            {
                return string.Format(TaskRunnerServiceRes.CouldNotAddNewSection_, "SSC_FRACTION_1");
            }

            PFSSection pfsSectionSourcePrev2 = pfsSectionSourcePrev.GetSection("SSC_FRACTION_1", 1);
            if (pfsSectionSourcePrev2 == null)
            {
                return string.Format(TaskRunnerServiceRes.CouldNotGetSection_, "SSC_FRACTION_1");
            }

            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "type_of_component", 1, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "type", 2, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "format", 3, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "constant_value", 4, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "file_name", 5, 1, "filename");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "item_name", 6, 1, "string");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "type_of_soft_start", 7, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "soft_time_interval", 8, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "reference_value", 9, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew2, pfsSectionSourcePrev2, "type_of_time_interpolation", 10, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;

            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "name", 1, 1, "string");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "MzSEPfsListItemCount", 1, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;
            ret = AddSourceKeywordAndParameter(pfsSectionNew, pfsSectionSourcePrev, "Touched", 1, 1, "int");
            if (!string.IsNullOrWhiteSpace(ret)) return ret;

            return "";
        }
        private string AddNewSourceInPFS(PFSFile pfsFile, string NewSource, int LastSourceIndex)
        {
            int SourceInt = int.Parse(NewSource.Substring("SOURCE_".Length));

            List<string> pathList = new List<string>()
            {
                "FemEngineHD/HYDRODYNAMIC_MODULE/SOURCES/",
                "FemEngineHD/HYDRODYNAMIC_MODULE/TEMPERATURE_SALINITY_MODULE/SOURCES/",
                "FemEngineHD/HYDRODYNAMIC_MODULE/TURBULENCE_MODULE/SOURCES/",
                "FemEngineHD/TRANSPORT_MODULE/SOURCES/",
                "FemEngineHD/ECOLAB_MODULE/SOURCES/",
                "FemEngineHD/MUD_TRANSPORT_MODULE/SOURCES/",
                "FemEngineHD/SAND_TRANSPORT_MODULE/SOURCES/",
            };

            foreach (string path in pathList)
            {
                PFSSection pfsSectionSources = pfsFile.GetSectionFromHandle(path);
                if (pfsSectionSources == null)
                {
                    continue;
                    //return string.Format(TaskRunnerServiceRes.CouldNotFind_, path);
                }

                PFSSection pfsSectionNew = pfsSectionSources.InsertNewSection(NewSource, SourceInt);
                if (pfsSectionNew == null)
                {
                    return string.Format(TaskRunnerServiceRes.CouldNotAddNewSection_, path + NewSource);
                }

                PFSSection pfsSectionSourcePrev = pfsFile.GetSectionFromHandle(path + "SOURCE_" + LastSourceIndex.ToString());

                switch (path)
                {
                    case "FemEngineHD/HYDRODYNAMIC_MODULE/SOURCES/":
                        {
                            AddHydroSourceKeyAndParam(pfsSectionNew, pfsSectionSourcePrev);
                        }
                        break;
                    case "FemEngineHD/HYDRODYNAMIC_MODULE/TEMPERATURE_SALINITY_MODULE/SOURCES/":
                        {
                            AddHydroTempSalSourceKeyAndParam(pfsSectionNew, pfsSectionSourcePrev);
                        }
                        break;
                    case "FemEngineHD/HYDRODYNAMIC_MODULE/TURBULENCE_MODULE/SOURCES/":
                        {
                            AddHydroTurbSourceKeyAndParam(pfsSectionNew, pfsSectionSourcePrev);
                        }
                        break;
                    case "FemEngineHD/TRANSPORT_MODULE/SOURCES/":
                        {
                            AddTransSourceKeyAndParam(pfsSectionNew, pfsSectionSourcePrev);
                        }
                        break;
                    case "FemEngineHD/ECOLAB_MODULE/SOURCES/":
                        {
                            AddEcolabSourceKeyAndParam(pfsSectionNew, pfsSectionSourcePrev);
                        }
                        break;
                    case "FemEngineHD/MUD_TRANSPORT_MODULE/SOURCES/":
                        {
                            AddMudTransSourceKeyAndParam(pfsSectionNew, pfsSectionSourcePrev);
                        }
                        break;
                    case "FemEngineHD/SAND_TRANSPORT_MODULE/SOURCES/":
                        {
                            AddSandTransSourceKeyAndParam(pfsSectionNew, pfsSectionSourcePrev);
                        }
                        break;
                    default:
                        break;
                }
            }

            return "";
        }
        private void AddFileToFileNameList(string ServerPath, string FullClientFileName, List<OtherFileInfo> FileNameList, bool IsOutput)
        {
            if (!string.IsNullOrWhiteSpace(ServerPath))
            {
                if (!string.IsNullOrWhiteSpace(FullClientFileName))
                {
                    FileInfo fi = new FileInfo(FullClientFileName);
                    string FileName = fi.Name;
                    FileNameList.Add(new OtherFileInfo() { TVFileID = 0, ClientFullFileName = FullClientFileName, ServerFullFileName = ServerPath + FileName, IsOutput = IsOutput });
                }
            }
            return;
        }
        private bool CreateDischargeSource(PFSFile pfsFile, MikeSourceModel msm, FileInfo fiBC, MikeScenarioModel mikeScenarioModel, FileInfo fiDischarge, TVItemLanguageModel tvItemLanguageModelSourceName, string NewFlowSourceFileName, TVItemService tvItemService, TVFileService tvFileService)
        {
            string NotUsed = "";

            MikeSourceStartEndService mikeSourceStartEndService = new MikeSourceStartEndService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            HydrometricSiteService hydrometricSiteService = new HydrometricSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            HydrometricDataValueService hydrometricDataValueService = new HydrometricDataValueService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MWQMRunService mwqmRunService = new MWQMRunService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            List<MikeSourceStartEndModel> mikeSourceStartEndModelList = mikeSourceStartEndService.GetMikeSourceStartEndModelListWithMikeSourceIDDB(msm.MikeSourceID);

            DfsFactory factory = new DfsFactory();

            IDfsFile dfsOldFile = DfsFileFactory.DfsGenericOpen(fiBC.FullName);

            DfsBuilder dfsNewFile = DfsBuilder.Create(tvItemLanguageModelSourceName.TVText + " - Discharge", dfsOldFile.FileInfo.ApplicationTitle + tvItemLanguageModelSourceName.TVText + " - Discharge", dfsOldFile.FileInfo.ApplicationVersion);

            double DischargeStepsInMinutes = ((double)((IDfsEqCalendarAxis)((dfsOldFile.FileInfo).TimeAxis)).TimeStep / 60);

            DateTime? dateTimeTemp = GetParameterStartTime(pfsFile, "FemEngineHD/TIME", "start_time");
            if (dateTimeTemp == null)
            {
                return false;
            }

            int? NumberOfTimeSteps = GetParameterInt(pfsFile, "FemEngineHD/TIME", "number_of_time_steps");
            if (NumberOfTimeSteps == null)
            {
                return false;
            }

            int? TimeStepInterval = GetParameterInt(pfsFile, "FemEngineHD/TIME", "time_step_interval");
            if (TimeStepInterval == null)
            {
                return false;
            }

            DateTime StartDate = ((DateTime)dateTimeTemp).AddHours(-1);
            DateTime EndDate = ((DateTime)dateTimeTemp).AddSeconds((int)NumberOfTimeSteps * (int)TimeStepInterval).AddHours(1);

            dfsNewFile.SetDataType(1);
            dfsNewFile.SetGeographicalProjection(dfsOldFile.FileInfo.Projection);
            dfsNewFile.SetTemporalAxis(factory.CreateTemporalEqCalendarAxis(eumUnit.eumUsec, StartDate, 0, DischargeStepsInMinutes * 60));
            dfsNewFile.SetItemStatisticsType(StatType.RegularStat);

            DfsDynamicItemBuilder item = dfsNewFile.CreateDynamicItemBuilder();
            item.Set(NewFlowSourceFileName, eumQuantity.Create(eumItem.eumIDischarge, eumUnit.eumUm3PerSec), DfsSimpleType.Float);
            item.SetValueType(DataValueType.Instantaneous);
            item.SetAxis(factory.CreateAxisEqD0());
            item.SetReferenceCoordinates(1f, 2f, 3f);
            dfsNewFile.AddDynamicItem(item.GetDynamicItemInfo());

            dfsOldFile.Close();

            string[] NewFileErrors = dfsNewFile.Validate();
            StringBuilder sbErr = new StringBuilder();
            foreach (string s in NewFileErrors)
            {
                sbErr.AppendLine(s);
            }

            if (NewFileErrors.Count() > 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateFile_Error_, fiDischarge.FullName, sbErr.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateFile_Error_", fiDischarge.FullName, sbErr.ToString());
                return false;
            }

            dfsNewFile.CreateFile(fiDischarge.FullName);
            IDfsFile DischargeFile = dfsNewFile.GetFile();

            DateTime NewDateTime = StartDate;

            if (msm.UseHydrometric)
            {
                DateTime RunDate = new DateTime();

                if (mikeScenarioModel.ForSimulatingMWQMRunTVItemID == null || mikeScenarioModel.ForSimulatingMWQMRunTVItemID == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MWQMRun, TaskRunnerServiceRes.MWQMRunTVItemID, mikeScenarioModel.ForSimulatingMWQMRunTVItemID.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MWQMRun, TaskRunnerServiceRes.MWQMRunTVItemID, mikeScenarioModel.ForSimulatingMWQMRunTVItemID.ToString());
                    return false;
                }

                MWQMRunModel mwqmRunModel = mwqmRunService.GetMWQMRunModelWithMWQMRunTVItemIDDB((int)mikeScenarioModel.ForSimulatingMWQMRunTVItemID);
                if (!string.IsNullOrWhiteSpace(mwqmRunModel.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MWQMRun, TaskRunnerServiceRes.MWQMRunTVItemID, mikeScenarioModel.ForSimulatingMWQMRunTVItemID.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MWQMRun, TaskRunnerServiceRes.MWQMRunTVItemID, mikeScenarioModel.ForSimulatingMWQMRunTVItemID.ToString());
                    return false;
                }

                RunDate = mwqmRunModel.DateTime_Local;

                if (msm.HydrometricTVItemID == null || msm.HydrometricTVItemID == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.HydrometricSite, TaskRunnerServiceRes.HydrometricTVItemID, msm.HydrometricTVItemID.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.HydrometricSite, TaskRunnerServiceRes.HydrometricTVItemID, msm.HydrometricTVItemID.ToString());
                    return false;
                }

                HydrometricSiteModel hydrometricSiteModel = hydrometricSiteService.GetHydrometricSiteModelWithHydrometricSiteTVItemIDDB((int)msm.HydrometricTVItemID);
                if (!string.IsNullOrWhiteSpace(hydrometricSiteModel.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.HydrometricSite, TaskRunnerServiceRes.HydrometricTVItemID, msm.HydrometricTVItemID.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.HydrometricSite, TaskRunnerServiceRes.HydrometricTVItemID, msm.HydrometricTVItemID.ToString());
                    return false;
                }

                List<HydrometricDataValueModel> hydrometricDataValueModelList = hydrometricDataValueService.GetHydrometricDataValueModelListWithHydrometricSiteIDAroundRunDateDB(hydrometricSiteModel.HydrometricSiteTVItemID, RunDate).OrderBy(c => c.DateTime_Local).ToList();
                if (hydrometricDataValueModelList.Count != 10)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.HydrometricSite, TaskRunnerServiceRes.HydrometricTVItemID, msm.HydrometricTVItemID.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.HydrometricSite, TaskRunnerServiceRes.HydrometricTVItemID, msm.HydrometricTVItemID.ToString());
                    return false;
                }

                for (int i = 0, count = hydrometricDataValueModelList.Count - 1; i < count; i++)
                {
                    DateTime StartDateAndTime_Local = hydrometricDataValueModelList[i].DateTime_Local;
                    DateTime EndDateAndTime_Local = hydrometricDataValueModelList[i + 1].DateTime_Local;

                    while (NewDateTime < StartDateAndTime_Local)
                    {
                        DischargeFile.WriteItemTimeStepNext(0, new float[] { 0f });  // Discharge
                        NewDateTime = NewDateTime.AddSeconds(DischargeStepsInMinutes * 60);
                    }

                    long TotalTicks = EndDateAndTime_Local.Ticks - StartDateAndTime_Local.Ticks;
                    while (NewDateTime < EndDate)
                    {
                        if (NewDateTime < StartDateAndTime_Local)
                        {
                            DischargeFile.WriteItemTimeStepNext(0, new float[] { 0f });  // Discharge
                        }
                        else if (NewDateTime >= StartDateAndTime_Local && NewDateTime <= EndDateAndTime_Local)
                        {
                            double TickFraction = (double)(NewDateTime.Ticks - StartDateAndTime_Local.Ticks) / (double)TotalTicks;
                            double DischargeValue = 0.0D;
                            if (hydrometricDataValueModelList[i + 1].Discharge_m3_s == hydrometricDataValueModelList[i].Discharge_m3_s)
                            {
                                DischargeValue = (double)hydrometricDataValueModelList[i].Discharge_m3_s;
                            }
                            else if (hydrometricDataValueModelList[i + 1].Discharge_m3_s > hydrometricDataValueModelList[i].Discharge_m3_s)
                            {
                                DischargeValue = (double)hydrometricDataValueModelList[i].Discharge_m3_s - ((double)(hydrometricDataValueModelList[i + 1].Discharge_m3_s - hydrometricDataValueModelList[i].Discharge_m3_s) * TickFraction);
                            }
                            else if (hydrometricDataValueModelList[i + 1].Discharge_m3_s < hydrometricDataValueModelList[i].Discharge_m3_s)
                            {
                                DischargeValue = (double)hydrometricDataValueModelList[i].Discharge_m3_s + ((double)(hydrometricDataValueModelList[i + 1].Discharge_m3_s - hydrometricDataValueModelList[i].Discharge_m3_s) * TickFraction);
                            }
                            DischargeFile.WriteItemTimeStepNext(0, new float[] { (float)DischargeValue });  // Discharge
                        }
                        else if (NewDateTime > EndDateAndTime_Local)
                        {
                            break;
                        }
                        NewDateTime = NewDateTime.AddSeconds(DischargeStepsInMinutes * 60);
                    }
                }

                while (NewDateTime < EndDate)
                {
                    DischargeFile.WriteItemTimeStepNext(0, new float[] { 0f });  // Discharge
                    NewDateTime = NewDateTime.AddSeconds(DischargeStepsInMinutes * 60);
                }
            }
            else
            {
                foreach (MikeSourceStartEndModel mssem in mikeSourceStartEndModelList.OrderBy(c => c.StartDateAndTime_Local))
                {
                    while (NewDateTime < (DateTime)mssem.StartDateAndTime_Local)
                    {
                        DischargeFile.WriteItemTimeStepNext(0, new float[] { 0f });  // Discharge
                        NewDateTime = NewDateTime.AddSeconds(DischargeStepsInMinutes * 60);
                    }

                    long TotalTicks = mssem.EndDateAndTime_Local.Ticks - mssem.StartDateAndTime_Local.Ticks;
                    while (NewDateTime < EndDate)
                    {
                        if (NewDateTime < mssem.StartDateAndTime_Local)
                        {
                            DischargeFile.WriteItemTimeStepNext(0, new float[] { 0f });  // Discharge
                        }
                        else if (NewDateTime >= mssem.StartDateAndTime_Local && NewDateTime <= mssem.EndDateAndTime_Local)
                        {
                            double TickFraction = (double)(NewDateTime.Ticks - mssem.StartDateAndTime_Local.Ticks) / (double)TotalTicks;
                            double DischargeValue = 0.0D;
                            if (mssem.SourceFlowEnd_m3_day == mssem.SourceFlowStart_m3_day)
                            {
                                DischargeValue = (double)mssem.SourceFlowStart_m3_day;
                            }
                            else if (mssem.SourceFlowEnd_m3_day > mssem.SourceFlowStart_m3_day)
                            {
                                DischargeValue = (double)mssem.SourceFlowStart_m3_day - ((double)(mssem.SourceFlowEnd_m3_day - mssem.SourceFlowStart_m3_day) * TickFraction);
                            }
                            else if (mssem.SourceFlowEnd_m3_day < mssem.SourceFlowStart_m3_day)
                            {
                                DischargeValue = (double)mssem.SourceFlowStart_m3_day + ((double)(mssem.SourceFlowEnd_m3_day - mssem.SourceFlowStart_m3_day) * TickFraction);
                            }
                            DischargeFile.WriteItemTimeStepNext(0, new float[] { (float)DischargeValue });  // Discharge
                        }
                        else if (NewDateTime > mssem.EndDateAndTime_Local)
                        {
                            break;
                        }
                        NewDateTime = NewDateTime.AddSeconds(DischargeStepsInMinutes * 60);
                    }
                }

                while (NewDateTime < EndDate)
                {
                    DischargeFile.WriteItemTimeStepNext(0, new float[] { 0f });  // Discharge
                    NewDateTime = NewDateTime.AddSeconds(DischargeStepsInMinutes * 60);
                }
            }

            DischargeFile.Close();

            fiDischarge = new FileInfo(fiDischarge.FullName);

            NotUsed = string.Format(TaskRunnerServiceRes.SavingSourceFlowDFSOFileInDBSourceName_Number_, msm.MikeSourceTVText, msm.SourceNumberString);
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat2List("SavingSourceFlowDFSOFileInDBSourceName_Number_", msm.MikeSourceTVText, msm.SourceNumberString));

            TVFileModel tvFileModelPolNew = tvFileService.GetTVFileModelWithServerFilePathAndServerFileNameDB(fiDischarge.Directory + "\\", fiDischarge.Name);
            if (!string.IsNullOrWhiteSpace(tvFileModelPolNew.Error))
            {
                if (!fiDischarge.Exists)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiDischarge.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiDischarge.FullName);
                    return false;
                }

                TVItemModel tvItemFileModel = tvItemService.PostAddChildTVItemDB(mikeScenarioModel.MikeScenarioTVItemID, NewFlowSourceFileName, TVTypeEnum.File);
                if (!string.IsNullOrEmpty(tvItemFileModel.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_For_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.MIKEScenarioDocumentation, TaskRunnerServiceRes.TVText, NewFlowSourceFileName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat4List("CouldNotCreate_For_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.MIKEScenarioDocumentation, TaskRunnerServiceRes.TVText, NewFlowSourceFileName);
                    return false;
                }

                tvFileModelPolNew = new TVFileModel()
                {
                    TVFileTVItemID = tvItemFileModel.TVItemID,
                    FilePurpose = FilePurposeEnum.MikeInput,
                    Language = _TaskRunnerBaseService._BWObj.appTaskModel.Language,
                    Year = DateTime.Now.Year,
                    FileDescription = "Mike Discharge File",
                    FileType = tvFileService.GetFileType(fiDischarge.Extension),
                    FileSize_kb = (int)(fiDischarge.Length / 1024),
                    FileInfo = "Mike Scenario Documentation",
                    FileCreatedDate_UTC = fiDischarge.CreationTime.ToUniversalTime(),
                    ServerFileName = NewFlowSourceFileName,
                    ServerFilePath = fiDischarge.Directory + "\\",
                };

                TVFileModel tvFileModelRet = tvFileService.PostAddTVFileDB(tvFileModelPolNew);
                if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                    return false;
                }
            }
            else
            {
                tvFileModelPolNew.FilePurpose = FilePurposeEnum.MikeInput;
                tvFileModelPolNew.FileDescription = "Mike Pollution File";
                tvFileModelPolNew.FileType = tvFileService.GetFileType(fiDischarge.Extension);
                tvFileModelPolNew.FileSize_kb = (int)(fiDischarge.Length / 1024);
                tvFileModelPolNew.FileCreatedDate_UTC = fiDischarge.CreationTime.ToUniversalTime();
                tvFileModelPolNew.ServerFileName = NewFlowSourceFileName;
                tvFileModelPolNew.ServerFilePath = fiDischarge.Directory + "\\";

                TVFileModel tvFileModelRet = tvFileService.PostUpdateTVFileDB(tvFileModelPolNew);
                if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                    return false;
                }
            }

            return true;
        }
        public string CreateInititalConditionFileSalAndTempFromTVFileItemID(int MikeScenarioTVItemID, int useSalinityAndTemperatureInitialConditionFromTVFileTVItemID, FileInfo fiM21_M3)
        {
            string NotUsed = "";
            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            string DFSUFileName = "";
            TVFileModel tvFileModelUseDFSUFile = tvFileService.GetTVFileModelWithTVFileTVItemIDDB(useSalinityAndTemperatureInitialConditionFromTVFileTVItemID);
            if (string.IsNullOrWhiteSpace(tvFileModelUseDFSUFile.Error))
            {
                DFSUFileName = tvFileModelUseDFSUFile.ServerFilePath + tvFileModelUseDFSUFile.ServerFileName;
            }

            List<string> ParamTypeList = new List<string>() { "Salinity", "Temperature" };
            foreach (string ParamType in ParamTypeList)
            {
                DfsuFile dfsuFile;
                try
                {
                    dfsuFile = DfsuFile.Open(DFSUFileName);
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotOpenFile_Error_, DFSUFileName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotOpenFile_Error_", DFSUFileName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    return NotUsed;
                }

                DfsuDynamicItemInfo itemInfoParam = null;
                foreach (DfsuDynamicItemInfo itemInfo in dfsuFile.ItemInfo)
                {
                    if (ParamType == "Salinity")
                    {
                        if (itemInfo.Quantity.Item == eumItem.eumISalinity)
                        {
                            itemInfoParam = itemInfo;
                            break;
                        }
                    }
                    else
                    {
                        if (itemInfo.Quantity.Item == eumItem.eumITemperature)
                        {
                            itemInfoParam = itemInfo;
                            break;
                        }
                    }
                }

                DfsuBuilder dfsBuilder = DfsuBuilder.Create(DfsuFileType.Dfsu2D);
                dfsBuilder.ApplicationTitle = dfsuFile.ApplicationTitle;
                dfsBuilder.ApplicationVersion = dfsuFile.ApplicationVersion;
                dfsBuilder.FileTitle = dfsuFile.FileTitle;
                dfsBuilder.SetNodeIds(dfsuFile.NodeIds);
                dfsBuilder.SetElementIds(dfsuFile.ElementIds);
                dfsBuilder.SetProjection(dfsuFile.Projection);
                dfsBuilder.SetNodes(dfsuFile.X, dfsuFile.Y, dfsuFile.Z, dfsuFile.Code);
                dfsBuilder.SetElements(dfsuFile.ElementTable);

                dfsBuilder.SetTimeInfo(dfsuFile.StartDateTime.AddSeconds((dfsuFile.NumberOfTimeSteps - 1) * dfsuFile.TimeStepInSeconds), dfsuFile.TimeStepInSeconds);
                dfsBuilder.AddDynamicItem(itemInfoParam.Name, itemInfoParam.Quantity);

                string ParamDfsuFileName = (ParamType == "Salinity" ? fiM21_M3.FullName.Replace(fiM21_M3.Name, "Salinity.dfsu") : fiM21_M3.FullName.Replace(fiM21_M3.Name, "Temperature.dfsu"));
                DfsuFile dfsuFileParam = null;
                try
                {
                    dfsuFileParam = dfsBuilder.CreateFile(ParamDfsuFileName);
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateFile_Error_, ParamDfsuFileName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateFile_Error_", ParamDfsuFileName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    return NotUsed;
                }

                IDfsItemData itemData = dfsuFile.ReadItemTimeStep(itemInfoParam.ItemNumber, dfsuFile.NumberOfTimeSteps - 1);

                try
                {
                    dfsuFileParam.WriteItemTimeStep(1, 0, 0, itemData.Data);
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateFile_Error_, ParamDfsuFileName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateFile_Error_", ParamDfsuFileName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    return NotUsed;
                }

                dfsuFileParam.Close();

                dfsuFile.Close();

                FileInfo fiParam = new FileInfo(ParamDfsuFileName);

                if (!fiParam.Exists)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, ParamDfsuFileName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", ParamDfsuFileName);
                    return NotUsed;
                }

                TVItemModel tvItemModelMikeScenario = tvItemService.GetTVItemModelWithTVItemIDDB(MikeScenarioTVItemID);
                if (!string.IsNullOrWhiteSpace(tvItemModelMikeScenario.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                    return NotUsed;
                }

                TVItemModel tvItemModelFileParam = null;
                List<TVItemModel> tvItemModelFileList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(MikeScenarioTVItemID, TVTypeEnum.File);
                foreach (TVItemModel tvItemModelFile in tvItemModelFileList)
                {
                    if (ParamType == "Salinity")
                    {
                        if (tvItemModelFile.TVText.ToLower().EndsWith("salinity.dfsu"))
                        {
                            tvItemModelFileParam = tvItemModelFile;
                            break;
                        }
                    }
                    else
                    {
                        if (tvItemModelFile.TVText.ToLower().EndsWith("temperature.dfsu"))
                        {
                            tvItemModelFileParam = tvItemModelFile;
                            break;
                        }
                    }
                }

                if (tvItemModelFileParam == null)
                {
                    TVItemModel tvItemModelRet = tvItemService.PostAddChildTVItemDB(MikeScenarioTVItemID, (ParamType == "Salinity" ? "Salinity.dfsu" : "Temperature.dfsu"), TVTypeEnum.File);
                    if (!string.IsNullOrWhiteSpace(tvItemModelRet.Error))
                    {
                        // already exist
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_For_With_Equal_, TaskRunnerServiceRes.TVItem, "MikeScenario", TaskRunnerServiceRes.TVText, (ParamType == "Salinity" ? "Salinity.dfsu" : "Temperature.dfsu"));
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat4List("CouldNotCreate_For_With_Equal_", TaskRunnerServiceRes.TVItem, "MikeScenario", TaskRunnerServiceRes.TVText, (ParamType == "Salinity" ? "Salinity.dfsu" : "Temperature.dfsu"));
                        return NotUsed;
                    }

                    TVFileModel tvFileModelNew = new TVFileModel()
                    {
                        ClientFilePath = "",
                        FileDescription = "Mike Scenario " + (ParamType == "Salinity" ? "Salinity" : "Temperature") + " DFSU file",
                        FileInfo = "Mike Scenario " + (ParamType == "Salinity" ? "Salinity" : "Temperature") + " DFSU file",
                        FileCreatedDate_UTC = fiParam.CreationTimeUtc,
                        FileSize_kb = (int)(fiParam.Length / 1024),
                        Language = LanguageEnum.enAndfr,
                        Year = DateTime.Now.Year,
                        FromWater = null,
                        FilePurpose = FilePurposeEnum.MikeInput,
                        FileType = FileTypeEnum.DFSU,
                        ServerFileName = fiParam.Name,
                        ServerFilePath = fiParam.DirectoryName + @"\",
                        TVFileTVItemID = tvItemModelRet.TVItemID,
                    };

                    TVFileModel tvFileModelRet = tvFileService.PostAddTVFileDB(tvFileModelNew);
                    if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_Error_, TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreate_Error_", TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                        return NotUsed;
                    }

                }
                else
                {
                    TVFileModel tvFileModelSal = tvFileService.GetTVFileModelWithTVFileTVItemIDDB(tvItemModelFileParam.TVItemID);
                    if (string.IsNullOrWhiteSpace(tvFileModelSal.Error))
                    {
                        tvFileModelSal.FileSize_kb = (int)(fiParam.Length / 1024);

                        TVFileModel tvFileModelRet = tvFileService.PostUpdateTVFileDB(tvFileModelSal);
                        if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                            return NotUsed;
                        }
                    }
                    else
                    {
                        TVFileModel tvFileModelNew = new TVFileModel()
                        {
                            ClientFilePath = "",
                            FileDescription = "Mike Scenario " + (ParamType == "Salinity" ? "Salinity" : "Temperature") + " DFSU file",
                            FileInfo = "Mike Scenario " + (ParamType == "Salinity" ? "Salinity" : "Temperature") + " DFSU file",
                            FileCreatedDate_UTC = fiParam.CreationTimeUtc,
                            FileSize_kb = (int)(fiParam.Length / 1024),
                            Language = LanguageEnum.enAndfr,
                            Year = DateTime.Now.Year,
                            FromWater = null,
                            FilePurpose = FilePurposeEnum.MikeInput,
                            FileType = FileTypeEnum.DFSU,
                            ServerFileName = fiParam.Name,
                            ServerFilePath = fiParam.DirectoryName + @"\",
                            TVFileTVItemID = tvItemModelFileParam.TVItemID,
                        };

                        TVFileModel tvFileModelRet = tvFileService.PostAddTVFileDB(tvFileModelNew);
                        if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_Error_, TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreate_Error_", TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                            return NotUsed;
                        }
                    }
                }
            }

            return "";
        }
        private bool CreatePollutionConcentrationSource(PFSFile pfsFile, MikeSourceModel msm, FileInfo fiBC, MikeScenarioModel mikeScenarioModel, FileInfo fiPolConc, TVItemLanguageModel tvItemLanguageModelSourceName, string NewPolSourceFileName, TVItemService tvItemService, TVFileService tvFileService)
        {
            string NotUsed = "";


            MikeSourceStartEndService mikeSourceStartEndService = new MikeSourceStartEndService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            List<MikeSourceStartEndModel> mikeSourceStartEndModelList = mikeSourceStartEndService.GetMikeSourceStartEndModelListWithMikeSourceIDDB(msm.MikeSourceID);

            DfsFactory factory = new DfsFactory();

            IDfsFile dfsOldFile = DfsFileFactory.DfsGenericOpen(fiBC.FullName);

            DfsBuilder dfsNewFile = DfsBuilder.Create(tvItemLanguageModelSourceName.TVText + " - Pollution Concentration", dfsOldFile.FileInfo.ApplicationTitle + tvItemLanguageModelSourceName.TVText + " - Pollution Concentration", dfsOldFile.FileInfo.ApplicationVersion);

            double PolConcStepsInMinutes = ((double)((IDfsEqCalendarAxis)((dfsOldFile.FileInfo).TimeAxis)).TimeStep / 60);

            DateTime? dateTimeTemp = GetParameterStartTime(pfsFile, "FemEngineHD/TIME", "start_time");
            if (dateTimeTemp == null)
            {
                return false;
            }

            int? NumberOfTimeSteps = GetParameterInt(pfsFile, "FemEngineHD/TIME", "number_of_time_steps");
            if (NumberOfTimeSteps == null)
            {
                return false;
            }

            int? TimeStepInterval = GetParameterInt(pfsFile, "FemEngineHD/TIME", "time_step_interval");
            if (TimeStepInterval == null)
            {
                return false;
            }

            DateTime StartDate = ((DateTime)dateTimeTemp).AddHours(-1);
            DateTime EndDate = ((DateTime)dateTimeTemp).AddSeconds((int)NumberOfTimeSteps * (int)TimeStepInterval).AddHours(1);

            dfsNewFile.SetDataType(1);
            dfsNewFile.SetGeographicalProjection(dfsOldFile.FileInfo.Projection);
            dfsNewFile.SetTemporalAxis(factory.CreateTemporalEqCalendarAxis(eumUnit.eumUsec, StartDate, 0, PolConcStepsInMinutes * 60));
            dfsNewFile.SetItemStatisticsType(StatType.RegularStat);

            DfsDynamicItemBuilder item = dfsNewFile.CreateDynamicItemBuilder();
            item.Set(NewPolSourceFileName, eumQuantity.Create(eumItem.eumIBacteriaConc, eumUnit.eumUPer100ml), DfsSimpleType.Float);
            item.SetValueType(DataValueType.Instantaneous);
            item.SetAxis(factory.CreateAxisEqD0());
            item.SetReferenceCoordinates(1f, 2f, 3f);
            dfsNewFile.AddDynamicItem(item.GetDynamicItemInfo());

            dfsOldFile.Close();

            string[] NewFileErrors = dfsNewFile.Validate();
            StringBuilder sbErr = new StringBuilder();
            foreach (string s in NewFileErrors)
            {
                sbErr.AppendLine(s);
            }

            if (NewFileErrors.Count() > 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateFile_Error_, fiPolConc.FullName, sbErr.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateFile_Error_", fiPolConc.FullName, sbErr.ToString());
                return false;
            }

            dfsNewFile.CreateFile(fiPolConc.FullName);
            IDfsFile PolConcFile = dfsNewFile.GetFile();

            DateTime NewDateTime = StartDate;

            foreach (MikeSourceStartEndModel mssem in mikeSourceStartEndModelList.OrderBy(c => c.StartDateAndTime_Local))
            {
                while (NewDateTime < (DateTime)mssem.StartDateAndTime_Local)
                {
                    PolConcFile.WriteItemTimeStepNext(0, new float[] { 0f });  // Pollution Concentration
                    NewDateTime = NewDateTime.AddSeconds(PolConcStepsInMinutes * 60);
                }

                long TotalTicks = mssem.EndDateAndTime_Local.Ticks - mssem.StartDateAndTime_Local.Ticks;
                while (NewDateTime < EndDate)
                {
                    if (NewDateTime < mssem.StartDateAndTime_Local)
                    {
                        PolConcFile.WriteItemTimeStepNext(0, new float[] { 0f });  // Pollution Concentration
                    }
                    else if (NewDateTime >= mssem.StartDateAndTime_Local && NewDateTime <= mssem.EndDateAndTime_Local)
                    {
                        double TickFraction = (double)(NewDateTime.Ticks - mssem.StartDateAndTime_Local.Ticks) / (double)TotalTicks;
                        double PolValue = 0.0D;
                        if (mssem.SourcePollutionEnd_MPN_100ml == mssem.SourcePollutionStart_MPN_100ml)
                        {
                            PolValue = (double)mssem.SourcePollutionStart_MPN_100ml;
                        }
                        else if (mssem.SourcePollutionEnd_MPN_100ml > mssem.SourcePollutionStart_MPN_100ml)
                        {
                            PolValue = (double)mssem.SourcePollutionStart_MPN_100ml - ((double)(mssem.SourcePollutionEnd_MPN_100ml - mssem.SourcePollutionStart_MPN_100ml) * TickFraction);
                        }
                        else if (mssem.SourcePollutionEnd_MPN_100ml < mssem.SourcePollutionStart_MPN_100ml)
                        {
                            PolValue = (double)mssem.SourcePollutionStart_MPN_100ml + ((double)(mssem.SourcePollutionEnd_MPN_100ml - mssem.SourcePollutionStart_MPN_100ml) * TickFraction);
                        }
                        PolConcFile.WriteItemTimeStepNext(0, new float[] { (float)PolValue });  // Pollution Concentration
                    }
                    else if (NewDateTime > mssem.EndDateAndTime_Local)
                    {
                        break;
                    }
                    NewDateTime = NewDateTime.AddSeconds(PolConcStepsInMinutes * 60);
                }
            }

            while (NewDateTime < EndDate)
            {
                PolConcFile.WriteItemTimeStepNext(0, new float[] { 0f });  // Pollution Concentration
                NewDateTime = NewDateTime.AddSeconds(PolConcStepsInMinutes * 60);
            }

            PolConcFile.Close();

            fiPolConc = new FileInfo(fiPolConc.FullName);

            NotUsed = string.Format(TaskRunnerServiceRes.SavingSourcePolDFSOFileInDBSourceName_Number_, msm.MikeSourceTVText, msm.SourceNumberString);
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat2List("SavingSourcePolDFSOFileInDBSourceName_Number_", msm.MikeSourceTVText, msm.SourceNumberString));

            TVFileModel tvFileModelPolNew = tvFileService.GetTVFileModelWithServerFilePathAndServerFileNameDB(fiPolConc.Directory + "\\", fiPolConc.Name);
            if (!string.IsNullOrWhiteSpace(tvFileModelPolNew.Error))
            {
                if (!fiPolConc.Exists)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiPolConc.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiPolConc.FullName);
                    return false;
                }

                TVItemModel tvItemFileModel = tvItemService.PostAddChildTVItemDB(mikeScenarioModel.MikeScenarioTVItemID, NewPolSourceFileName, TVTypeEnum.File);
                if (!string.IsNullOrEmpty(tvItemFileModel.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_For_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.MIKEScenarioDocumentation, TaskRunnerServiceRes.TVText, NewPolSourceFileName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat4List("CouldNotCreate_For_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.MIKEScenarioDocumentation, TaskRunnerServiceRes.TVText, NewPolSourceFileName);
                    return false;
                }

                tvFileModelPolNew = new TVFileModel()
                {
                    TVFileTVItemID = tvItemFileModel.TVItemID,
                    FilePurpose = FilePurposeEnum.MikeInput,
                    Language = _TaskRunnerBaseService._BWObj.appTaskModel.Language,
                    Year = DateTime.Now.Year,
                    FileDescription = "Mike Pollution File",
                    FileType = tvFileService.GetFileType(fiPolConc.Extension),
                    FileSize_kb = (int)(fiPolConc.Length / 1024),
                    FileInfo = "Mike Scenario Documentation",
                    FileCreatedDate_UTC = fiPolConc.CreationTime.ToUniversalTime(),
                    ServerFileName = NewPolSourceFileName,
                    ServerFilePath = fiPolConc.Directory + "\\",
                };

                TVFileModel tvFileModelRet = tvFileService.PostAddTVFileDB(tvFileModelPolNew);
                if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                    return false;
                }
            }
            else
            {
                tvFileModelPolNew.FilePurpose = FilePurposeEnum.MikeInput;
                tvFileModelPolNew.FileDescription = "Mike Pollution Concentration File";
                tvFileModelPolNew.FileType = tvFileService.GetFileType(fiPolConc.Extension);
                tvFileModelPolNew.FileSize_kb = (int)(fiPolConc.Length / 1024);
                tvFileModelPolNew.FileCreatedDate_UTC = fiPolConc.CreationTime.ToUniversalTime();
                tvFileModelPolNew.ServerFileName = NewPolSourceFileName;
                tvFileModelPolNew.ServerFilePath = fiPolConc.Directory + "\\";

                TVFileModel tvFileModelRet = tvFileService.PostUpdateTVFileDB(tvFileModelPolNew);
                if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                    return false;
                }
            }

            return true;
        }
        private void CreateWebTideDfs1Files(int MikeScenarioTVItemID)
        {
            string NotUsed = "";
            int TotalWebTideNodes = 0;
            int CurrentWebTideNode = 0;
            if (MikeScenarioTVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.MikeScenarioTVItemID, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                return;
            }

            NotUsed = TaskRunnerServiceRes.CreatingWebTideBoundaryConditionFiles;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("CreatingWebTideBoundaryConditionFiles"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 2);

            MikeScenarioService MikeScenarioService = new CSSPDBDLL.Services.MikeScenarioService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MikeScenarioModel mikeScenarioModel = MikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(MikeScenarioTVItemID);
            if (!string.IsNullOrWhiteSpace(mikeScenarioModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                return;
            }

            MikeBoundaryConditionService mikeBoundaryConditionService = new MikeBoundaryConditionService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            List<MikeBoundaryConditionModel> mbcModelList = mikeBoundaryConditionService.GetMikeBoundaryConditionModelListWithMikeScenarioTVItemIDAndTVTypeDB(mikeScenarioModel.MikeScenarioTVItemID, TVTypeEnum.MikeBoundaryConditionWebTide);
            if (mbcModelList.Count == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_StartingWith_, TaskRunnerServiceRes.MIKEBoundaryCondition, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_StartingWith_", TaskRunnerServiceRes.MIKEBoundaryCondition, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                return;
            }

            foreach (MikeBoundaryConditionModel mbcm in mbcModelList)
            {
                MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                int WebTideCount = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(mbcm.MikeBoundaryConditionTVItemID, TVTypeEnum.MikeBoundaryConditionWebTide, MapInfoDrawTypeEnum.Polyline).Count();
                TotalWebTideNodes += WebTideCount;
            }

            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVFileModel tvFileModel = tvFileService.GetTVFileModelWithTVItemIDAndTVFileTypeM21FMOrM3FMDB(MikeScenarioTVItemID);
            if (!string.IsNullOrWhiteSpace(tvFileModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVFileM21_3FM, TaskRunnerServiceRes.MikeScenarioTVItemID, MikeScenarioTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVFileM21_3FM, TaskRunnerServiceRes.MikeScenarioTVItemID, MikeScenarioTVItemID.ToString());
                return;
            }

            string ServerPath = tvFileService.GetServerFilePath(MikeScenarioTVItemID);

            FileInfo fiM21_M3 = new FileInfo(ServerPath + tvFileModel.ServerFileName);

            if (!fiM21_M3.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiM21_M3.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiM21_M3.FullName);
                return;
            }

            PFSFile pfsFile = new PFSFile(fiM21_M3.FullName);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
            {
                return;
            }

            GoogleTimeZoneJSON googleTimeZoneJSON = new GoogleTimeZoneJSON();
            bool GoogleLoaded = false;
            float offset = 0.0f;
            string offsetText = "";

            int NumberOfMBC = 0;
            foreach (MikeBoundaryConditionModel mbcm in mbcModelList)
            {
                NumberOfMBC += 1;
                _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 10);

                List<OtherFileInfo> FileList = new List<OtherFileInfo>();
                string FileName = GetParameterFileName(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/BOUNDARY_CONDITIONS/" + mbcm.MikeBoundaryConditionCode, "file_name");
                FileInfo fiBC = new FileInfo(FileName);

                List<MapInfoPointModel> mapInfoPointModelBC = new List<MapInfoPointModel>();

                string FilePathBC = fiBC.Directory + @"\";
                string FileNameBC = fiBC.Name;
                string FilePathBC_EDrive = FilePathBC.Replace(@"C:\", @"E:\");
                string newServerPathBC = fiBC.Directory + @"\";
                TVFileModel tvFileModelBC = tvFileService.GetTVFileModelWithServerFilePathAndServerFileNameDB(FilePathBC_EDrive, FileNameBC);
                if (!string.IsNullOrWhiteSpace(tvFileModelBC.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.ServerFilePath + "," + TaskRunnerServiceRes.ServerFileName, FilePathBC_EDrive + FileNameBC);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.ServerFilePath + "," + TaskRunnerServiceRes.ServerFileName, FilePathBC_EDrive + FileNameBC);
                    return;
                }

                MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                mapInfoPointModelBC = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(mbcm.MikeBoundaryConditionTVItemID, TVTypeEnum.MikeBoundaryConditionWebTide, MapInfoDrawTypeEnum.Polyline);
                if (mapInfoPointModelBC.Count == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBeNullOrEmpty, TaskRunnerServiceRes.mapInfoPointModelBC);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_ShouldNotBeNullOrEmpty", TaskRunnerServiceRes.mapInfoPointModelBC);
                    return;
                }

                List<eumItem> eumItemList = new List<eumItem>();

                DfsFactory factory = new DfsFactory();

                FileInfo fi = new FileInfo(newServerPathBC + tvFileModelBC.ServerFileName);

                IDfsFile dfsOldFile = DfsFileFactory.DfsGenericOpen(tvFileService.ChoseEDriveOrCDrive(tvFileModelBC.ServerFilePath) + tvFileModelBC.ServerFileName);

                DfsBuilder dfsNewFile = DfsBuilder.Create(tvFileModelBC.ServerFileName, dfsOldFile.FileInfo.ApplicationTitle, dfsOldFile.FileInfo.ApplicationVersion);

                double WebTideStepsInMinutes = ((double)((IDfsEqCalendarAxis)((dfsOldFile.FileInfo).TimeAxis)).TimeStep / 60);

                DateTime? dateTimeTemp = GetParameterStartTime(pfsFile, "FemEngineHD/TIME", "start_time");
                if (dateTimeTemp == null)
                {
                    return;
                }

                int? NumberOfTimeSteps = GetParameterInt(pfsFile, "FemEngineHD/TIME", "number_of_time_steps");
                if (NumberOfTimeSteps == null)
                {
                    return;
                }

                int? TimeStepInterval = GetParameterInt(pfsFile, "FemEngineHD/TIME", "time_step_interval");
                if (TimeStepInterval == null)
                {
                    return;
                }

                DateTime StartDate = ((DateTime)dateTimeTemp).AddHours(-1);
                DateTime EndDate = ((DateTime)dateTimeTemp).AddSeconds((int)NumberOfTimeSteps * (int)TimeStepInterval).AddHours(1);

                dfsNewFile.SetDataType(dfsOldFile.FileInfo.DataType);
                dfsNewFile.SetGeographicalProjection(dfsOldFile.FileInfo.Projection);
                dfsNewFile.SetTemporalAxis(factory.CreateTemporalEqCalendarAxis(eumUnit.eumUsec, StartDate, 0, WebTideStepsInMinutes * 60));
                dfsNewFile.SetItemStatisticsType(StatType.RegularStat);

                foreach (IDfsDynamicItemInfo di in dfsOldFile.ItemInfo)
                {
                    DfsDynamicItemBuilder ddib = dfsNewFile.CreateDynamicItemBuilder();
                    ddib.Set(di.Name, eumQuantity.Create(di.Quantity.Item, di.Quantity.Unit), di.DataType);
                    ddib.SetValueType(di.ValueType);
                    ddib.SetAxis(factory.CreateAxisEqD1(eumUnit.eumUsec, mapInfoPointModelBC.Count, 0, 1));
                    ddib.SetReferenceCoordinates(di.ReferenceCoordinateX, di.ReferenceCoordinateY, di.ReferenceCoordinateZ);
                    dfsNewFile.AddDynamicItem(ddib.GetDynamicItemInfo());
                    eumItemList.Add(di.Quantity.Item);
                }

                dfsOldFile.Close();

                string[] NewFileErrors = dfsNewFile.Validate();
                StringBuilder sbErr = new StringBuilder();
                foreach (string s in NewFileErrors)
                {
                    sbErr.AppendLine(s);
                }

                if (NewFileErrors.Count() > 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateFile_Error_, tvFileModelBC.ServerFilePath + tvFileModelBC.ServerFileName, sbErr.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreateFile_Error_", tvFileModelBC.ServerFilePath + tvFileModelBC.ServerFileName, sbErr.ToString());
                    return;
                }

                if (eumItemList.Count == 1)
                {
                    if (eumItemList[0] == eumItem.eumIWaterLevel || eumItemList[0] == eumItem.eumIWaterDepth)
                    {
                        List<IEnumerable<WaterLevelResult>> AllWLResults = new List<IEnumerable<WaterLevelResult>>();
                        IEnumerable<WaterLevelResult> WLResults = null;

                        int NodeCount = 0;
                        List<float> InitialWL = new List<float>();
                        for (int i = 0; i < mapInfoPointModelBC.Count; i++)
                        {
                            if (!GoogleLoaded)
                            {
                                using (WebClient webClient = new WebClient() { Encoding = Encoding.UTF8 })
                                {
                                    WebProxy webProxy = new WebProxy();
                                    webClient.Proxy = webProxy;

                                    webClient.UseDefaultCredentials = true;

                                    var json_data = string.Empty;
                                    // attempt to download JSON data as a string
                                    try
                                    {
                                        DateTime epoch = new DateTime(1970, 1, 1, 0, 0, 0, 0).ToLocalTime();
                                        TimeSpan span = (mikeScenarioModel.MikeScenarioStartDateTime_Local - epoch);
                                        string timeStamp = ((long)Convert.ToDouble(span.TotalSeconds)).ToString();
                                        string url = "https://maps.googleapis.com/maps/api/timezone/json?location=" +
                                            mapInfoPointModelBC[i].Lat.ToString("F5") + "," + mapInfoPointModelBC[i].Lng.ToString("F5") + "&timestamp=" + timeStamp + "&key=AIzaSyAwPGpdSM6z0A7DFdWPbS3vIDTk2mxINaA";
                                        byte[] responseBytes = webClient.DownloadData(url);
                                        json_data = Encoding.UTF8.GetString(responseBytes);
                                    }
                                    catch (Exception)
                                    {
                                        json_data = "Error";
                                    }
                                    JsonSerializerSettings jsonSerializerSettings = new JsonSerializerSettings
                                    {
                                        DateTimeZoneHandling = DateTimeZoneHandling.Local
                                    };
                                    googleTimeZoneJSON = JsonConvert.DeserializeObject<GoogleTimeZoneJSON>(json_data, jsonSerializerSettings);
                                }

                                offset = (float)((googleTimeZoneJSON.rawOffset + googleTimeZoneJSON.dstOffset) / 3600.0f);

                                offsetText = offset.ToString("F1");
                                if (offsetText.EndsWith("0"))
                                {
                                    offsetText = offsetText.Substring(0, offsetText.Length - 2);
                                }
                            }

                            NodeCount += 1;
                            CurrentWebTideNode += 1;
                            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (CurrentWebTideNode * 100 / TotalWebTideNodes));

                            TideModel tideModel = new TideModel(tvFileService.ChoseEDriveOrCDrive(tvFileService.BasePath), mbcm.WebTideDataSet)
                            {
                                StartDate = StartDate.AddHours(-offset),
                                EndDate = EndDate.AddHours(-offset),
                                Lng = (double)mapInfoPointModelBC[i].Lng,
                                Lat = (double)mapInfoPointModelBC[i].Lat,
                                Steps_min = WebTideStepsInMinutes,
                                DoWaterLevels = true,
                            };
                            WLResults = tidesAndCurrentsService.GetTides(tideModel);
                            if (WLResults != null)
                            {
                                foreach (WaterLevelResult wlr in WLResults)
                                {
                                    wlr.Date = wlr.Date.AddHours(offset);
                                }
                                AllWLResults.Add(WLResults);
                                InitialWL.Add((float)WLResults.First().WaterLevel);
                            }
                        }
                        if (AllWLResults.Count > 0)
                        {
                            if (!SetParameterDouble(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/INITIAL_CONDITIONS", (float)InitialWL.Average(), "surface_elevation_constant"))
                            {
                                return;
                            }
                        }

                        dfsNewFile.CreateFile(newServerPathBC + tvFileModelBC.ServerFileName);
                        IDfsFile file = dfsNewFile.GetFile();
                        for (int i = 0; i < WLResults.ToList().Count; i++)
                        {
                            float[] floatArray = new float[AllWLResults.Count];

                            for (int j = 0; j < AllWLResults.Count; j++)
                            {
                                floatArray[j] = ((float)((List<WaterLevelResult>)AllWLResults[j].ToList())[i].WaterLevel);
                            }

                            file.WriteItemTimeStepNext(0, floatArray);  // water level array
                        }
                        file.Close();

                        StringBuilder sb = new StringBuilder();
                        sb.AppendLine(TaskRunnerServiceRes.TimeOffsetFromGMT + ": " + offsetText + " " + TaskRunnerServiceRes.Hours + " --- " + googleTimeZoneJSON.timeZoneName);
                        sb.AppendLine(TaskRunnerServiceRes.From + " " + mikeScenarioModel.MikeScenarioStartDateTime_Local.ToString("yyyy MMMM dd HH:mm") +
                            " " + TaskRunnerServiceRes.To + " " + mikeScenarioModel.MikeScenarioEndDateTime_Local.ToString("yyyy MMMM dd HH:mm") + " (" +
                            TaskRunnerServiceRes.LocalTime + ")");
                        sb.AppendLine("");
                        sb.Append(TaskRunnerServiceRes.DateAndTime + "\t");
                        for (int i = 0; i < NodeCount; i++)
                        {
                            sb.Append("\t" + TaskRunnerServiceRes.Node + i.ToString());
                        }
                        sb.AppendLine("");
                        for (int i = 0; i < ((List<WaterLevelResult>)AllWLResults[0]).Count; i++)
                        {
                            sb.Append(((List<WaterLevelResult>)AllWLResults[0])[i].Date.ToString("yyyy MMMM dd HH:mm"));

                            for (int j = 0; j < AllWLResults.Count; j++)
                            {
                                float wl = ((float)((List<WaterLevelResult>)AllWLResults[j].ToList())[i].WaterLevel);
                                string xText = (wl < 0 ? wl.ToString("F4") : " " + wl.ToString("F4"));
                                sb.Append("\t" + xText);
                            }
                            sb.AppendLine();
                        }
                        mbcm.WebTideDataFromStartToEndDate = sb.ToString();

                        MikeBoundaryConditionModel mikeBoundaryConditionModelRet = mikeBoundaryConditionService.PostUpdateMikeBoundaryConditionDB(mbcm);
                        if (!string.IsNullOrWhiteSpace(mikeBoundaryConditionModelRet.Error))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.MIKEBoundaryCondition, mikeBoundaryConditionModelRet.Error);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.MIKEBoundaryCondition, mikeBoundaryConditionModelRet.Error);
                            return;
                        }
                    }
                    else
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.FileContainsOneParamButItsNotOfTypeWLOrWDItIs_, eumItemList[0].ToString());
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("FileContainsOneParamButItsNotOfTypeWLOrWDItIs_", eumItemList[0].ToString());
                        return;
                    }
                }
                else if (eumItemList.Count == 2)
                {
                    if (eumItemList[0] == eumItem.eumIuVelocity && eumItemList[1] == eumItem.eumIvVelocity)
                    {
                        // read web tide for the required time
                        List<IEnumerable<CurrentResult>> AllCurrentResults = new List<IEnumerable<CurrentResult>>();
                        IEnumerable<CurrentResult> CurrentResults = null;

                        int NodeCount = 0;
                        for (int i = 0; i < mapInfoPointModelBC.Count; i++)
                        {
                            if (!GoogleLoaded)
                            {
                                using (WebClient webClient = new WebClient() { Encoding = Encoding.UTF8 })
                                {
                                    WebProxy webProxy = new WebProxy();
                                    webClient.Proxy = webProxy;

                                    webClient.UseDefaultCredentials = true;

                                    var json_data = string.Empty;
                                    // attempt to download JSON data as a string
                                    try
                                    {
                                        DateTime epoch = new DateTime(1970, 1, 1, 0, 0, 0, 0).ToLocalTime();
                                        TimeSpan span = (mikeScenarioModel.MikeScenarioStartDateTime_Local - epoch);
                                        string timeStamp = ((long)Convert.ToDouble(span.TotalSeconds)).ToString();
                                        string url = "https://maps.googleapis.com/maps/api/timezone/json?location=" +
                                            mapInfoPointModelBC[i].Lat.ToString("F5") + "," + mapInfoPointModelBC[i].Lng.ToString("F5") + "&timestamp=" + timeStamp + "&key=AIzaSyAwPGpdSM6z0A7DFdWPbS3vIDTk2mxINaA";
                                        byte[] responseBytes = webClient.DownloadData(url);
                                        json_data = Encoding.UTF8.GetString(responseBytes);
                                    }
                                    catch (Exception)
                                    {
                                        json_data = "Error";
                                    }
                                    JsonSerializerSettings jsonSerializerSettings = new JsonSerializerSettings
                                    {
                                        DateTimeZoneHandling = DateTimeZoneHandling.Local
                                    };
                                    googleTimeZoneJSON = JsonConvert.DeserializeObject<GoogleTimeZoneJSON>(json_data, jsonSerializerSettings);
                                }

                                offset = (float)((googleTimeZoneJSON.rawOffset + googleTimeZoneJSON.dstOffset) / 3600.0f);

                                offsetText = offset.ToString("F1");
                                if (offsetText.EndsWith("0"))
                                {
                                    offsetText = offsetText.Substring(0, offsetText.Length - 2);
                                }
                            }

                            NodeCount += 1;
                            CurrentWebTideNode += 1;
                            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (CurrentWebTideNode * 100 / TotalWebTideNodes));
                            TideModel tideModel = new TideModel(tvFileService.ChoseEDriveOrCDrive(tvFileService.BasePath), mbcm.WebTideDataSet)
                            {
                                StartDate = StartDate.AddHours(-offset),
                                EndDate = EndDate.AddHours(-offset),
                                Lng = (double)mapInfoPointModelBC[i].Lng,
                                Lat = (double)mapInfoPointModelBC[i].Lat,
                                Steps_min = WebTideStepsInMinutes,
                                DoWaterLevels = false,
                            };
                            CurrentResults = tidesAndCurrentsService.GetCurrents(tideModel);
                            if (CurrentResults != null)
                            {
                                foreach (CurrentResult cr in CurrentResults)
                                {
                                    cr.Date = cr.Date.AddHours(offset);
                                }
                                AllCurrentResults.Add(CurrentResults);
                            }
                        }
                        dfsNewFile.CreateFile(newServerPathBC + tvFileModelBC.ServerFileName);
                        IDfsFile file = dfsNewFile.GetFile();
                        for (int i = 0; i < CurrentResults.ToList().Count; i++)
                        {
                            float[] floatArrayX = new float[AllCurrentResults.Count];
                            float[] floatArrayY = new float[AllCurrentResults.Count];

                            for (int j = 0; j < AllCurrentResults.Count; j++)
                            {
                                floatArrayX[j] = ((float)((List<CurrentResult>)AllCurrentResults[j].ToList())[i].x_velocity);
                                floatArrayY[j] = ((float)((List<CurrentResult>)AllCurrentResults[j].ToList())[i].y_velocity);
                            }

                            file.WriteItemTimeStepNext(0, floatArrayX);  // Current xVelocity
                            file.WriteItemTimeStepNext(0, floatArrayY);  // Current yVelocity
                        }
                        file.Close();

                        StringBuilder sb = new StringBuilder();
                        sb.AppendLine(TaskRunnerServiceRes.TimeOffsetFromGMT + ": " + offsetText + " " + TaskRunnerServiceRes.Hours + " --- " + googleTimeZoneJSON.timeZoneName);
                        sb.AppendLine(TaskRunnerServiceRes.From + " " + mikeScenarioModel.MikeScenarioStartDateTime_Local.ToString("yyyy MMMM dd HH:mm") +
                            " " + TaskRunnerServiceRes.To + " " + mikeScenarioModel.MikeScenarioEndDateTime_Local.ToString("yyyy MMMM dd HH:mm") + " (" +
                            TaskRunnerServiceRes.LocalTime + ")");
                        sb.AppendLine("");
                        sb.Append("\t");
                        for (int i = 0; i < NodeCount; i++)
                        {
                            sb.Append("\t\t" + TaskRunnerServiceRes.Node + i.ToString());
                        }
                        sb.AppendLine("");
                        sb.Append(TaskRunnerServiceRes.DateAndTime + "\t");
                        for (int i = 0; i < NodeCount; i++)
                        {
                            sb.Append("\txVel\tyVel");
                        }
                        sb.AppendLine("");
                        for (int i = 0; i < ((List<CurrentResult>)AllCurrentResults[0]).Count; i++)
                        {
                            sb.Append(((List<CurrentResult>)AllCurrentResults[0])[i].Date.ToString("yyyy MMMM dd HH:mm"));

                            for (int j = 0; j < AllCurrentResults.Count; j++)
                            {
                                float x_velocity = ((float)((List<CurrentResult>)AllCurrentResults[j].ToList())[i].x_velocity);
                                float y_velocity = ((float)((List<CurrentResult>)AllCurrentResults[j].ToList())[i].y_velocity);
                                string xText = (x_velocity < 0 ? x_velocity.ToString("F4") : " " + x_velocity.ToString("F4"));
                                string yText = (y_velocity < 0 ? y_velocity.ToString("F4") : " " + y_velocity.ToString("F4"));
                                sb.Append("\t" + xText + " " + yText);
                            }
                            sb.AppendLine();
                        }

                        mbcm.WebTideDataFromStartToEndDate = sb.ToString();

                        MikeBoundaryConditionModel mikeBoundaryConditionModelRet = mikeBoundaryConditionService.PostUpdateMikeBoundaryConditionDB(mbcm);
                        if (!string.IsNullOrWhiteSpace(mikeBoundaryConditionModelRet.Error))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.MIKEBoundaryCondition, mikeBoundaryConditionModelRet.Error);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.MIKEBoundaryCondition, mikeBoundaryConditionModelRet.Error);
                            return;
                        }
                    }
                    else
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.FileContainsTwoParamsButItsNotOfTypeUVAndVVItIs_And_, eumItemList[0].ToString(), eumItemList[1].ToString());
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("FileContainsTwoParamsButItsNotOfTypeUVAndVVItIs_And_", eumItemList[0].ToString(), eumItemList[1].ToString());
                        return;
                    }
                }
                else
                {
                    // this is not a file that is used for Water Level or Currents
                }

                FileInfo NewFileCreated = new FileInfo(newServerPathBC + tvFileModelBC.ServerFileName);

                if (!NewFileCreated.Exists)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, NewFileCreated.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", NewFileCreated.FullName);
                    return;
                }

                tvFileModelBC.FileSize_kb = (int)(NewFileCreated.Length / 1024);
                tvFileModelBC.FileCreatedDate_UTC = NewFileCreated.CreationTime;

                TVFileModel tvFileModelRet2 = tvFileService.PostUpdateTVFileDB(tvFileModelBC);
                if (!string.IsNullOrWhiteSpace(tvFileModelRet2.Error))
                {
                    // do something
                }
            }

            try
            {
                pfsFile.Write(fiM21_M3.FullName);
            }
            catch (Exception)
            {
                // nothing
            }
            pfsFile.Close();
            FixPFSFileSystemPart(fiM21_M3.FullName);

        }
        private void DoDecay(PFSFile pfsFile, FileInfo fiM21_M3, MikeScenarioModel mikeScenarioModel, TVFileModel tvFileModelM21_3FM)
        {
            string NotUsed = "";

            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            string BCFileName = "";

            BCFileName = GetParameterFileName(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/BOUNDARY_CONDITIONS/CODE_2", "file_name");
            if (string.IsNullOrWhiteSpace(BCFileName))
            {
                NotUsed = TaskRunnerServiceRes.CouldNotFindABoundaryConditionWithFileNameInM21_3FM;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("CouldNotFindABoundaryConditionWithFileNameInM21_3FM");
                return;
            }

            FileInfo fiBC = new FileInfo(BCFileName);

            List<TVFileModel> tvFileModelList = tvFileService.GetTVFileModelListWithParentTVItemIDDB(mikeScenarioModel.MikeScenarioTVItemID);
            TVFileModel tvFileModelBC = new TVFileModel();
            foreach (TVFileModel tvFileModel in tvFileModelList)
            {
                if (tvFileModel.ServerFileName == fiBC.Name)
                {
                    tvFileModelBC = tvFileModel;
                    break;
                }
            }

            fiBC = new FileInfo(tvFileService.ChoseEDriveOrCDrive(tvFileModelBC.ServerFilePath) + tvFileModelBC.ServerFileName);

            string NewDecayFileName = "Decay.dfs0";

            if (mikeScenarioModel.DecayIsConstant)
            {
                NotUsed = TaskRunnerServiceRes.SettingTransportModuleDecayContinuous;
                _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("SettingTransportModuleDecayContinuous"));
                _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 34);

                if (!SetParameterDouble(pfsFile, "FemEngineHD/TRANSPORT_MODULE/DECAY/COMPONENT_1", 0D, "format"))
                {
                    return;
                }

                NotUsed = TaskRunnerServiceRes.SettingTransportModuleDecayConstantValue;
                _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("SettingTransportModuleDecayConstantValue"));
                _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 35);

                if (!SetParameterDouble(pfsFile, "FemEngineHD/TRANSPORT_MODULE/DECAY/COMPONENT_1", ((float)mikeScenarioModel.DecayFactor_per_day) / 24 / 3600, "constant_value"))
                {
                    return;
                }

                NotUsed = TaskRunnerServiceRes.SettingTransportModuleDecayFileNameToEmpty;
                _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("SettingTransportModuleDecayFileNameToEmpty"));
                _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 36);

                if (!SetParameterFileName(pfsFile, "FemEngineHD/TRANSPORT_MODULE/DECAY/COMPONENT_1", @"", "file_name"))
                {
                    return;
                }

                FileInfo fi = new FileInfo(fiBC.Directory + "\\" + NewDecayFileName);

                foreach (TVFileModel tvFileModel in tvFileModelList)
                {
                    if (tvFileModel.ServerFileName == NewDecayFileName)
                    {
                        TVFileModel tvFileModelRet = tvFileService.PostDeleteTVFileWithTVItemIDDB(tvFileModel.TVFileTVItemID);
                        if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDelete_With_Equal_Error_, TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.TVFileTVItemID, tvFileModel.TVFileTVItemID.ToString(), tvFileModelRet.Error);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat4List("CouldNotDelete_With_Equal_Error_", TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.TVFileTVItemID, tvFileModel.TVFileTVItemID.ToString(), tvFileModelRet.Error);
                            return;
                        }

                        try
                        {
                            fi.Delete();
                        }
                        catch (Exception ex)
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDeleteFile_Error_, fi.FullName, ex.Message + " - " + (ex.InnerException != null ? ex.InnerException.Message : ""));
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDeleteFile_Error_", fi.FullName, ex.Message + " - " + (ex.InnerException != null ? ex.InnerException.Message : ""));
                            return;
                        }

                        break;
                    }
                }

            }
            else
            {
                NotUsed = TaskRunnerServiceRes.SettingTransportModuleDecayNotContinuous;
                _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("SettingTransportModuleDecayNotContinuous"));
                _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 34);

                if (!SetParameterDouble(pfsFile, "FemEngineHD/TRANSPORT_MODULE/DECAY/COMPONENT_1", 1 /* 0 = continuous, 1 = Not continuous*/, "format"))
                {
                    return;
                }

                NotUsed = TaskRunnerServiceRes.SettingTransportModuleDecayConstantValue;
                _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("SettingTransportModuleDecayConstantValue"));
                _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 35);

                if (!SetParameterDouble(pfsFile, "FemEngineHD/TRANSPORT_MODULE/DECAY/COMPONENT_1", ((float)mikeScenarioModel.DecayFactor_per_day) / 24 / 3600, "constant_value"))
                {
                    return;
                }

                // creating the varying decay file

                NotUsed = TaskRunnerServiceRes.CreatingFileDecayDotDFS0;
                _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("CreatingFileDecayDotDFS0"));
                _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 36);

                DfsFactory factory = new DfsFactory();

                FileInfo fi = new FileInfo(fiBC.Directory + "\\" + NewDecayFileName);

                IDfsFile dfsOldFile = DfsFileFactory.DfsGenericOpen(fiBC.Directory + "\\" + tvFileModelBC.ServerFileName);

                DfsBuilder dfsNewFile = DfsBuilder.Create("Decay", dfsOldFile.FileInfo.ApplicationTitle + " - Decay", dfsOldFile.FileInfo.ApplicationVersion);

                double DecayStepsInMinutes = ((double)((IDfsEqCalendarAxis)((dfsOldFile.FileInfo).TimeAxis)).TimeStep / 60);

                DateTime? DateTimeTemp = GetParameterStartTime(pfsFile, "FemEngineHD/TIME", "start_time");
                if (DateTimeTemp == null)
                {
                    return;
                }

                int? NumberOfTimeSteps = GetParameterInt(pfsFile, "FemEngineHD/TIME", "number_of_time_steps");
                if (NumberOfTimeSteps == null)
                {
                    return;
                }

                int? TimeStepInterval = GetParameterInt(pfsFile, "FemEngineHD/TIME", "time_step_interval");
                if (TimeStepInterval == null)
                {
                    return;
                }

                DateTime StartDate = ((DateTime)DateTimeTemp).AddHours(-1);
                DateTime EndDate = ((DateTime)DateTimeTemp).AddSeconds((int)NumberOfTimeSteps * (int)TimeStepInterval).AddHours(1);

                dfsNewFile.SetDataType(1);
                dfsNewFile.SetGeographicalProjection(dfsOldFile.FileInfo.Projection);
                dfsNewFile.SetTemporalAxis(factory.CreateTemporalEqCalendarAxis(eumUnit.eumUsec, StartDate, 0, DecayStepsInMinutes * 60));
                dfsNewFile.SetItemStatisticsType(StatType.RegularStat);

                DfsDynamicItemBuilder item = dfsNewFile.CreateDynamicItemBuilder();
                item.Set("Decay", eumQuantity.Create(eumItem.eumIDecayFactor, eumUnit.eumUperSec), DfsSimpleType.Float);
                item.SetValueType(DataValueType.Instantaneous);
                item.SetAxis(factory.CreateAxisEqD0());
                item.SetReferenceCoordinates(1f, 2f, 3f);
                dfsNewFile.AddDynamicItem(item.GetDynamicItemInfo());

                dfsOldFile.Close();

                string[] NewFileErrors = dfsNewFile.Validate();
                StringBuilder sbErr = new StringBuilder();
                foreach (string s in NewFileErrors)
                {
                    sbErr.AppendLine(s);
                }

                if (NewFileErrors.Count() > 0)
                {
                    throw new Exception(sbErr.ToString());
                }

                dfsNewFile.CreateFile(fi.FullName);
                IDfsFile DecayFile = dfsNewFile.GetFile();

                //float AccumulatedTime = 0f;
                DateTime NewDateTime = StartDate;
                while (NewDateTime < EndDate)
                {
                    double DecayValue = (float)mikeScenarioModel.DecayFactor_per_day + (float)mikeScenarioModel.DecayFactorAmplitude * Math.Cos((double)(((double)(NewDateTime.Hour * 3600 + NewDateTime.Minute * 60 + NewDateTime.Second) / (double)86400) + (double)0.5) * 2 * Math.PI);

                    DecayFile.WriteItemTimeStepNext(0, new float[] { (float)(DecayValue / 24 / 3600) });  // Decay

                    NewDateTime = NewDateTime.AddSeconds(DecayStepsInMinutes * 60);
                }

                DecayFile.Close();

                FileInfo fiDecay = new FileInfo(fiBC.Directory + "\\" + NewDecayFileName);

                if (!fiDecay.Exists)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiBC.Directory + "\\" + NewDecayFileName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiBC.Directory + "\\" + NewDecayFileName);
                    return;
                }

                NotUsed = string.Format(TaskRunnerServiceRes.StoringDecayFile_InDB, NewDecayFileName);
                _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat1List("StoringDecayFile_InDB", NewDecayFileName));
                _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 37);

                TVFileModel tvFileModelDecay = tvFileService.GetTVFileModelWithServerFilePathAndServerFileNameDB(tvFileModelBC.ServerFilePath, NewDecayFileName);
                if (!string.IsNullOrWhiteSpace(tvFileModelDecay.Error))
                {
                    TVItemModel tvItemFileModel = tvItemService.PostAddChildTVItemDB(mikeScenarioModel.MikeScenarioTVItemID, NewDecayFileName, TVTypeEnum.File);
                    if (!string.IsNullOrEmpty(tvItemFileModel.Error))
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_For_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.MIKEScenarioDocumentation, TaskRunnerServiceRes.TVText, NewDecayFileName);
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat4List("CouldNotCreate_For_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.MIKEScenarioDocumentation, TaskRunnerServiceRes.TVText, NewDecayFileName);
                        return;
                    }

                    TVFileModel tvFileModelDecayNew = new TVFileModel()
                    {
                        TVFileTVItemID = tvItemFileModel.TVItemID,
                        FilePurpose = FilePurposeEnum.MikeInput,
                        Language = _TaskRunnerBaseService._BWObj.appTaskModel.Language,
                        Year = DateTime.Now.Year,
                        FileDescription = "Mike Decay File",
                        FileType = tvFileService.GetFileType(fiDecay.Extension),
                        FileSize_kb = (int)(fiDecay.Length / 1024),
                        FileInfo = "Mike Scenario Documentation Decay File",
                        FileCreatedDate_UTC = fiDecay.CreationTime.ToUniversalTime(),
                        ServerFileName = NewDecayFileName,
                        ServerFilePath = tvFileModelBC.ServerFilePath,
                    };

                    TVFileModel tvFileModelRet = tvFileService.PostAddTVFileDB(tvFileModelDecayNew);
                    if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.TVFileDecayNew, tvFileModelRet.Error);
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.TVFileDecayNew, tvFileModelRet.Error);
                        return;
                    }
                }
                else
                {
                    tvFileModelDecay.ServerFileName = fiDecay.Name;
                    tvFileModelDecay.FilePurpose = FilePurposeEnum.MikeInput;
                    tvFileModelDecay.FileDescription = "Mike Decay File";
                    tvFileModelDecay.FileType = tvFileService.GetFileType(fiDecay.Extension);
                    tvFileModelDecay.FileSize_kb = (int)(fiDecay.Length / 1024);
                    tvFileModelDecay.FileCreatedDate_UTC = fiDecay.CreationTime.ToUniversalTime();
                    tvFileModelDecay.ServerFileName = NewDecayFileName;
                    tvFileModelDecay.ServerFilePath = tvFileModelBC.ServerFilePath;

                    TVFileModel tvFileModelRet = tvFileService.PostUpdateTVFileDB(tvFileModelDecay);
                    if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, "TVFileDecayNew", tvFileModelRet.Error);
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.TVFileDecayNew, tvFileModelRet.Error);
                        return;
                    }
                }

                NotUsed = TaskRunnerServiceRes.SettingTransportModuleDecayFileName;
                _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat1List("SettingTransportModuleDecayFileName", NewDecayFileName));
                _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 38);

                if (!SetParameterFileName(pfsFile, "FemEngineHD/TRANSPORT_MODULE/DECAY/COMPONENT_1", @"|.\" + NewDecayFileName + "|", "file_name"))
                {
                    return;
                }
            }
        }
        private void DoSource(PFSFile pfsFile, MikeSourceModel msm, FileInfo fiBC, MikeScenarioModel mikeScenarioModel, int LastSourceIndex)
        {
            string NotUsed = "";

            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            TVItemLanguageService tvItemLanguageService = new TVItemLanguageService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemLanguageModel tvItemLanguageModelSourceName = tvItemLanguageService.GetTVItemLanguageModelWithTVItemIDAndLanguageDB(msm.MikeSourceTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.Language);
            if (!string.IsNullOrWhiteSpace(tvItemLanguageModelSourceName.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItemLanguage, TaskRunnerServiceRes.TVItemID, msm.MikeSourceTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItemLanguage, TaskRunnerServiceRes.TVItemID, msm.MikeSourceTVItemID.ToString());
                return;
            }

            MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            List<MapInfoPointModel> mapInfoPointModelSourceList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(msm.MikeSourceTVItemID, TVTypeEnum.MikeSource, MapInfoDrawTypeEnum.Point);
            if (mapInfoPointModelSourceList.Count == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItemLanguage, TaskRunnerServiceRes.TVItemID, msm.MikeSourceTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItemLanguage, TaskRunnerServiceRes.TVItemID, msm.MikeSourceTVItemID.ToString());
                return;
            }

            PFSSection pfsSectionSource = null;
            try
            {
                pfsSectionSource = pfsFile.GetSectionFromHandle("FemEngineHD/HYDRODYNAMIC_MODULE/SOURCES/" + msm.SourceNumberString);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetSection", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                return;
            }

            if (pfsSectionSource == null)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.AddNewSource_Number_InM21FMOrM3MFFile, msm.MikeSourceTVText, msm.SourceNumberString);
                _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat2List("AddNewSource_Number_InM21FMOrM3MFFile", msm.MikeSourceTVText, msm.SourceNumberString));

                string ret = AddNewSourceInPFS(pfsFile, msm.SourceNumberString, LastSourceIndex);
                if (!string.IsNullOrWhiteSpace(ret))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAddNewSection_, ret);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotAddNewSource_", ret);
                    return;
                }
            }

            NotUsed = string.Format(TaskRunnerServiceRes.SettingHydrodynamicModuleSourceName_Number_, msm.MikeSourceTVText, msm.SourceNumberString);
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat2List("SettingHydrodynamicModuleSourceName_Number_", msm.MikeSourceTVText, msm.SourceNumberString));

            if (!SetParameterString(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/SOURCES/" + msm.SourceNumberString, tvItemLanguageModelSourceName.TVText, "Name"))
            {
                return;
            }

            NotUsed = string.Format(TaskRunnerServiceRes.SettingHydrodynamicModuleSourceIncludeSourceName_Number_, msm.MikeSourceTVText, msm.SourceNumberString);
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat2List("SettingHydrodynamicModuleSourceIncludeSourceName_Number_", msm.MikeSourceTVText, msm.SourceNumberString));

            if (!SetParameterInt(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/SOURCES/" + msm.SourceNumberString, msm.Include == true ? (int)1 : (int)0, "include"))
            {
                return;
            }

            NotUsed = string.Format(TaskRunnerServiceRes.SettingHydrodynamicModuleSourceCoordinatesSourceName_Number_, msm.MikeSourceTVText, msm.SourceNumberString);
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat2List("SettingHydrodynamicModuleSourceCoordinatesSourceName_Number_", msm.MikeSourceTVText, msm.SourceNumberString));

            Coord coord = new Coord() { Lat = (float)mapInfoPointModelSourceList[0].Lat, Lng = (float)mapInfoPointModelSourceList[0].Lng, Ordinal = 0 };
            if (!SetParameterCoordinates(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/SOURCES/" + msm.SourceNumberString, coord, "coordinates"))
            {
                return;
            }

            string NewPolConcSourceFileName = tvItemLanguageModelSourceName.TVText + "_PollutionConcentration.dfs0";
            FileInfo fiPolConc = new FileInfo(fiBC.Directory + "\\" + NewPolConcSourceFileName);

            string NewDischargeSourceFileName = tvItemLanguageModelSourceName.TVText + "_Discharge.dfs0";
            FileInfo fiDischarge = new FileInfo(fiBC.Directory + "\\" + NewDischargeSourceFileName);

            if (msm.IsContinuous && !msm.UseHydrometric)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.SettingTransportModuleSourceFormatSourceName_Number_, msm.MikeSourceTVText, msm.SourceNumberString);
                _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat2List("SettingTransportModuleSourceFormatSourceName_Number_", msm.MikeSourceTVText, msm.SourceNumberString));

                if (!SetParameterInt(pfsFile, "FemEngineHD/TRANSPORT_MODULE/SOURCES/" + msm.SourceNumberString + @"/COMPONENT_1", 0 /* continuoues */, "format"))
                {
                    return;
                }

                NotUsed = string.Format(TaskRunnerServiceRes.SettingTransportModuleSourceName_Number_, msm.MikeSourceTVText, msm.SourceNumberString);
                _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat2List("SettingTransportModuleSourceName_Number_", msm.MikeSourceTVText, msm.SourceNumberString));

                if (!SetParameterString(pfsFile, "FemEngineHD/TRANSPORT_MODULE/SOURCES/" + msm.SourceNumberString, tvItemLanguageModelSourceName.TVText, "Name"))
                {
                    return;
                }

                MikeSourceStartEndService mikeSourceStartEndService = new MikeSourceStartEndService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                MikeSourceStartEndModel mikeSourceStartEndModel = mikeSourceStartEndService.GetMikeSourceStartEndModelListWithMikeSourceIDDB(msm.MikeSourceID).FirstOrDefault();
                if (mikeSourceStartEndModel == null)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeSourceStartEnd, TaskRunnerServiceRes.MikeSourceID, msm.MikeSourceID.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeSourceStartEnd, TaskRunnerServiceRes.MikeSourceID, msm.MikeSourceID.ToString());
                    return;
                }

                NotUsed = string.Format(TaskRunnerServiceRes.SettingHydrodynamicModuleSourceConstantValueSourceName_Number_, msm.MikeSourceTVText, msm.SourceNumberString);
                _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat2List("SettingHydrodynamicModuleSourceConstantValueSourceName_Number_", msm.MikeSourceTVText, msm.SourceNumberString));

                if (!SetParameterDouble(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/SOURCES/" + msm.SourceNumberString, (float)(mikeSourceStartEndModel.SourceFlowStart_m3_day / 24 / 3600), "constant_value"))
                {
                    return;
                }

                NotUsed = string.Format(TaskRunnerServiceRes.SettingTransportModuleSourceConstantValueSourceName_Number_, msm.MikeSourceTVText, msm.SourceNumberString);
                _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat2List("SettingTransportModuleSourceConstantValueSourceName_Number_", msm.MikeSourceTVText, msm.SourceNumberString));

                if (!SetParameterDouble(pfsFile, "FemEngineHD/TRANSPORT_MODULE/SOURCES/" + msm.SourceNumberString + "/COMPONENT_1", (float)mikeSourceStartEndModel.SourcePollutionStart_MPN_100ml, "constant_value"))
                {
                    return;
                }

                NotUsed = string.Format(TaskRunnerServiceRes.SettingTransportModuleTemperatureConstantValueSourceName_Number_, msm.MikeSourceTVText, msm.SourceNumberString);
                _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat2List("SettingTransportModuleTemperatureConstantValueSourceName_Number_", msm.MikeSourceTVText, msm.SourceNumberString));

                if (!SetParameterDouble(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/TEMPERATURE_SALINITY_MODULE/SOURCES/" + msm.SourceNumberString + "/TEMPERATURE", (float)(mikeSourceStartEndModel.SourceTemperatureStart_C), "constant_value"))
                {
                    return;
                }

                NotUsed = string.Format(TaskRunnerServiceRes.SettingTransportModuleSalinityConstantValueSourceName_Number_, msm.MikeSourceTVText, msm.SourceNumberString);
                _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat2List("SettingTransportModuleSalinityConstantValueSourceName_Number_", msm.MikeSourceTVText, msm.SourceNumberString));

                if (!SetParameterDouble(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/TEMPERATURE_SALINITY_MODULE/SOURCES/" + msm.SourceNumberString + "/SALINITY", (float)(mikeSourceStartEndModel.SourceSalinityStart_PSU), "constant_value"))
                {
                    return;
                }

                List<TVFileModel> tvFileModelList = tvFileService.GetTVFileModelListWithParentTVItemIDDB(mikeScenarioModel.MikeScenarioTVItemID);

                foreach (TVFileModel tvFileModel in tvFileModelList)
                {
                    if (tvFileModel.ServerFileName == NewPolConcSourceFileName)
                    {
                        TVFileModel tvFileModelRet = tvFileService.PostDeleteTVFileWithTVItemIDDB(tvFileModel.TVFileTVItemID);
                        if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDelete_With_Equal_Error_, TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.TVFileTVItemID, tvFileModel.TVFileTVItemID.ToString(), tvFileModelRet.Error);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat4List("CouldNotDelete_With_Equal_Error_", TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.TVFileTVItemID, tvFileModel.TVFileTVItemID.ToString(), tvFileModelRet.Error);
                            return;
                        }

                        try
                        {
                            fiPolConc.Delete();
                        }
                        catch (Exception ex)
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDeleteFile_Error_, fiPolConc.FullName, ex.Message + " - " + (ex.InnerException != null ? ex.InnerException.Message : ""));
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDeleteFile_Error_", fiPolConc.FullName, ex.Message + " - " + (ex.InnerException != null ? ex.InnerException.Message : ""));
                            return;
                        }

                        break;
                    }
                }

            }
            else
            {
                NotUsed = string.Format(TaskRunnerServiceRes.SettingTransportModuleSourceFormatSourceName_Number_, msm.MikeSourceTVText, msm.SourceNumberString);
                _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat2List("SettingTransportModuleSourceFormatSourceName_Number_", msm.MikeSourceTVText, msm.SourceNumberString));

                if (!SetParameterInt(pfsFile, "FemEngineHD/TRANSPORT_MODULE/SOURCES/" + msm.SourceNumberString + @"/COMPONENT_1", 1 /* not continuoues */, "format"))
                {
                    return;
                }

                NotUsed = string.Format(TaskRunnerServiceRes.CreatingSourcePolDFSOFileSourceName_Number_, msm.MikeSourceTVText, msm.SourceNumberString);
                _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat2List("CreatingSourcePolDFSOFileSourceName_Number_", msm.MikeSourceTVText, msm.SourceNumberString));

                if (!CreatePollutionConcentrationSource(pfsFile, msm, fiBC, mikeScenarioModel, fiPolConc, tvItemLanguageModelSourceName, NewPolConcSourceFileName, tvItemService, tvFileService))
                {
                    return;
                }

                NotUsed = string.Format(TaskRunnerServiceRes.SettingTransportModuleSourceFileNameSourceName_Number_, msm.MikeSourceTVText, msm.SourceNumberString);
                _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat2List("SettingTransportModuleSourceFileNameSourceName_Number_", msm.MikeSourceTVText, msm.SourceNumberString));

                if (!SetParameterFileName(pfsFile, "FemEngineHD/TRANSPORT_MODULE/SOURCES/" + msm.SourceNumberString + "/COMPONENT_1", @".\" + NewPolConcSourceFileName, "file_name"))
                {
                    return;
                }

                NotUsed = string.Format(TaskRunnerServiceRes.CreatingSourceFlowDFSOFileSourceName_Number_, msm.MikeSourceTVText, msm.SourceNumberString);
                _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat2List("CreatingSourceFlowDFSOFileSourceName_Number_", msm.MikeSourceTVText, msm.SourceNumberString));

                if (!CreateDischargeSource(pfsFile, msm, fiBC, mikeScenarioModel, fiDischarge, tvItemLanguageModelSourceName, NewDischargeSourceFileName, tvItemService, tvFileService))
                {
                    return;
                }

                NotUsed = string.Format(TaskRunnerServiceRes.SettingHydrodynamicModuleSourceFileNameSourceName_Number_, msm.MikeSourceTVText, msm.SourceNumberString);
                _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat2List("SettingHydrodynamicModuleSourceFileNameSourceName_Number_", msm.MikeSourceTVText, msm.SourceNumberString));

                if (!SetParameterFileName(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/SOURCES/" + msm.SourceNumberString, @".\" + NewPolConcSourceFileName, "file_name"))
                {
                    return;
                }
            }
        }
        public void DoSources(PFSFile pfsFile, FileInfo fiM21_M3, MikeScenarioModel mikeScenarioModel, int TVItemID, TVFileModel tvFileModelM21_3FM)
        {
            string NotUsed = "";

            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            string BCFileName = GetParameterFileName(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/BOUNDARY_CONDITIONS/CODE_2", "file_name");

            if (string.IsNullOrWhiteSpace(BCFileName))
            {
                NotUsed = TaskRunnerServiceRes.CouldNotFindABoundaryConditionWithFileNameInM21_3FM;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("CouldNotFindABoundaryConditionWithFileNameInM21_3FM");
                return;
            }

            FileInfo fiBC = new FileInfo(BCFileName);

            List<TVFileModel> tvFileModelList = tvFileService.GetTVFileModelListWithParentTVItemIDDB(mikeScenarioModel.MikeScenarioTVItemID);
            TVFileModel tvFileModelBC = new TVFileModel();
            foreach (TVFileModel tvFileModel in tvFileModelList)
            {
                if (tvFileModel.ServerFileName == fiBC.Name)
                {
                    tvFileModelBC = tvFileModel;
                    break;
                }
            }

            fiBC = new FileInfo(tvFileService.ChoseEDriveOrCDrive(tvFileModelBC.ServerFilePath) + tvFileModelBC.ServerFileName);

            MikeSourceService mikeSourceService = new MikeSourceService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            List<MikeSourceModel> mikeSourceModelList = mikeSourceService.GetMikeSourceModelListWithMikeScenarioTVItemIDDB(mikeScenarioModel.MikeScenarioTVItemID);

            int LastSourceIndex = 0;
            for (int i = 1; i < 200; i++)
            {
                try
                {
                    PFSSection pfsSectionSource = pfsFile.GetSectionFromHandle("FemEngineHD/HYDRODYNAMIC_MODULE/SOURCES/SOURCE_" + i.ToString());
                    if (pfsSectionSource != null)
                    {
                        LastSourceIndex = i;
                    }
                }
                catch (Exception)
                {
                    break;
                }
            }

            int SourceCount = 0;
            foreach (MikeSourceModel msm in mikeSourceModelList.OrderBy(c => c.SourceNumberString))
            {
                SourceCount += 1;

                NotUsed = string.Format(TaskRunnerServiceRes.DoingSource_, msm.MikeSourceTVText);
                _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat1List("DoingSource_", msm.MikeSourceTVText));
                _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 46 + SourceCount);

                DoSource(pfsFile, msm, fiBC, mikeScenarioModel, LastSourceIndex);
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                {
                    return;
                }
            }

            if (SourceCount < LastSourceIndex)
            {
                SourceCount += 1;

                for (int i = SourceCount; i < LastSourceIndex + 1; i++)
                {
                    RemoveSource(pfsFile, i);
                    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                    {
                        return;
                    }
                }
            }

            UpdateSourceCount(pfsFile, mikeSourceModelList.Count());
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
            {
                return;
            }
        }
        private void UpdateSourceCount(PFSFile pfsFile, int SourceCount)
        {
            string NotUsed = "";

            List<string> pathList = new List<string>()
            {
                "FemEngineHD/HYDRODYNAMIC_MODULE/SOURCES/",
                "FemEngineHD/HYDRODYNAMIC_MODULE/TEMPERATURE_SALINITY_MODULE/SOURCES/",
                "FemEngineHD/HYDRODYNAMIC_MODULE/TURBULENCE_MODULE/SOURCES/",
                "FemEngineHD/TRANSPORT_MODULE/SOURCES/",
                "FemEngineHD/ECOLAB_MODULE/SOURCES/",
                "FemEngineHD/MUD_TRANSPORT_MODULE/SOURCES/",
                "FemEngineHD/SAND_TRANSPORT_MODULE/SOURCES/",
            };

            foreach (string path in pathList)
            {
                PFSSection pfsSectionSources = null;
                try
                {
                    pfsSectionSources = pfsFile.GetSectionFromHandle(path);
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetSection", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return;
                }

                PFSKeyword pfsKeywork = null;
                try
                {
                    pfsKeywork = pfsSectionSources.GetKeyword("MzSEPfsListItemCount");

                    pfsKeywork.DeleteParameter(1);
                    PFSParameter pfsParameter = pfsKeywork.InsertNewParameterInt(SourceCount, 1);
                }
                catch (Exception)
                {
                    continue;
                    //NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    //_TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    //return;
                }

                try
                {
                    pfsKeywork.DeleteParameter(1);
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "DeleteKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return;
                }

                try
                {
                    PFSParameter pfsParameter = pfsKeywork.InsertNewParameterInt(SourceCount, 1);
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "InsertKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return;
                }

                if (path == "FemEngineHD/HYDRODYNAMIC_MODULE/SOURCES/")
                {
                    try
                    {
                        pfsKeywork = pfsSectionSources.GetKeyword("number_of_sources");

                        pfsKeywork.DeleteParameter(1);
                        PFSParameter pfsParameter = pfsKeywork.InsertNewParameterInt(SourceCount, 1);
                    }
                    catch (Exception ex)
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                        return;
                    }

                    try
                    {
                        pfsKeywork.DeleteParameter(1);
                    }
                    catch (Exception ex)
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "DeleteKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                        return;
                    }

                    try
                    {
                        PFSParameter pfsParameter = pfsKeywork.InsertNewParameterInt(SourceCount, 1);
                    }
                    catch (Exception ex)
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "InsertKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                        return;
                    }

                }
            }
        }
        public void FixPFSFileSystemPart(string FullFileName)
        {
            StreamReader sr = new StreamReader(FullFileName, Encoding.GetEncoding("iso-8859-1"));

            FileInfo fi = new FileInfo(FullFileName);
            string FileText = sr.ReadToEnd();
            sr.Close();

            if (FileText.IndexOf("[SYSTEM]") > 0)
            {
                FileText = FileText.Substring(0, FileText.IndexOf("[SYSTEM]"));
            }
            StringBuilder sb = new StringBuilder(FileText);
            sb.AppendLine(@"[SYSTEM]");
            sb.AppendLine(@"   ResultRootFolder = ||");
            sb.AppendLine(@"   UseCustomResultFolder = true");
            sb.AppendLine(@"   CustomResultFolder = |.\|");
            sb.AppendLine(@"EndSect  // SYSTEM");

            StreamWriter sw = new StreamWriter(FullFileName, false, Encoding.GetEncoding("iso-8859-1"));
            sw.Write(sb.ToString());
            sw.Close();
        }
        private void GetAllInputFilesToUpload(string ServerPath, string m21_3fmFileName, List<OtherFileInfo> FileNameList, MikeScenarioModel NewMikeScenarioModel, PFSFile pfsFile)
        {
            string NotUsed = "";
            FileInfo fiOld = new FileInfo(m21_3fmFileName);
            FileInfo fi = new FileInfo(ServerPath + fiOld.Name);

            if (!fi.Exists)
            {
                NotUsed = TaskRunnerServiceRes.CouldNotFindFile_;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", ServerPath + fiOld.Name);
                return;
            }

            try
            {
                string FileName = GetParameterFileNameOnlyText(pfsFile, "FemEngineHD/DOMAIN", "file_name");
                if (!string.IsNullOrWhiteSpace(FileName))
                {
                    fi = new FileInfo(fiOld.DirectoryName + @"\" + FileName);

                    if (!SetParameterFileName(pfsFile, "FemEngineHD/DOMAIN", @".\" + fi.Name, "file_name"))
                    {
                        return;
                    }

                    AddFileToFileNameList(ServerPath, fi.FullName, FileNameList, false);
                    // also want to load the .mdf file (the mesh generator file)
                    AddFileToFileNameList(ServerPath, fi.FullName.Replace(".mesh", ".mdf"), FileNameList, false);
                }


                PFSSection pfsSectionBoundaryCondition = pfsFile.GetSectionFromHandle("FemEngineHD/HYDRODYNAMIC_MODULE/BOUNDARY_CONDITIONS");
                if (pfsSectionBoundaryCondition != null)
                {
                    for (int i = 2; i < 1000; i++)
                    {
                        PFSSection pfsSectionBoundaryConditionCode = pfsFile.GetSectionFromHandle("FemEngineHD/HYDRODYNAMIC_MODULE/BOUNDARY_CONDITIONS/CODE_" + i.ToString());
                        if (pfsSectionBoundaryConditionCode == null)
                        {
                            break;
                        }
                        FileName = GetParameterFileNameOnlyText(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/BOUNDARY_CONDITIONS/CODE_" + i.ToString(), "file_name");
                        if (!string.IsNullOrWhiteSpace(FileName))
                        {
                            fi = new FileInfo(fiOld.DirectoryName + @"\" + FileName);

                            if (!SetParameterFileName(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/BOUNDARY_CONDITIONS/CODE_" + i.ToString(), @".\" + fi.Name, "file_name"))
                            {
                                return;
                            }
                            AddFileToFileNameList(ServerPath, fi.FullName, FileNameList, false);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotReadInputFileError_, ex.Message + (ex.InnerException != null ? ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotReadInputFileError_", ex.Message + (ex.InnerException != null ? ex.InnerException.Message : ""));
                return;
            }

        }
        private void GetAllResultFilesToUpload(string ServerPath, string m21_3fmFileName, List<OtherFileInfo> FileNameList, MikeScenarioModel NewMikeScenarioModel, PFSFile pfsFile)
        {
            string NotUsed = "";
            FileInfo fiOld = new FileInfo(m21_3fmFileName);
            FileInfo fi = new FileInfo(ServerPath + fiOld.Name);

            if (!fi.Exists)
            {
                NotUsed = TaskRunnerServiceRes.CouldNotFindFile_;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", ServerPath + fiOld.Name);
                return;
            }

            try
            {
                string ResultRootFolder = GetParameterString(pfsFile, "SYSTEM", "ResultRootFolder");
                if (string.IsNullOrWhiteSpace(ResultRootFolder))
                {
                    return;
                }

                // already verified if the number of output is 1
                string FileName = GetParameterFileNameOnlyText(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/OUTPUTS/OUTPUT_1", "file_name");
                if (string.IsNullOrWhiteSpace(FileName))
                {
                    return;
                }
                fi = new FileInfo(fiOld.DirectoryName + @"\" + ResultRootFolder + @"\" + fiOld.Name + @" - Result Files\" + FileName);
                if (!SetParameterFileName(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/OUTPUTS/OUTPUT_1", @".\" + fi.Name, "file_name"))
                {
                    return;
                }
                AddFileToFileNameList(ServerPath, fi.FullName, FileNameList, true);

                // already verified if the number of output is 1
                FileName = GetParameterFileNameOnlyText(pfsFile, "FemEngineHD/TRANSPORT_MODULE/OUTPUTS/OUTPUT_1", "file_name");
                if (string.IsNullOrWhiteSpace(FileName))
                {
                    return;
                }
                fi = new FileInfo(fiOld.DirectoryName + @"\" + ResultRootFolder + @"\" + fiOld.Name + @" - Result Files\" + FileName);
                if (!SetParameterFileName(pfsFile, "FemEngineHD/TRANSPORT_MODULE/OUTPUTS/OUTPUT_1", @".\" + fi.Name, "file_name"))
                {
                    return;
                }
                AddFileToFileNameList(ServerPath, fi.FullName, FileNameList, true);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotReadResultFileError_, ex.Message + (ex.InnerException != null ? ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotReadResultFileError_", ex.Message + (ex.InnerException != null ? ex.InnerException.Message : ""));
                return;
            }

        }
        private string GetParameterFileName(PFSFile pfsFile, string Path, string Keyword)
        {
            string NotUsed = "";
            string FileName = "";

            PFSSection pfsSection = pfsFile.GetSectionFromHandle(Path);

            if (pfsSection == null)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindPFSSectionWithPath_, Path);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindPFSSectionWithPath_", Path);
                return null;
            }

            PFSKeyword keyword = null;
            try
            {
                keyword = pfsSection.GetKeyword(Keyword);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                return FileName;
            }

            if (keyword != null)
            {
                try
                {
                    FileName = keyword.GetParameter(1).ToFileNamePath();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return FileName;
                }
            }

            return FileName;
        }
        private string GetParameterFileNameOnlyText(PFSFile pfsFile, string Path, string Keyword)
        {
            string NotUsed = "";
            string FileName = "";

            PFSSection pfsSection = pfsFile.GetSectionFromHandle(Path);

            if (pfsSection == null)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindPFSSectionWithPath_, Path);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindPFSSectionWithPath_", Path);
                return null;
            }

            PFSKeyword keyword = null;
            try
            {
                keyword = pfsSection.GetKeyword(Keyword);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                return FileName;
            }

            if (keyword != null)
            {
                try
                {
                    FileName = keyword.GetParameter(1).ToString();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return FileName;
                }
            }

            return FileName;
        }
        private double? GetParameterDouble(PFSFile pfsFile, string Path, string Keyword)
        {
            string NotUsed = "";
            double? MannningNumber = null;
            PFSSection pfsSection = pfsFile.GetSectionFromHandle(Path);

            if (pfsSection == null)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindPFSSectionWithPath_, Path);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindPFSSectionWithPath_", Path);
                return null;
            }

            PFSKeyword keyword = null;
            try
            {
                keyword = pfsSection.GetKeyword(Keyword);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                return MannningNumber;
            }

            if (keyword != null)
            {
                PFSParameter param = null;
                try
                {
                    param = keyword.GetParameter(1);
                    MannningNumber = param.ToDouble();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return MannningNumber;
                }
            }

            return MannningNumber;
        }
        private void GetMoreFilesToImportFromMDFFile(int MikeScenarioTVItemID, int TVFileTVItemID, string ClientFullPath, string ServerFullPath)
        {
            string NotUsed = "";
            if (MikeScenarioTVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBe0, TaskRunnerServiceRes.MikeScenarioTVItemID, MikeScenarioTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("_ShouldNotBe0", TaskRunnerServiceRes.MikeScenarioTVItemID, MikeScenarioTVItemID.ToString());
                return;
            }

            if (TVFileTVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBe0, TaskRunnerServiceRes.TVFileID, TVFileTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("_ShouldNotBe0", TaskRunnerServiceRes.TVFileID, TVFileTVItemID.ToString());
                return;
            }

            if (string.IsNullOrEmpty(ClientFullPath))
            {
                NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBeNullOrEmpty, TaskRunnerServiceRes.ClientFullPath);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_ShouldNotBeNullOrEmpty", TaskRunnerServiceRes.ClientFullPath);
                return;
            }

            if (string.IsNullOrEmpty(ServerFullPath))
            {
                NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBeNullOrEmpty, TaskRunnerServiceRes.ServerFullPath);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_ShouldNotBeNullOrEmpty", TaskRunnerServiceRes.ServerFullPath);
                return;
            }

            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemModel tvItemModelMikeScenario = tvItemService.GetTVItemModelWithTVItemIDDB(MikeScenarioTVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelMikeScenario.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                return;
            }


            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVFileModel tvFileModel = tvFileService.GetTVFileModelWithTVFileTVItemIDDB(TVFileTVItemID);
            if (!string.IsNullOrWhiteSpace(tvFileModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.TVFileID, TVFileTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.TVFileID, TVFileTVItemID.ToString());
                return;
            }

            if ((tvFileModel.ServerFilePath + tvFileModel.ServerFileName) != ServerFullPath)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._NotEqualTo_, ServerFullPath, tvFileModel.ServerFilePath + tvFileModel.ServerFileName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("_NotEqualTo_", ServerFullPath, tvFileModel.ServerFilePath + tvFileModel.ServerFileName);
                return;
            }

            // ready to save the file uploaded
            string ServerFilePath = tvFileService.GetServerFilePath(MikeScenarioTVItemID);

            FileInfo fi = new FileInfo(ServerFilePath + tvFileModel.ServerFileName);

            if (!fi.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, ServerFilePath + tvFileModel.ServerFileName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", ServerFilePath + tvFileModel.ServerFileName);
                return;
            }

            PFSFile pfsFile = new PFSFile(fi.FullName);

            try
            {
                PFSSection pfsSection = pfsFile.GetTarget("MESH_EDITOR_SCATTERDATAFILE_SETTINGS", 1);
                PFSSection pfsSection2 = pfsSection.GetSection("MESH_EDITOR_SCATTERDATAFILE", 1);
                for (int i = 1; i < pfsSection2.GetSectionsCount() + 1; i++)
                {
                    PFSSection pfsSection3 = pfsSection2.GetSection("SCATTER_DATA_I", i);
                    PFSKeyword keyword = pfsSection3.GetKeyword("File_name");
                    PFSParameter param = keyword.GetParameter(1);
                    string PartialFileName = param.ToString();
                    string ClientFullPathNew = ClientFullPath.Substring(0, ClientFullPath.LastIndexOf(@"\") + 1);

                    FileInfo fiClient = new FileInfo(ClientFullPathNew + @"\" + PartialFileName);

                    string ClientPath = fiClient.FullName.Substring(0, fiClient.FullName.LastIndexOf(@"\") + 1);

                    string ServerFullPathNew = ServerFullPath.Substring(0, ServerFullPath.LastIndexOf(@"\") + 1);

                    FileInfo fiServer = new FileInfo(ServerFullPathNew + @"\" + PartialFileName);

                    keyword.DeleteParameter(1);
                    keyword.InsertNewParameterFileName(@".\" + fiServer.Name + "", 1);

                    string FileName = fiClient.FullName.Substring(fiClient.FullName.LastIndexOf(@"\") + 1);

                    TVItemModel tvItemFileModel = tvItemService.PostAddChildTVItemDB(tvItemModelMikeScenario.TVItemID, FileName, TVTypeEnum.File);
                    if (!string.IsNullOrEmpty(tvItemFileModel.Error))
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_For_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.MIKEScenarioDocumentation, TaskRunnerServiceRes.TVText, FileName);
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat4List("CouldNotCreate_For_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.MIKEScenarioDocumentation, TaskRunnerServiceRes.TVText, FileName);
                        pfsFile.Close();
                        return;
                    }

                    TVFileModel tvFileModelNew = new TVFileModel()
                    {
                        TVFileTVItemID = tvItemFileModel.TVItemID,
                        FilePurpose = FilePurposeEnum.MikeInputMDF,
                        Language = _TaskRunnerBaseService._BWObj.appTaskModel.Language,
                        Year = DateTime.Now.Year,
                        FileDescription = "Mike Scenario File",
                        FileType = tvFileService.GetFileType(fiClient.Extension),
                        FileSize_kb = 0,
                        FileInfo = "Mike Scenario File",
                        FileCreatedDate_UTC = DateTime.UtcNow,
                        ClientFilePath = ClientPath,
                        ServerFileName = FileName,
                        ServerFilePath = ServerFilePath.Replace(@"C:\", @"E:\"),
                    };

                    TVFileModel tvFileModelRet = tvFileService.PostAddTVFileDB(tvFileModelNew);
                    if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreate_For_With_Equal_", TaskRunnerServiceRes.TVFile, tvFileModelRet.Error);
                        pfsFile.Close();
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                NotUsed = ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                pfsFile.Close();
                return;
            }

            try
            {
                pfsFile.Write(fi.FullName);
            }
            catch (Exception)
            {
                // nothing
            }
            pfsFile.Close();
        }
        private int? GetParameterInt(PFSFile pfsFile, string Path, string Keyword)
        {
            string NotUsed = "";
            int? NumberOfOutputs = null;

            PFSSection pfsSection = pfsFile.GetSectionFromHandle(Path);

            if (pfsSection == null)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindPFSSectionWithPath_, Path);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindPFSSectionWithPath_", Path);
                return null;
            }

            PFSKeyword keyword = null;
            try
            {
                keyword = pfsSection.GetKeyword(Keyword);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                return NumberOfOutputs;
            }

            if (keyword != null)
            {
                try
                {
                    NumberOfOutputs = keyword.GetParameter(1).ToInt();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return NumberOfOutputs;
                }
            }

            return NumberOfOutputs;
        }
        private Coord GetParameterCoord(PFSFile pfsFile, string Path, string Keyword)
        {
            string NotUsed = "";
            Coord SourceCoord = new Coord() { Lat = 0.0f, Lng = 0.0f, Ordinal = 0 };
            PFSKeyword pfsKeywordCoord = null;

            PFSSection pfsSection = pfsFile.GetSectionFromHandle(Path);

            if (pfsSection == null)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindPFSSectionWithPath_, Path);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindPFSSectionWithPath_", Path);
                return null;
            }

            try
            {
                pfsKeywordCoord = pfsSection.GetKeyword(Keyword);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                return SourceCoord;
            }

            if (pfsKeywordCoord != null)
            {
                try
                {
                    float Lng = (float)pfsKeywordCoord.GetParameter(1).ToDouble();
                    float Lat = (float)pfsKeywordCoord.GetParameter(2).ToDouble();
                    SourceCoord = new Coord() { Lat = (float)Lat, Lng = (float)Lng, Ordinal = 0 };
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return SourceCoord;
                }
            }

            return SourceCoord;
        }
        private string GetParameterString(PFSFile pfsFile, string Path, string Keyword)
        {
            string NotUsed = "";
            string Name = "";
            PFSKeyword pfsKeywordName = null;

            PFSSection pfsSection = pfsFile.GetSectionFromHandle(Path);

            if (pfsSection == null)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindPFSSectionWithPath_, Path);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindPFSSectionWithPath_", Path);
                return null;
            }

            try
            {
                pfsKeywordName = pfsSection.GetKeyword(Keyword);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                return Name;
            }

            if (pfsKeywordName != null)
            {
                try
                {
                    Name = pfsKeywordName.GetParameter(1).ToString();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return Name;
                }
            }

            if (Name.StartsWith("'"))
            {
                Name = Name.Substring(1);
            }
            if (Name.EndsWith("'"))
            {
                Name = Name.Substring(0, Name.Length - 1);
            }

            return Name;
        }
        private DateTime? GetParameterStartTime(PFSFile pfsFile, string Path, string Keyword)
        {
            string NotUsed = "";
            DateTime? dateTime = null;

            PFSSection pfsSection = pfsFile.GetSectionFromHandle(Path);

            if (pfsSection == null)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindPFSSectionWithPath_, Path);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindPFSSectionWithPath_", Path);
                return null;
            }

            PFSKeyword keyword = null;
            try
            {
                keyword = pfsSection.GetKeyword("start_time");
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                return dateTime;
            }

            if (keyword != null)
            {
                try
                {
                    int Year = keyword.GetParameter(1).ToInt();
                    int Month = keyword.GetParameter(2).ToInt();
                    int Day = keyword.GetParameter(3).ToInt();
                    int Hour = keyword.GetParameter(4).ToInt();
                    int Minute = keyword.GetParameter(5).ToInt();
                    int Second = keyword.GetParameter(6).ToInt();
                    dateTime = new DateTime(Year, Month, Day, Hour, Minute, Second);
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return dateTime;
                }
            }

            return dateTime;
        }
        private bool IsEachSourceContinuous(PFSFile pfsFile)
        {
            string NotUsed = "";
            string SectionSourcePath = "";
            string SectionComponentPath = "";

            for (int i = 1; i < 1000; i++)
            {
                SectionSourcePath = "FemEngineHD/TRANSPORT_MODULE/SOURCES/SOURCE_" + i.ToString();
                PFSSection pfsSectionSource = pfsFile.GetSectionFromHandle(SectionSourcePath);

                if (pfsSectionSource == null)
                {
                    break;
                }

                for (int j = 1; j < 1000; j++)
                {
                    SectionComponentPath = SectionSourcePath + "/COMPONENT_" + j.ToString();
                    PFSSection pfsSectionComponent = pfsFile.GetSectionFromHandle(SectionComponentPath);

                    if (pfsSectionComponent == null)
                    {
                        break;
                    }

                    PFSKeyword keyword = null;
                    try
                    {
                        keyword = pfsSectionComponent.GetKeyword("format");
                    }
                    catch (Exception ex)
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                        return false;
                    }

                    if (keyword != null)
                    {
                        PFSParameter param = null;
                        int format = -1;
                        try
                        {
                            param = keyword.GetParameter(1);
                            format = param.ToInt();
                        }
                        catch (Exception ex)
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                            return false;
                        }

                        if (format != 0)
                        {
                            NotUsed = TaskRunnerServiceRes.OnlyAllowedConstantPollutionFlow;
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("OnlyAllowedContantPollutionFlow");
                            return false;
                        }
                    }
                }
            }

            return true;
        }
        private bool IsTransportModuleDecayConstant(PFSFile pfsFile)
        {
            string NotUsed = "";
            string SectionDecayComponentPath = "";

            for (int i = 1; i < 1000; i++)
            {
                SectionDecayComponentPath = "FemEngineHD/TRANSPORT_MODULE/DECAY/COMPONENT_" + i.ToString();
                PFSSection pfsSectionDecayComponent = pfsFile.GetSectionFromHandle(SectionDecayComponentPath);

                if (pfsSectionDecayComponent == null)
                {
                    break;
                }

                PFSKeyword keyword = null;
                try
                {
                    keyword = pfsSectionDecayComponent.GetKeyword("format");
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return false;
                }

                if (keyword != null)
                {
                    PFSParameter param = null;
                    int format = -1;
                    try
                    {
                        param = keyword.GetParameter(1);
                        format = param.ToInt();
                    }
                    catch (Exception ex)
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                        return false;
                    }

                    if (format != 0)
                    {
                        NotUsed = TaskRunnerServiceRes.OnlyAllowedConstantDecayRate;
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("OnlyAllowedContantDecayRate");
                        return false;
                    }
                }
            }

            return true;
        }
        private void RemoveSource(PFSFile pfsFile, int SourceIndex)
        {
            string NotUsed = "";

            List<string> pathList = new List<string>()
            {
                "FemEngineHD/HYDRODYNAMIC_MODULE/SOURCES/",
                "FemEngineHD/HYDRODYNAMIC_MODULE/TEMPERATURE_SALINITY_MODULE/SOURCES/",
                "FemEngineHD/HYDRODYNAMIC_MODULE/TURBULENCE_MODULE/SOURCES/",
                "FemEngineHD/TRANSPORT_MODULE/SOURCES/",
                "FemEngineHD/ECOLAB_MODULE/SOURCES/",
                "FemEngineHD/MUD_TRANSPORT_MODULE/SOURCES/",
                "FemEngineHD/SAND_TRANSPORT_MODULE/SOURCES/",
            };

            foreach (string path in pathList)
            {
                PFSSection pfsSectionSources = null;
                try
                {
                    pfsSectionSources = pfsFile.GetSectionFromHandle(path);
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetSection", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return;
                }

                try
                {
                    pfsSectionSources.DeleteSection("SOURCE_" + SourceIndex.ToString(), 1);
                }
                catch (Exception)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDeleteSection_, path + "SOURCE_" + SourceIndex.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotDeleteSection_", path + "SOURCE_" + SourceIndex.ToString());
                    return;
                }
            }
        }
        private string ReturnStrLimit(string TempStr, int NumbOfChar)
        {
            StringBuilder RetString = new StringBuilder();
            if (TempStr != null)
            {
                if (TempStr.Length < NumbOfChar)
                {
                    for (int i = 0; i < (NumbOfChar - TempStr.Length); i++)
                    {
                        RetString.Append(" ");
                    }
                    RetString.Append(TempStr);
                }
                else
                {
                    RetString.Append(TempStr);
                }
            }
            return RetString.ToString();
        }
        private void RunScenario(int AppTaskID, int MikeScenarioTVItemID, int ContactTVItemID)
        {
            string NotUsed = "";

            MikeScenarioService mikeScenarioService = new MikeScenarioService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            MikeScenarioModel mikeScenarioModel = mikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(MikeScenarioTVItemID);
            if (!string.IsNullOrWhiteSpace(mikeScenarioModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.MikeScenarioTVItemID, MikeScenarioTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.MikeScenarioTVItemID, MikeScenarioTVItemID.ToString());
                return;
            }

            if (mikeScenarioModel.ScenarioStatus == ScenarioStatusEnum.Running)
            {
                NotUsed = TaskRunnerServiceRes.MikeScenarioAlreadyRunning;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("MikeScenarioAlreadyRunning");
                return;
            }

            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVFileModel tvFileModelM21_3FM = tvFileService.GetTVFileModelWithTVItemIDAndTVFileTypeM21FMOrM3FMDB(MikeScenarioTVItemID);
            if (!string.IsNullOrWhiteSpace(tvFileModelM21_3FM.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_StartingWith_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_StartingWith_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                return;
            }

            string ServerPath = tvFileService.ChoseEDriveOrCDrive(tvFileService.GetServerFilePath(MikeScenarioTVItemID));

            FileInfo fiM21_M3 = new FileInfo(ServerPath + tvFileModelM21_3FM.ServerFileName);

            if (!fiM21_M3.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiM21_M3.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiM21_M3.FullName);
                return;
            }

            PFSFile pfsFile = new PFSFile(fiM21_M3.FullName);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
            {
                return;
            }

            string FileNameTemp = GetParameterFileNameOnlyText(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/OUTPUTS/OUTPUT_1", "file_name");
            if (string.IsNullOrWhiteSpace(FileNameTemp))
            {
                return;
            }

            FileInfo fi = new FileInfo(ServerPath + FileNameTemp);

            FileInfo fiServer = new FileInfo(tvFileService.GetServerFilePath(MikeScenarioTVItemID) + "\\" + fi.Name);
            string ServerHydroName = tvFileService.ChoseEDriveOrCDrive(fiServer.FullName);

            string FileNameTemp2 = GetParameterFileNameOnlyText(pfsFile, "FemEngineHD/TRANSPORT_MODULE/OUTPUTS/OUTPUT_1", "file_name");
            if (string.IsNullOrWhiteSpace(FileNameTemp2))
            {
                return;
            }

            fi = new FileInfo(ServerPath + FileNameTemp2);

            FileInfo fiServer2 = new FileInfo(tvFileService.GetServerFilePath(MikeScenarioTVItemID) + "\\" + fi.Name);

            string ServerTransName = tvFileService.ChoseEDriveOrCDrive(fiServer2.FullName);

            if (string.IsNullOrEmpty(ServerHydroName) && string.IsNullOrEmpty(ServerTransName))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, ServerHydroName + ", " + ServerTransName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", ServerHydroName + " | " + ServerTransName);
                return;
            }

            _TaskRunnerBaseService.SendStatusToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, AppTaskStatusEnum.Running);
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 3);
            _TaskRunnerBaseService.SendTaskCommandToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, AppTaskCommandEnum.MikeScenarioRunning);

            mikeScenarioModel.ScenarioStatus = ScenarioStatusEnum.Running;
            MikeScenarioModel mikeScenarioModelRet = mikeScenarioService.PostUpdateMikeScenarioDB(mikeScenarioModel);
            if (!string.IsNullOrWhiteSpace(mikeScenarioModelRet.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.MikeScenario, mikeScenarioModelRet.Error);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.MikeScenario, mikeScenarioModelRet.Error);
                return;
            }

            // starting the Updater.exe application
            ProcessStartInfo pInfoUpdater = new ProcessStartInfo();
            pInfoUpdater.Arguments = @" """ + AppTaskID + @""" "
                + @" """ + MikeScenarioTVItemID + @""" "
                + @" """ + ContactTVItemID + @""" ";
            pInfoUpdater.WindowStyle = ProcessWindowStyle.Minimized;
            pInfoUpdater.UseShellExecute = true;

            Process processUpdater = new Process();
            processUpdater.StartInfo = pInfoUpdater;
            try
            {
                //pInfoUpdater.FileName = @"C:\CSSP_Execs\Updater\Debug\Updater.exe";
                pInfoUpdater.FileName = @"C:\CSSP Latest Code\Updater\bin\Debug\Updater.exe";
                processUpdater.Start();
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotRun_Error_, pInfoUpdater.FileName, ex.Message);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotRun_Error_", pInfoUpdater.FileName, ex.Message);
                return;
            }

            processUpdater.WaitForInputIdle(2000);

        }
        private void SaveDBInfoToMikeScenarioFile(int MikeScenarioTVItemID)
        {
            string NotUsed = "";
            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            DfsuFile dfsuFile = null;
            if (MikeScenarioTVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                return;
            }

            NotUsed = TaskRunnerServiceRes.ReadingUserInput;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("ReadingUserInput"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 5);

            MikeScenarioService mikeScenarioService = new MikeScenarioService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MikeScenarioModel mikeScenarioModel = mikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(MikeScenarioTVItemID);
            if (!string.IsNullOrWhiteSpace(mikeScenarioModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                return;
            }

            TVItemLanguageService tvItemLanguageService = new TVItemLanguageService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemLanguageModel tvItemLangModelMikeScenario = tvItemLanguageService.GetTVItemLanguageModelWithTVItemIDAndLanguageDB(MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.Language);
            if (!string.IsNullOrWhiteSpace(tvItemLangModelMikeScenario.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItemLanguage, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItemLanguage, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                return;
            }

            TVFileModel tvFileModelM21_3FM = tvFileService.GetTVFileModelWithTVItemIDAndTVFileTypeM21FMOrM3FMDB(MikeScenarioTVItemID);
            if (!string.IsNullOrWhiteSpace(tvFileModelM21_3FM.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_StartingWith_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_StartingWith_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                return;
            }

            NotUsed = TaskRunnerServiceRes.TweakScenarioStartTimeToHave60PercHighTideAndFalling;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("TweakScenarioStartTimeToHave60PercHighTideAndFalling"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 10);

            if (!TweakScenarioStartTimeToHave60PercHighTideAndFalling(mikeScenarioModel))
            {
                return;
            }

            string ServerPath = tvFileService.GetServerFilePath(MikeScenarioTVItemID);

            FileInfo fiM21_M3 = new FileInfo(ServerPath + tvFileModelM21_3FM.ServerFileName);

            if (!fiM21_M3.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiM21_M3.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiM21_M3.FullName);
                return;
            }

            if (mikeScenarioModel.UseSalinityAndTemperatureInitialConditionFromTVFileTVItemID != null && mikeScenarioModel.UseSalinityAndTemperatureInitialConditionFromTVFileTVItemID > 0)
            {
                NotUsed = TaskRunnerServiceRes.CreatingSalinityAndTemperatureInitialConditionsFiles;
                _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("CreatingSalinityAndTemperatureInitialConditionsFiles"));
                _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 15);

                if (!string.IsNullOrWhiteSpace(CreateInititalConditionFileSalAndTempFromTVFileItemID(MikeScenarioTVItemID, (int)mikeScenarioModel.UseSalinityAndTemperatureInitialConditionFromTVFileTVItemID, fiM21_M3)))
                {
                    return;
                }
            }

            PFSFile pfsFile = new PFSFile(fiM21_M3.FullName);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
            {
                return;
            }

            TimeSpan timeSpan = new TimeSpan(((DateTime)mikeScenarioModel.MikeScenarioEndDateTime_Local).Ticks - ((DateTime)mikeScenarioModel.MikeScenarioStartDateTime_Local).Ticks);

            NotUsed = TaskRunnerServiceRes.SettingStartTimeInM21FMOrM3FMFile;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("SettingStartTimeInM21FMOrM3FMFile"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 17);

            if (!SetParameterTimeStart(pfsFile, "FemEngineHD/TIME", (DateTime)mikeScenarioModel.MikeScenarioStartDateTime_Local, "start_time"))
            {
                return;
            }

            NotUsed = TaskRunnerServiceRes.SettingNumberOfTimeStepsInM21FMOrM3FMFile;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("SettingNumberOfTimeStepsInM21FMOrM3FMFile"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 20);

            if (!SetParameterInt(pfsFile, "FemEngineHD/TIME", (int)(timeSpan.Days * 24 * 60 + timeSpan.Hours * 60 + timeSpan.Minutes), "number_of_time_steps"))
            {
                return;
            }

            NotUsed = TaskRunnerServiceRes.GettingApproximateSizeOfHydrodynamicResultFile;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("GettingApproximateSizeOfHydrodynamicResultFile"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 21);

            //string DfsuFileName = fiM21_M3.FullName;

            string FileName = GetParameterFileNameOnlyText(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/OUTPUTS/OUTPUT_1", "file_name");

            FileInfo fiDfsu = new FileInfo(ServerPath + FileName);

            string delimStr = ",";
            char[] delimiter = delimStr.ToCharArray();
            int NumberOfElements = 0;
            int NumberOfTimeSteps = 0;
            int NumberOfSigmaLayers = 0;
            int NumberOfZLayers = 0;
            int NumberOfHydroOutputParameters = 0;
            int NumberOfTransOutputParameters = 0;
            long EstimatedHydroFileSize = 0;
            long EstimatedTransFileSize = 0;
            bool HydroExist = false;
            bool TransExist = false;
            string HydroDfsu = "";
            string TransDfsu = "";

            if (fiDfsu.Exists)
            {
                HydroDfsu = fiDfsu.FullName;
                try
                {
                    dfsuFile = DfsuFile.Open(fiDfsu.FullName);
                    HydroExist = true;

                    NumberOfElements = dfsuFile.NumberOfElements;
                    NumberOfTimeSteps = (int)(((timeSpan.Days * 24 * 60) + timeSpan.Hours * 60 + timeSpan.Minutes) / (int)mikeScenarioModel.ResultFrequency_min);
                    NumberOfSigmaLayers = dfsuFile.NumberOfSigmaLayers;
                    NumberOfZLayers = dfsuFile.NumberOfLayers - NumberOfSigmaLayers;
                    NumberOfHydroOutputParameters = dfsuFile.ItemInfo.Count();
                    if (tvFileModelM21_3FM.FileType == FileTypeEnum.M21FM)
                    {
                        EstimatedHydroFileSize = 9 * (NumberOfElements) + ((NumberOfElements * NumberOfTimeSteps * NumberOfHydroOutputParameters) * 4);
                    }
                    else // should be of type .m3fm
                    {
                        EstimatedHydroFileSize = 9 * (NumberOfElements) + ((NumberOfElements * NumberOfTimeSteps * (NumberOfHydroOutputParameters + 1)) * 4);
                    }

                    dfsuFile.Close();

                }
                catch (Exception)
                {
                    // nothing for now
                }
            }

            NotUsed = TaskRunnerServiceRes.GettingApproximateSizeOfTransportResultFile;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("GettingApproximateSizeOfTransportResultFile"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 23);

            // try transport module result file
            FileName = GetParameterFileNameOnlyText(pfsFile, "FemEngineHD/TRANSPORT_MODULE/OUTPUTS/OUTPUT_1", "file_name");

            FileInfo fiDfsuTrans = new FileInfo(ServerPath + FileName);

            if (fiDfsuTrans.Exists)
            {
                TransDfsu = fiDfsuTrans.FullName;
                try
                {
                    dfsuFile = DfsuFile.Open(fiDfsuTrans.FullName);
                    TransExist = true;

                    if (!HydroExist)
                    {
                        NumberOfElements = dfsuFile.NumberOfElements;
                        NumberOfTimeSteps = (int)(((timeSpan.Days * 24 * 60) + timeSpan.Hours * 60 + timeSpan.Minutes) / (int)mikeScenarioModel.ResultFrequency_min);
                        NumberOfSigmaLayers = dfsuFile.NumberOfSigmaLayers;
                        NumberOfZLayers = dfsuFile.NumberOfLayers - NumberOfSigmaLayers;
                    }
                    NumberOfTransOutputParameters = dfsuFile.ItemInfo.Count();
                    if (tvFileModelM21_3FM.FileType == FileTypeEnum.M21FM)
                    {
                        EstimatedTransFileSize = 9 * (NumberOfElements) + ((NumberOfElements * NumberOfTimeSteps * NumberOfTransOutputParameters) * 4);
                    }
                    else // should be of type .m3fm
                    {
                        EstimatedTransFileSize = 9 * (NumberOfElements) + ((NumberOfElements * NumberOfTimeSteps * (NumberOfTransOutputParameters + 1)) * 4);
                    }

                }
                catch (Exception)
                {
                    // nothing for now
                }
            }


            if (!TransExist)
            {
                FileName = GetParameterFileNameOnlyText(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/OUTPUT/OUTPUT_1", "file_name");

                fiDfsu = new FileInfo(ServerPath + FileName);

                if (!fiDfsu.Exists)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiDfsu.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiDfsu.FullName);
                    return;
                }

                try
                {
                    dfsuFile = DfsuFile.Open(fiDfsu.FullName);
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotOpenFile_Error_, fiDfsu.FullName, ex.Message);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotOpenFile_Error_", fiDfsu.FullName, ex.Message);
                    return;
                }

            }

            if (dfsuFile == null)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, TransDfsu + " | " + HydroDfsu);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", TransDfsu + " | " + HydroDfsu);
                return;
            }

            dfsuFile.Close();

            try
            {
                NotUsed = TaskRunnerServiceRes.PartialSaveOfM21FMOrM3FMFile;
                _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("PartialSaveOfM21FMOrM3FMFile"));
                _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 25);

                pfsFile.Write(fiM21_M3.FullName);
                pfsFile.Close();
            }
            catch (Exception ex)
            {
                string InnerException = ex.InnerException != null ? $" InnerException: { ex.InnerException.Message }" : "";
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotWriteFile_Error_, $"[{ fiM21_M3.FullName }]", $"{ ex.Message } { InnerException }");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotWriteFile_Error_", $"[{ fiM21_M3.FullName }]", $"{ ex.Message } { InnerException }");
                return;
            }

            if (tvFileModelM21_3FM.ServerFileName.Replace(".m21fm", "").Replace(".m3fm", "") != tvItemLangModelMikeScenario.TVText)
            {
                // should rename the m21fm or m3fm file on the server to reflect the new ScenarioName

                fiM21_M3 = new FileInfo(fiM21_M3.FullName);
                //fiM21_M3.CopyTo(tvFileModelM21_3FM.ServerFilePath + tvFileModelM21_3FM.ServerFileName);

                //fiM21_M3.Delete();

                if (!fiM21_M3.Exists)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, ServerPath + tvItemLangModelMikeScenario.TVText + fiM21_M3.Extension);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", ServerPath + tvItemLangModelMikeScenario.TVText + fiM21_M3.Extension);
                    return;
                }

                //string OldFileName = tvFileModelM21_3FM.ServerFileName;

                tvFileModelM21_3FM.ServerFileName = mikeScenarioModel.MikeScenarioTVText + fiM21_M3.Extension;

                TVFileModel tvFileModelRet = tvFileService.PostUpdateTVFileDB(tvFileModelM21_3FM);
                if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.TVFile, tvFileModelRet.Error.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.TVFile, tvFileModelRet.Error.ToString());
                    return;
                }
            }

            pfsFile = new PFSFile(ServerPath + mikeScenarioModel.MikeScenarioTVText + fiM21_M3.Extension);

            fiM21_M3 = new FileInfo(ServerPath + mikeScenarioModel.MikeScenarioTVText + fiM21_M3.Extension);

            int? NumberOfTimeSteps2 = GetParameterInt(pfsFile, "FemEngineHD/TIME", "number_of_time_steps");
            if (NumberOfTimeSteps2 == null)
            {
                return;
            }

            NotUsed = TaskRunnerServiceRes.SettingHydrodynamicOutputLastTimeStep;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("SettingHydrodynamicOutputLastTimeStep"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 27);

            if (!SetParameterInt(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/OUTPUTS/OUTPUT_1", (int)NumberOfTimeSteps2, "last_time_step"))
            {
                return;
            }

            NotUsed = TaskRunnerServiceRes.SettingHydrodynamicOutputTimeStepFrequency;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("SettingHydrodynamicOutputTimeStepFrequency"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 28);

            if (!SetParameterInt(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/OUTPUTS/OUTPUT_1", (int)mikeScenarioModel.ResultFrequency_min, "time_step_frequency"))
            {
                return;
            }

            NotUsed = TaskRunnerServiceRes.SettingTransportOutputLastTimeStep;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("SettingTransportOutputLastTimeStep"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 29);

            if (!SetParameterInt(pfsFile, "FemEngineHD/TRANSPORT_MODULE/OUTPUTS/OUTPUT_1", (int)NumberOfTimeSteps2, "last_time_step"))
            {
                return;
            }

            NotUsed = TaskRunnerServiceRes.SettingTransportOutputTimeStepFrequency;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("SettingTransportOutputTimeStepFrequency"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 30);

            if (!SetParameterInt(pfsFile, "FemEngineHD/TRANSPORT_MODULE/OUTPUTS/OUTPUT_1", (int)mikeScenarioModel.ResultFrequency_min, "time_step_frequency"))
            {
                return;
            }

            mikeScenarioModel.MikeScenarioStartExecutionDateTime_Local = null;
            mikeScenarioModel.MikeScenarioExecutionTime_min = null;


            NotUsed = TaskRunnerServiceRes.SettingHydrodynamicModuleWindSpeed;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("SettingHydrodynamicModuleWindSpeed"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 31);

            if (!SetParameterDouble(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/WIND_FORCING", (float)mikeScenarioModel.WindSpeed_km_h / (float)3.6, "constant_speed"))
            {
                return;
            }

            NotUsed = TaskRunnerServiceRes.SettingHydrodynamicModuleWindDirection;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("SettingHydrodynamicModuleWindDirection"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 32);

            if (!SetParameterDouble(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/WIND_FORCING", (float)mikeScenarioModel.WindDirection_deg, "constant_direction"))
            {
                return;
            }

            NotUsed = TaskRunnerServiceRes.DoingDecay;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("DoingDecay"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 33);

            DoDecay(pfsFile, fiM21_M3, mikeScenarioModel, tvFileModelM21_3FM);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            int? NumberOfTimeSteps3 = GetParameterInt(pfsFile, "FemEngineHD/TIME", "number_of_time_steps");
            if (NumberOfTimeSteps3 == null)
            {
                return;
            }

            NotUsed = TaskRunnerServiceRes.SettingTransportModuleLastTimeStep;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("SettingTransportModuleLastTimeStep"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 39);

            if (!SetParameterInt(pfsFile, "FemEngineHD/TRANSPORT_MODULE/OUTPUTS/OUTPUT_1", (int)NumberOfTimeSteps3, "last_time_step"))
            {
                return;
            }

            NotUsed = TaskRunnerServiceRes.SettingTransportModuleTimeStepFrequency;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("SettingTransportModuleTimeStepFrequency"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 40);

            if (!SetParameterInt(pfsFile, "FemEngineHD/TRANSPORT_MODULE/OUTPUTS/OUTPUT_1", (int)mikeScenarioModel.ResultFrequency_min, "time_step_frequency"))
            {
                return;
            }

            NotUsed = TaskRunnerServiceRes.SettingHydrodynamicModuleSalinity;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("SettingHydrodynamicModuleSalinity"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 41);

            if (!SetParameterDouble(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/DENSITY", (float)mikeScenarioModel.AmbientSalinity_PSU, "salinity_reference"))
            {
                return;
            }

            NotUsed = TaskRunnerServiceRes.SettingHydrodynamicModuleTemperature;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("SettingHydrodynamicModuleTemperature"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 42);

            if (!SetParameterDouble(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/DENSITY", (float)mikeScenarioModel.AmbientTemperature_C, "temperature_reference"))
            {
                return;
            }

            NotUsed = TaskRunnerServiceRes.SettingHydrodynamicModuleManning;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("SettingHydrodynamicModuleManning"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 43);

            if (!SetParameterDouble(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/BED_RESISTANCE/MANNING_NUMBER", (float)mikeScenarioModel.ManningNumber, "constant_value"))
            {
                return;
            }

            string TempFileName = "";
            string SalFileName = "";
            int format = 0;
            if (mikeScenarioModel.UseSalinityAndTemperatureInitialConditionFromTVFileTVItemID != null)
            {
                TempFileName = @".\temperature.dfsu";
                SalFileName = @".\salinity.dfsu";
                format = 2;

            }

            NotUsed = TaskRunnerServiceRes.SettingHydrodynamicModuleInitialConditionTemperatureFileName;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("SettingHydrodynamicModuleInitialConditionTemperatureFileName"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 43);

            if (!SetParameterFileName(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/TEMPERATURE_SALINITY_MODULE/INITIAL_CONDITIONS/TEMPERATURE", TempFileName, "file_name"))
            {
                return;
            }

            NotUsed = TaskRunnerServiceRes.SettingHydrodynamicModuleInitialConditionTemperatureFormat;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("SettingHydrodynamicModuleInitialConditionTemperatureFormat"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 43);

            if (!SetParameterInt(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/TEMPERATURE_SALINITY_MODULE/INITIAL_CONDITIONS/TEMPERATURE", format, "format"))
            {
                return;
            }

            NotUsed = TaskRunnerServiceRes.SettingHydrodynamicModuleInitialConditionSalinityFileName;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("SettingHydrodynamicModuleInitialConditionSalinityFileName"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 43);

            if (!SetParameterFileName(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/TEMPERATURE_SALINITY_MODULE/INITIAL_CONDITIONS/SALINITY", SalFileName, "file_name"))
            {
                return;
            }


            NotUsed = TaskRunnerServiceRes.SettingHydrodynamicModuleInitialConditionSalinityFormat;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("SettingHydrodynamicModuleInitialConditionSalinityFormat"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 43);

            if (!SetParameterInt(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/TEMPERATURE_SALINITY_MODULE/INITIAL_CONDITIONS/SALINITY", format, "format"))
            {
                return;
            }

            mikeScenarioModel.NumberOfElements = NumberOfElements;
            mikeScenarioModel.NumberOfTimeSteps = (int)NumberOfTimeSteps3;
            mikeScenarioModel.NumberOfSigmaLayers = NumberOfSigmaLayers;
            mikeScenarioModel.NumberOfZLayers = NumberOfZLayers;
            mikeScenarioModel.NumberOfHydroOutputParameters = NumberOfHydroOutputParameters;
            mikeScenarioModel.NumberOfTransOutputParameters = NumberOfTransOutputParameters;
            mikeScenarioModel.EstimatedHydroFileSize = EstimatedHydroFileSize;
            mikeScenarioModel.EstimatedTransFileSize = EstimatedTransFileSize;

            TVItemModel tvItemModelMikeScenario = tvItemService.GetTVItemModelWithTVItemIDDB(mikeScenarioModel.MikeScenarioTVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelMikeScenario.Error))
            {
                // do something
            }

            TVItemModel tvItemModelMikeScenarioParent = tvItemService.GetTVItemModelWithTVItemIDDB((int)tvItemModelMikeScenario.ParentID);
            if (!string.IsNullOrWhiteSpace(tvItemModelMikeScenarioParent.Error))
            {
                // do something
            }

            NotUsed = TaskRunnerServiceRes.UpdatingMikeScenarioInfoInDB;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("UpdatingMikeScenarioInfoInDB"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 44);

            MikeScenarioModel mikeScenarioModelRet = mikeScenarioService.PostUpdateMikeScenarioDB(mikeScenarioModel);
            if (!string.IsNullOrWhiteSpace(mikeScenarioModelRet.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.MikeScenario, mikeScenarioModelRet.Error);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.MikeScenario, mikeScenarioModelRet.Error);
                return;
            }

            NotUsed = TaskRunnerServiceRes.DoingSources;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("DoingSources"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 45);

            DoSources(pfsFile, fiM21_M3, mikeScenarioModel, MikeScenarioTVItemID, tvFileModelM21_3FM);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
            {
                return;
            }

            try
            {
                pfsFile.Write(fiM21_M3.FullName);
                pfsFile.Close();
            }
            catch (Exception ex)
            {
                string InnerException = ex.InnerException != null ? $" InnerException: { ex.InnerException.Message }" : "";
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotWriteFile_Error_, $"[{ fiM21_M3.FullName }]", $"{ ex.Message } { InnerException }");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotWriteFile_Error_", $"[{ fiM21_M3.FullName }]", $"{ ex.Message } { InnerException }");
                return;
            }

            //try
            //{
            //    pfsFile.Write(ServerPath + mikeScenarioModel.MikeScenarioTVText + fiM21_M3.Extension);
            //}
            //catch (Exception)
            //{
            //    // nothing
            //}
            //pfsFile.Close();
            FixPFSFileSystemPart(fiM21_M3.FullName);

            FileInfo fiLog = new FileInfo(fiM21_M3.FullName.Replace(".m21fm", ".log").Replace(".m3fm", ".log"));

            StreamWriter sw = fiLog.CreateText();
            sw.Write("empty log");
            sw.Close();

            NotUsed = TaskRunnerServiceRes.CreatingALogFileIfItDoesNotYetExist;
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageList("CreatingALogFileIfItDoesNotYetExist"));
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 80);

            // does the log file exist if not create it
            TVFileModel tvFileModelLog = tvFileService.GetTVFileModelWithTVItemIDAndTVFileTypeLogDB(MikeScenarioTVItemID);
            if (!string.IsNullOrWhiteSpace(tvFileModelLog.Error))
            {
                TVItemModel tvItemModelExist = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(MikeScenarioTVItemID, fiLog.Name, TVTypeEnum.File);
                if (!string.IsNullOrWhiteSpace(tvItemModelExist.Error))
                {
                    tvItemModelExist = tvItemService.PostAddChildTVItemDB(MikeScenarioTVItemID, fiLog.Name, TVTypeEnum.File);
                    if (!string.IsNullOrWhiteSpace(tvItemModelExist.Error))
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_For_With_Equal_, TaskRunnerServiceRes.TVItem, "log", TaskRunnerServiceRes.TVText, fiLog.Name);
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat4List("CouldNotCreate_For_With_Equal_", TaskRunnerServiceRes.TVFile, "log", TaskRunnerServiceRes.TVText, fiLog.Name);
                        return;
                    }
                }
                fiLog = new FileInfo(fiM21_M3.FullName.Replace(".m21fm", ".log").Replace(".m3fm", ".log"));
                TVFileModel tvFileModelNew = new TVFileModel()
                {
                    ClientFilePath = "",
                    FileDescription = "Mike Scenario Log file",
                    FileInfo = "Mike Scenario Log file",
                    FileCreatedDate_UTC = fiLog.CreationTimeUtc,
                    FileSize_kb = 1,
                    Language = LanguageEnum.enAndfr,
                    Year = DateTime.Now.Year,
                    FromWater = null,
                    FilePurpose = FilePurposeEnum.MikeResultDFSU,
                    FileType = FileTypeEnum.LOG,
                    ServerFileName = fiLog.Name,
                    ServerFilePath = fiLog.DirectoryName + @"\",
                    TVFileTVItemID = tvItemModelExist.TVItemID,
                };

                TVFileModel tvFileModelRet = tvFileService.PostAddTVFileDB(tvFileModelNew);
                if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_For_With_Equal_, TaskRunnerServiceRes.TVFile, "log", TaskRunnerServiceRes.TVText, fiLog.Name);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat4List("CouldNotCreate_For_With_Equal_", TaskRunnerServiceRes.TVFile, "log", TaskRunnerServiceRes.TVText, fiLog.Name);
                    return;
                }
            }
        }
        private bool SetParameterFileName(PFSFile pfsFile, string Path, string FileName, string Keyword)
        {
            string NotUsed = "";
            PFSSection pfsSection = pfsFile.GetSectionFromHandle(Path);

            if (pfsSection == null)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindPFSSectionWithPath_, Path);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindPFSSectionWithPath_", Path);
                return false;
            }

            PFSKeyword keyword = null;
            try
            {
                keyword = pfsSection.GetKeyword(Keyword);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                return false;
            }

            if (keyword != null)
            {
                PFSParameter param = null;
                try
                {
                    keyword.DeleteParameter(1);
                    param = keyword.InsertNewParameterFileName(FileName, 1);
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return false;
                }
            }

            return true;
        }
        private bool SetParameterInt(PFSFile pfsFile, string Path, int value, string Keyword)
        {
            string NotUsed = "";
            PFSSection pfsSection = pfsFile.GetSectionFromHandle(Path);

            if (pfsSection == null)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindPFSSectionWithPath_, Path);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindPFSSectionWithPath_", Path);
                return false;
            }

            PFSKeyword keyword = null;
            try
            {
                keyword = pfsSection.GetKeyword(Keyword);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                return false;
            }

            if (keyword != null)
            {
                PFSParameter param = null;
                try
                {
                    keyword.DeleteParameter(1);
                    param = keyword.InsertNewParameterInt(value, 1);
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return false;
                }
            }

            return true;
        }
        private bool SetParameterDouble(PFSFile pfsFile, string Path, double value, string Keyword)
        {
            string NotUsed = "";
            PFSSection pfsSection = pfsFile.GetSectionFromHandle(Path);

            if (pfsSection == null)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindPFSSectionWithPath_, Path);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindPFSSectionWithPath_", Path);
                return false;
            }

            PFSKeyword keyword = null;
            try
            {
                keyword = pfsSection.GetKeyword(Keyword);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                return false;
            }

            if (keyword != null)
            {
                PFSParameter param = null;
                try
                {
                    keyword.DeleteParameter(1);
                    param = keyword.InsertNewParameterDouble(value, 1);
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return false;
                }
            }

            return true;
        }
        private bool SetParameterCoordinates(PFSFile pfsFile, string Path, Coord coord, string Keyword)
        {
            string NotUsed = "";
            PFSSection pfsSection = pfsFile.GetSectionFromHandle(Path);

            if (pfsSection == null)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindPFSSectionWithPath_, Path);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindPFSSectionWithPath_", Path);
                return false;
            }

            PFSKeyword keyword = null;
            try
            {
                keyword = pfsSection.GetKeyword(Keyword);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                return false;
            }

            if (keyword != null)
            {
                PFSParameter param = null;
                try
                {
                    double temp = keyword.GetParameter(3).ToDouble();
                    keyword.DeleteParameter(1);
                    keyword.DeleteParameter(1);
                    keyword.DeleteParameter(1);
                    param = keyword.InsertNewParameterDouble(coord.Lng, 1);
                    param = keyword.InsertNewParameterDouble(coord.Lat, 2);
                    param = keyword.InsertNewParameterDouble(temp, 3);
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return false;
                }
            }


            return true;
        }
        private bool SetParameterString(PFSFile pfsFile, string Path, string value, string Keyword)
        {
            string NotUsed = "";
            PFSSection pfsSection = pfsFile.GetSectionFromHandle(Path);

            if (pfsSection == null)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindPFSSectionWithPath_, Path);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindPFSSectionWithPath_", Path);
                return false;
            }

            PFSKeyword keyword = null;
            try
            {
                keyword = pfsSection.GetKeyword(Keyword);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                return false;
            }

            if (keyword != null)
            {
                PFSParameter param = null;
                try
                {
                    keyword.DeleteParameter(1);
                    param = keyword.InsertNewParameterString(value, 1);
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return false;
                }
            }

            return true;
        }
        private bool SetParameterTimeStart(PFSFile pfsFile, string Path, DateTime NewDateTime, string Keyword)
        {
            string NotUsed = "";
            PFSSection pfsSection = pfsFile.GetSectionFromHandle(Path);

            if (pfsSection == null)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindPFSSectionWithPath_, Path);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindPFSSectionWithPath_", Path);
                return false;
            }

            PFSKeyword keyword = null;
            try
            {
                keyword = pfsSection.GetKeyword(Keyword);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                return false;
            }

            if (keyword != null)
            {
                PFSParameter param = null;
                try
                {
                    keyword.DeleteParameter(6);
                    keyword.DeleteParameter(5);
                    keyword.DeleteParameter(4);
                    keyword.DeleteParameter(3);
                    keyword.DeleteParameter(2);
                    keyword.DeleteParameter(1);
                    param = keyword.InsertNewParameterDouble(NewDateTime.Year, 1);
                    param = keyword.InsertNewParameterDouble(NewDateTime.Month, 2);
                    param = keyword.InsertNewParameterDouble(NewDateTime.Day, 3);
                    param = keyword.InsertNewParameterDouble(NewDateTime.Hour, 4);
                    param = keyword.InsertNewParameterDouble(NewDateTime.Minute, 5);
                    param = keyword.InsertNewParameterDouble(NewDateTime.Second, 6);
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return false;
                }
            }

            return true;
        }
        private bool TweakScenarioStartTimeToHave60PercHighTideAndFalling(MikeScenarioModel mikeScenarioModel)
        {
            string NotUsed = "";
            MikeBoundaryConditionService mikeBoundaryConditionService = new MikeBoundaryConditionService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            List<MikeBoundaryConditionModel> mbcModelList = mikeBoundaryConditionService.GetMikeBoundaryConditionModelListWithMikeScenarioTVItemIDAndTVTypeDB(mikeScenarioModel.MikeScenarioTVItemID, TVTypeEnum.MikeBoundaryConditionWebTide);
            if (mbcModelList.Count == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_StartingWith_, TaskRunnerServiceRes.MIKEBoundaryCondition, TaskRunnerServiceRes.TVItemID, mikeScenarioModel.MikeScenarioTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_StartingWith_", TaskRunnerServiceRes.MIKEBoundaryCondition, TaskRunnerServiceRes.TVItemID, mikeScenarioModel.MikeScenarioTVItemID.ToString());
                return false;
            }

            MikeBoundaryConditionModel mbcm = mbcModelList.Where(c => c.MikeBoundaryConditionLevelOrVelocity == MikeBoundaryConditionLevelOrVelocityEnum.Level && c.TVType == TVTypeEnum.MikeBoundaryConditionWebTide).FirstOrDefault();
            if (mbcm == null)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MIKEBoundaryCondition, TaskRunnerServiceRes.MikeBoundaryConditionLevelOrVelocity + "," + TaskRunnerServiceRes.TVType, MikeBoundaryConditionLevelOrVelocityEnum.Level.ToString() + "," + TVTypeEnum.MikeBoundaryConditionWebTide.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MIKEBoundaryCondition, TaskRunnerServiceRes.MikeBoundaryConditionLevelOrVelocity + "," + TaskRunnerServiceRes.TVType, MikeBoundaryConditionLevelOrVelocityEnum.Level.ToString() + "," + TVTypeEnum.MikeBoundaryConditionWebTide.ToString());
                return false;
            }

            MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            List<MapInfoPointModel> mapInfoPointList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(mbcm.MikeBoundaryConditionTVItemID, TVTypeEnum.MikeBoundaryConditionWebTide, MapInfoDrawTypeEnum.Polyline);
            if (mapInfoPointList.Count > 0)
            {
                List<WaterLevelResult> WLResults = null;

                GoogleTimeZoneJSON googleTimeZoneJSON = new GoogleTimeZoneJSON();
                using (WebClient webClient = new WebClient() { Encoding = Encoding.UTF8 })
                {
                    WebProxy webProxy = new WebProxy();
                    webClient.Proxy = webProxy;

                    webClient.UseDefaultCredentials = true;

                    var json_data = string.Empty;
                    // attempt to download JSON data as a string
                    try
                    {
                        DateTime epoch = new DateTime(1970, 1, 1, 0, 0, 0, 0).ToLocalTime();
                        TimeSpan span = (mikeScenarioModel.MikeScenarioStartDateTime_Local - epoch);
                        string timeStamp = ((long)Convert.ToDouble(span.TotalSeconds)).ToString();
                        string url = "https://maps.googleapis.com/maps/api/timezone/json?location=" +
                            mapInfoPointList[0].Lat.ToString("F5") + "," + mapInfoPointList[0].Lng.ToString("F5") + "&timestamp=" + timeStamp + "&key=AIzaSyAwPGpdSM6z0A7DFdWPbS3vIDTk2mxINaA";
                        byte[] responseBytes = webClient.DownloadData(url);
                        json_data = Encoding.UTF8.GetString(responseBytes);
                    }
                    catch (Exception)
                    {
                        json_data = "Error";
                    }
                    JsonSerializerSettings jsonSerializerSettings = new JsonSerializerSettings
                    {
                        DateTimeZoneHandling = DateTimeZoneHandling.Local
                    };
                    googleTimeZoneJSON = JsonConvert.DeserializeObject<GoogleTimeZoneJSON>(json_data, jsonSerializerSettings);
                }

                float offset = (float)((googleTimeZoneJSON.rawOffset + googleTimeZoneJSON.dstOffset) / 3600.0f);

                string offsetText = offset.ToString("F1");
                if (offsetText.EndsWith("0"))
                {
                    offsetText = offsetText.Substring(0, offsetText.Length - 2);
                }

                TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                TideModel tideModel = new TideModel(tvFileService.ChoseEDriveOrCDrive(tvFileService.BasePath), mbcm.WebTideDataSet)
                {
                    StartDate = mikeScenarioModel.MikeScenarioStartDateTime_Local.AddHours(-offset - 24),
                    EndDate = mikeScenarioModel.MikeScenarioEndDateTime_Local.AddHours(-offset),
                    Lng = (double)mapInfoPointList[0].Lng,
                    Lat = (double)mapInfoPointList[0].Lat,
                    Steps_min = 60,
                    DoWaterLevels = true,
                };
                WLResults = tidesAndCurrentsService.GetTides(tideModel);
                if (WLResults != null)
                {
                    foreach (WaterLevelResult wlr in WLResults)
                    {
                        wlr.Date = wlr.Date.AddHours(offset);
                    }
                }

                if (WLResults.Count > 0)
                {
                    DateTime StartTime = mikeScenarioModel.MikeScenarioStartDateTime_Local;
                    double Max = -1000;
                    double Min = 1000;
                    foreach (WaterLevelResult wl in WLResults)
                    {
                        if (Max < wl.WaterLevel)
                        {
                            Max = wl.WaterLevel;
                        }
                        if (Min > wl.WaterLevel)
                        {
                            Min = wl.WaterLevel;
                        }
                    }

                    double Perc60Value = (Max + Min) / 0.6D;

                    DateTime NewStartDate = mikeScenarioModel.MikeScenarioStartDateTime_Local;

                    for (int i = 1, count = WLResults.Count; i < count; i++)
                    {
                        if (WLResults[i].Date >= NewStartDate)
                        {
                            if (WLResults[i - 1].WaterLevel > WLResults[i].WaterLevel)
                            {
                                if (WLResults[i - 1].WaterLevel >= Perc60Value && WLResults[i].WaterLevel <= Perc60Value)
                                {
                                    mikeScenarioModel.MikeScenarioStartDateTime_Local = WLResults[i].Date;
                                    break;
                                }
                            }
                        }
                    }

                }


            }

            return true;
        }
        #endregion
    }

    public class GoogleTimeZoneJSON
    {
        public int dstOffset { get; set; }
        public int rawOffset { get; set; }
        public string status { get; set; }
        public string timeZoneId { get; set; }
        public string timeZoneName { get; set; }
    }
}
