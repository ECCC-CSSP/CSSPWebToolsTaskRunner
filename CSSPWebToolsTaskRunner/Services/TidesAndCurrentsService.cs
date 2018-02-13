using CSSPDHI;
using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;
using CSSPReportWriterHelperDLL.Services;
using CSSPWebToolsDBDLL.Services;
using CSSPWebToolsTaskRunner.Services.Resources;
using DHI.Generic.MikeZero.DFS.dfsu;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Transactions;
using System.Windows.Forms;

namespace CSSPWebToolsTaskRunner.Services
{
    public class TidesAndCurrentsService
    {
        #region Variables
        public bool TideCancel { get; set; }

        public List<InputAndResult> InputAndResultList { get; set; }
        public List<TideNode> TideNodeList { get; set; }
        public List<TideElement> TideElementList { get; set; }
        public List<Constituent> ConstituentNameList { get; set; }
        public List<MainConstituent> MainConstituentList { get; set; }
        public List<ShallowConstituent> ShallowConstituentList { get; set; }

        // Holder for Main Constituents
        public class MainConstituent
        {
            public MainConstituent()
            {
                dood = new List<int>();
                satellites = new List<Satellite>();
            }
            public string Name { get; set; }
            public List<int> dood { get; set; }
            public double phase { get; set; }
            public List<Satellite> satellites { get; set; }
            public double f { get; set; }
            public double freq { get; set; }
            public double vu { get; set; }

            // Holder for Satellites for each main constituent 
            public class Satellite
            {
                public Satellite()
                {
                    deldList = new List<int>();
                }
                public List<int> deldList { get; set; }
                public double phase { get; set; }
                public double ratio { get; set; }
                public int corr { get; set; }
            }
        }
        // Holder for Shallow water constituents 
        public class ShallowConstituent
        {
            public ShallowConstituent()
            {
                factors = new List<Factor>();
            }
            public string Name { get; set; }
            public List<Factor> factors { get; set; }
            public double f { get; set; }
            public double freq { get; set; }
            public double vu { get; set; }

            // Holder for Constituent factors for the shallow water constituents 
            public class Factor
            {
                public Factor()
                {

                }
                public string Name { get; set; }
                public double factor { get; set; }
            }

        }
        // Holder for TideNodes 
        public class TideNode
        {
            public TideNode()
            {

            }
            public int ID { get; set; }
            public double x { get; set; }
            public double y { get; set; }
            public double MinimumTide { get; set; }
            public double MinimumTide2 { get; set; }
        }
        // Holder for TideElements
        public class TideElement
        {
            public TideElement()
            {
                this.NodeList = new List<TideNode>();
            }
            public int ID { get; set; }
            public List<TideNode> NodeList { get; set; }
        }
        // Holder for Constituent names and amp and phase values
        public class Constituent
        {
            public Constituent()
            {
                amp = new List<double>();
                phase = new List<double>();
                amp2 = new List<double>();
                phase2 = new List<double>();
            }
            public string Name { get; set; }
            public List<double> amp { get; set; }
            public List<double> phase { get; set; }
            public List<double> amp2 { get; set; }
            public List<double> phase2 { get; set; }
        }
        // Holder for Config file information
        private class Config
        {
            public Config()
            {

            }
            public string NodeFileName { get; set; }
            public string ElementFileName { get; set; }
            public string IOS_FileName { get; set; }
        }
        // Holder for location and date (memory input)
        public class InputAndResult
        {
            public InputAndResult()
            {

            }
            public double Longitude { get; set; }
            public double Latitude { get; set; }
            public DateTime Date { get; set; }
            public double Reslt { get; set; }
            public double Reslt2 { get; set; }
        }
        // Holder for Data Files name and Path
        private class DataFilePath
        {
            public DataFilePath()
            {

            }
            public string Name { get; set; }
            public string Path { get; set; }
        }
        // Holder for Astro information and parameters
        private class Astro
        {
            public Astro()
            {

            }
            public double d1 { get; set; }
            public double h { get; set; }
            public double pp { get; set; }
            public double s { get; set; }
            public double p { get; set; }
            public double enp { get; set; }
            public double dh { get; set; }
            public double dpp { get; set; }
            public double ds { get; set; }
            public double dp { get; set; }
            public double dnp { get; set; }
        }

        List<double> basis;                   // Factors for interpolation 
        TideNode closestNode = new TideNode();        // Closest node number to an arbitrary point 
        Config config = new Config();
        List<string> ParsedValues = new List<string>();

        public enum Direction
        {
            Up,
            Down
        }
        public class Peaks
        {
            public DateTime Date { get; set; }
            public float Value { get; set; }
        }
        public enum TideType
        {
            Low,
            High
        }

        #endregion Variables

        #region Properties
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        #endregion Properties

        #region Constructors
        public TidesAndCurrentsService(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
        }
        #endregion Constructors

        #region Functions public
        public List<PeakDifference> CreateHighAndLowTide(TideModel tideModel, int Days)
        {
            CheckAllDataOK(tideModel);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return new List<PeakDifference>();

            LoadBase(tideModel);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return new List<PeakDifference>();

            CreateMemoryInput(tideModel);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return new List<PeakDifference>();

            CalculateResults(tideModel);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return new List<PeakDifference>();

            List<PeakDifference> HighPeakDifference = FindMonthlyHighAndLowTide(tideModel.StartDate.Year, InputAndResultList, TideType.High, Days);
            List<PeakDifference> LowPeakDifference = FindMonthlyHighAndLowTide(tideModel.StartDate.Year, InputAndResultList, TideType.Low, Days);

            return HighPeakDifference.Concat(LowPeakDifference).ToList(); // first 12 values are the High Peaks from 13 to 24 are the Low Peaks
        }
        public void GenerateWebTideNodes()
        {
            string NotUsed = "";

            MikeScenarioService mikeScenarioService = new MikeScenarioService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MikeBoundaryConditionService mikeBoundaryConditionService = new MikeBoundaryConditionService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TidesAndCurrentsService tidesAndCurrentsService = new TidesAndCurrentsService(_TaskRunnerBaseService);
            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            string Parameters = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters;
            string[] ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            int MikeScenarioTVItemID = 0;
            int BCMeshTVItemID = 0;
            int WebTideNodeNumb = 0;

            if (ParamValueList.Length != 3)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.ParameterCount_NotEqual_, ParamValueList.Length.ToString(), "5");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("ParameterCount_NotEqual_", ParamValueList.Length.ToString(), "5");
                return;
            }

            foreach (string s in ParamValueList)
            {
                string[] ParamValue = s.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                if (ParamValue.Length != 2)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.ProblemWhenParsing_, Parameters);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ProblemWhenParsing_", Parameters);
                    return;
                }

                if (ParamValue[0] == "MikeScenarioTVItemID")
                {
                    MikeScenarioTVItemID = int.Parse(ParamValue[1]);
                }
                else if (ParamValue[0] == "BCMeshTVItemID")
                {
                    BCMeshTVItemID = int.Parse(ParamValue[1]);
                }
                else if (ParamValue[0] == "WebTideNodeNumb")
                {
                    WebTideNodeNumb = int.Parse(ParamValue[1]);
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
                NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBeNullOrEmpty, TaskRunnerServiceRes.TVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_ShouldNotBeNullOrEmpty", TaskRunnerServiceRes.TVItemID);
                return;
            }


            if (WebTideNodeNumb < 1)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBe0, TaskRunnerServiceRes.WTNodeNumb);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_ShouldNotBe0", TaskRunnerServiceRes.WTNodeNumb);
                return;
            }

            if (BCMeshTVItemID < 1)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBeNullOrEmpty, TaskRunnerServiceRes.BCMeshTVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_ShouldNotBeNullOrEmpty", TaskRunnerServiceRes.BCMeshTVItemID);
                return;
            }

            MikeScenarioModel mikeScenarioModel = mikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(MikeScenarioTVItemID);
            if (!string.IsNullOrWhiteSpace(mikeScenarioModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                return;
            }

            MikeBoundaryConditionModel mikeBoundaryConditionModelMesh = mikeBoundaryConditionService.GetMikeBoundaryConditionModelWithMikeBoundaryConditionTVItemIDDB(BCMeshTVItemID);
            if (!string.IsNullOrWhiteSpace(mikeBoundaryConditionModelMesh.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MIKEBoundaryCondition, TaskRunnerServiceRes.BCMeshTVItemID, BCMeshTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MIKEBoundaryCondition, TaskRunnerServiceRes.BCMeshTVItemID, BCMeshTVItemID.ToString());
                return;
            }

            List<MikeBoundaryConditionModel> mikeBoundaryConditionModelWebTideList = mikeBoundaryConditionService.GetMikeBoundaryConditionModelListWithMikeScenarioTVItemIDAndTVTypeDB(mikeScenarioModel.MikeScenarioTVItemID, TVTypeEnum.MikeBoundaryConditionWebTide);
            MikeBoundaryConditionModel mikeBoundaryConditionModelWebTide = null;
            foreach (MikeBoundaryConditionModel mikeBoundaryConditionModel in mikeBoundaryConditionModelWebTideList)
            {
                if (mikeBoundaryConditionModelMesh.MikeBoundaryConditionCode == mikeBoundaryConditionModel.MikeBoundaryConditionCode)
                {
                    mikeBoundaryConditionModelWebTide = mikeBoundaryConditionModel;
                }
            }

            List<MapInfoPointModel> mapInfoPointModelList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(BCMeshTVItemID, TVTypeEnum.MikeBoundaryConditionMesh, MapInfoDrawTypeEnum.Polyline);
            if (mapInfoPointModelList.Count == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MapInfoPolyPoint, TaskRunnerServiceRes.TVItemID, BCMeshTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MapInfoPolyPoint, TaskRunnerServiceRes.TVItemID, BCMeshTVItemID.ToString());
                return;
            }

            float Distance = 0.0f;
            float TotalDistance = 0.0f;
            for (int i = 0; i < mapInfoPointModelList.Count(); i++)
            {
                if ((i + 1) != mapInfoPointModelList.Count())
                {
                    Distance = (float)mapInfoService.CalculateDistance((double)mapInfoPointModelList[i].Lat * mapInfoService.d2r, (double)mapInfoPointModelList[i].Lng * mapInfoService.d2r, (double)mapInfoPointModelList[i + 1].Lat * mapInfoService.d2r, (double)mapInfoPointModelList[i + 1].Lng * mapInfoService.d2r, mapInfoService.R);
                    TotalDistance += Distance;
                }
            }
            if (TotalDistance == 0.0f)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBe0, TaskRunnerServiceRes.TotalDistance);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_ShouldNotBe0", TaskRunnerServiceRes.TotalDistance);
                return;
            }

            List<Coord> CoordList = new List<Coord>();

            float WebTideDistance = TotalDistance / (WebTideNodeNumb * 1.0f);
            float CurrentDistance = WebTideDistance / 2.0f;

            Distance = 0.0f;
            TotalDistance = 0.0f;
            int countWTNode = 0;
            int BCCount = 0;
            for (int i = 0; i < mapInfoPointModelList.Count(); i++)
            {
                BCCount += 1;
                if ((i + 1) != mapInfoPointModelList.Count())
                {
                    Distance = (float)mapInfoService.CalculateDistance((double)mapInfoPointModelList[i].Lat * mapInfoService.d2r, (double)mapInfoPointModelList[i].Lng * mapInfoService.d2r, (double)mapInfoPointModelList[i + 1].Lat * mapInfoService.d2r, (double)mapInfoPointModelList[i + 1].Lng * mapInfoService.d2r, mapInfoService.R);
                    TotalDistance += Distance;
                }

                while (TotalDistance >= CurrentDistance)
                {
                    double LatToTry = 0.0f;
                    double LngToTry = 0.0f;
                    double Fraction = (CurrentDistance - (TotalDistance - Distance)) / Distance;

                    LatToTry = (double)mapInfoPointModelList[i].Lat + ((double)mapInfoPointModelList[i + 1].Lat - (double)mapInfoPointModelList[i].Lat) * Fraction;
                    LngToTry = (double)mapInfoPointModelList[i].Lng + ((double)mapInfoPointModelList[i + 1].Lng - (double)mapInfoPointModelList[i].Lng) * Fraction;

                    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)(CurrentDistance * 100 / TotalDistance));
                    TideModel tideModel = new TideModel(tvFileService.ChoseEDriveOrCDrive(tvFileService.BasePath), mikeBoundaryConditionModelMesh.WebTideDataSet)
                    {
                        StartDate = DateTime.Now,
                        EndDate = DateTime.Now.AddHours(1),
                        Lng = (double)LngToTry,
                        Lat = (double)LatToTry,
                        Steps_min = 60,
                        DoWaterLevels = false,
                    };
                    List<WaterLevelResult> WLResults = tidesAndCurrentsService.GetTides(tideModel);

                    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                        return;

                    if (WLResults != null)
                    {
                        Coord coordNew = new Coord();
                        coordNew.Lat = (float)LatToTry;
                        coordNew.Lng = (float)LngToTry;
                        coordNew.Ordinal = CoordList.Count();

                        CoordList.Add(coordNew);

                        countWTNode += 1;
                    }

                    CurrentDistance += WebTideDistance;
                }
            }

            using (TransactionScope ts = new TransactionScope())
            {
                if (mikeBoundaryConditionModelWebTide == null)
                {
                    TVItemModel tvItemModelMBC = tvItemService.PostAddChildTVItemDB(MikeScenarioTVItemID, mikeBoundaryConditionModelMesh.MikeBoundaryConditionName + " (Webtide)", TVTypeEnum.MikeBoundaryConditionWebTide);
                    if (!string.IsNullOrWhiteSpace(tvItemModelMBC.Error))
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_For_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.MIKEBoundaryCondition, TaskRunnerServiceRes.BoundaryConditionCode, mikeBoundaryConditionModelMesh.MikeBoundaryConditionCode);
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat4List("CouldNotCreate_For_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.MIKEBoundaryCondition, TaskRunnerServiceRes.BoundaryConditionCode, mikeBoundaryConditionModelMesh.MikeBoundaryConditionCode);
                        return;
                    }

                    mikeBoundaryConditionModelWebTide = new MikeBoundaryConditionModel();
                    mikeBoundaryConditionModelWebTide.MikeBoundaryConditionTVItemID = tvItemModelMBC.TVItemID;
                    mikeBoundaryConditionModelWebTide.MikeBoundaryConditionCode = mikeBoundaryConditionModelMesh.MikeBoundaryConditionCode;
                    mikeBoundaryConditionModelWebTide.MikeBoundaryConditionFormat = mikeBoundaryConditionModelMesh.MikeBoundaryConditionFormat;
                    mikeBoundaryConditionModelWebTide.MikeBoundaryConditionName = mikeBoundaryConditionModelMesh.MikeBoundaryConditionName;
                    mikeBoundaryConditionModelWebTide.MikeBoundaryConditionLevelOrVelocity = mikeBoundaryConditionModelMesh.MikeBoundaryConditionLevelOrVelocity;
                    mikeBoundaryConditionModelWebTide.NumberOfWebTideNodes = WebTideNodeNumb;
                    mikeBoundaryConditionModelWebTide.WebTideDataSet = mikeBoundaryConditionModelMesh.WebTideDataSet;
                    mikeBoundaryConditionModelWebTide.MikeBoundaryConditionTVText = mikeBoundaryConditionModelMesh.MikeBoundaryConditionTVText;
                    mikeBoundaryConditionModelWebTide.TVType = TVTypeEnum.MikeBoundaryConditionWebTide;
                    mikeBoundaryConditionModelWebTide.WebTideDataFromStartToEndDate = "";

                    TVItemModel tvItemModelMikeScenario = tvItemService.GetTVItemModelWithTVItemIDDB(MikeScenarioTVItemID);
                    if (!string.IsNullOrWhiteSpace(tvItemModelMikeScenario.Error))
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                        return;
                    }

                    MikeBoundaryConditionModel mikeBoundaryConditionModelRet = mikeBoundaryConditionService.PostAddMikeBoundaryConditionDB(mikeBoundaryConditionModelWebTide);
                    if (!string.IsNullOrWhiteSpace(mikeBoundaryConditionModelRet.Error))
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.MIKEBoundaryCondition, mikeBoundaryConditionModelRet.Error);
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.MIKEBoundaryCondition, mikeBoundaryConditionModelRet.Error);
                        return;
                    }
                }

                MapInfoModel mapInfoModel = mapInfoService.PostDeleteMapInfoWithTVItemIDDB(mikeBoundaryConditionModelWebTide.MikeBoundaryConditionTVItemID);
                if (!string.IsNullOrWhiteSpace(mapInfoModel.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDelete_With_Equal_Error_, TaskRunnerServiceRes.MapInfo, TaskRunnerServiceRes.TVItemID, mikeBoundaryConditionModelWebTide.MikeBoundaryConditionTVItemID, mapInfoModel.Error);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(string.Format(TaskRunnerServiceRes.CouldNotDelete_With_Equal_Error_, TaskRunnerServiceRes.MapInfo, TaskRunnerServiceRes.TVItemID, mikeBoundaryConditionModelWebTide.MikeBoundaryConditionTVItemID, mapInfoModel.Error));
                    return;
                }

                mapInfoModel = mapInfoService.CreateMapInfoObjectDB(CoordList, MapInfoDrawTypeEnum.Polyline, TVTypeEnum.MikeBoundaryConditionWebTide, mikeBoundaryConditionModelWebTide.MikeBoundaryConditionTVItemID);
                if (!string.IsNullOrWhiteSpace(mapInfoModel.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_Error_, TaskRunnerServiceRes.MapInfo, mapInfoModel.Error);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreate_Error_", TaskRunnerServiceRes.MapInfo, mapInfoModel.Error);
                    return;
                }

                ts.Complete();

            } // TransactionScope end

            TVFileModel tvFileModelM21_3 = tvFileService.GetTVFileModelWithTVItemIDAndTVFileTypeM21FMOrM3FMDB(MikeScenarioTVItemID);
            if (!string.IsNullOrWhiteSpace(tvFileModelM21_3.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_M21OrM3MDBWith_Equal_, TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(string.Format(TaskRunnerServiceRes.CouldNotFind_M21OrM3MDBWith_Equal_, TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString()));
                return;
            }

            FileInfo fiServer = new FileInfo(tvFileService.ChoseEDriveOrCDrive(tvFileModelM21_3.ServerFilePath) + tvFileModelM21_3.ServerFileName);

            if (!fiServer.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiServer.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiServer.FullName);
                return;
            }

            //TVFileModel tvFileModelDFS0 = new TVFileModel();
            List<TVFileModel> tvFileModelList = tvFileService.GetTVFileModelListWithParentTVItemIDDB(MikeScenarioTVItemID);

            for (int t = 2; t < 100; t++)
            {
                if (mikeBoundaryConditionModelMesh.MikeBoundaryConditionCode != "CODE_" + t.ToString())
                {
                    continue;
                }

                DateTime? dateTimeTemp = null;
                int? NumberOfTimeSteps = null;
                int? TimeStepInterval = null;
                using (PFS pfs = new PFS(fiServer))
                {
                    dateTimeTemp = pfs.GetVariableDateTime("FemEngineHD/TIME", "start_time");
                    if (dateTimeTemp == null)
                    {
                        return;
                    }

                    NumberOfTimeSteps = pfs.GetVariable<int>("FemEngineHD/TIME", "number_of_time_steps", 1);
                    if (NumberOfTimeSteps == null)
                    {
                        return;
                    }

                    TimeStepInterval = pfs.GetVariable<int>("FemEngineHD/TIME", "time_step_interval", 1);
                    if (TimeStepInterval == null)
                    {
                        return;
                    }
                }

                double WebTideStepsInMinutes = 0.0D;

                FileInfo fiBC = null;
                using (PFS pfs = new PFS(fiServer))
                {
                    fiBC = pfs.GetVariableFileInfo("FemEngineHD/HYDRODYNAMIC_MODULE/BOUNDARY_CONDITIONS/CODE_" + t.ToString(), "file_name", 1);
                    if (fiBC == null)
                    {
                        return;
                    }
                }

                string BCFileName = fiBC.Name;
                TVFileModel TVFileModelBC = (from c in tvFileModelList
                                             where c.ServerFileName == BCFileName
                                             select c).FirstOrDefault();
                if (TVFileModelBC == null)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_StartingWith_, TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.ServerFileName, BCFileName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_StartingWith_", TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.ServerFileName, BCFileName);
                    return;
                }

                FileInfo fi = new FileInfo(TVFileModelBC.ServerFilePath + TVFileModelBC.ServerFileName);
                if (!fi.Exists)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fi.FullName);
                    return;
                }

                using (DFS0 dfs0 = new DFS0(fi))
                {
                    WebTideStepsInMinutes = dfs0.GetWebTideStepsInMinutes();
                }

                DateTime StartDate = ((DateTime)dateTimeTemp).AddHours(-1);
                DateTime EndDate = ((DateTime)dateTimeTemp).AddSeconds((int)NumberOfTimeSteps * (int)TimeStepInterval).AddHours(1);

                List<List<WaterLevelResult>> AllWLResults = new List<List<WaterLevelResult>>();
                List<WaterLevelResult> WLResults = null;

                if (mikeBoundaryConditionModelWebTide.MikeBoundaryConditionLevelOrVelocity == MikeBoundaryConditionLevelOrVelocityEnum.Level)
                {
                    for (int i = 0; i < CoordList.Count; i++)
                    {
                        TideModel tideModel = new TideModel(tvFileService.ChoseEDriveOrCDrive(tvFileService.BasePath), mikeBoundaryConditionModelWebTide.WebTideDataSet)
                        {
                            StartDate = StartDate,
                            EndDate = EndDate,
                            Lng = (double)CoordList[i].Lng,
                            Lat = (double)CoordList[i].Lat,
                            Steps_min = WebTideStepsInMinutes,
                            DoWaterLevels = false,
                        };
                        WLResults = GetTides(tideModel);

                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                            return;

                        if (WLResults != null)
                        {
                            AllWLResults.Add(WLResults);
                        }
                    }
                }

                List<IEnumerable<CurrentResult>> AllCurrentResults = new List<IEnumerable<CurrentResult>>();
                List<CurrentResult> CurrentResults = null;

                if (mikeBoundaryConditionModelWebTide.MikeBoundaryConditionLevelOrVelocity == MikeBoundaryConditionLevelOrVelocityEnum.Velocity)
                {
                    for (int i = 0; i < CoordList.Count; i++)
                    {
                        TideModel tideModel = new TideModel(tvFileService.ChoseEDriveOrCDrive(tvFileService.BasePath), mikeBoundaryConditionModelWebTide.WebTideDataSet)
                        {
                            StartDate = StartDate,
                            EndDate = EndDate,
                            Lng = (double)CoordList[i].Lng,
                            Lat = (double)CoordList[i].Lat,
                            Steps_min = WebTideStepsInMinutes,
                            DoWaterLevels = false,
                        };
                        CurrentResults = GetCurrents(tideModel);

                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                            return;

                        if (CurrentResults != null)
                        {
                            AllCurrentResults.Add(CurrentResults);
                        }
                    }
                }

                FileInfo NewFileCreated = new FileInfo(TVFileModelBC.ServerFilePath + TVFileModelBC.ServerFileName);

                TVFileModelBC.ServerFileName = TVFileModelBC.ServerFileName.Replace(".dfs0", ".dfs1");
                TVFileModelBC.FileType = FileTypeEnum.DFS1;
                TVFileModelBC.FileSize_kb = (int)(NewFileCreated.Length / 1024);
                TVFileModelBC.FileCreatedDate_UTC = (NewFileCreated.CreationTime).ToUniversalTime();
                TVFileModelBC.ServerFilePath = TVFileModelBC.ServerFilePath;

                TVFileModel tvFileModelBCRet = tvFileService.PostUpdateTVFileDB(TVFileModelBC);
                if (!string.IsNullOrWhiteSpace(tvFileModelBCRet.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.TVFile, tvFileModelBCRet.Error);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.TVFile, tvFileModelBCRet.Error);
                    return;
                }

                using (PFS pfs = new PFS(fiServer))
                {
                    FileInfo fiBCdfs1 = new FileInfo(TVFileModelBC.ServerFilePath + TVFileModelBC.ServerFileName);
                    FileInfo fiRet = pfs.SetVariableFileInfo("FemEngineHD/HYDRODYNAMIC_MODULE/BOUNDARY_CONDITIONS/CODE_" + t.ToString(), "file_name", 1, fiBCdfs1);
                    if (fiRet == null)
                    {
                        return;
                    }
                    int? intRet = pfs.SetVariable<int>("FemEngineHD/HYDRODYNAMIC_MODULE/BOUNDARY_CONDITIONS/CODE_" + t.ToString(), "format", 1, 3 /* format 1 = VaryingInTimeConstantAlongBoundary, 3 = VaryingInTimeAndAlongBoundary */);
                    if (intRet == null)
                    {
                        return;
                    }
                }
            }
        }
        public List<CurrentResult> GetCurrents(TideModel tideModel)
        {
            CheckAllDataOK(tideModel);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return new List<CurrentResult>();

            LoadBase(tideModel);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return new List<CurrentResult>();

            CreateMemoryInput(tideModel);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return new List<CurrentResult>();

            CalculateResults(tideModel);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return new List<CurrentResult>();

            List<CurrentResult> CurRes = (from c in InputAndResultList
                                          select new CurrentResult
                                          {
                                              Date = c.Date,
                                              x_velocity = c.Reslt,
                                              y_velocity = c.Reslt2
                                          }).ToList<CurrentResult>();
            return CurRes;
        }
        public List<WaterLevelResult> GetTides(TideModel tideModel)
        {
            CheckAllDataOK(tideModel);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return new List<WaterLevelResult>();

            LoadBase(tideModel);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return new List<WaterLevelResult>();

            CreateMemoryInput(tideModel);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return new List<WaterLevelResult>();

            CalculateResults(tideModel);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return new List<WaterLevelResult>();

            List<WaterLevelResult> WLRes = (from c in InputAndResultList
                                            select new WaterLevelResult
                                            {
                                                Date = c.Date,
                                                WaterLevel = c.Reslt
                                            }).ToList<WaterLevelResult>();
            return WLRes;

        }
        public void SetupWebTide()
        {
            KmzServiceMikeScenario kmzServiceMikeScenario = new KmzServiceMikeScenario(_TaskRunnerBaseService);
            MikeScenarioService mikeScenarioService = new MikeScenarioService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MikeScenarioModel mikeScenarioModel = mikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MikeBoundaryConditionService mikeBoundaryConditionService = new MikeBoundaryConditionService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            string NotUsed = "";
            string Parameters = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters;
            string[] ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            int MikeScenarioTVItemID = 0;
            WebTideDataSetEnum WebTideDataSet = WebTideDataSetEnum.Error;
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
                else if (ParamValue[0] == "WebTideDataSet")
                {
                    WebTideDataSet = (WebTideDataSetEnum)(int.Parse(ParamValue[1]));
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
                NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBeNullOrEmpty, "MikeScenarioTVItemID");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_ShouldNotBeNullOrEmpty", TaskRunnerServiceRes.MikeScenarioTVItemID);
                return;
            }

            if (WebTideDataSet == WebTideDataSetEnum.Error)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBeNullOrEmpty, "WebTideDataSet");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_ShouldNotBeNullOrEmpty", TaskRunnerServiceRes.WebTideDataSet);
                return;
            }

            if (!string.IsNullOrWhiteSpace(mikeScenarioModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, "MikeScenario", "MikeScenarioTVItemID", MikeScenarioTVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.TVItemID, MikeScenarioTVItemID.ToString());
                return;
            }

            if (!ResetWebTideDB())
            {
                return;
            }

            TVFileModel tvFileModelm21fm = tvFileService.GetTVFileModelWithTVItemIDAndTVFileTypeM21FMOrM3FMDB(mikeScenarioModel.MikeScenarioTVItemID);
            if (!string.IsNullOrWhiteSpace(tvFileModelm21fm.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, ".m21fm | .m3fm");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", ".m21fm | .m3fm");
                return;
            }

            string ServerFilePath = tvFileService.GetServerFilePath(mikeScenarioModel.MikeScenarioTVItemID);

            FileInfo fiServer = new FileInfo(tvFileService.ChoseEDriveOrCDrive(ServerFilePath) + tvFileModelm21fm.ServerFileName);

            if (!fiServer.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, fiServer.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", fiServer.FullName);
                return;
            }

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 10);

            //PFSFile pfsFile = new PFSFile(fiServer.FullName);
            //if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
            //{
            //    return;
            //}

            FileInfo fiTemp = null;
            using (PFS pfs = new PFS(fiServer))
            {
                fiTemp = pfs.GetVariableFileInfo("FemEngineHD/TRANSPORT_MODULE/OUTPUTS/OUTPUT_1", "file_name", 1);
                if (fiTemp == null)
                {
                    return;
                }
            }

            FileInfo fiDfsu = new FileInfo(tvFileService.ChoseEDriveOrCDrive(fiTemp.FullName));

            if (!fiDfsu.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiDfsu.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiDfsu.FullName);
                return;
            }

            DfsuFile dfsuFile;

            try
            {
                dfsuFile = DfsuFile.Open(fiDfsu.FullName);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotOpenFile_Error_, fiServer.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotOpenFile_Error_", fiServer.FullName, ex.Message);
                return;
            }

            ReportBaseService reportBaseService = new ReportBaseService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, new TreeView(), _TaskRunnerBaseService._User);

            List<ElementLayer> elementLayerList = new List<ElementLayer>();
            List<Element> elementList = new List<Element>();
            List<Node> nodeList = new List<Node>();
            List<NodeLayer> topNodeLayerList = new List<NodeLayer>();
            List<NodeLayer> bottomNodeLayerList = new List<NodeLayer>();

            ReportTag reportTag = new ReportTag();

            ParametersService parametersService = new ParametersService(_TaskRunnerBaseService);

            //string retStr = parametersService.FillRequiredList(dfsuFile, elementLayerList, elementList, nodeList, topNodeLayerList, bottomNodeLayerList);

            if (!parametersService.FillRequiredList(dfsuFile, elementList, elementLayerList, nodeList, topNodeLayerList, bottomNodeLayerList))
            {
                return;
            }

            dfsuFile.Close();

            Dictionary<String, Vector> ForwardVector = new Dictionary<String, Vector>();
            Dictionary<String, Vector> BackwardVector = new Dictionary<String, Vector>();

            float MaxDepth = Math.Abs(nodeList.Min(n => n.Z));

            List<Node> AllNodeList = new List<Node>();
            List<ContourPolygon> ContourPolygonList = new List<ContourPolygon>();

            List<Node> AboveNodeList = (from n in nodeList
                                        where n.Code != 0
                                        select n).ToList<Node>();

            foreach (Node sn in AboveNodeList)
            {
                List<Node> EndNodeList = (from n in sn.ConnectNodeList
                                          where n.Code != 0
                                          select n).ToList<Node>();

                foreach (Node en in EndNodeList)
                {
                    AllNodeList.Add(en);
                }

                if (sn.Code != 0)
                {
                    AllNodeList.Add(sn);
                }

            }

            List<Element> UniqueElementList = new List<Element>();


            // filling UniqueElementList triangle
            UniqueElementList = (from el in elementLayerList.Where(c => c.Layer == 1)
                                 where (el.Element.Type == 21 || el.Element.Type == 32)
                                 && (el.Element.NodeList[0].Code != 0 && el.Element.NodeList[1].Code != 0)
                                 || (el.Element.NodeList[0].Code != 0 && el.Element.NodeList[2].Code != 0)
                                 || (el.Element.NodeList[1].Code != 0 && el.Element.NodeList[2].Code != 0)
                                 select el.Element).Distinct().ToList<Element>();

            UniqueElementList.Concat((from el in elementLayerList.Where(c => c.Layer == 1)
                                      where (el.Element.Type == 25 || el.Element.Type == 33)
                                      && (el.Element.NodeList[0].Code != 0 && el.Element.NodeList[1].Code != 0)
                                      || (el.Element.NodeList[0].Code != 0 && el.Element.NodeList[2].Code != 0)
                                      || (el.Element.NodeList[1].Code != 0 && el.Element.NodeList[2].Code != 0)
                                      select el.Element).Distinct().ToList<Element>());

            List<Node> UniqueNodeList = (from n in AllNodeList orderby n.ID select n).Distinct().ToList<Node>();


            ForwardVector = new Dictionary<String, Vector>();
            BackwardVector = new Dictionary<String, Vector>();

            //UpdateTask(AppTaskID, "30 %");

            foreach (Element el in UniqueElementList)
            {
                if (el.Type == 21 || el.Type == 32)
                {
                    List<Node> OrderedNodes = (from nv in el.NodeList
                                               select nv).ToList<Node>();
                    Node Node0 = OrderedNodes[0];
                    Node Node1 = OrderedNodes[1];
                    Node Node2 = OrderedNodes[2];

                    int ElemCount01 = (from el1 in UniqueElementList
                                       from el2 in Node0.ElementList
                                       from el3 in Node1.ElementList
                                       where el1 == el2 && el1 == el3
                                       select el1).Count();

                    int ElemCount02 = (from el1 in UniqueElementList
                                       from el2 in Node0.ElementList
                                       from el3 in Node2.ElementList
                                       where el1 == el2 && el1 == el3
                                       select el1).Count();

                    int ElemCount12 = (from el1 in UniqueElementList
                                       from el2 in Node1.ElementList
                                       from el3 in Node2.ElementList
                                       where el1 == el2 && el1 == el3
                                       select el1).Count();


                    if (Node0.Code != 0 && Node1.Code != 0)
                    {
                        if (ElemCount01 == 1)
                        {
                            ForwardVector.Add(Node0.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node1 });
                            BackwardVector.Add(Node1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node0 });
                        }
                    }
                    if (Node0.Code != 0 && Node2.Code != 0)
                    {
                        if (ElemCount02 == 1)
                        {
                            ForwardVector.Add(Node0.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node2 });
                            BackwardVector.Add(Node2.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node2, EndNode = Node0 });
                        }
                    }
                    if (Node1.Code != 0 && Node2.Code != 0)
                    {
                        if (ElemCount12 == 1)
                        {
                            ForwardVector.Add(Node1.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node2 });
                            BackwardVector.Add(Node2.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node2, EndNode = Node1 });
                        }
                    }
                }
                else if (el.Type == 25 || el.Type == 33)
                {
                    List<Node> OrderedNodes = (from nv in el.NodeList
                                               select nv).ToList<Node>();
                    Node Node0 = OrderedNodes[0];
                    Node Node1 = OrderedNodes[1];
                    Node Node2 = OrderedNodes[2];
                    Node Node3 = OrderedNodes[3];

                    int ElemCount01 = (from el1 in UniqueElementList
                                       from el2 in Node0.ElementList
                                       from el3 in Node1.ElementList
                                       where el1 == el2 && el1 == el3
                                       select el1).Count();

                    int ElemCount03 = (from el1 in UniqueElementList
                                       from el2 in Node0.ElementList
                                       from el3 in Node3.ElementList
                                       where el1 == el2 && el1 == el3
                                       select el1).Count();

                    int ElemCount12 = (from el1 in UniqueElementList
                                       from el2 in Node1.ElementList
                                       from el3 in Node2.ElementList
                                       where el1 == el2 && el1 == el3
                                       select el1).Count();

                    int ElemCount23 = (from el1 in UniqueElementList
                                       from el2 in Node2.ElementList
                                       from el3 in Node3.ElementList
                                       where el1 == el2 && el1 == el3
                                       select el1).Count();


                    if (Node0.Code != 0 && Node1.Code != 0)
                    {
                        if (ElemCount01 == 1)
                        {
                            ForwardVector.Add(Node0.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node1 });
                            BackwardVector.Add(Node1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node0 });
                        }
                    }
                    if (Node0.Code != 0 && Node3.Code != 0)
                    {
                        if (ElemCount03 == 1)
                        {
                            ForwardVector.Add(Node0.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node3 });
                            BackwardVector.Add(Node3.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node3, EndNode = Node0 });
                        }
                    }
                    if (Node1.Code != 0 && Node2.Code != 0)
                    {
                        if (ElemCount12 == 1)
                        {
                            ForwardVector.Add(Node1.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node2 });
                            BackwardVector.Add(Node2.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node2, EndNode = Node1 });
                        }
                    }
                    if (Node2.Code != 0 && Node3.Code != 0)
                    {
                        if (ElemCount23 == 1)
                        {
                            ForwardVector.Add(Node2.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = Node2, EndNode = Node3 });
                            BackwardVector.Add(Node3.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = Node3, EndNode = Node2 });
                        }
                    }
                }
            }

            bool MoreStudyAreaLine = true;
            while (MoreStudyAreaLine)
            {
                if (ForwardVector.Count > 0)
                {
                    List<Node> FinalContourNodeList = new List<Node>();
                    Vector LastVector = new Vector();
                    LastVector = ForwardVector.First().Value;
                    FinalContourNodeList.Add(LastVector.StartNode);
                    FinalContourNodeList.Add(LastVector.EndNode);
                    ForwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
                    BackwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
                    bool PolygonCompleted = false;
                    while (!PolygonCompleted)
                    {
                        List<string> KeyStrList = (from k in ForwardVector.Keys
                                                   where k.StartsWith(LastVector.EndNode.ID.ToString() + ",")
                                                   && !k.EndsWith("," + LastVector.StartNode.ID.ToString())
                                                   select k).ToList<string>();

                        if (KeyStrList.Count != 1)
                        {
                            KeyStrList = (from k in BackwardVector.Keys
                                          where k.StartsWith(LastVector.EndNode.ID.ToString() + ",")
                                          && !k.EndsWith("," + LastVector.StartNode.ID.ToString())
                                          select k).ToList<string>();

                            if (KeyStrList.Count != 1)
                            {
                                PolygonCompleted = true;
                                break;
                            }
                            else
                            {
                                LastVector = BackwardVector[KeyStrList[0]];
                                BackwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
                                ForwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
                            }
                        }
                        else
                        {
                            LastVector = ForwardVector[KeyStrList[0]];
                            ForwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
                            BackwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
                        }
                        FinalContourNodeList.Add(LastVector.EndNode);
                        if (FinalContourNodeList[FinalContourNodeList.Count - 1] == FinalContourNodeList[0])
                        {
                            PolygonCompleted = true;
                        }
                    }

                    double Area = mapInfoService.CalculateAreaOfPolygon(FinalContourNodeList);
                    if (Area < 0)
                    {
                        FinalContourNodeList.Reverse();
                        Area = mapInfoService.CalculateAreaOfPolygon(FinalContourNodeList);
                    }

                    FinalContourNodeList.Add(FinalContourNodeList[0]);

                    ContourPolygon contourLine = new ContourPolygon() { };
                    contourLine.ContourNodeList = FinalContourNodeList;
                    contourLine.ContourValue = 0;
                    ContourPolygonList.Add(contourLine);

                }

                MoreStudyAreaLine = false;
            }

            if (ContourPolygonList.Count != 1)
            {
                NotUsed = TaskRunnerServiceRes.ContourPolygonListCountShouldBeOne;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("ContourPolygonListCountShouldBe1");
                return;
            }

            bool Found = true;
            while (Found)
            {
                if (ContourPolygonList[0].ContourNodeList[0].Code != 1)
                {
                    Node n = ContourPolygonList[0].ContourNodeList[0];
                    ContourPolygonList[0].ContourNodeList.RemoveAt(0);
                    ContourPolygonList[0].ContourNodeList.Add(ContourPolygonList[0].ContourNodeList[0]);
                }
                else
                {
                    Found = false;
                }
            }

            List<Node> FinalNodes = new List<Node>();
            FinalNodes = (from c in ContourPolygonList[0].ContourNodeList
                          select c).ToList<Node>();
            int CountCode = (from c in ContourPolygonList[0].ContourNodeList
                             where c.Code > 1
                             select c.Code).Distinct<int>().Count();


            Dictionary<string, List<Node>> BCListNodes = new Dictionary<string, List<Node>>();

            for (int i = 0; i < CountCode; i++)
            {
                BCListNodes.Add("CODE_" + (i + 2).ToString(), new List<Node>());
            }

            for (int i = 0; i < FinalNodes.Count - 1; i++)
            {
                if (FinalNodes[i].Code > 1)
                {
                    try
                    {
                        if ((FinalNodes[i - 1].Code != FinalNodes[i].Code))
                        {
                            if (!BCListNodes["CODE_" + FinalNodes[i].Code].Contains(FinalNodes[i - 1]))
                            {
                                BCListNodes["CODE_" + FinalNodes[i].Code].Add(FinalNodes[i - 1]);
                            }
                        }
                    }
                    catch (Exception)
                    {
                    }
                    if (!BCListNodes["CODE_" + FinalNodes[i].Code].Contains(FinalNodes[i]))
                    {
                        BCListNodes["CODE_" + FinalNodes[i].Code].Add(FinalNodes[i]);
                    }
                    try
                    {
                        if ((FinalNodes[i + 1].Code != FinalNodes[i].Code))
                        {
                            if (!BCListNodes["CODE_" + FinalNodes[i].Code].Contains(FinalNodes[i + 1]))
                            {
                                BCListNodes["CODE_" + FinalNodes[i].Code].Add(FinalNodes[i + 1]);
                            }
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
            }

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 50);

            int TotalNodes = 0;
            int AtNodeCount = 0;
            float CurrentStatus = 50f;
            foreach (string key in BCListNodes.Keys)
            {
                TotalNodes += BCListNodes[key].Count;
            }

            //mikeScenarioModel.UseWebTide = true;

            // calculating the length of the boundary conditions
            //Dictionary<string, double> DistList = new Dictionary<string, double>();
            int BCCount = 0;
            foreach (string key in BCListNodes.Keys)
            {
                BCCount += 1;
                float Distance = 0;
                float TotalDistance = 0;
                List<Node> bcNodes = BCListNodes[key];

                // Doing Mike Boundary Condition
                //TVTypeEnum TVType = TVTypeEnum.MikeBoundaryCondition;

                string strBC = "";
                using (PFS pfs = new PFS(fiServer))
                {
                    strBC = pfs.GetVariableString("FemEngineHD/DOMAIN/BOUNDARY_NAMES/" + key, "name", 1);
                    if (string.IsNullOrWhiteSpace(strBC))
                    {
                        return;
                    }
                }
                string TVText = _TaskRunnerBaseService.CleanText(strBC.Replace("'", ""));

                TVItemModel tvItemModelBC = tvItemService.GetChildTVItemModelWithTVItemIDAndTVTextStartWithAndTVTypeDB(MikeScenarioTVItemID, TVText, TVTypeEnum.MikeBoundaryConditionMesh);
                if (!string.IsNullOrWhiteSpace(tvItemModelBC.Error))
                {
                    tvItemModelBC = tvItemService.PostAddChildTVItemDB(MikeScenarioTVItemID, TVText, TVTypeEnum.MikeBoundaryConditionMesh);
                    if (!string.IsNullOrWhiteSpace(tvItemModelBC.Error)) // would indicate there was an error
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_, TaskRunnerServiceRes.MIKEBoundaryCondition);
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotCreate_", TaskRunnerServiceRes.MIKEBoundaryCondition);
                        return;
                    }
                }

                MikeBoundaryConditionModel mbcmExist = mikeBoundaryConditionService.GetMikeBoundaryConditionModelWithMikeBoundaryConditionTVItemIDDB(tvItemModelBC.TVItemID);
                if (string.IsNullOrWhiteSpace(mbcmExist.Error))
                {
                    MikeBoundaryConditionModel mbcmRet = mikeBoundaryConditionService.PostDeleteMikeBoundaryConditionDB(tvItemModelBC.TVItemID);
                    if (!string.IsNullOrWhiteSpace(mbcmRet.Error))
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDelete_With_Equal_Error_, TaskRunnerServiceRes.MIKEBoundaryCondition, TaskRunnerServiceRes.MikeBoundaryConditionTVItemID, tvItemModelBC.TVItemID, mbcmRet.Error);
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat4List("CouldNotDelete_With_Equal_Error_", TaskRunnerServiceRes.MIKEBoundaryCondition, TaskRunnerServiceRes.MikeBoundaryConditionTVItemID, tvItemModelBC.TVItemID.ToString(), mbcmRet.Error);
                        return;
                    }
                }

                MikeBoundaryConditionModel mbcm = new MikeBoundaryConditionModel();
                mbcm.MikeBoundaryConditionTVItemID = tvItemModelBC.TVItemID;
                mbcm.MikeBoundaryConditionCode = key;
                mbcm.MikeBoundaryConditionName = TVText;
                mbcm.MikeBoundaryConditionLength_m = TotalDistance;

                int? formatBC = null;
                using (PFS pfs = new PFS(fiServer))
                {
                    formatBC = pfs.GetVariable<int>("FemEngineHD/HYDRODYNAMIC_MODULE/BOUNDARY_CONDITIONS/" + key, "format", 1);
                    if (formatBC == null)
                    {
                        return;
                    }
                }
                if (formatBC == 0) // Constant
                {
                    mbcm.MikeBoundaryConditionFormat = "Contant";
                }
                else if (formatBC == 1) // VaryingInTimeConstantAlongBoundary
                {
                    mbcm.MikeBoundaryConditionFormat = "VaryingInTimeConstantAlongBoundary";
                }
                else if (formatBC == 3) // VaryingInTimeAndAlongBoundary
                {
                    mbcm.MikeBoundaryConditionFormat = "VaryingInTimeAndAlongBoundary";
                }
                else
                {
                    NotUsed = TaskRunnerServiceRes.MikeBoundaryConditionFormatShouldBeVaryingInTimeConstantAlongBoundaryToStart;
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("MikeBoundaryConditionFormatShouldBeVaryingInTimeConstantAlongBoundaryToStart");
                    return;
                }

                //if (formatBC != 1)
                //{
                //    NotUsed = TaskRunnerServiceRes.MikeBoundaryConditionFormatShouldBeVaryingInTimeConstantAlongBoundaryToStart;
                //    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("MikeBoundaryConditionFormatShouldBeVaryingInTimeConstantAlongBoundaryToStart");
                //    return;
                //}

                mbcm.MikeBoundaryConditionTVText = TVText;

                int? typeBC = null;
                using (PFS pfs = new PFS(fiServer))
                {
                    typeBC = pfs.GetVariable<int>("FemEngineHD/HYDRODYNAMIC_MODULE/BOUNDARY_CONDITIONS/" + key, "type", 1);
                    if (typeBC == null)
                    {
                        return;
                    }
                }

                if (typeBC == 4) // 4 == Velocity
                {
                    mbcm.MikeBoundaryConditionLevelOrVelocity = MikeBoundaryConditionLevelOrVelocityEnum.Velocity;
                }
                else if (typeBC == 6) // 6 == Level
                {
                    mbcm.MikeBoundaryConditionLevelOrVelocity = MikeBoundaryConditionLevelOrVelocityEnum.Level;
                }
                else
                {
                    NotUsed = TaskRunnerServiceRes.MikeBoundaryConditionTypeShouldBeVelocityOrLevel;
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("MikeBoundaryConditionTypeShouldBeVelocityOrLevel");
                    return;
                }

                mbcm.WebTideDataSet = WebTideDataSet;
                mbcm.NumberOfWebTideNodes = 5;
                mbcm.TVType = TVTypeEnum.MikeBoundaryConditionMesh;
                mbcm.WebTideDataFromStartToEndDate = "";

                int BCNodeCount = 0;
                List<Coord> coordList = new List<Coord>();
                for (int i = 0; i < bcNodes.Count(); i++)
                {
                    AtNodeCount += 1;
                    float Perc = (float)(50.0f + (AtNodeCount * 50f / TotalNodes));
                    if (Perc > CurrentStatus)
                    {
                        _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)Perc);
                        CurrentStatus += 5;
                    }

                    BCNodeCount += 1;
                    if ((i + 1) != bcNodes.Count())
                    {
                        Distance = (float)mapInfoService.CalculateDistance((double)bcNodes[i].Y * mapInfoService.d2r, (double)bcNodes[i].X * mapInfoService.d2r, (double)bcNodes[i + 1].Y * mapInfoService.d2r, (double)bcNodes[i + 1].X * mapInfoService.d2r, mapInfoService.R);
                        TotalDistance += Distance;
                    }

                    coordList.Add(new Coord() { Lat = bcNodes[i].Y, Lng = bcNodes[i].X, Ordinal = i });
                }

                MapInfoModel mapInfoModel = mapInfoService.PostDeleteMapInfoWithTVItemIDDB(tvItemModelBC.TVItemID);
                if (!string.IsNullOrWhiteSpace(mapInfoModel.Error))
                {
                    // do something
                }

                MapInfoModel mapInfoModel2 = mapInfoService.CreateMapInfoObjectDB(coordList, MapInfoDrawTypeEnum.Polyline, TVTypeEnum.MikeBoundaryConditionMesh, tvItemModelBC.TVItemID);
                if (!string.IsNullOrWhiteSpace(mapInfoModel2.Error))
                {
                    // do something
                }
                mbcm.MikeBoundaryConditionLength_m = TotalDistance;

                TVItemModel tvItemModelMikeScenario = tvItemService.GetTVItemModelWithTVItemIDDB(MikeScenarioTVItemID);
                if (!string.IsNullOrWhiteSpace(tvItemModelMikeScenario.Error))
                {
                    // do something
                }

                MikeBoundaryConditionModel mikeBoundaryConditionModelRet = mikeBoundaryConditionService.PostAddMikeBoundaryConditionDB(mbcm);
                if (!string.IsNullOrWhiteSpace(mikeBoundaryConditionModelRet.Error))
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_Error_, TaskRunnerServiceRes.MikeBoundaryConditionName, mikeBoundaryConditionModelRet.Error);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreate_Error_", TaskRunnerServiceRes.MikeBoundaryConditionName, mikeBoundaryConditionModelRet.Error);
                    return;
                }
            }
        }
        #endregion Functions public

        #region Functions private
        private bool ResetWebTideDB()
        {
            string NotUsed = "";

            MikeBoundaryConditionService mikeBoundaryConditionService = new MikeBoundaryConditionService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            int TVItemID = _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID;


            using (TransactionScope ts = new TransactionScope())
            {
                List<int> TVItemToDeleteList = new List<int>();

                // doing MikeBoundaryCondition Mesh
                List<MikeBoundaryConditionModel> mikeBoundaryConditionModelToDeleteList = mikeBoundaryConditionService.GetMikeBoundaryConditionModelListWithMikeScenarioTVItemIDAndTVTypeDB(TVItemID, TVTypeEnum.MikeBoundaryConditionMesh);
                foreach (MikeBoundaryConditionModel mbcm in mikeBoundaryConditionModelToDeleteList)
                {
                    List<MapInfoModel> mapInfoModelListToDelete = mapInfoService.GetMapInfoModelListWithTVItemIDDB(mbcm.MikeBoundaryConditionTVItemID);
                    foreach (MapInfoModel mapInfoModel in mapInfoModelListToDelete)
                    {
                        MapInfoModel mapInfoModel2 = mapInfoService.PostDeleteMapInfoDB(mapInfoModel.MapInfoID);
                        if (!string.IsNullOrWhiteSpace(mapInfoModel2.Error))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDelete_With_Equal_Error_, TaskRunnerServiceRes.MapInfo, TaskRunnerServiceRes.MapInfoID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString(), mapInfoModel2.Error);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat4List("CouldNotDelete_With_Equal_Error_", TaskRunnerServiceRes.MapInfo, TaskRunnerServiceRes.MapInfoID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString(), mapInfoModel2.Error);
                            return false;
                        }
                    }
                }

                foreach (MikeBoundaryConditionModel mbcm in mikeBoundaryConditionModelToDeleteList)
                {
                    TVItemToDeleteList.Add(mbcm.MikeBoundaryConditionTVItemID);
                }

                foreach (int tvitd in TVItemToDeleteList)
                {
                    TVItemModel tvItemModel = tvItemService.PostDeleteTVItemWithTVItemIDDB(tvitd);
                    if (!string.IsNullOrWhiteSpace(tvItemModel.Error))
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDelete_With_Equal_Error_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString(), tvItemModel.Error);
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat4List("CouldNotDelete_With_Equal_Error_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString(), tvItemModel.Error);
                        return false;
                    }
                }

                // doing MikeBoundaryCondition WebTide
                mikeBoundaryConditionModelToDeleteList = mikeBoundaryConditionService.GetMikeBoundaryConditionModelListWithMikeScenarioTVItemIDAndTVTypeDB(TVItemID, TVTypeEnum.MikeBoundaryConditionWebTide);
                foreach (MikeBoundaryConditionModel mbcm in mikeBoundaryConditionModelToDeleteList)
                {
                    List<MapInfoModel> mapInfoModelListToDelete = mapInfoService.GetMapInfoModelListWithTVItemIDDB(mbcm.MikeBoundaryConditionTVItemID);
                    foreach (MapInfoModel mapInfoModel in mapInfoModelListToDelete)
                    {
                        MapInfoModel mapInfoModel2 = mapInfoService.PostDeleteMapInfoDB(mapInfoModel.MapInfoID);
                        if (!string.IsNullOrWhiteSpace(mapInfoModel2.Error))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDelete_With_Equal_Error_, TaskRunnerServiceRes.MapInfo, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString(), mapInfoModel2.Error);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat4List("CouldNotDelete_With_Equal_Error_", TaskRunnerServiceRes.MapInfo, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString(), mapInfoModel2.Error);
                            return false;
                        }
                    }
                }

                foreach (MikeBoundaryConditionModel mbcm in mikeBoundaryConditionModelToDeleteList)
                {
                    TVItemToDeleteList.Add(mbcm.MikeBoundaryConditionTVItemID);
                }

                foreach (int tvitd in TVItemToDeleteList)
                {
                    TVItemModel tvItemModel = tvItemService.PostDeleteTVItemWithTVItemIDDB(tvitd);
                    if (!string.IsNullOrWhiteSpace(tvItemModel.Error))
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDelete_With_Equal_Error_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString(), tvItemModel.Error);
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat4List("CouldNotDelete_With_Equal_Error_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString(), tvItemModel.Error);
                        return false;
                    }
                }
                ts.Complete();
            }
            return true;
        }
        private List<PeakDifference> FindMonthlyHighAndLowTide(int Year, List<InputAndResult> InputAndResultList, TideType tideType, int Days)
        {

            DateTime StartUpDate = new DateTime(Year, 1, 1);

            List<Peaks> PeakList = new List<Peaks>();

            Direction direction = new Direction();

            if (InputAndResultList[0].Reslt > InputAndResultList[1].Reslt)
            {
                direction = Direction.Down;
            }
            else
            {
                direction = Direction.Up;
            }

            for (int i = 1; i < InputAndResultList.Count - 2; i++)
            {
                if (InputAndResultList[i].Reslt > InputAndResultList[i + 1].Reslt)
                {
                    if (direction == Direction.Up)
                    {
                        PeakList.Add(new Peaks() { Date = InputAndResultList[i].Date, Value = (float)InputAndResultList[i].Reslt });
                        direction = Direction.Down;
                    }
                }
                else
                {
                    if (direction == Direction.Down)
                    {
                        PeakList.Add(new Peaks() { Date = InputAndResultList[i].Date, Value = (float)InputAndResultList[i].Reslt });
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

        private MainConstituent.Satellite AddSatellite(string satin)
        {
            /* Gets satellite info from input line, puts the info into a new
                satellite structure and returns the structure. */
            MainConstituent.Satellite newsat = new MainConstituent.Satellite();

            ParsedValues = satin.Split(default(string[]), StringSplitOptions.RemoveEmptyEntries).ToList<string>();

            if (Thread.CurrentThread.CurrentCulture.Name == "fr-CA")
            {
                newsat.deldList.Add(int.Parse(ParsedValues[0].Replace(".", ",")));
                newsat.deldList.Add(int.Parse(ParsedValues[1].Replace(".", ",")));
                newsat.deldList.Add(int.Parse(ParsedValues[2].Replace(".", ",")));
                newsat.phase = double.Parse(ParsedValues[3].Replace(".", ","));
                if (ParsedValues[4].Contains("R"))
                {
                    newsat.ratio = double.Parse(ParsedValues[4].Substring(0, ParsedValues[4].IndexOf("R")).Replace(".", ","));
                }
                else
                {
                    newsat.ratio = double.Parse(ParsedValues[4].Replace(".", ","));
                }
            }
            else
            {
                newsat.deldList.Add(int.Parse(ParsedValues[0]));
                newsat.deldList.Add(int.Parse(ParsedValues[1]));
                newsat.deldList.Add(int.Parse(ParsedValues[2]));
                newsat.phase = double.Parse(ParsedValues[3]);
                if (ParsedValues[4].Contains("R"))
                {
                    newsat.ratio = double.Parse(ParsedValues[4].Substring(0, ParsedValues[4].IndexOf("R")));
                }
                else
                {
                    newsat.ratio = double.Parse(ParsedValues[4]);
                }
            }

            if (!satin.Contains("R"))
            {
                newsat.corr = 0;
            }
            else
            {
                if (satin.Contains("R1"))
                {
                    newsat.corr = 1;
                }
                else if (satin.Contains("R2"))
                {
                    newsat.corr = 2;
                }
                else
                {
                    newsat.corr = 0;
                }
            }
            return newsat;
        }
        private void AstroAngles(ref Astro astro)
        {
            /* Calculates the following ephermides of the sun and moon:
                astro.h  = mean longitude of the sun;
                astro.pp = mean longitude of the solar perigee;
                astro.s  = mean longitude of the moon;
                astro.p  = mean longitude of the lunar perigee; and
                astro.enp = negative of the longitude of the mean ascending node.
                Also calculates their rate of change.
                Units are cycles ( cycles / 365 days for rates of change ).
                These formulae were taken from pp 98 and 107 of the Explanatory
                Supplement to the Astronomical Ephermeris and the American
                Ephermis and Nautical Almanac (1961) */
            double d12, d2, d22, d23, f, f2;

            d12 = astro.d1 * astro.d1;
            d2 = astro.d1 * 1.0E-04;
            d22 = d2 * d2;
            d23 = Math.Pow(d2, 3.0);
            f = 360.0;
            f2 = f / 365.0;

            astro.h = (2.79696678E+02 + astro.d1 * 9.856473354E-01 + d22 * 2.267E-05) / f;
            astro.h -= (int)astro.h;

            astro.pp = (2.81220833E+02 + astro.d1 * 4.70684E-05 + d22 * 3.39E-05 + d23 * 7.0E-08) / f;
            astro.pp -= (int)astro.pp;

            astro.s = (2.70434164E+02 + astro.d1 * 1.31763965268E+01 - d22 * 8.5E-05 + d23 * 3.9E-08) / f;
            astro.s -= (int)astro.s;

            astro.p = (3.34329556E+02 + astro.d1 * 1.114040803E-01 - d22 * 7.739E-04 - d23 * 2.6E-07) / f;
            astro.p -= (int)astro.p;

            astro.enp = (-2.59183275E+02 + astro.d1 * 5.29539222E-02 - d22 * 1.557E-04 - d23 * 5.0E-08) / f;
            astro.enp -= (int)astro.enp;

            astro.dh = (9.856473354E-01 + astro.d1 * 2.267E-05 * 2.0E-08) / f2;

            astro.dpp = (4.70684E-05 + astro.d1 * 3.39E-05 * 2.0E-08 + d12 * 7.0E-08 * 3.0E-12) / f2;

            astro.ds = (1.31763965268E+01 - astro.d1 * 8.5E-05 * 2.0E-08 + d12 * 3.9E-08 * 3.0E-12) / f2;

            astro.dp = (1.114040803E-01 - astro.d1 * 7.739E-04 * 2.0E-08 - d12 * 2.6E-07 * 3.0E-12) / f2;

            astro.dnp = (5.29539222E-02 - astro.d1 * 1.557E-04 * 2.0E-08 - d12 * 5.0E-08 * 3.0E-12) / f2;
        }
        private TideElement Basis2d(InputAndResult mi)
        {
            basis = new List<double>();

            // Finds the closest node to a point (ptx, pty) and the element containing
            // that point, if one exists.
            // Also gets the basis functions for interpolations to that point.
            // Returns the element number contaning the point (ptx, pty).
            // Returns -999 if no element found containing the point (ptx, pty ). 

            int flag;
            List<double> xlocal = new List<double>();
            List<double> ylocal = new List<double>();

            for (int i = 0; i < 3; i++)
            {
                xlocal.Add(0.0);
                ylocal.Add(0.0);
            }

            closestNode = FindClosestNode(mi);

            //  Try to find an element that contains the point and the closest node.
            TideElement ele = new TideElement();
            ele = null;

            foreach (TideElement element in TideElementList.Where(el => el.NodeList[0] == closestNode
                                            || el.NodeList[1] == closestNode
                                            || el.NodeList[2] == closestNode).Distinct<TideElement>().ToList<TideElement>())
            {
                int i = 0;
                foreach (TideNode n in element.NodeList)
                {
                    xlocal[i] = n.x;
                    ylocal[i] = n.y;
                    i += 1;
                }

                // See if the point is within this element 
                flag = RayBound(xlocal, ylocal, mi);
                if (flag == 1)
                {
                    // The point is within the element 
                    ele = element;
                    basis = Phi2d(xlocal, ylocal, mi);
                    break;
                }
            }

            /* if the closest node's elements don't work, search through all elements */
            if (ele == null)
            {
                foreach (TideElement element in TideElementList)
                {
                    int i = 0;
                    foreach (TideNode n in element.NodeList)
                    {
                        xlocal[i] = n.x;
                        ylocal[i] = n.y;
                        i += 1;
                    }

                    /* See if the point is within this element */
                    flag = RayBound(xlocal, ylocal, mi);
                    if (flag == 1)
                    {
                        ele = element;
                        basis = Phi2d(xlocal, ylocal, mi);
                        return (element);
                    }
                }
            }
            return (ele);
        }
        private void CalculateResults(TideModel tideModel)
        {
            string NotUsed = "";

            DateTime CurrentDate = DateTime.Now;
            TideElement ele = new TideElement();
            ele = null;
            List<double> elem_res = new List<double>();
            List<double> elem_res2 = new List<double>();
            List<double> elem_min = new List<double>();
            List<double> elem_min2 = new List<double>();

            // Loop through the input file. For each line calculate the tidal correction 

            int CountRefresh = 0;
            int CountAt = 0;
            int UpdateAfter = (int)(InputAndResultList.Count() / 100);
            foreach (InputAndResult mi in InputAndResultList)
            {
                if (TideCancel)
                {
                    TideCancel = false;
                    break;
                }

                if (CurrentDate.Day != mi.Date.Day)
                {
                    CountRefresh += 1;
                    CountAt += 1;
                    if (CountRefresh > UpdateAfter)
                    {
                        //SendPercentToDB(bwObj, (int)((CountAt * 100) / InputAndResultList.Count()));
                        CountRefresh = 0;
                    }
                }

                ele = Basis2d(mi);

                if (ele == null)
                {
                    // No element was found that contained the new position. 
                    // Calculate the tidal correction for the closest node. 

                    // removed by Charles LeBlanc, results should be -999 if no elements were found
                    mi.Reslt = -999;
                    mi.Reslt2 = -999;

                    NotUsed = string.Format(TaskRunnerServiceRes.ElementNotFoundWithElem_, ele.ID);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ElementNotFoundWithElem_", ele.ID.ToString());
                    return;

                    //__________________________________________________
                    //reslt = TideP(mi, closestNode, true);
                    //reslt = reslt + closestNode.MinimumTide;
                    //if (dataType == DataType.Current)
                    //{
                    //    reslt2 = TideP(mi, closestNode, false);
                    //    reslt2 = reslt2 + closestNode.MinimumTide2;
                    //}
                    //__________________________________________________

                }
                else
                {
                    /* If the point is inside and element, calculate the tidal correction for
                        each node of the element and interpolate to get the tidal correction
                        at the new position. */
                    elem_res.Clear();
                    elem_res2.Clear();
                    elem_min.Clear();
                    elem_min2.Clear();

                    foreach (TideNode n in ele.NodeList)
                    {
                        double tideP = TideP(mi, n, true);

                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                            return;

                        elem_res.Add(tideP);
                        elem_min.Add(n.MinimumTide);
                        if (!tideModel.DoWaterLevels)
                        {
                            double tideP2 = TideP(mi, n, false);

                            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                                return;

                            elem_res2.Add(tideP2);
                            elem_min2.Add(n.MinimumTide2);
                        }
                    }

                    mi.Reslt = elem_res[0] * basis[0]
                            + elem_res[1] * basis[1]
                        + elem_res[2] * basis[2]
                        + elem_min[0] * basis[0]
                            + elem_min[1] * basis[1]
                        + elem_min[2] * basis[2];
                    if (!tideModel.DoWaterLevels)
                    {
                        mi.Reslt2 = elem_res2[0] * basis[0]
                                + elem_res2[1] * basis[1]
                            + elem_res2[2] * basis[2]
                            + elem_min2[0] * basis[0]
                                + elem_min2[1] * basis[1]
                            + elem_min2[2] * basis[2];
                    }
                }
            }
        }
        private void CheckAllDataOK(TideModel tideModel)
        {
            string NotUsed = "";
            if (tideModel.Lng < -180 || tideModel.Lng > 180)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.PleaseEnterAValid_, TaskRunnerServiceRes.Longitude) + TaskRunnerServiceRes.LongitudeInfo;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("PleaseEnterAValid_", TaskRunnerServiceRes.Longitude + TaskRunnerServiceRes.LongitudeInfo);
                return;
            }

            if (tideModel.Lat < -90 || tideModel.Lat > 90)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.PleaseEnterAValid_, TaskRunnerServiceRes.Latitude) + TaskRunnerServiceRes.LatitudeInfo;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("PleaseEnterAValid_", TaskRunnerServiceRes.Latitude + TaskRunnerServiceRes.LatitudeInfo);
                return;
            }

            if (tideModel.Steps_min < 1)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.PleaseEnterAValid_, TaskRunnerServiceRes.Steps) + TaskRunnerServiceRes.StepsInfo;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("PleaseEnterAValid_", TaskRunnerServiceRes.Steps + TaskRunnerServiceRes.StepsInfo);
                return;
            }

            if (tideModel.StartDate.Year < 1900)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.PleaseEnterAValid_, TaskRunnerServiceRes.StartDate) + TaskRunnerServiceRes.StartDateInfo;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("PleaseEnterAValid_", TaskRunnerServiceRes.StartDate + TaskRunnerServiceRes.StartDateInfo);
                return;
            }

            if (tideModel.StartDate >= tideModel.EndDate)
            {
                NotUsed = TaskRunnerServiceRes.StartDateChronologicallyEarlierThanEndDate;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("StartDateChronologicallyEarlierThanEndDate");
                return;
            }
        }
        private bool CreateMemoryInput(TideModel tideModel)
        {
            InputAndResultList = new List<InputAndResult>();

            while (tideModel.EndDate > tideModel.StartDate)
            {
                InputAndResultList.Add(new InputAndResult()
                {
                    Longitude = tideModel.Lng,
                    Latitude = tideModel.Lat,
                    Date = tideModel.StartDate
                });
                tideModel.StartDate = tideModel.StartDate.AddMinutes(tideModel.Steps_min);
            }

            return true;
        }
        private TideNode FindClosestNode(InputAndResult mi)
        {
            // Find the node that is closest to the point (mi.ptx, mi.pty). 
            double currdist, closedist;
            TideNode n = new TideNode();
            n = null;

            closedist = Math.Abs(mi.Latitude - TideNodeList[0].y) + Math.Abs(mi.Longitude - TideNodeList[0].x);

            foreach (TideNode node in TideNodeList)
            {
                currdist = Math.Abs(mi.Latitude - node.y) + Math.Abs(mi.Longitude - node.x);
                if (currdist < closedist)
                {
                    closedist = currdist;
                    n = node;
                }
            }
            return (n);
        }
        private long GetJulianDay(DateTime Date)
        {
            string NotUsed = "";
            /* Calculate the Julian day number.
                Accounts for the change to the Gregorian calandar. */
            long jul, ja, jy, jm;

            jy = Date.Year;
            if (jy == 0)
            {
                NotUsed = TaskRunnerServiceRes.JulianDayThereIsNoYear0;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("JulianDayThereIsNoYear0");
                return 0;
            }
            if (jy < 0) ++jy;
            if (Date.Month > 2)
            {
                jm = Date.Month + 1;
            }
            else
            {
                --jy;
                jm = Date.Month + 13;
            }

            jul = (long)(Math.Floor(365.25 * jy) + Math.Floor(30.6001 * jm) +
                            Date.Day + 1720995);
            if ((Date.Day + 31L * (Date.Month + 12L * Date.Year)) >= (15 + (long)31 * (10 + (long)12 * 1582)))
            {
                ja = (long)(0.01 * jy);
                jul += 2 - ja + (long)(0.25 * ja);
            }
            return jul;
        }
        public void LoadBase(TideModel tideModel)
        {
            ReadConfig(tideModel);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            ReadNodes(tideModel);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            ReadElements(tideModel);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            ReadConstituents(tideModel);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            ReadIOS_TideTable(tideModel);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            LoadConstituentsValues(tideModel);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;
        }
        private void LoadConstituentsValues(TideModel tideModel)
        {
            string NotUsed = "";
            foreach (Constituent ct in ConstituentNameList)
            {
                StreamReader sr;

                FileInfo fiConstituent = new FileInfo(tideModel.TideDataPath + ct.Name + ".barotropic." + (tideModel.DoWaterLevels ? "s2c" : "v2c"));

                if (!fiConstituent.Exists)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiConstituent.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiConstituent.FullName);
                    return;
                }

                try
                {
                    sr = fiConstituent.OpenText();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotOpenFile_Error_, fiConstituent.FullName, ex.Message);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotOpenFile_Error_", fiConstituent.FullName, ex.Message);
                    return;
                }

                sr.ReadLine();
                sr.ReadLine();
                if (ct.Name != "Z0")
                {
                    sr.ReadLine();
                }

                foreach (TideNode node in TideNodeList)
                {
                    ParsedValues = sr.ReadLine().Split(default(string[]), StringSplitOptions.RemoveEmptyEntries).ToList<string>();
                    if (tideModel.DoWaterLevels)
                    {
                        if (ParsedValues.Count() != 3)
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, fiConstituent.FullName);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", fiConstituent.FullName);
                            return;
                        }
                        if (Thread.CurrentThread.CurrentCulture.Name == "fr-CA")
                        {
                            ct.amp.Add(double.Parse(ParsedValues[1].Replace(".", ",")));
                            ct.phase.Add(double.Parse(ParsedValues[2].Replace(".", ",")));
                        }
                        else
                        {
                            ct.amp.Add(double.Parse(ParsedValues[1]));
                            ct.phase.Add(double.Parse(ParsedValues[2]));
                        }
                    }
                    else
                    {
                        if (ct.Name == "Z0")
                        {
                            if (ParsedValues.Count() != 3)
                            {
                                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, fiConstituent.FullName);
                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", fiConstituent.FullName);
                                return;
                            }
                            if (Thread.CurrentThread.CurrentCulture.Name == "fr-CA")
                            {
                                ct.amp.Add(double.Parse(ParsedValues[1].Replace(".", ",")));
                                ct.amp2.Add(double.Parse(ParsedValues[2].Replace(".", ",")));
                                ct.phase.Add(double.Parse(("0.0").Replace(".", ",")));
                                ct.phase2.Add(double.Parse(("0.0").Replace(".", ",")));
                            }
                            else
                            {
                                ct.amp.Add(double.Parse(ParsedValues[1]));
                                ct.amp2.Add(double.Parse(ParsedValues[2]));
                                ct.phase.Add((double)0.0);
                                ct.phase2.Add((double)0.0);
                            }
                        }
                        else
                        {
                            if (ParsedValues.Count() != 5)
                            {
                                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, fiConstituent.FullName);
                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", fiConstituent.FullName);
                                return;
                            }
                            if (Thread.CurrentThread.CurrentCulture.Name == "fr-CA")
                            {
                                ct.amp.Add(double.Parse(ParsedValues[1].Replace(".", ",")));
                                ct.phase.Add(double.Parse(ParsedValues[2].Replace(".", ",")));
                                ct.amp2.Add(double.Parse(ParsedValues[3].Replace(".", ",")));
                                ct.phase2.Add(double.Parse(ParsedValues[4].Replace(".", ",")));
                            }
                            else
                            {
                                ct.amp.Add(double.Parse(ParsedValues[1]));
                                ct.phase.Add(double.Parse(ParsedValues[2]));
                                ct.amp2.Add(double.Parse(ParsedValues[3]));
                                ct.phase2.Add(double.Parse(ParsedValues[4]));
                            }
                        }
                    }
                    if (sr.EndOfStream)
                    {
                        break;
                    }
                }

                sr.Close();
            }
        }
        private List<double> Phi2d(List<double> xloc, List<double> yloc, InputAndResult mi)
        {
            /* Calculates the basis functions for interpolating to a point inside
                an element. */
            List<double> phi = new List<double>();
            double area, a, b, c;
            int j = 0, k = 0;

            area = 0.5 * (xloc[0] * (yloc[1] - yloc[2]) +
                            xloc[1] * (yloc[2] - yloc[0]) +
                            xloc[2] * (yloc[0] - yloc[1]));

            /* Calculate the Basis function... */
            for (int i = 0; i <= 2; i++)
            {
                switch (i)
                {
                    case 0: j = 1; k = 2; break;
                    case 1: j = 2; k = 0; break;
                    case 2: j = 0; k = 1; break;
                }
                a = (xloc[j] * yloc[k] - xloc[k] * yloc[j]) / (area * 2);
                b = (yloc[j] - yloc[k]) / (area * 2);
                c = -1 * (xloc[j] - xloc[k]) / (area * 2);
                phi.Add(a + b * mi.Longitude + c * mi.Latitude);
            }
            return (phi);
        }
        private int RayBound(List<double> xd, List<double> yd, InputAndResult mi)
        {
            /*  Subroutine to check wether or not a point is inside a polygon.
            The process is as follows:
                Use an arbitrary ray (here, y = constant and x >= xref), starting from 
            the point and going off to infinity.
                Count the number of polygon boundaries it crosses.
                If an odd number, the point is inside the polygon, otherwise it is
            outside.   */
            int j, bcross;
            double b, m, x;

            bcross = 0; /* Number of boundary crossings. */

            /* Check to see if the element side crosses the International Dateline
                (changes sign at +180/-180 degrees) and if so, change the longitudes
                so that they all have the same sign. */

            if (mi.Longitude > 0.0)
            {
                if ((xd[0] < -170.0) && ((xd[1] > 170.0) || (xd[2] > 170.0)))
                {
                    xd[0] += 360.0;
                }
                if ((xd[1] < -170.0) && ((xd[0] > 170.0) || (xd[2] > 170.0)))
                {
                    xd[1] += 360.0;
                }
                if ((xd[2] < -170.0) && ((xd[1] > 170.0) || (xd[0] > 170.0)))
                {
                    xd[2] += 360.0;
                }
            }
            else
            {
                if ((xd[0] > 170.0) && ((xd[1] < -170.0) || (xd[2] < -170.0)))
                {
                    xd[0] -= 360.0;
                }
                if ((xd[1] > 170.0) && ((xd[0] < -170.0) || (xd[2] < -170.0)))
                {
                    xd[1] -= 360.0;
                }
                if ((xd[2] > 170.0) && ((xd[1] < -170.0) || (xd[0] < -170.0)))
                {
                    xd[2] -= 360.0;
                }
            }

            /* As above, except for the Greenwich meridian, for longitude coordinates
                that range from 0 to 360 degrees. */

            if (mi.Longitude > 350.0)
            {
                if ((xd[0] < 10.0) && ((xd[1] > 350.0) || (xd[2] > 350.0)))
                {
                    xd[0] += 360.0;
                }
                if ((xd[1] < 10.0) && ((xd[0] > 350.0) || (xd[2] > 350.0)))
                {
                    xd[1] += 360.0;
                }
                if ((xd[2] < 10.0) && ((xd[1] > 350.0) || (xd[0] > 350.0)))
                {
                    xd[2] += 360.0;
                }
            }
            else
            {
                if ((xd[0] > 350.0) && ((xd[1] < 10.0) || (xd[2] < 10.0)))
                {
                    xd[0] -= 360.0;
                }
                if ((xd[1] > 350.0) && ((xd[0] < 10.0) || (xd[2] < 10.0)))
                {
                    xd[1] -= 360.0;
                }
                if ((xd[2] > 350.0) && ((xd[1] < 10.0) || (xd[0] < 10.0)))
                {
                    xd[2] -= 360.0;
                }
            }

            for (int i = 0; i <= 2; i++)
            {

                /* for each line segment around the element */
                j = ((i == 2) ? 0 : i + 1);

                /* If both endpoints of the line segment are on the same (vertical)
                side of the ray, do nothing.
                    Otherwise, count the number of times the ray intersects the segment. */
                if (!(((yd[i] < mi.Latitude) && (yd[j] < mi.Latitude)) ||
                        ((yd[i] >= mi.Latitude) && (yd[j] >= mi.Latitude))))
                {

                    if (xd[i] != xd[j])
                    {
                        m = (yd[j] - yd[i]) / (xd[j] - xd[i]);
                        b = yd[i] - m * xd[i];
                        x = (mi.Latitude - b) / m;
                        if (x > mi.Longitude)
                        {
                            bcross++;
                        }
                    }
                    else
                    {
                        if (xd[j] > mi.Longitude)
                        {
                            bcross++;
                        }
                    }
                }
            }

            /*  Return the evenness/oddness of the boundary crossings
                    i.e. the remainder from division by two. */
            return (bcross % 2);
        }
        private void ReadConfig(TideModel tideModel)
        {
            string NotUsed = "";
            string TheLine = "";
            string Last4Letters = "";
            StreamReader sr;

            FileInfo fiCfg = new FileInfo(tideModel.TideDataPath + "tidecor.cfg");

            //RaiseMessageEvent(string.Format("Reading Config File [{0}]", fiCfg.FullName));
            //RaiseStatusEvent(string.Format("Reading Config File [{0}]", fiCfg.FullName));

            if (!fiCfg.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiCfg.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiCfg.FullName);
                return;
            }

            try
            {
                sr = fiCfg.OpenText();
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotOpenFile_Error_, fiCfg.FullName, ex.Message);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotOpenFile_Error_", fiCfg.FullName, ex.Message);
                return;
            }

            while (true)
            {
                TheLine = sr.ReadLine().Trim();
                Last4Letters = TheLine.Substring(TheLine.Length - 4).ToLower();
                switch (Last4Letters)
                {
                    case ".nod":
                        {
                            config.NodeFileName = TheLine;
                        }
                        break;
                    case ".ele":
                        {
                            config.ElementFileName = TheLine;
                        }
                        break;
                    case "etbl":
                        {
                            config.IOS_FileName = TheLine;
                        }
                        break;
                    default:
                        break;
                }
                if (sr.EndOfStream)
                {
                    break;
                }
            }

            sr.Close();
        }
        private void ReadConstituents(TideModel tideModel)
        {
            string NotUsed = "";
            ConstituentNameList = new List<Constituent>();
            StreamReader sr;

            FileInfo fiConst = new FileInfo(tideModel.TideDataPath + "constituents.txt");

            if (!fiConst.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiConst.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiConst.FullName);
                return;
            }

            try
            {
                sr = fiConst.OpenText();
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotOpenFile_Error_, fiConst.FullName, ex.Message);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotOpenFile_Error_", fiConst.FullName, ex.Message);
                return;
            }

            while (true)
            {
                ParsedValues = sr.ReadLine().Split(default(string[]), StringSplitOptions.RemoveEmptyEntries).ToList<string>();
                if (ParsedValues.Count() != 1)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, fiConst.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", fiConst.FullName);
                    return;
                }
                if (ParsedValues[0].Trim().ToUpper() == "NONE")
                {
                    break;
                }
                ConstituentNameList.Add(new Constituent() { Name = ParsedValues[0].Trim().ToUpper() });
                if (sr.EndOfStream)
                {
                    break;
                }
            }
            sr.Close();
        }
        private void ReadElements(TideModel tideModel)
        {
            string NotUsed = "";
            TideElementList = new List<TideElement>();
            StreamReader sr;

            FileInfo fiElem = new FileInfo(tideModel.TideDataPath + config.ElementFileName);

            if (!fiElem.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiElem.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiElem.FullName);
                return;
            }

            try
            {
                sr = fiElem.OpenText();
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotOpenFile_Error_, fiElem.FullName, ex.Message);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotOpenFile_Error_", fiElem.FullName, ex.Message);
                return;
            }

            while (true)
            {
                ParsedValues = sr.ReadLine().Split(default(string[]), StringSplitOptions.RemoveEmptyEntries).ToList<string>();
                if (ParsedValues.Count() != 4)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, fiElem.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", fiElem.FullName);
                    return;
                }
                List<TideNode> TempNodeList = new List<TideNode>();
                TempNodeList.Add(TideNodeList[int.Parse(ParsedValues[1]) - 1]);
                TempNodeList.Add(TideNodeList[int.Parse(ParsedValues[2]) - 1]);
                TempNodeList.Add(TideNodeList[int.Parse(ParsedValues[3]) - 1]);

                TideElementList.Add(new TideElement() { ID = int.Parse(ParsedValues[0]), NodeList = TempNodeList });
                if (sr.EndOfStream)
                {
                    break;
                }
            }
            sr.Close();
        }
        private void ReadIOS_TideTable(TideModel tideModel)
        {
            string NotUsed = "";
            MainConstituentList = new List<MainConstituent>();
            ShallowConstituentList = new List<ShallowConstituent>();

            /* Read in the constituent data from the IOS_tidetbl file */
            int cnt, nln;
            string satin1, satin2, satin3;

            FileInfo fiIOS = new FileInfo(tideModel.TideDataPath + config.IOS_FileName);

            if (!fiIOS.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiIOS.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiIOS.FullName);
                return;
            }

            /* Read in the main constituents*/

            StreamReader sr = fiIOS.OpenText();

            while (true) // this while will not hit the end of stream, it will be stopped by a blank line.
            {
                ParsedValues = sr.ReadLine().Split(default(string[]), StringSplitOptions.RemoveEmptyEntries).ToList<string>();

                if (ParsedValues.Count == 0)
                {
                    break;
                }
                MainConstituent TempNewCon = new MainConstituent();
                TempNewCon.Name = ParsedValues[0];
                List<int> dood = new List<int>();
                TempNewCon.dood.Add(int.Parse(ParsedValues[1]));
                TempNewCon.dood.Add(int.Parse(ParsedValues[2]));
                TempNewCon.dood.Add(int.Parse(ParsedValues[3]));
                TempNewCon.dood.Add(int.Parse(ParsedValues[4]));
                TempNewCon.dood.Add(int.Parse(ParsedValues[5]));
                if (ParsedValues[6].LastIndexOf("-") > 0)
                {
                    ParsedValues.Add(ParsedValues[7]);
                    ParsedValues[7] = ParsedValues[6].Substring(ParsedValues[6].LastIndexOf("-"));
                    ParsedValues[6] = ParsedValues[6].Substring(0, ParsedValues[6].LastIndexOf("-"));
                }
                TempNewCon.dood.Add(int.Parse(ParsedValues[6]));
                if (Thread.CurrentThread.CurrentCulture.Name == "fr-CA")
                {
                    TempNewCon.phase = double.Parse(ParsedValues[7].Replace(".", ","));
                }
                else
                {
                    TempNewCon.phase = double.Parse(ParsedValues[7]);
                }

                int NumberOfSatellite = int.Parse(ParsedValues[8]);
                /* Read in the satellites for this constituent, if any */
                if (NumberOfSatellite > 0)
                {
                    nln = ((NumberOfSatellite - 1) / 3) + 1;
                    for (int i = 0; i < nln; i++)
                    {
                        string TheLine = sr.ReadLine();
                        cnt = NumberOfSatellite - (i * 3); /* # of satellites on this line */
                        switch (cnt)
                        {
                            case 1:
                                satin1 = TheLine.Substring(12);
                                TempNewCon.satellites.Add(AddSatellite(satin1));
                                break;
                            case 2:
                                satin1 = TheLine.Substring(12, 23);
                                satin2 = TheLine.Substring(12 + 23);
                                TempNewCon.satellites.Add(AddSatellite(satin1));
                                TempNewCon.satellites.Add(AddSatellite(satin2));
                                break;
                            default:
                                satin1 = TheLine.Substring(12, 23);
                                satin2 = TheLine.Substring(12 + 23, 23);
                                satin3 = TheLine.Substring(12 + 23 + 23);
                                TempNewCon.satellites.Add(AddSatellite(satin1));
                                TempNewCon.satellites.Add(AddSatellite(satin2));
                                TempNewCon.satellites.Add(AddSatellite(satin3));
                                break;
                        }
                    }
                }
                else
                {
                    /*  NO satellites */
                    TempNewCon.satellites = null;
                }
                MainConstituentList.Add(TempNewCon);

                if (sr.EndOfStream)
                {
                    break;
                }
                if (ParsedValues.Count() == 0)
                {
                    return;
                }
            }



            /* Read in the shallow water constiuents */

            while (true)
            {
                ParsedValues = sr.ReadLine().Split(default(string[]), StringSplitOptions.RemoveEmptyEntries).ToList<string>();

                if (ParsedValues.Count == 0)
                {
                    break;
                }

                ShallowConstituent ns = new ShallowConstituent();

                ns.Name = ParsedValues[0];
                int NumberOfFactors = int.Parse(ParsedValues[1]);

                switch (NumberOfFactors)
                {
                    case 4:
                        {
                            if (Thread.CurrentThread.CurrentCulture.Name == "fr-CA")
                            {
                                ns.factors.Add(new ShallowConstituent.Factor() { Name = ParsedValues[3], factor = double.Parse(ParsedValues[2].Replace(".", ",")) });
                                ns.factors.Add(new ShallowConstituent.Factor() { Name = ParsedValues[5], factor = double.Parse(ParsedValues[4].Replace(".", ",")) });
                                ns.factors.Add(new ShallowConstituent.Factor() { Name = ParsedValues[7], factor = double.Parse(ParsedValues[6].Replace(".", ",")) });
                                ns.factors.Add(new ShallowConstituent.Factor() { Name = ParsedValues[9], factor = double.Parse(ParsedValues[8].Replace(".", ",")) });
                            }
                            else
                            {
                                ns.factors.Add(new ShallowConstituent.Factor() { Name = ParsedValues[3], factor = double.Parse(ParsedValues[2]) });
                                ns.factors.Add(new ShallowConstituent.Factor() { Name = ParsedValues[5], factor = double.Parse(ParsedValues[4]) });
                                ns.factors.Add(new ShallowConstituent.Factor() { Name = ParsedValues[7], factor = double.Parse(ParsedValues[6]) });
                                ns.factors.Add(new ShallowConstituent.Factor() { Name = ParsedValues[9], factor = double.Parse(ParsedValues[8]) });
                            }
                        }
                        break;
                    case 3:
                        {
                            if (Thread.CurrentThread.CurrentCulture.Name == "fr-CA")
                            {
                                ns.factors.Add(new ShallowConstituent.Factor() { Name = ParsedValues[3], factor = double.Parse(ParsedValues[2].Replace(".", ",")) });
                                ns.factors.Add(new ShallowConstituent.Factor() { Name = ParsedValues[5], factor = double.Parse(ParsedValues[4].Replace(".", ",")) });
                                ns.factors.Add(new ShallowConstituent.Factor() { Name = ParsedValues[7], factor = double.Parse(ParsedValues[6].Replace(".", ",")) });
                            }
                            else
                            {
                                ns.factors.Add(new ShallowConstituent.Factor() { Name = ParsedValues[3], factor = double.Parse(ParsedValues[2]) });
                                ns.factors.Add(new ShallowConstituent.Factor() { Name = ParsedValues[5], factor = double.Parse(ParsedValues[4]) });
                                ns.factors.Add(new ShallowConstituent.Factor() { Name = ParsedValues[7], factor = double.Parse(ParsedValues[6]) });
                            }
                        }
                        break;
                    case 2:
                        {
                            if (Thread.CurrentThread.CurrentCulture.Name == "fr-CA")
                            {
                                ns.factors.Add(new ShallowConstituent.Factor() { Name = ParsedValues[3], factor = double.Parse(ParsedValues[2].Replace(".", ",")) });
                                ns.factors.Add(new ShallowConstituent.Factor() { Name = ParsedValues[5], factor = double.Parse(ParsedValues[4].Replace(".", ",")) });
                            }
                            else
                            {
                                ns.factors.Add(new ShallowConstituent.Factor() { Name = ParsedValues[3], factor = double.Parse(ParsedValues[2]) });
                                ns.factors.Add(new ShallowConstituent.Factor() { Name = ParsedValues[5], factor = double.Parse(ParsedValues[4]) });
                            }
                        }
                        break;
                    case 1:
                        {
                            if (Thread.CurrentThread.CurrentCulture.Name == "fr-CA")
                            {
                                ns.factors.Add(new ShallowConstituent.Factor() { Name = ParsedValues[3], factor = double.Parse(ParsedValues[2].Replace(".", ",")) });
                            }
                            else
                            {
                                ns.factors.Add(new ShallowConstituent.Factor() { Name = ParsedValues[3], factor = double.Parse(ParsedValues[2]) });
                            }
                        }
                        break;
                }

                ShallowConstituentList.Add(ns);

                if (sr.EndOfStream)
                {
                    break;
                }
            }

            sr.Close();
        }
        private void ReadNodes(TideModel tideModel)
        {
            string NotUsed = "";
            TideNodeList = new List<TideNode>();
            StreamReader sr;
            FileInfo fiNode = new FileInfo(tideModel.TideDataPath + config.NodeFileName);

            //RaiseMessageEvent(string.Format("Reading Node File [{0}]", fiNode.FullName));
            //RaiseStatusEvent(string.Format("Reading Node File [{0}]", fiNode.FullName));

            if (!fiNode.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiNode.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiNode.FullName);
                return;
            }

            try
            {
                sr = fiNode.OpenText();
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotOpenFile_Error_, fiNode.FullName, ex.Message);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotOpenFile_Error_", fiNode.FullName, ex.Message);
                return;
            }

            while (true)
            {
                ParsedValues = sr.ReadLine().Split(default(string[]), StringSplitOptions.RemoveEmptyEntries).ToList<string>();
                if (ParsedValues.Count() != 3)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, fiNode.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", fiNode.FullName);
                    return;
                }
                if (Thread.CurrentThread.CurrentCulture.Name == "fr-CA")
                {
                    TideNodeList.Add(new TideNode() { ID = int.Parse(ParsedValues[0]), x = double.Parse(ParsedValues[1].Replace(".", ",")), y = double.Parse(ParsedValues[2].Replace(".", ",")), MinimumTide = 0, MinimumTide2 = 0 });
                }
                else
                {
                    TideNodeList.Add(new TideNode() { ID = int.Parse(ParsedValues[0]), x = double.Parse(ParsedValues[1]), y = double.Parse(ParsedValues[2]), MinimumTide = 0, MinimumTide2 = 0 });
                }
                if (sr.EndOfStream)
                {
                    break;
                }
            }

            sr.Close();
        }
        private void SetVuf(long kh, double xlat)
        {
            Astro astro = new Astro();

            /* Calculate the amplitudes, phases, etc. for each of the constituents */
            double hh, tau, dtau, slat, sumc, sums, v, vdbl;
            double adj = 0, uu, uudbl;
            long kd, ktmp;

            kd = GetJulianDay(new DateTime(1899, 12, 31));

            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            astro.d1 = (double)(kh - kd) - 0.5;
            ktmp = kh * 24;


            AstroAngles(ref astro);

            hh = (double)ktmp - (Math.Floor((double)(ktmp / 24.0)) * 24.0);
            tau = hh / 24.0 + astro.h - astro.s;
            dtau = 365.00 + astro.dh - astro.ds;
            slat = Math.Sin((Math.PI) * (xlat / 180.0));

            /* The main constituents */
            foreach (MainConstituent mc in MainConstituentList)
            {
                mc.freq = (mc.dood[0] * dtau + mc.dood[1] * astro.ds + mc.dood[2] * astro.dh + mc.dood[3] * astro.dp +
                                    mc.dood[4] * astro.dnp + mc.dood[5] * astro.dpp) / (24.0 * 365.0);

                vdbl = mc.dood[0] * tau + mc.dood[1] * astro.s + mc.dood[2] * astro.h + mc.dood[3] * astro.p +
                                    mc.dood[4] * astro.enp + mc.dood[5] * astro.pp + mc.phase;

                v = vdbl - (Math.Floor(Math.Floor(vdbl) / 2.0) * 2.0);

                sumc = 1.0;
                sums = 0.0;

                if (mc.satellites != null)
                {
                    foreach (MainConstituent.Satellite satellite in mc.satellites)
                    {
                        switch (satellite.corr)
                        {
                            case 0:
                                adj = satellite.ratio;
                                break;
                            case 1:
                                adj = satellite.ratio * 0.36309 * (1.0 - 5.0 * slat * slat) / slat;
                                break;
                            case 2:
                                adj = satellite.ratio * 2.59808 * slat;
                                break;
                        }
                        uudbl = (double)satellite.deldList[0] * astro.p + (double)satellite.deldList[1] * astro.enp +
                                (double)satellite.deldList[2] * astro.pp + (double)satellite.phase;
                        uu = uudbl - (int)uudbl;

                        sumc += (adj * Math.Cos(uu * Math.PI * 2));
                        sums += (adj * Math.Sin(uu * Math.PI * 2));
                    }
                }
                mc.f = Math.Sqrt((sumc * sumc) + (sums * sums));
                mc.vu = v + Math.Atan2(sums, sumc) / (Math.PI * 2);
            }

            /* The shallow water constituents */
            foreach (ShallowConstituent shallowConstituent in ShallowConstituentList)
            {
                shallowConstituent.f = 1.0;
                shallowConstituent.vu = 0.0;
                shallowConstituent.freq = 0.0;
                foreach (ShallowConstituent.Factor factor in shallowConstituent.factors)
                {
                    foreach (MainConstituent mainConstituent in MainConstituentList)
                    {
                        if (mainConstituent.Name == factor.Name)
                        {
                            shallowConstituent.f *= Math.Pow(mainConstituent.f, Math.Abs(factor.factor));
                            shallowConstituent.vu += (factor.factor * mainConstituent.vu);
                            shallowConstituent.freq += (factor.factor * mainConstituent.freq);
                            break;
                        }
                    }
                }
            }
        }
        private double TideP(InputAndResult mi, TideNode n, bool IsWL)
        {
            string NotUsed = "";
            /* Calculates and returns the tidal correction */
            int indx;
            long kd;
            double dthr, radgmt, revgmt, res;

            kd = GetJulianDay(mi.Date);

            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return -999;

            SetVuf(kd, mi.Latitude);
            dthr = (((double)mi.Date.Hour * 3600.0) + ((double)mi.Date.Minute * 60.0) +
                        (double)mi.Date.Second) / 3600.0;
            res = 0.0;

            /* For each of the desired constituents ... (See top of program) */
            foreach (Constituent constituent in ConstituentNameList)
            {
                /* Find the constituent from those loaded from IOS_tidetbl */
                indx = Vuf(constituent.Name);
                if (indx < 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.BadInputConstituant_, constituent.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("BadInputConstituant_", constituent.ToString());
                    return -999;
                }

                if (indx < MainConstituentList.Count)
                {                        // Main constituent 
                    if (IsWL)
                        revgmt = MainConstituentList[indx].freq * dthr + MainConstituentList[indx].vu - constituent.phase[n.ID - 1] / 360.0;
                    else
                        revgmt = MainConstituentList[indx].freq * dthr + MainConstituentList[indx].vu - constituent.phase2[n.ID - 1] / 360.0;

                    radgmt = Math.PI * 2 * (revgmt - (int)revgmt);

                    if (IsWL)
                        res += MainConstituentList[indx].f * constituent.amp[n.ID - 1] * Math.Cos(radgmt);
                    else
                        res += MainConstituentList[indx].f * constituent.amp2[n.ID - 1] * Math.Cos(radgmt);
                }
                else if ((indx - MainConstituentList.Count) < ShallowConstituentList.Count)
                {     // Shallow water constituent 
                    indx -= MainConstituentList.Count;
                    if (IsWL)
                        revgmt = ShallowConstituentList[indx].freq * dthr + ShallowConstituentList[indx].vu - constituent.phase[n.ID - 1] / 360.0;
                    else
                        revgmt = ShallowConstituentList[indx].freq * dthr + ShallowConstituentList[indx].vu - constituent.phase2[n.ID - 1] / 360.0;

                    radgmt = Math.PI * 2 * (revgmt - (int)revgmt);
                    if (IsWL)
                        res += ShallowConstituentList[indx].f * constituent.amp[n.ID - 1] * Math.Cos(radgmt);
                    else
                        res += ShallowConstituentList[indx].f * constituent.amp2[n.ID - 1] * Math.Cos(radgmt);
                }
                else
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.ErrorInIndex___, indx.ToString(), MainConstituentList.Count.ToString(), ShallowConstituentList.Count.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("ErrorInIndex___", indx.ToString(), MainConstituentList.Count.ToString(), ShallowConstituentList.Count.ToString());
                    return -999;
                }
            }

            return (res);
        }
        private int Vuf(string inname)
        {
            string NotUsed = "";
            /* Finds constituent info corresponding to inname and returns the index to
                the node containing the info. Shallow water constituent indices are
                returned as their number greater than the max # of main constituents.
                e.g. if we want the 2nd shallow water constituent and there are 45
                main constituents, then the indice returned is 46, since constituents
                are counted from zero. ( 45 - 1 + 2 = 46 ) */
            int i, j;

            i = 0;
            j = 0;
            while (MainConstituentList[i].Name != inname)
            {
                i++;
                if (i == MainConstituentList.Count) break;
            }
            if (i == MainConstituentList.Count)
            {
                j = 0;
                while (ShallowConstituentList[j].Name != inname)
                {
                    j++;
                    if (j == ShallowConstituentList.Count) break;
                }
            }

            if (i < MainConstituentList.Count)
            {
                return (i);
            }
            else if (j < ShallowConstituentList.Count)
            {
                return (i + j);
            }
            else
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, inname);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", inname);
                return -999;
            }
        }
        #endregion Functions private

    }


    public class TideModel
    {
        public TideModel(string BasePath, WebTideDataSetEnum WebTideDataSet)
        {
            TideDataPath = BasePath + @"WebTide\data\" + WebTideDataSet.ToString() + @"\";

        }

        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public double Lng { get; set; }
        public double Lat { get; set; }
        public double Steps_min { get; set; }
        public WebTideDataSetEnum WebTideDataSet { get; set; }
        public bool DoWaterLevels { get; set; }
        public string TideDataPath { get; set; }
    }
}
