using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CSSPWebToolsTaskRunner.Services.Resources;
using CSSPWebToolsTaskRunner.Services;
using System.IO;
using System.Transactions;
using System.Windows.Forms;
using System.Threading;
using System.ComponentModel;
using CSSPDBDLL.Models;
using CSSPDBDLL.Services;
using System.Net;
using CSSPModelsDLL.Models;
using CSSPEnumsDLL.Enums;
using CSSPDBDLL;
using System.Xml;
using System.Drawing;

namespace CSSPWebToolsTaskRunner.Services
{
    public class ClassificationGenerateService
    {
        #region Variables
        public List<string> ProvList = new List<string>() { "BC", "NB", "NL", "NS", "PE", "QC" };
        public List<string> ProvNameListEN = new List<string>() { "British Columbia", "New Brunswick", "Newfoundland and Labrador", "Nova Scotia", "Prince Edward Island", "Québec" };
        public List<string> ProvNameListFR = new List<string>() { "Colombie-Britannique", "Nouveau-Brunswick", "Terre-Neuve-et-Labrador", "Nouvelle-Écosse", "Île-du-Prince-Édouard", "Québec" };
        #endregion Variables

        #region Properties
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        #endregion Properties

        #region Constructors
        public ClassificationGenerateService(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
        }
        #endregion Constructors

        #region Events
        #endregion Events

        #region Functions public
        public void GenerateClassificationForCSSPWebToolsVisualization()
        {
            string NotUsed = "";

            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            ClassificationService classificationService = new ClassificationService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            ProvinceToolsService provinceToolsService = new ProvinceToolsService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);

            if (_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._Required, TaskRunnerServiceRes.TVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_Required", TaskRunnerServiceRes.TVItemID);
                return;
            }
            if (_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID2 == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._Required, TaskRunnerServiceRes.TVItemID2);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_Required", TaskRunnerServiceRes.TVItemID2);
                return;
            }

            TVItemModel tvItemModelProvince = tvItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelProvince.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                return;
            }

            if (tvItemModelProvince.TVType != TVTypeEnum.Province)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.TVTypeShouldBe_, TVTypeEnum.Province.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("TVTypeShouldBe_", TVTypeEnum.Province.ToString());
                return;
            }

            string Parameters = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters;
            string[] ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            int ProvinceTVItemID = 0;
            foreach (string s in ParamValueList)
            {
                string[] ParamValue = s.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                if (ParamValue.Length != 2)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, TaskRunnerServiceRes.Parameters);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", TaskRunnerServiceRes.Parameters);
                    return;
                }

                if (ParamValue[0] == "ProvinceTVItemID")
                {
                    ProvinceTVItemID = int.Parse(ParamValue[1]);
                }
                else
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, ParamValue[0]);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", ParamValue[0].ToString());
                    return;
                }
            }

            if (tvItemModelProvince.TVItemID != ProvinceTVItemID)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._NotEqualTo_, "tvItemModelProvince.TVItemID[" + tvItemModelProvince.TVItemID.ToString() + "]", "ProvinceTVItemID[" + ProvinceTVItemID.ToString() + "]");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("_NotEqualTo_", "tvItemModelProvince.TVItemID[" + tvItemModelProvince.TVItemID.ToString() + "]", "ProvinceTVItemID[" + ProvinceTVItemID.ToString() + "]");
                return;
            }

            string ServerPath = tvFileService.GetServerFilePath(ProvinceTVItemID);
            string Init = provinceToolsService.GetInit(ProvinceTVItemID);

            #region Reading the ClassificationPolygons_XX.KML
            List<PolyObj> polyObjList = new List<PolyObj>();

            string FileName = $"ClassificationPolygons_{Init}.kml";

            FileInfo fi = new FileInfo(ServerPath + FileName);

            if (!fi.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fi.FullName);
                return;
            }

            StreamReader sr = fi.OpenText();
            StringBuilder sb = new StringBuilder(sr.ReadToEnd());
            sb.Replace("xsi:schemaLocation", "xmlns");
            sr.Close();

            StreamWriter sw = fi.CreateText();
            sw.Write(sb.ToString());
            sw.Close();

            XmlDocument doc = new XmlDocument();
            doc.Load(fi.FullName);

            foreach (XmlNode xmlNode in doc.ChildNodes)
            {
                GetCoordinates(polyObjList, xmlNode);
            }

            #endregion Reading the ClassificationPolygons_XX.KML

            #region Reading the ClassificationInputs_XX.KML
            List<PolyObj> polyObjList2 = new List<PolyObj>();

            string FileName2 = $"ClassificationInputs_{Init}.kml";

            FileInfo fi2 = new FileInfo(ServerPath + FileName2);

            if (!fi2.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fi2.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fi2.FullName);
                return;
            }

            XmlDocument doc2 = new XmlDocument();
            doc2.Load(fi2.FullName);

            string CurrentSubsector = "";
            string CurrentClassification = "";

            XmlNode StartNode2 = doc2.ChildNodes[1].ChildNodes[0];
            foreach (XmlNode n in StartNode2.ChildNodes)
            {
                if (n.Name == "Folder")
                {
                    foreach (XmlNode n1 in n.ChildNodes)
                    {
                        if (n1.Name == "Folder")
                        {
                            CurrentSubsector = "";

                            foreach (XmlNode n22 in n1)
                            {
                                if (n22.Name == "name")
                                {
                                    CurrentSubsector = n22.InnerText;
                                }

                                if (n22.Name == "Placemark")
                                {
                                    CurrentClassification = "";

                                    foreach (XmlNode n2 in n22)
                                    {

                                        if (n2.Name == "name")
                                        {
                                            CurrentClassification = n2.InnerText;
                                        }

                                        if (n2.Name == "LineString")
                                        {
                                            PolyObj polyObj = new PolyObj();

                                            polyObj.Subsector = CurrentSubsector;
                                            polyObj.Classification = CurrentClassification.ToUpper().Trim();

                                            foreach (XmlNode n3 in n2.ChildNodes)
                                            {
                                                if (n3.Name == "coordinates")
                                                {
                                                    string coordText = n3.InnerText.Trim();

                                                    List<string> pointListText = coordText.Split(" ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

                                                    int ordinal = 0;
                                                    foreach (string pointText in pointListText)
                                                    {
                                                        string pointTxt = pointText.Trim();

                                                        if (!string.IsNullOrWhiteSpace(pointTxt))
                                                        {
                                                            List<string> valListText = pointTxt.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

                                                            if (valListText.Count != 3)
                                                            {
                                                                NotUsed = "valListText.Count != 3";
                                                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("valListText.Count != 3");
                                                                return;
                                                            }

                                                            float Lng = float.Parse(valListText[0]);
                                                            float Lat = float.Parse(valListText[1]);

                                                            List<PolyObj> polyObjCloseList = (from c in polyObjList
                                                                                              where c.MinLat <= Lat
                                                                                              && c.MaxLat >= Lat
                                                                                              && c.MinLng <= Lng
                                                                                              && c.MaxLng >= Lng
                                                                                              select c).ToList();

                                                            double distMin = 10000000D;
                                                            foreach (PolyObj polyObjClose in polyObjCloseList)
                                                            {
                                                                foreach (Coord coordClose in polyObjClose.coordList)
                                                                {
                                                                    double dist = mapInfoService.CalculateDistance(Lat * mapInfoService.d2r, Lng * mapInfoService.d2r, coordClose.Lat * mapInfoService.d2r, coordClose.Lng * mapInfoService.d2r, mapInfoService.R);

                                                                    if (dist < 20)
                                                                    {
                                                                        if (distMin > dist)
                                                                        {
                                                                            Lat = coordClose.Lat;
                                                                            Lng = coordClose.Lng;
                                                                            distMin = dist;
                                                                        }
                                                                    }
                                                                }
                                                            }

                                                            Coord coord = new Coord() { Lat = Lat, Lng = Lng, Ordinal = ordinal };

                                                            polyObj.coordList.Add(coord);

                                                            ordinal += 1;
                                                        }
                                                    }
                                                }
                                            }
                                            polyObjList2.Add(polyObj);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            #endregion Reading the ClassificationInputs_XX.KML

            #region Uploading MapInfo to CSSPWebTools
            TVItemModel tvItemModelProv = tvItemService.GetTVItemModelWithTVItemIDDB(ProvinceTVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelProv.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, ProvinceTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, ProvinceTVItemID.ToString());
                return;
            }

            List<TVItemModel> tvitemModelSSList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelProv.TVItemID, TVTypeEnum.Subsector);

            int CountSS = 0;
            string Status = appTaskModel.StatusText;
            foreach (TVItemModel tvItemModelSS in tvitemModelSSList)
            {
                CountSS += 1;
                if (CountSS % 1 == 0)
                {
                    appTaskModel.PercentCompleted = (100 * CountSS) / tvitemModelSSList.Count;
                    appTaskModel.StatusText = Status + " --- " + tvItemModelSS.TVText;
                    appTaskService.PostUpdateAppTask(appTaskModel);
                }
                Application.DoEvents();

                string TVTextSS = tvItemModelSS.TVText.Substring(0, tvItemModelSS.TVText.IndexOf(" "));

                List<TVItemModel> tvItemModelClassificationList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSS.TVItemID, TVTypeEnum.Classification);
                List<int> TVItemIDClassificationList = new List<int>();

                int Ordinal = 0;
                foreach (PolyObj polyObj in polyObjList2.Where(c => c.Subsector == TVTextSS).OrderBy(c => c.Classification))
                {
                    Ordinal += 1;

                    string TVTextClass = "";
                    TVTypeEnum tvType = TVTypeEnum.Error;
                    ClassificationTypeEnum classificationType = ClassificationTypeEnum.Error;

                    switch (polyObj.Classification)
                    {
                        case "R":
                            {
                                TVTextClass = "R " + Ordinal;
                                tvType = TVTypeEnum.Restricted;
                                classificationType = ClassificationTypeEnum.Restricted;
                            }
                            break;
                        case "P":
                            {
                                TVTextClass = "P " + Ordinal;
                                tvType = TVTypeEnum.Prohibited;
                                classificationType = ClassificationTypeEnum.Prohibited;
                            }
                            break;
                        case "A":
                            {
                                TVTextClass = "A " + Ordinal;
                                tvType = TVTypeEnum.Approved;
                                classificationType = ClassificationTypeEnum.Approved;
                            }
                            break;
                        case "CA":
                            {
                                TVTextClass = "CA " + Ordinal;
                                tvType = TVTypeEnum.ConditionallyApproved;
                                classificationType = ClassificationTypeEnum.ConditionallyApproved;
                            }
                            break;
                        case "CR":
                            {
                                TVTextClass = "CR " + Ordinal;
                                tvType = TVTypeEnum.ConditionallyRestricted;
                                classificationType = ClassificationTypeEnum.ConditionallyRestricted;
                            }
                            break;
                        default:
                            break;
                    }

                    TVItemModel tvItemModelClass = tvItemService.GetChildTVItemModelWithParentIDAndTVTextAndTVTypeDB(tvItemModelSS.TVItemID, TVTextClass, TVTypeEnum.Classification);
                    if (!string.IsNullOrWhiteSpace(tvItemModelClass.Error))
                    {
                        tvItemModelClass = tvItemService.PostAddChildTVItemDB(tvItemModelSS.TVItemID, TVTextClass, TVTypeEnum.Classification);
                        if (!string.IsNullOrWhiteSpace(tvItemModelClass.Error))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVItem, tvItemModelClass.Error);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVItem, tvItemModelClass.Error);
                            return;
                        }
                    }

                    TVItemIDClassificationList.Add(tvItemModelClass.TVItemID);

                    bool CoordListIsDifferent = false;
                    List<MapInfoModel> mapInfoModelList = mapInfoService.GetMapInfoModelListWithTVItemIDDB(tvItemModelClass.TVItemID);
                    if (mapInfoModelList.Count == 0)
                    {
                        MapInfoModel mapInfoModelRet = tvItemService.CreateMapInfoObjectDB(polyObj.coordList, MapInfoDrawTypeEnum.Polyline, tvType, tvItemModelClass.TVItemID);
                        if (!string.IsNullOrWhiteSpace(mapInfoModelRet.Error))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.MapInfo, mapInfoModelRet.Error);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.MapInfo, mapInfoModelRet.Error);
                            return;
                        }
                    }
                    else
                    {
                        foreach (MapInfoModel mapInfoModel in mapInfoModelList)
                        {
                            if (mapInfoModel.TVType == tvType)
                            {
                                List<MapInfoPointModel> mapInfoPointModelList = mapInfoService._MapInfoPointService.GetMapInfoPointModelListWithMapInfoIDDB(mapInfoModel.MapInfoID);
                                if (mapInfoPointModelList.Count != polyObj.coordList.Count)
                                {
                                    CoordListIsDifferent = true;
                                }
                                else
                                {
                                    for (int i = 0, count = mapInfoPointModelList.Count; i < count; i++)
                                    {
                                        if (!(mapInfoPointModelList[i].Lat == polyObj.coordList[i].Lat && mapInfoPointModelList[i].Lng == polyObj.coordList[i].Lng))
                                        {
                                            CoordListIsDifferent = true;
                                            break;
                                        }
                                    }
                                }

                                if (CoordListIsDifferent)
                                {
                                    MapInfoModel mapInfoModelDeletRet = mapInfoService.PostDeleteMapInfoDB(mapInfoModel.MapInfoID);
                                    if (!string.IsNullOrWhiteSpace(mapInfoModelDeletRet.Error))
                                    {
                                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDelete_Error_, TaskRunnerServiceRes.MapInfo, mapInfoModelDeletRet.Error);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDelete_Error_", TaskRunnerServiceRes.MapInfo, mapInfoModelDeletRet.Error);
                                        return;
                                    }
                                }
                            }
                        }
                    }

                    if (CoordListIsDifferent)
                    {
                        MapInfoModel mapInfoModelRet = tvItemService.CreateMapInfoObjectDB(polyObj.coordList, MapInfoDrawTypeEnum.Polyline, tvType, tvItemModelClass.TVItemID);
                        if (!string.IsNullOrWhiteSpace(mapInfoModelRet.Error))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.MapInfo, mapInfoModelRet.Error);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.MapInfo, mapInfoModelRet.Error);
                            return;
                        }
                    }

                    ClassificationModel classificationModel = classificationService.GetClassificationModelWithClassificationTVItemIDDB(tvItemModelClass.TVItemID);
                    if (!string.IsNullOrWhiteSpace(classificationModel.Error))
                    {
                        ClassificationModel classificationModelNew = new ClassificationModel();
                        classificationModelNew.ClassificationTVItemID = tvItemModelClass.TVItemID;
                        classificationModelNew.ClassificationType = classificationType;
                        classificationModelNew.ClassificationTVText = TVTextClass;
                        classificationModelNew.Ordinal = Ordinal;

                        ClassificationModel classificationModelRet = classificationService.PostAddClassificationDB(classificationModelNew);
                        if (!string.IsNullOrWhiteSpace(classificationModelRet.Error))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.Classification, classificationModelRet.Error);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.Classification, classificationModelRet.Error);
                            return;
                        }
                    }
                    else
                    {
                        classificationModel.ClassificationType = classificationType;
                        classificationModel.Ordinal = Ordinal;

                        ClassificationModel classificationModelRet = classificationService.PostUpdateClassificationDB(classificationModel);
                        if (!string.IsNullOrWhiteSpace(classificationModelRet.Error))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotUpdate_Error_, TaskRunnerServiceRes.Classification, classificationModelRet.Error);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotUpdate_Error_", TaskRunnerServiceRes.Classification, classificationModelRet.Error);
                            return;
                        }
                    }

                    //foreach (Coord coord in polyObj.coordList)
                    //{
                    //    richTextBoxStatus.AppendText($"\t{coord.Lat}\t{coord.Lng}\t{coord.Ordinal}\r\n");
                    //}

                }

                foreach (TVItemModel tvItemModel in tvItemModelClassificationList)
                {
                    if (!TVItemIDClassificationList.Where(c => c == tvItemModel.TVItemID).Any())
                    {
                        ClassificationModel classificationModelToDelete = classificationService.GetClassificationModelWithClassificationTVItemIDDB(tvItemModel.TVItemID);
                        if (!string.IsNullOrWhiteSpace(classificationModelToDelete.Error))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.Classification, TaskRunnerServiceRes.TVItemID, tvItemModel.TVItemID.ToString());
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.Classification, TaskRunnerServiceRes.TVItemID, tvItemModel.TVItemID.ToString());
                            return;
                        }

                        ClassificationModel classificationModelToDeleteRet = classificationService.PostDeleteClassificationDB(classificationModelToDelete.ClassificationID);
                        if (!string.IsNullOrWhiteSpace(classificationModelToDeleteRet.Error))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDelete_Error_, TaskRunnerServiceRes.Classification, classificationModelToDelete.Error);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDelete_Error_", TaskRunnerServiceRes.Classification, classificationModelToDelete.Error);
                            return;
                        }

                        MapInfoModel mapInfoModelToDelete = mapInfoService.PostDeleteMapInfoWithTVItemIDDB(tvItemModel.TVItemID);
                        if (!string.IsNullOrWhiteSpace(mapInfoModelToDelete.Error))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDelete_Error_, TaskRunnerServiceRes.MapInfo, mapInfoModelToDelete.Error);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDelete_Error_", TaskRunnerServiceRes.MapInfo, mapInfoModelToDelete.Error);
                            return;
                        }

                        TVItemModel tvItemModelToDelete = tvItemService.PostDeleteTVItemWithTVItemIDDB(tvItemModel.TVItemID);
                        if (!string.IsNullOrWhiteSpace(tvItemModelToDelete.Error))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDelete_Error_, TaskRunnerServiceRes.TVItem, tvItemModelToDelete.Error);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDelete_Error_", TaskRunnerServiceRes.TVItem, tvItemModelToDelete.Error);
                            return;
                        }

                    }
                }


            }
            #endregion Uploading MapInfo to CSSPWebTools

            List<TVItemModel> tvItemModelSubsectorList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(ProvinceTVItemID, TVTypeEnum.Subsector).OrderBy(c => c.TVText).ToList();
            if (tvItemModelSubsectorList.Count == 0)
            {
                return;
            }

            appTaskModel.PercentCompleted = 100;
            appTaskService.PostUpdateAppTask(appTaskModel);
        }
        public void GenerateKMLFileClassificationForCSSPWebToolsVisualization()
        {
            string NotUsed = "";
            string ProvInit = "";
            List<string> ProvInitList = new List<string>()
            {
                "BC", "ME", "NB", "NL", "NS", "PE", "QC",
            };
            List<string> ProvList = new List<string>()
            {
                "British Columbia", "Maine", "New Brunswick", "Newfoundland and Labrador", "Nova Scotia", "Prince Edward Island", "Québec",
            };

            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            int TVItemID = _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID;

            if (TVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.Parameter_NotFound, "TVItemID");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Parameter_NotFound", "TVItemID");
                return;
            }


            TVItemModel tvItemModel = tvItemService.GetTVItemModelWithTVItemIDDB(TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotDeleteFile_Error_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                return;
            }

            for (int i = 0, countProv = ProvList.Count; i < countProv; i++)
            {
                if (ProvList[i] == tvItemModel.TVText)
                {
                    ProvInit = ProvInitList[i];
                    break;
                }
            }

            string ServerFilePath = tvFileService.GetServerFilePath(TVItemID);

            FileInfo fi = new FileInfo(tvFileService.ChoseEDriveOrCDrive(ServerFilePath) + $"ClassViewCheck_{ProvInit}.kml");

            TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            // loop through all the MWQMSites etc...


            if (fi.Exists)
            {
                try
                {
                    fi.Delete();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDeleteFile_Error_, fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDeleteFile_Error_", fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    return;
                }
            }

            StringBuilder sb = new StringBuilder();

            sb.AppendLine(@"<?xml version=""1.0"" encoding=""UTF-8""?>");
            sb.AppendLine(@"<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom"">");
            sb.AppendLine(@"<Document>");
            sb.AppendLine(@"	<name>" + fi.Name + "</name>");

            sb.AppendLine(@"	<Style id=""A"">");
            sb.AppendLine(@"        <LineStyle>");
            sb.AppendLine(@"			<color>ff00ff00</color>");
            sb.AppendLine(@"			<width>1</width>");
            sb.AppendLine(@"		</LineStyle>");
            sb.AppendLine(@"		<PolyStyle>");
            sb.AppendLine(@"			<color>ff00ff00</color>");
            sb.AppendLine(@"		</PolyStyle>");
            sb.AppendLine(@"	</Style>");
            sb.AppendLine(@"	<Style id=""R"">");
            sb.AppendLine(@"		<LineStyle>");
            sb.AppendLine(@"			<color>ff0000ff</color>");
            sb.AppendLine(@"			<width>1</width>");
            sb.AppendLine(@"		</LineStyle>");
            sb.AppendLine(@"		<PolyStyle>");
            sb.AppendLine(@"			<color>ff0000ff</color>");
            sb.AppendLine(@"		</PolyStyle>");
            sb.AppendLine(@"	</Style>");
            sb.AppendLine(@"	<Style id=""P"">");
            sb.AppendLine(@"		<LineStyle>");
            sb.AppendLine(@"			<color>ffcccccc</color>");
            sb.AppendLine(@"			<width>1</width>");
            sb.AppendLine(@"		</LineStyle>");
            sb.AppendLine(@"		<PolyStyle>");
            sb.AppendLine(@"			<color>ffcccccc</color>");
            sb.AppendLine(@"		</PolyStyle>");
            sb.AppendLine(@"	</Style>");
            sb.AppendLine(@"	<Style id=""CA"">");
            sb.AppendLine(@"		<LineStyle>");
            sb.AppendLine(@"			<color>ff00ffff</color>");
            sb.AppendLine(@"			<width>1</width>");
            sb.AppendLine(@"		</LineStyle>");
            sb.AppendLine(@"		<PolyStyle>");
            sb.AppendLine(@"			<color>ff00ffff</color>");
            sb.AppendLine(@"		</PolyStyle>");
            sb.AppendLine(@"	</Style>");
            sb.AppendLine(@"	<Style id=""CR"">");
            sb.AppendLine(@"		<LineStyle>");
            sb.AppendLine(@"			<color>ffff00aa</color>");
            sb.AppendLine(@"			<width>1</width>");
            sb.AppendLine(@"		</LineStyle>");
            sb.AppendLine(@"		<PolyStyle>");
            sb.AppendLine(@"			<color>ffff00aa</color>");
            sb.AppendLine(@"		</PolyStyle>");
            sb.AppendLine(@"	</Style>");
            sb.AppendLine(@"	<Style id=""E"">");
            sb.AppendLine(@"		<LineStyle>");
            sb.AppendLine(@"			<color>ffffffff</color>");
            sb.AppendLine(@"			<width>1</width>");
            sb.AppendLine(@"		</LineStyle>");
            sb.AppendLine(@"		<PolyStyle>");
            sb.AppendLine(@"			<color>ffffffff</color>");
            sb.AppendLine(@"		</PolyStyle>");
            sb.AppendLine(@"	</Style>");


            using (CSSPDBEntities db = new CSSPDBEntities())
            {
                var tvItemProv = (from c in db.TVItems
                                  from cl in db.TVItemLanguages
                                  where c.TVItemID == cl.TVItemID
                                  && c.TVItemID == TVItemID
                                  && cl.Language == (int)LanguageEnum.en
                                  && c.TVType == (int)TVTypeEnum.Province
                                  select new { c, cl }).FirstOrDefault();

                if (tvItemProv == null)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                    return;
                }

                for (int i = 0, countProv = ProvList.Count; i < countProv; i++)
                {
                    if (ProvList[i] == tvItemProv.cl.TVText)
                    {
                        ProvInit = ProvInitList[i];
                        break;
                    }
                }

                NotUsed = string.Format(TaskRunnerServiceRes.Creating_, fi.Name);
                List<TextLanguage> TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", fi.Name);

                _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);

                var tvItemSSList = (from t in db.TVItems
                                    from tl in db.TVItemLanguages
                                    where t.TVItemID == tl.TVItemID
                                    && tl.Language == (int)LanguageEnum.en
                                    && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                    && t.TVType == (int)TVTypeEnum.Subsector
                                    orderby tl.TVText
                                    select new { t, tl }).ToList();


                int TotalCount2 = tvItemSSList.Count;
                int Count2 = 0;
                foreach (var tvItemSS in tvItemSSList)
                {
                    sb.AppendLine("     <Folder>");
                    sb.AppendLine($"        <name>{ tvItemSS.tl.TVText }</name>");

                    if (Count2 % 20 == 0)
                    {
                        _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)(100.0f * ((float)Count2 / (float)TotalCount2)));

                        NotUsed = string.Format(TaskRunnerServiceRes.Creating_, fi.Name + " --- doing " + tvItemSS.tl.TVText + "");
                        TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", fi.Name + " --- doing " + tvItemSS.tl.TVText + "");

                        _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);
                    }

                    Count2 += 1;

                    var ClassificationList = (from t in db.TVItems
                                              from c in db.Classifications
                                              from mi in db.MapInfos
                                              let mipList = (from mip in db.MapInfoPoints
                                                             where mip.MapInfoID == mi.MapInfoID
                                                             select mip).ToList()
                                              where c.ClassificationTVItemID == t.TVItemID
                                              && mi.TVItemID == t.TVItemID
                                              && t.TVPath.StartsWith(tvItemSS.t.TVPath + "p")
                                              && t.TVType == (int)TVTypeEnum.Classification
                                              && mi.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Polyline
                                              select new { t, mipList, c }).ToList();


                    foreach (var classification in ClassificationList)
                    {
                        string TVText = "";

                        switch (((ClassificationTypeEnum)classification.c.ClassificationType))
                        {
                            case ClassificationTypeEnum.Approved:
                                {
                                    TVText = "A";
                                }
                                break;
                            case ClassificationTypeEnum.Restricted:
                                {
                                    TVText = "R";
                                }
                                break;
                            case ClassificationTypeEnum.Prohibited:
                                {
                                    TVText = "P";
                                }
                                break;
                            case ClassificationTypeEnum.ConditionallyApproved:
                                {
                                    TVText = "CA";
                                }
                                break;
                            case ClassificationTypeEnum.ConditionallyRestricted:
                                {
                                    TVText = "CR";
                                }
                                break;
                            default:
                                {
                                    TVText = "E";
                                }
                                break;
                        }

                        //string style = "msn_ylw-pushpin";

                        List<MapInfoPoint> mapInfoPointList = classification.mipList.OrderBy(c => c.Ordinal).ToList();

                        for (int i = 0, count = mapInfoPointList.Count - 1; i < count; i++)
                        {
                            GeoLocation geo1 = new GeoLocation() { Latitude = (float)classification.mipList[i].Lat, Longitude = (float)classification.mipList[i].Lng };
                            GeoLocation geo2 = new GeoLocation() { Latitude = (float)classification.mipList[i + 1].Lat, Longitude = (float)classification.mipList[i + 1].Lng };
                            GeoLocation geo3 = new GeoLocation();
                            GeoLocation geo4 = new GeoLocation();

                            geo3.Longitude = (geo1.Longitude + geo2.Longitude) / 2.0D;
                            geo3.Latitude = (geo1.Latitude + geo2.Latitude) / 2.0D;

                            double xDiff = geo2.Longitude - geo1.Longitude;
                            double yDiff = geo2.Latitude - geo1.Latitude;
                            double rad = Math.Atan2(yDiff, xDiff);

                            double ABY = geo2.Latitude - geo1.Latitude;
                            double ABX = geo2.Longitude - geo1.Longitude;
                            double Len = Math.Sqrt(ABY * ABY + ABX * ABX);
                            geo4.Longitude = geo3.Longitude + 0.001D * ABY / Len;
                            geo4.Latitude = geo3.Latitude - 0.001D * ABX / Len;

                            sb.AppendLine(@"        <Placemark>");
                            sb.AppendLine($@"            <name>{ TVText }</name>");
                            sb.AppendLine($@"            <styleUrl>#{ TVText }</styleUrl>");
                            sb.AppendLine(@"            <Polygon>");
                            sb.AppendLine(@"		    	<tessellate>1</tessellate>");
                            sb.AppendLine(@"		    	<outerBoundaryIs>");
                            sb.AppendLine(@"		    		<LinearRing>");
                            sb.AppendLine(@"		    			<coordinates>");
                            sb.AppendLine($@"		    				{ geo1.Longitude },{ geo1.Latitude },0 { geo2.Longitude },{ geo2.Latitude },0 { geo4.Longitude },{ geo4.Latitude },0 { geo1.Longitude },{ geo1.Latitude },0 ");
                            sb.AppendLine(@"		    			</coordinates>");
                            sb.AppendLine(@"		    		</LinearRing>");
                            sb.AppendLine(@"		    	</outerBoundaryIs>");
                            sb.AppendLine(@"		    </Polygon>");
                            sb.AppendLine(@"        </Placemark>");

                        }
                    }
                    sb.AppendLine("     </Folder>");
                }

                sb.AppendLine(@"</Document>");
                sb.AppendLine(@"</kml>");

                StreamWriter sw = fi.CreateText();
                sw.Write(sb.ToString());
                sw.Close();

                _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.KMZClassViewCheck, FilePurposeEnum.Information);
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                    return;
            }
        }

        public void GenerateClassificationInputs_XX_FromDB_kmlFromDataInCSSPDB()
        {
            string NotUsed = "";
            string ProvInit = "";
            List<string> ProvInitList = new List<string>()
            {
                "BC", "ME", "NB", "NL", "NS", "PE", "QC",
            };
            List<string> ProvList = new List<string>()
            {
                "British Columbia", "Maine", "New Brunswick", "Newfoundland and Labrador", "Nova Scotia", "Prince Edward Island", "Québec",
            };

            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            int TVItemID = _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID;

            if (TVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.Parameter_NotFound, "TVItemID");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Parameter_NotFound", "TVItemID");
                return;
            }


            TVItemModel tvItemModel = tvItemService.GetTVItemModelWithTVItemIDDB(TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotDeleteFile_Error_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                return;
            }

            for (int i = 0, countProv = ProvList.Count; i < countProv; i++)
            {
                if (ProvList[i] == tvItemModel.TVText)
                {
                    ProvInit = ProvInitList[i];
                    break;
                }
            }

            string ServerFilePath = tvFileService.GetServerFilePath(TVItemID);

            FileInfo fi = new FileInfo(tvFileService.ChoseEDriveOrCDrive(ServerFilePath) + $"ClassificationInputs_{ProvInit}_FromDB.kml");

            TVItemModel tvItemModelFile = _TaskRunnerBaseService.CreateFileTVItem(fi);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return;

            // loop through all the MWQMSites etc...


            if (fi.Exists)
            {
                try
                {
                    fi.Delete();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDeleteFile_Error_, fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDeleteFile_Error_", fi.FullName, ex.Message + (ex.InnerException != null ? " InnerException: " + ex.InnerException.Message : ""));
                    return;
                }
            }

            StringBuilder sb = new StringBuilder();

            sb.AppendLine(@"<?xml version=""1.0"" encoding=""UTF-8""?>");
            sb.AppendLine(@"<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom"">");
            sb.AppendLine(@"<Document>");
            sb.AppendLine(@"  <name>" + fi.Name.Replace("_FromDB", "") + "</name>");
            sb.AppendLine(@"  <open>1</open>");
            sb.AppendLine(@"  <Style id=""A"">");
            sb.AppendLine(@"    <LineStyle>");
            sb.AppendLine(@"      <color>ff00ff00</color>");
            sb.AppendLine(@"      <width>1</width>");
            sb.AppendLine(@"  	</LineStyle>");
            sb.AppendLine(@"    <PolyStyle>");
            sb.AppendLine(@"  	  <color>ff00ff00</color>");
            sb.AppendLine(@"    </PolyStyle>");
            sb.AppendLine(@"  </Style>");
            sb.AppendLine(@"  <Style id=""R"">");
            sb.AppendLine(@"    <LineStyle>");
            sb.AppendLine(@"      <color>ff0000ff</color>");
            sb.AppendLine(@"      <width>1</width>");
            sb.AppendLine(@"    </LineStyle>");
            sb.AppendLine(@"    <PolyStyle>");
            sb.AppendLine(@"      <color>ff0000ff</color>");
            sb.AppendLine(@"    </PolyStyle>");
            sb.AppendLine(@"  </Style>");
            sb.AppendLine(@"  <Style id=""P"">");
            sb.AppendLine(@"    <LineStyle>");
            sb.AppendLine(@"      <color>ffcccccc</color>");
            sb.AppendLine(@"      <width>1</width>");
            sb.AppendLine(@"    </LineStyle>");
            sb.AppendLine(@"    <PolyStyle>");
            sb.AppendLine(@"      <color>ffcccccc</color>");
            sb.AppendLine(@"    </PolyStyle>");
            sb.AppendLine(@"  </Style>");
            sb.AppendLine(@"  <Style id=""CA"">");
            sb.AppendLine(@"    <LineStyle>");
            sb.AppendLine(@"      <color>ff00ffff</color>");
            sb.AppendLine(@"      <width>1</width>");
            sb.AppendLine(@"    </LineStyle>");
            sb.AppendLine(@"    <PolyStyle>");
            sb.AppendLine(@"     <color>ff00ffff</color>");
            sb.AppendLine(@"    </PolyStyle>");
            sb.AppendLine(@"  </Style>");
            sb.AppendLine(@"  <Style id=""CR"">");
            sb.AppendLine(@"    <LineStyle>");
            sb.AppendLine(@"      <color>ffff00aa</color>");
            sb.AppendLine(@"      <width>1</width>");
            sb.AppendLine(@"    </LineStyle>");
            sb.AppendLine(@"    <PolyStyle>");
            sb.AppendLine(@"      <color>ffff00aa</color>");
            sb.AppendLine(@"    </PolyStyle>");
            sb.AppendLine(@"  </Style>");
            sb.AppendLine(@"  <Style id=""SS"">");
            sb.AppendLine(@"    <LineStyle>");
            sb.AppendLine(@"      <color>ff00ff00</color>");
            sb.AppendLine(@"      <width>2</width>");
            sb.AppendLine(@"    </LineStyle>");
            sb.AppendLine(@"    <PolyStyle>");
            sb.AppendLine(@"      <color>ff00ff00</color>");
            sb.AppendLine(@"      <fill>0</fill>");
            sb.AppendLine(@" </PolyStyle>");
            sb.AppendLine(@"  </Style>");
            sb.AppendLine(@"  <Folder>");
            sb.AppendLine($@"    <name>{ProvInit} Subsectors</name>");
            sb.AppendLine(@"    <open>1</open>");

            using (CSSPDBEntities db = new CSSPDBEntities())
            {
                var tvItemProv = (from c in db.TVItems
                                  from cl in db.TVItemLanguages
                                  where c.TVItemID == cl.TVItemID
                                  && c.TVItemID == TVItemID
                                  && cl.Language == (int)LanguageEnum.en
                                  && c.TVType == (int)TVTypeEnum.Province
                                  select new { c, cl }).FirstOrDefault();

                if (tvItemProv == null)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                    return;
                }

                for (int i = 0, countProv = ProvList.Count; i < countProv; i++)
                {
                    if (ProvList[i] == tvItemProv.cl.TVText)
                    {
                        ProvInit = ProvInitList[i];
                        break;
                    }
                }

                NotUsed = string.Format(TaskRunnerServiceRes.Creating_, fi.Name);
                List<TextLanguage> TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", fi.Name);

                _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);

                var tvItemSSList = (from t in db.TVItems
                                    from tl in db.TVItemLanguages
                                    where t.TVItemID == tl.TVItemID
                                    && tl.Language == (int)LanguageEnum.en
                                    && t.TVPath.StartsWith(tvItemProv.c.TVPath + "p")
                                    && t.TVType == (int)TVTypeEnum.Subsector
                                    orderby tl.TVText
                                    select new { t, tl }).ToList();


                int TotalCount2 = tvItemSSList.Count;
                int Count2 = 0;
                foreach (var tvItemSS in tvItemSSList)
                {
                    sb.AppendLine("  <Folder>");
                    sb.AppendLine($"  <name>{ tvItemSS.tl.TVText }</name>");

                    if (Count2 % 20 == 0)
                    {
                        _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)(100.0f * ((float)Count2 / (float)TotalCount2)));

                        NotUsed = string.Format(TaskRunnerServiceRes.Creating_, fi.Name + " --- doing " + tvItemSS.tl.TVText + "");
                        TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", fi.Name + " --- doing " + tvItemSS.tl.TVText + "");

                        _TaskRunnerBaseService.SendStatusTextToDB(TextLanguageList);
                    }

                    List<MapInfoPoint> mapInfoPointList = (from t in db.TVItems
                                                           from mi in db.MapInfos
                                                           from mip in db.MapInfoPoints
                                                           where mi.TVItemID == t.TVItemID
                                                           && mi.MapInfoID == mip.MapInfoID
                                                           && mi.TVItemID == tvItemSS.t.TVItemID
                                                           && t.TVType == (int)TVTypeEnum.Subsector
                                                           && mi.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Polygon
                                                           select mip).ToList();


                    sb.AppendLine("    <Placemark>");
                    sb.AppendLine("      <name>Subsector Polygon</name>");
                    sb.AppendLine("      <styleUrl>#SS</styleUrl>");
                    sb.AppendLine("      <Polygon>");
                    sb.AppendLine("    	   <tessellate>1</tessellate>");
                    sb.AppendLine("    	   <outerBoundaryIs>");
                    sb.AppendLine("    	     <LinearRing>");
                    sb.AppendLine("    	       <coordinates>");
                    foreach (MapInfoPoint mapInfoPoint in mapInfoPointList)
                    {
                        sb.AppendLine($"{mapInfoPoint.Lng},{mapInfoPoint.Lat},0 ");
                    }
                    sb.AppendLine("            </coordinates>");
                    sb.AppendLine("          </LinearRing>");
                    sb.AppendLine("        </outerBoundaryIs>");
                    sb.AppendLine("      </Polygon>");
                    sb.AppendLine("    </Placemark>");

                    Count2 += 1;

                    var ClassificationList = (from t in db.TVItems
                                              from c in db.Classifications
                                              from mi in db.MapInfos
                                              let mipList = (from mip in db.MapInfoPoints
                                                             where mip.MapInfoID == mi.MapInfoID
                                                             select mip).ToList()
                                              where c.ClassificationTVItemID == t.TVItemID
                                              && mi.TVItemID == t.TVItemID
                                              && t.TVPath.StartsWith(tvItemSS.t.TVPath + "p")
                                              && t.TVType == (int)TVTypeEnum.Classification
                                              && mi.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Polyline
                                              select new { t, mipList, c }).ToList();


                    foreach (var classification in ClassificationList)
                    {
                        string TVText = "";

                        switch (((ClassificationTypeEnum)classification.c.ClassificationType))
                        {
                            case ClassificationTypeEnum.Approved:
                                {
                                    TVText = "A";
                                }
                                break;
                            case ClassificationTypeEnum.Restricted:
                                {
                                    TVText = "R";
                                }
                                break;
                            case ClassificationTypeEnum.Prohibited:
                                {
                                    TVText = "P";
                                }
                                break;
                            case ClassificationTypeEnum.ConditionallyApproved:
                                {
                                    TVText = "CA";
                                }
                                break;
                            case ClassificationTypeEnum.ConditionallyRestricted:
                                {
                                    TVText = "CR";
                                }
                                break;
                            default:
                                {
                                    TVText = "E";
                                }
                                break;
                        }

                        //string style = "msn_ylw-pushpin";

                        mapInfoPointList = classification.mipList.OrderBy(c => c.Ordinal).ToList();

                        sb.AppendLine(@"    <Placemark>");
                        sb.AppendLine($@"      <name>{ TVText }</name>");
                        sb.AppendLine($@"      <styleUrl>#{ TVText }</styleUrl>");
                        sb.AppendLine("        <LineString>");
                        sb.AppendLine("          <tessellate>1</tessellate>");
                        sb.AppendLine("          <coordinates>");
                        foreach (MapInfoPoint mapInfoPoint in mapInfoPointList)
                        {
                            sb.AppendLine($"{mapInfoPoint.Lng},{mapInfoPoint.Lat},0 ");
                        }
                        sb.AppendLine("          </coordinates>");
                        sb.AppendLine("        </LineString>");
                        sb.AppendLine("      </Placemark>");
                    }
                    sb.AppendLine("    </Folder>");
                }

                sb.AppendLine("  </Folder>");
                sb.AppendLine(@"</Document>");
                sb.AppendLine(@"</kml>");

                StreamWriter sw = fi.CreateText();
                sw.Write(sb.ToString());
                sw.Close();

                _TaskRunnerBaseService.UpdateOrCreateTVFile(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID, fi, tvItemModelFile, TaskRunnerServiceRes.KMZClassViewCheck, FilePurposeEnum.Information);
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                    return;
            }
        }

        public struct GeoLocation
        {
            public double Latitude { get; set; }
            public double Longitude { get; set; }
        }
        #endregion Functions public

        #region Functions private
        private void GetCoordinates(List<PolyObj> polyObsList, XmlNode node)
        {
            string NotUsed = "";

            if (node.Name == "coordinates")
            {
                PolyObj polyObj = new PolyObj();

                List<string> pointListText = node.InnerText.Split(" ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

                int ordinal = 0;
                foreach (string pointText in pointListText)
                {
                    string pointTxt = pointText.Replace("\r\n", "");

                    if (!string.IsNullOrWhiteSpace(pointTxt))
                    {
                        List<string> valListText = pointTxt.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

                        if (valListText.Count != 3)
                        {
                            NotUsed = "valListText.Count != 3";
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("valListText.Count != 3");
                            return;
                        }

                        float Lng = float.Parse(valListText[0]);
                        float Lat = float.Parse(valListText[1]);

                        Coord coord = new Coord() { Lat = Lat, Lng = Lng, Ordinal = ordinal };

                        polyObj.coordList.Add(coord);

                        ordinal += 1;
                    }
                }

                polyObj.MinLat = polyObj.coordList.Min(c => c.Lat) - 0.0001f;
                polyObj.MinLng = polyObj.coordList.Min(c => c.Lng) - 0.0001f;
                polyObj.MaxLat = polyObj.coordList.Max(c => c.Lat) + 0.0001f;
                polyObj.MaxLng = polyObj.coordList.Max(c => c.Lng) + 0.0001f;


                polyObsList.Add(polyObj);
            }



            foreach (XmlNode n in node.ChildNodes)
            {
                GetCoordinates(polyObsList, n);
            }

        }
        private void GetTVItemIDAndLatLng(List<TVItemIDAndLatLng> TVItemIDAndLatLngList, XmlNode node)
        {
            string NotUsed = "";
            if (node.Name == "Placemark")
            {
                TVItemIDAndLatLng tvItemIDAndLatLng = new TVItemIDAndLatLng();

                foreach (XmlNode n2 in node.ChildNodes)
                {
                    if (n2.Name == "description")
                    {
                        string desc = n2.InnerText;
                        desc = desc.Substring(desc.IndexOf(@"data-tvitemid=""") + @"data-tvitemid=""".Length);
                        desc = desc.Substring(0, desc.IndexOf(@""""));

                        tvItemIDAndLatLng.TVItemID = int.Parse(desc);
                    }

                    if (n2.Name == "Point")
                    {
                        foreach (XmlNode n3 in n2.ChildNodes)
                        {
                            if (n3.Name == "coordinates")
                            {
                                List<string> strList = node.InnerText.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

                                foreach (string pointText in strList)
                                {
                                    string pointTxt = pointText.Replace("\r\n", "");


                                    if (strList.Count != 3)
                                    {
                                        NotUsed = "valListText.Count != 3";
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("valListText.Count != 3");
                                        return;
                                    }

                                    float Lng = float.Parse(strList[0]);
                                    float Lat = float.Parse(strList[1]);

                                    tvItemIDAndLatLng.Lng = Lng;
                                    tvItemIDAndLatLng.Lat = Lat;

                                }

                            }
                        }
                    }
                }
                TVItemIDAndLatLngList.Add(tvItemIDAndLatLng);
            }



            foreach (XmlNode n in node.ChildNodes)
            {
                GetTVItemIDAndLatLng(TVItemIDAndLatLngList, n);
            }

        }

        #endregion Functions private

    }

    public class PolyObj
    {
        public string Classification { get; set; } = "";
        public float MinLat { get; set; } = 0.0F;
        public float MaxLat { get; set; } = 0.0F;
        public float MinLng { get; set; } = 0.0F;
        public float MaxLng { get; set; } = 0.0F;
        public List<Coord> coordList { get; set; } = new List<Coord>();
        public string Subsector { get; set; } = "";
    }
    public class TVItemIDAndLatLng
    {
        public int TVItemID { get; set; } = 0;
        public float Lat { get; set; } = 0.0F;
        public float Lng { get; set; } = 0.0F;
    }
}
