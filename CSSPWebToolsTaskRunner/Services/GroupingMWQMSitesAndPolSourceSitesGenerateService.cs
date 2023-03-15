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

namespace CSSPWebToolsTaskRunner.Services
{
    public class GroupingMWQMSitesAndPolSourceSitesGenerateService
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
        public GroupingMWQMSitesAndPolSourceSitesGenerateService(TaskRunnerBaseService taskRunnerBaseService)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
        }
        #endregion Constructors

        #region Events
        #endregion Events

        #region Functions public
        public void GenerateLinksBetweenMWQMSitesAndPolSourceSitesForCSSPWebToolsVisualization()
        {
            string NotUsed = "";

            TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            ProvinceToolsService provinceToolsService = new ProvinceToolsService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
            PolSourceSiteEffectService polSourceSiteEffectService = new PolSourceSiteEffectService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            PolSourceSiteEffectTermService polSourceSiteEffectTermService = new PolSourceSiteEffectTermService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

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

            #region Reading the MWQMSitesAndPolSourceSites_XX.KML
            List<TVItemIDAndLatLng> TVItemIDAndLatLngList = new List<TVItemIDAndLatLng>();

            string FileName = $"MWQMSitesAndPolSourceSites_{Init}.kml";

            FileInfo fi = new FileInfo(ServerPath + FileName);

            if (!fi.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fi.FullName);
                return;
            }

            XmlDocument doc = new XmlDocument();
            doc.Load(fi.FullName);

            foreach (XmlNode node in doc.ChildNodes)
            {
                GetTVItemIDAndLatLng(TVItemIDAndLatLngList, node);
            }

            #endregion Reading the MWQMSitesAndPolSourceSites_XX.KML

            #region Reading the GroupingInputs__XX.KML

            string FileName2 = $"GroupingInputs_{Init}.kml";

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
            string CurrentGroupingMWQMSitesAndPolSourceSites = "";

            List<PolyObj> polyObjList = new List<PolyObj>();

            XmlNode StartNode2 = doc2.ChildNodes[1].ChildNodes[0];
            foreach (XmlNode n in StartNode2.ChildNodes)
            {
                if (n.Name == "Folder")
                {
                    CurrentSubsector = "";

                    foreach (XmlNode n22 in n)
                    {
                        if (n22.Name == "name")
                        {
                            CurrentSubsector = n22.InnerText;
                        }

                        if (n22.Name == "Placemark")
                        {
                            CurrentGroupingMWQMSitesAndPolSourceSites = "";

                            foreach (XmlNode n2 in n22)
                            {

                                if (n2.Name == "name")
                                {
                                    CurrentGroupingMWQMSitesAndPolSourceSites = n2.InnerText;
                                }

                                if (n2.Name == "Polygon")
                                {
                                    foreach (XmlNode n222 in n2.ChildNodes)
                                    {
                                        if (n222.Name == "outerBoundaryIs")
                                        {
                                            foreach (XmlNode n2222 in n222.ChildNodes)
                                            {
                                                if (n2222.Name == "LinearRing")
                                                {
                                                    PolyObj polyObj = new PolyObj();

                                                    polyObj.Subsector = CurrentSubsector;
                                                    polyObj.Classification = CurrentGroupingMWQMSitesAndPolSourceSites.ToUpper().Trim();

                                                    foreach (XmlNode n3 in n2222.ChildNodes)
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

                                                                    Coord coord = new Coord() { Lat = Lat, Lng = Lng, Ordinal = ordinal };

                                                                    polyObj.coordList.Add(coord);

                                                                    ordinal += 1;
                                                                }
                                                            }
                                                        }
                                                    }
                                                    if (polyObj.coordList.Count > 0)
                                                    {
                                                        polyObj.MinLat = polyObj.coordList.Min(c => c.Lat);
                                                        polyObj.MaxLat = polyObj.coordList.Max(c => c.Lat);
                                                        polyObj.MinLng = polyObj.coordList.Min(c => c.Lng);
                                                        polyObj.MaxLng = polyObj.coordList.Max(c => c.Lng);
                                                    }
                                                    polyObjList.Add(polyObj);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            #endregion Reading the GroupingInputs__XX.KML

            #region Saving PolSourceSiteEffect in CSSPDB
            TVItemModel tvItemModelProv = tvItemService.GetTVItemModelWithTVItemIDDB(ProvinceTVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelProv.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, ProvinceTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, ProvinceTVItemID.ToString());
                return;
            }

            // getting all active MWQMSites and PolSourceSites under the province

            List<TVItem> tvItemMWQMSiteList = new List<TVItem>();
            List<TVItem> tvItemPSSList = new List<TVItem>();
            List<TVItem> tvItemInfraList = new List<TVItem>();
            using (CSSPDBEntities db2 = new CSSPDBEntities())
            {
                tvItemMWQMSiteList = (from c in db2.TVItems
                                      where c.TVPath.StartsWith(tvItemModelProv.TVPath + "p")
                                      && c.TVType == (int)TVTypeEnum.MWQMSite
                                      && c.IsActive == true
                                      select c).ToList();


                tvItemPSSList = (from c in db2.TVItems
                                 where c.TVPath.StartsWith(tvItemModelProv.TVPath + "p")
                                 && c.TVType == (int)TVTypeEnum.PolSourceSite
                                 && c.IsActive == true
                                 select c).ToList();

                tvItemInfraList = (from c in db2.TVItems
                                   where c.TVPath.StartsWith(tvItemModelProv.TVPath + "p")
                                   && c.TVType == (int)TVTypeEnum.Infrastructure
                                   && c.IsActive == true
                                   select c).ToList();

            }

            List<TVItemModel> tvitemModelSSList = tvItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelProv.TVItemID, TVTypeEnum.Subsector);

            List<PolSourceSiteEffect> existingPolSourceSiteEffectSubsectorList = new List<PolSourceSiteEffect>();
            List<PolSourceSiteEffect> newPolSourceSiteEffectSubsectorList = new List<PolSourceSiteEffect>();

            int CountSS = 0;
            int TotalCount = tvitemModelSSList.Count;
            string Status = appTaskModel.StatusText;
            foreach (TVItemModel tvItemModelSS in tvitemModelSSList)
            {
                CountSS += 1;
                if (CountSS % 1 == 0)
                {
                    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)(100.0f * ((float)CountSS / (float)TotalCount)));
                }
                Application.DoEvents();

                string TVTextSS = "";
                if (tvItemModelSS.TVText.Contains(" "))
                {
                    TVTextSS = tvItemModelSS.TVText.Substring(0, tvItemModelSS.TVText.IndexOf(" "));
                }
                else
                {
                    TVTextSS = tvItemModelSS.TVText;
                }

                foreach (PolyObj polyObj in polyObjList.Where(c => c.Subsector == TVTextSS))
                {
                    List<MapInfo> mapInfoMWQMSiteList2 = new List<MapInfo>();
                    List<MapInfo> mapInfoPSSList2 = new List<MapInfo>();
                    List<MapInfo> mapInfoInfraList2 = new List<MapInfo>();

                    using (CSSPDBEntities db2 = new CSSPDBEntities())
                    {

                        var mapInfoMWQMSiteList = (from c in db2.MapInfos
                                                   let lat = (from d in db2.MapInfoPoints where c.MapInfoID == d.MapInfoID select d).FirstOrDefault().Lat
                                                   let lng = (from d in db2.MapInfoPoints where c.MapInfoID == d.MapInfoID select d).FirstOrDefault().Lng
                                                   where c.TVType == (int)TVTypeEnum.MWQMSite
                                                   && c.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Point
                                                   && lat >= polyObj.MinLat
                                                   && lat <= polyObj.MaxLat
                                                   && lng >= polyObj.MinLng
                                                   && lng <= polyObj.MaxLng
                                                   select new { c, lat, lng }).ToList();

                        var mapInfoPSSList = (from c in db2.MapInfos
                                              let lat = (from d in db2.MapInfoPoints where c.MapInfoID == d.MapInfoID select d).FirstOrDefault().Lat
                                              let lng = (from d in db2.MapInfoPoints where c.MapInfoID == d.MapInfoID select d).FirstOrDefault().Lng
                                              where c.TVType == (int)TVTypeEnum.PolSourceSite
                                              && c.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Point
                                              && lat >= polyObj.MinLat
                                              && lat <= polyObj.MaxLat
                                              && lng >= polyObj.MinLng
                                              && lng <= polyObj.MaxLng
                                              select new { c, lat, lng }).ToList();

                        var mapInfoInfraList = (from c in db2.MapInfos
                                                let lat = (from d in db2.MapInfoPoints where c.MapInfoID == d.MapInfoID select d).FirstOrDefault().Lat
                                                let lng = (from d in db2.MapInfoPoints where c.MapInfoID == d.MapInfoID select d).FirstOrDefault().Lng
                                                where (c.TVType == (int)TVTypeEnum.WasteWaterTreatmentPlant
                                                || c.TVType == (int)TVTypeEnum.LiftStation
                                                || c.TVType == (int)TVTypeEnum.LineOverflow)
                                                && c.MapInfoDrawType == (int)MapInfoDrawTypeEnum.Point
                                                && lat >= polyObj.MinLat
                                                && lat <= polyObj.MaxLat
                                                && lng >= polyObj.MinLng
                                                && lng <= polyObj.MaxLng
                                                select new { c, lat, lng }).ToList();


                        foreach (var mapInfo in mapInfoMWQMSiteList)
                        {
                            if (mapInfoService.CoordInPolygon(polyObj.coordList, new Coord() { Lat = (float)mapInfo.lat, Lng = (float)mapInfo.lng, Ordinal = 0 }))
                            {
                                mapInfoMWQMSiteList2.Add(mapInfo.c);
                            }
                        }

                        foreach (var mapInfo in mapInfoPSSList)
                        {
                            if (mapInfoService.CoordInPolygon(polyObj.coordList, new Coord() { Lat = (float)mapInfo.lat, Lng = (float)mapInfo.lng, Ordinal = 0 }))
                            {
                                mapInfoPSSList2.Add(mapInfo.c);
                            }
                        }

                        foreach (var mapInfo in mapInfoInfraList)
                        {
                            if (mapInfoService.CoordInPolygon(polyObj.coordList, new Coord() { Lat = (float)mapInfo.lat, Lng = (float)mapInfo.lng, Ordinal = 0 }))
                            {
                                mapInfoInfraList2.Add(mapInfo.c);
                            }
                        }

                    }

                    List<int> PSSTVItemIDList = (from c in mapInfoPSSList2
                                                 select c.TVItemID).Distinct().ToList();
                    List<int> MWQMTVItemIDList = (from c in mapInfoMWQMSiteList2
                                                  select c.TVItemID).Distinct().ToList();
                    List<int> InfraTVItemIDList = (from c in mapInfoInfraList2
                                                   select c.TVItemID).Distinct().ToList();

                    using (CSSPDBEntities db2 = new CSSPDBEntities())
                    {
                        List<int> TVItemIDMWQMSiteWithinSubsector = tvItemMWQMSiteList.Where(c => c.TVPath.Contains(tvItemModelSS.TVPath + "p")).Select(c => c.TVItemID).ToList();

                        existingPolSourceSiteEffectSubsectorList = (from c in db2.PolSourceSiteEffects
                                                                    from m in TVItemIDMWQMSiteWithinSubsector
                                                                    where c.MWQMSiteTVItemID == m
                                                                    select c).ToList();

                    }

                    foreach (TVItem tvItemMWQMSite in tvItemMWQMSiteList.Where(c => c.TVPath.Contains(tvItemModelSS.TVPath + "p")))
                    {
                        if (MWQMTVItemIDList.Contains(tvItemMWQMSite.TVItemID))
                        {
                            List<PolSourceSiteEffect> MWQMSiteExistingPolSourceSiteEffectList = (from c in existingPolSourceSiteEffectSubsectorList
                                                                                                 where c.MWQMSiteTVItemID == tvItemMWQMSite.TVItemID
                                                                                                 select c).ToList();

                            // doing pollution source site
                            foreach (int PSSTVItemID in PSSTVItemIDList)
                            {
                                newPolSourceSiteEffectSubsectorList.Add(new PolSourceSiteEffect() { PolSourceSiteOrInfrastructureTVItemID = PSSTVItemID, MWQMSiteTVItemID = tvItemMWQMSite.TVItemID, });

                                if (!(MWQMSiteExistingPolSourceSiteEffectList.Where(c => c.PolSourceSiteOrInfrastructureTVItemID == PSSTVItemID).Any()))
                                {
                                    PolSourceSiteEffectModel polSourceSiteEffectModelNew = new PolSourceSiteEffectModel()
                                    {
                                        PolSourceSiteOrInfrastructureTVItemID = PSSTVItemID,
                                        MWQMSiteTVItemID = tvItemMWQMSite.TVItemID,
                                    };

                                    PolSourceSiteEffectModel polSourceSiteEffectModelAlreadyExist = polSourceSiteEffectService.GetPolSourceSiteEffectModelWithPolSourceSiteOrInfrastructureTVItemIDAndMWQMSiteTVItemIDDB(PSSTVItemID, tvItemMWQMSite.TVItemID);
                                    if (!string.IsNullOrWhiteSpace(polSourceSiteEffectModelAlreadyExist.Error))
                                    {

                                        PolSourceSiteEffectModel polSourceSiteEffectModelRet = polSourceSiteEffectService.PostAddPolSourceSiteEffectDB(polSourceSiteEffectModelNew);
                                        if (!string.IsNullOrWhiteSpace(polSourceSiteEffectModelRet.Error))
                                        {
                                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.PolSourceSiteEffect, polSourceSiteEffectModelRet.Error);
                                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.PolSourceSiteEffect, polSourceSiteEffectModelRet.Error);
                                            return;
                                        }
                                    }
                                }
                            }

                            // doing infrastructure
                            foreach (int InfraTVItemID in InfraTVItemIDList)
                            {
                                newPolSourceSiteEffectSubsectorList.Add(new PolSourceSiteEffect() { PolSourceSiteOrInfrastructureTVItemID = InfraTVItemID, MWQMSiteTVItemID = tvItemMWQMSite.TVItemID, });

                                if (!(MWQMSiteExistingPolSourceSiteEffectList.Where(c => c.PolSourceSiteOrInfrastructureTVItemID == InfraTVItemID).Any()))
                                {
                                    PolSourceSiteEffectModel polSourceSiteEffectModelNew = new PolSourceSiteEffectModel()
                                    {
                                        PolSourceSiteOrInfrastructureTVItemID = InfraTVItemID,
                                        MWQMSiteTVItemID = tvItemMWQMSite.TVItemID,
                                    };

                                    PolSourceSiteEffectModel polSourceSiteEffectModelAlreadyExist = polSourceSiteEffectService.GetPolSourceSiteEffectModelWithPolSourceSiteOrInfrastructureTVItemIDAndMWQMSiteTVItemIDDB(InfraTVItemID, tvItemMWQMSite.TVItemID);
                                    if (!string.IsNullOrWhiteSpace(polSourceSiteEffectModelAlreadyExist.Error))
                                    {

                                        PolSourceSiteEffectModel polSourceSiteEffectModelRet = polSourceSiteEffectService.PostAddPolSourceSiteEffectDB(polSourceSiteEffectModelNew);
                                        if (!string.IsNullOrWhiteSpace(polSourceSiteEffectModelRet.Error))
                                        {
                                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.PolSourceSiteEffect, polSourceSiteEffectModelRet.Error);
                                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.PolSourceSiteEffect, polSourceSiteEffectModelRet.Error);
                                            return;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            foreach (PolSourceSiteEffect polSourceSiteEffect in existingPolSourceSiteEffectSubsectorList)
            {
                if (!newPolSourceSiteEffectSubsectorList.Where(c => c.PolSourceSiteOrInfrastructureTVItemID == polSourceSiteEffect.PolSourceSiteOrInfrastructureTVItemID 
                && c.MWQMSiteTVItemID == polSourceSiteEffect.MWQMSiteTVItemID).Any())
                {
                    PolSourceSiteEffectModel polSourceSiteEffectModelAlreadyExist = polSourceSiteEffectService.GetPolSourceSiteEffectModelWithPolSourceSiteOrInfrastructureTVItemIDAndMWQMSiteTVItemIDDB(polSourceSiteEffect.PolSourceSiteOrInfrastructureTVItemID, polSourceSiteEffect.MWQMSiteTVItemID);
                    if (string.IsNullOrWhiteSpace(polSourceSiteEffectModelAlreadyExist.Error))
                    {

                        PolSourceSiteEffectModel polSourceSiteEffectModelRet = polSourceSiteEffectService.PostDeletePolSourceSiteEffectDB(polSourceSiteEffectModelAlreadyExist.PolSourceSiteEffectID);
                        if (!string.IsNullOrWhiteSpace(polSourceSiteEffectModelRet.Error))
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDelete_Error_, TaskRunnerServiceRes.PolSourceSiteEffect, polSourceSiteEffectModelRet.Error);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDelete_Error_", TaskRunnerServiceRes.PolSourceSiteEffect, polSourceSiteEffectModelRet.Error);
                            return;
                        }
                    }
                }
            }

            #endregion Saving PolSourceSiteEffect in CSSPDB

            appTaskModel.PercentCompleted = 100;
            appTaskService.PostUpdateAppTask(appTaskModel);
        }
        #endregion Functions public

        #region Functions private
        private void GetCoordinates(List<PolyObj> polyObsList, XmlNode node)
        {
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
                                List<string> strList = n3.InnerText.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

                                foreach (string pointText in strList)
                                {
                                    string pointTxt = pointText.Replace("\r\n", "");


                                    if (strList.Count != 3)
                                    {
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
}
